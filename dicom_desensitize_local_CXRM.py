#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DICOM 脱敏引擎 v8.0-local（原地修改版）
=============================================================
基于 dicom_desensitize.py（v8），改为直接在原文件上进行修改（in-place）。

与原版区别：
  - 原版：--src 源目录 + --dst 输出目录，文件写到 dst
  - 本版：只有 --src，脱敏后直接覆盖原文件（原子写入：先写 .tmp → os.replace）
  - 日志/审计/断点续传文件放在 <src>/_desensitize_log/ 子目录中
  - 未成功处理的文件（非 DICOM / 结构异常 / 处理失败）清单写入
    <src>/_desensitize_log/unparsed_files.csv，便于事后核对
  - 采用 pydicom 宽容解析：厂商私有标签长度与 VR 不一致时也能正常脱敏，
    不会因单个标签导致整份文件被判失败
  - 不做 .dcm 扩展名转换（原地覆盖，保留原文件名）

用法:
  python dicom_desensitize_local.py --src <源目录> [--workers N]
可选:
  --no-purge-private    不清空私有标签（默认清空）
  --purge-nested-ids    递归清空嵌套字段中的类影像号数字
  --datetime-empty      DA/DT 直接置空而非保留年-月
  --force               跳过源目录安全检查
"""

import os
import sys
import io
import re
import time
import signal
import hashlib
import argparse
import warnings
import multiprocessing as mp
from pathlib import Path
from queue import Empty, Full
from datetime import datetime
from logging.handlers import RotatingFileHandler
from collections import defaultdict

try:
    import pydicom
    from pydicom.errors import InvalidDicomError
except ImportError:
    print("致命错误: 缺少 pydicom，请运行 pip install pydicom")
    sys.exit(2)

warnings.filterwarnings('ignore', category=UserWarning, module='pydicom')

# 宽容解析（关键）：部分厂商会写入「长度与 VR 不匹配」的私有标签（例如 (01F1,1026)，
# pydicom 查不到私有标签的 VR → 按 "0 bytes per value" 直接抛错），严格模式下会导致整份
# 文件被判失败。打开该开关后按 UN 原始字节保留，随后由 purge_private 规则正常清掉，
# 落盘仍是脱敏后的影像；读取校验降级为警告可避免同类不规范标签中断流程。
# 注意：本模块在 worker 子进程中会被重新导入，模块级设置对每个 worker 都生效。
try:
    pydicom.config.convert_wrong_length_to_UN = True
    pydicom.config.settings.reading_validation_mode = pydicom.config.WARN
except Exception:
    pass


# ============================================================
# 冻结 exe 控制台绑定
# ============================================================
def _ensure_console():
    if not getattr(sys, 'frozen', False):
        return
    _is_worker = False
    try:
        _is_worker = mp.current_process().name.startswith('Worker')
    except Exception:
        pass
    _attach_console(allow_alloc=not _is_worker)


def _attach_console(allow_alloc=True):
    try:
        import ctypes
        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        if not kernel32.AttachConsole(-1):
            if allow_alloc:
                if not kernel32.AllocConsole():
                    return
            else:
                return
        STD_OUTPUT_HANDLE = -11
        STD_ERROR_HANDLE  = -12
        STD_INPUT_HANDLE  = -10
        h_out = kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
        h_err = kernel32.GetStdHandle(STD_ERROR_HANDLE)
        h_in  = kernel32.GetStdHandle(STD_INPUT_HANDLE)
        INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value
        if h_out == INVALID_HANDLE_VALUE or h_out is None:
            return
        try:
            import msvcrt
            fd_out = msvcrt.open_osfhandle(h_out, 0)
            fd_err = msvcrt.open_osfhandle(h_err, 0) if h_err not in (INVALID_HANDLE_VALUE, None) else None
            fd_in  = msvcrt.open_osfhandle(h_in,  0) if h_in  not in (INVALID_HANDLE_VALUE, None) else None
            if fd_out != -1:
                sys.stdout = io.TextIOWrapper(
                    os.fdopen(fd_out, 'wb', closefd=False),
                    encoding='utf-8', errors='replace', line_buffering=True
                )
            if fd_err != -1 and fd_err is not None:
                sys.stderr = io.TextIOWrapper(
                    os.fdopen(fd_err, 'wb', closefd=False),
                    encoding='utf-8', errors='replace', line_buffering=True
                )
            if fd_in != -1 and fd_in is not None:
                sys.stdin = io.TextIOWrapper(
                    os.fdopen(fd_in, 'rb', closefd=False),
                    encoding='utf-8', errors='replace'
                )
        except Exception:
            pass
    except ImportError:
        pass


# ============================================================
# 全局状态与配置
# ============================================================

RUN_LOG_MAX_BYTES = 50 * 1024 * 1024
RUN_LOG_BACKUP_COUNT = 3

SHA512_SALT = os.environ.get("DESALT_SALT", "")
SHA512_LEN = 16

CSV_COLUMNS = [
    "file_path",
    "PatientID_after",
    "desensitize_time",
]

# 未成功处理的文件清单（非 DICOM / 结构异常 / 处理失败），便于事后核对与补处理
FAILED_CSV_COLUMNS = [
    "file_path",
    "status",
    "reason",
]


# ============================================================
# 脱敏规则配置区
# ============================================================

REPLACE_RULES = {
    "PatientName":                 "ANONYMOUS",
    "InstitutionName":             "Shandong_Class3B_2_Hospital",
    "InstitutionAddress":          "",
    "ReferringPhysicianName":      "",
    "OperatorsName":               "",
    "RequestingPhysician":         "",
    "OtherPatientIDs":             "",
    "PatientAddress":              "",
    "PatientTelephoneNumbers":     "",
    "InstitutionalDepartmentName": "",
    "StationName":                 "",
    "StudyDescription":            "",
}

SHA_ENCODE_RULES = {
    "AccessionNumber":   16,
}

RECURSIVE_SHA_RULES = {
    "PatientID":                16,
    "ScheduledProcedureStepID": 16,
    "PerformedProcedureStepID": 16,
    "StudyID":                  16,
}

SPECIAL_DATE_RULES = {
    "PatientBirthDate": "year_0101",
}

KEYWORD_TO_TAG = {
    "PatientName":                 (0x0010, 0x0010),
    "PatientID":                   (0x0010, 0x0020),
    "PatientBirthDate":            (0x0010, 0x0030),
    "PatientAddress":              (0x0010, 0x1040),
    "PatientTelephoneNumbers":      (0x0010, 0x2154),
    "OtherPatientIDs":             (0x0010, 0x1000),
    "InstitutionName":             (0x0008, 0x0080),
    "InstitutionAddress":          (0x0008, 0x0081),
    "InstitutionalDepartmentName": (0x0008, 0x1040),
    "StationName":                 (0x0008, 0x1010),
    "ReferringPhysicianName":      (0x0008, 0x0090),
    "OperatorsName":               (0x0008, 0x1070),
    "RequestingPhysician":         (0x0032, 0x1032),
    "AccessionNumber":             (0x0008, 0x0050),
    "StudyDescription":            (0x0008, 0x1030),
    "StudyID":                     (0x0020, 0x0010),
    "ScheduledProcedureStepID":    (0x0040, 0x0009),
    "PerformedProcedureStepID":    (0x0040, 0x0253),
}

_RECURSIVE_SHA_TAGS = {}
for _kw, _n in RECURSIVE_SHA_RULES.items():
    _tag = KEYWORD_TO_TAG.get(_kw)
    if _tag:
        _RECURSIVE_SHA_TAGS[_tag] = _n

_REPLACE_TAGS = {}
for _kw, _val in REPLACE_RULES.items():
    _tag = KEYWORD_TO_TAG.get(_kw)
    if _tag:
        _REPLACE_TAGS[_tag] = _val

SAFE_ID_TAGS = {
    0x00100020, 0x00080050, 0x00200013, 0x00200011,
    0x0020000D, 0x0020000E, 0x00080018, 0x00100010,
}

_ID_RE = re.compile(r'^\d{4,12}$')

STRING_VRS = {'SH', 'LO', 'UI', 'PN', 'CS', 'IS', 'DS', 'ST', 'UT',
              'DA', 'TM', 'DT', 'AS', 'AE', 'UC', 'UR'}


# ============================================================
# 日志器
# ============================================================
class DuoLogger:
    def __init__(self, log_file):
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        import logging as _logging
        fmt = '%(asctime)s [%(levelname)s] %(message)s'
        self._logger = _logging.getLogger('dcm_local')
        self._logger.setLevel(_logging.DEBUG)
        self._logger.propagate = False
        fh = RotatingFileHandler(
            log_file, maxBytes=RUN_LOG_MAX_BYTES,
            backupCount=RUN_LOG_BACKUP_COUNT, encoding='utf-8'
        )
        fh.setFormatter(_logging.Formatter(fmt))
        self._err_cache = defaultdict(lambda: [0, 0.0])
        self._logger.addHandler(fh)

    def info(self, msg):    self._emit('INFO', msg)
    def warning(self, msg): self._emit('WARNING', msg)
    def error(self, msg):   self._emit('ERROR', msg)

    def debug(self, msg):
        try:
            self._logger.debug(msg)
        except Exception:
            pass

    def _emit(self, level, msg):
        try:
            if level == 'INFO':
                self._logger.info(msg)
            elif level == 'WARNING':
                self._logger.warning(msg)
            elif level == 'ERROR':
                key = hash(msg)
                now = time.time()
                cnt, last = self._err_cache[key]
                if now - last > 60:
                    self._err_cache[key] = [1, now]
                    self._logger.error(msg)
                elif cnt < 10:
                    self._err_cache[key] = [cnt + 1, last]
                    self._logger.error(msg)
        except Exception:
            pass
        if level in ('INFO', 'WARNING', 'ERROR'):
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            line = f"{ts} [{level}] {msg}"
            try:
                print(line, flush=True)
            except UnicodeEncodeError:
                try:
                    enc = sys.stdout.encoding or 'utf-8'
                    safe = line.encode(enc, errors='replace').decode(enc, errors='replace')
                    print(safe, flush=True)
                except Exception:
                    pass
            except Exception:
                pass


# ============================================================
# 工具函数
# ============================================================

def sha512_truncate(value, n=SHA512_LEN, salt=SHA512_SALT):
    if value is None:
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    if salt:
        raw = raw + salt
    return hashlib.sha512(raw.encode('utf-8')).hexdigest()[:n]


def _gen_da(s):
    s = str(s).strip()
    y = m = None
    if len(s) >= 8 and s[:8].isdigit():
        y, m = s[:4], s[4:6]
    elif len(s) >= 6 and s[:6].isdigit():
        y, m = s[:6][:4], s[:6][4:]
    if y and m:
        try:
            datetime(int(y), int(m), 1)
            return f"{y}{m}01"
        except ValueError:
            pass
    return ""


def _gen_dt(s):
    s = str(s).strip()
    if len(s) >= 8 and s[:8].isdigit():
        y, m = s[:4], s[4:6]
        try:
            datetime(int(y), int(m), 1)
            return f"{y}{m}01"
        except ValueError:
            pass
    return ""


def _transform_by_vr(vr, val, datetime_empty):
    if vr == 'TM':
        return ""
    if vr == 'DA':
        return "" if datetime_empty else _gen_da(val)
    if vr == 'DT':
        return "" if datetime_empty else _gen_dt(val)
    return val


def _collect_safe_ids(ds):
    safe = set()
    for t in SAFE_ID_TAGS:
        if t in ds:
            v = ds[t].value
            if v is not None:
                safe.add(str(v).strip())
    return safe


def _tag_tuple(tag):
    return (tag.group, tag.element)


def _is_sequence_value(v):
    if isinstance(v, (list, tuple)):
        return True
    try:
        if type(v).__name__ == 'Sequence':
            return True
    except Exception:
        pass
    try:
        if hasattr(v, '__len__') and hasattr(v, '__iter__'):
            if len(v) == 0:
                return True
            first = next(iter(v))
            return hasattr(first, 'data_element')
    except Exception:
        pass
    return False


def _deep_walk(ds, opts, safe_ids):
    to_delete = []
    for elem in list(ds):
        vr = elem.VR
        tag = elem.tag
        is_seq = (vr == 'SQ') or _is_sequence_value(elem.value)

        # 0. 递归 SHA512
        n = _RECURSIVE_SHA_TAGS.get(_tag_tuple(tag))
        if n is not None:
            if not is_seq:
                old = elem.value
                new = sha512_truncate(old, n)
                if new and new != str(old).strip():
                    elem.value = new
                continue

        # 0b. 递归固定值替换
        rep_val = _REPLACE_TAGS.get(_tag_tuple(tag))
        if rep_val is not None:
            if not is_seq:
                if str(elem.value).strip() != rep_val:
                    elem.value = rep_val
                continue

        # 1. 日期/时间 VR
        if not is_seq and vr in ('DA', 'DT', 'TM'):
            old = elem.value
            new = _transform_by_vr(vr, old, opts.datetime_empty)
            if new != old:
                elem.value = new

        # 2. 序列递归
        if is_seq:
            seq_items = elem.value
            if not isinstance(seq_items, (list, tuple)):
                try:
                    seq_items = list(seq_items)
                except Exception:
                    seq_items = []
            for item in seq_items:
                if hasattr(item, 'data_element'):
                    _deep_walk(item, opts, safe_ids)
            continue

        # 3. 私有标签
        if opts.purge_private and elem.is_private:
            to_delete.append(tag)

        # 4. 类影像号
        elif opts.purge_nested_ids and vr in STRING_VRS:
            vals = elem.value if isinstance(elem.value, (list, tuple)) else [elem.value]
            newvals = []
            changed = False
            for v in vals:
                s = str(v).strip() if v is not None else ""
                if _ID_RE.match(s) and s not in safe_ids:
                    newvals.append("")
                    changed = True
                else:
                    newvals.append(v)
            if changed:
                elem.value = newvals[0] if len(newvals) == 1 else newvals

    for t in to_delete:
        try:
            del ds[t]
        except KeyError:
            pass


# ============================================================
# 核心脱敏函数
# ============================================================

class _DefaultOpts:
    purge_private = True
    purge_nested_ids = False
    datetime_empty = False


def desensitize_dataset(ds, audit, opts=None):
    if opts is None:
        opts = _DefaultOpts()

    audit["desensitize_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    safe_ids = _collect_safe_ids(ds) if opts.purge_nested_ids else set()

    for keyword, n in SHA_ENCODE_RULES.items():
        tag = KEYWORD_TO_TAG.get(keyword)
        if tag and tag in ds:
            old = ds[tag].value
            new = sha512_truncate(old, n)
            ds[tag].value = new

    for keyword, val in REPLACE_RULES.items():
        tag = KEYWORD_TO_TAG.get(keyword)
        if tag and tag in ds:
            ds[tag].value = val

    for keyword, method in SPECIAL_DATE_RULES.items():
        tag = KEYWORD_TO_TAG.get(keyword)
        if tag and tag in ds:
            old = ds[tag].value
            if old:
                raw = str(old).strip()
                if method == "year_0101" and len(raw) >= 4:
                    ds[tag].value = f"{raw[:4]}0101"

    _deep_walk(ds, opts, safe_ids)

    for keyword in ("PatientID",):
        tag = KEYWORD_TO_TAG.get(keyword)
        if tag and tag in ds:
            cur_val = ds[tag].value
            if cur_val is not None:
                audit[f"{keyword}_after"] = str(cur_val).strip()

    return ds


# ============================================================
# CSV 审计表
# ============================================================

def _csv_line(values):
    """按 CSV 规则转义一行（逗号/换行/引号），返回带换行的字符串。"""
    fields = []
    for v in values:
        v = "" if v is None else str(v)
        if any(ch in v for ch in [',', '\n', '\r', '"']):
            v = '"' + v.replace('"', '""') + '"'
        fields.append(v)
    return ','.join(fields) + '\n'


def _csv_row(audit):
    return _csv_line([audit.get(col, "") for col in CSV_COLUMNS])


def _init_csv(csv_path):
    header = ','.join(CSV_COLUMNS) + '\n'
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(header)


def _append_csv(csv_path, rows):
    if not rows:
        return
    with open(csv_path, 'a', encoding='utf-8-sig', newline='') as f:
        for r in rows:
            f.write(_csv_row(r))


def _init_failed_csv(csv_path):
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        f.write(','.join(FAILED_CSV_COLUMNS) + '\n')


def _append_failed_csv(csv_path, rows):
    """rows: [(file_path, status, reason), ...]"""
    if not rows:
        return
    with open(csv_path, 'a', encoding='utf-8-sig', newline='') as f:
        for r in rows:
            f.write(_csv_line(r))


# ============================================================
# Worker 选项
# ============================================================

class WorkerOpts:
    __slots__ = ('purge_private', 'purge_nested_ids', 'datetime_empty')
    def __init__(self, purge_private, purge_nested_ids, datetime_empty):
        self.purge_private = purge_private
        self.purge_nested_ids = purge_nested_ids
        self.datetime_empty = datetime_empty


# ============================================================
# 工作进程（原地修改版）
# ============================================================

def worker_process(task_queue, result_queue, src_dir, opts):
    """
    Worker 主循环：取任务 → 脱敏 → 原子写入覆盖原文件 → 返回结果。
    原地修改：先写 <原文件>.desens_tmp → os.replace 覆盖原文件。
    """
    while True:
        try:
            rel_path = task_queue.get(timeout=3)
        except Empty:
            continue

        if rel_path is None:
            break

        in_path = Path(src_dir) / rel_path

        audit = {col: "" for col in CSV_COLUMNS}
        audit["file_path"] = rel_path

        try:
            ds = pydicom.dcmread(str(in_path), force=True, defer_size=256)

            if (0x0008, 0x0016) not in ds and (0x0010, 0x0020) not in ds:
                try:
                    result_queue.put(
                        ("SKIP_INVALID", rel_path, "无 SOPClassUID/PatientID", None),
                        timeout=2
                    )
                except Full:
                    pass
                continue

            ds = desensitize_dataset(ds, audit, opts)

            # 写入临时文件 → 原子替换
            buf = io.BytesIO()
            with warnings.catch_warnings():
                warnings.simplefilter('ignore', DeprecationWarning)
                warnings.simplefilter('ignore', UserWarning)
                try:
                    pydicom.dcmwrite(buf, ds, enforce_file_format=True)
                except TypeError:
                    pydicom.dcmwrite(buf, ds, write_like_original=False)
            data = buf.getvalue()

            if len(data) < 134 or data[128:132] != b'DICM':
                try:
                    result_queue.put(
                        ("ERROR", rel_path,
                         f"dcmwrite 产出非标准 DICOM (大小={len(data)})", None),
                        timeout=2
                    )
                except Full:
                    pass
                continue

            # 原子写入：先写临时文件，再替换原文件
            tmp_path = str(in_path) + '.desens_tmp'
            with open(tmp_path, 'wb') as f:
                f.write(data)
            os.replace(tmp_path, str(in_path))

            try:
                result_queue.put(("OK", rel_path, "", audit), timeout=2)
            except Full:
                pass

        except InvalidDicomError as e:
            try:
                result_queue.put(("SKIP", rel_path, str(e), None), timeout=2)
            except Full:
                pass

        except Exception as e:
            err = str(e) if str(e) else type(e).__name__
            try:
                result_queue.put(("ERROR", rel_path, err, None), timeout=2)
            except Full:
                pass
            # 清理可能残留的临时文件
            tmp_path = str(in_path) + '.desens_tmp'
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass


# ============================================================
# 核心引擎（原地修改版）
# ============================================================

def run_engine(src_dir, num_workers,
               purge_private=True, purge_nested_ids=False, datetime_empty=False):
    """
    DICOM 原地脱敏引擎 v8.0-local
    日志/审计/断点续传放在 <src>/_desensitize_log/ 子目录
    """
    log_dir = os.path.join(src_dir, "_desensitize_log")
    progress_log = os.path.join(log_dir, "desensitize_progress.log")
    run_log = os.path.join(log_dir, "desensitize_run.log")
    csv_path = os.path.join(log_dir, "desensitize_audit.csv")
    failed_csv = os.path.join(log_dir, "unparsed_files.csv")

    os.makedirs(log_dir, exist_ok=True)
    log = DuoLogger(run_log)

    interrupted = {'flag': False}

    def on_signal(signum, frame):
        interrupted['flag'] = True
        log.warning(f"[INTERRUPT] 收到信号 {signum}，停止派发新任务...")

    signal.signal(signal.SIGINT, on_signal)
    signal.signal(signal.SIGTERM, on_signal)

    _init_csv(csv_path)
    log.info(f"[CSV] 审计表已初始化: {csv_path}")
    _init_failed_csv(failed_csv)
    log.info(f"[CSV] 未处理文件清单已初始化: {failed_csv}")
    if SHA512_SALT:
        log.info("[SALT] 检测到盐值，SHA512 哈希已加盐")
    else:
        log.warning("[SALT] 未检测到 DESALT_SALT 环境变量，建议设置！")

    log.info(f"[MODE] 原地修改模式（in-place）：脱敏后直接覆盖原文件")

    # 断点续传
    processed_set = set()
    if os.path.exists(progress_log):
        with open(progress_log, 'r', encoding='utf-8') as f:
            processed_set = {line.strip() for line in f if line.strip()}
        log.info(f"[RESUME] 断点续传: 已加载 {len(processed_set)} 条历史记录")
        if processed_set:
            log.info(f"[RESUME] 如需重新处理已跳过的文件，请删除: {progress_log}")

    task_queue = mp.Queue(maxsize=max(num_workers * 8, 32))
    result_queue = mp.Queue(maxsize=2000)

    opts = WorkerOpts(purge_private, purge_nested_ids, datetime_empty)

    log.info(f"[START] 启动引擎 v8.0-local | Workers: {num_workers} | 源: {src_dir}")
    log.info(f"[CFG]   purge_private={purge_private} | purge_nested_ids={purge_nested_ids} "
             f"| datetime_empty={datetime_empty}")
    log.info(f"[CFG]   RECURSIVE_SHA_RULES={len(RECURSIVE_SHA_RULES)}条 "
             f"({','.join(RECURSIVE_SHA_RULES.keys())})")
    start_time = time.time()

    workers = []
    for _ in range(num_workers):
        p = mp.Process(
            target=worker_process,
            args=(task_queue, result_queue, src_dir, opts)
        )
        p.daemon = True
        p.start()
        workers.append(p)

    def _drain_quick():
        ok = err = 0
        csv_batch = []
        failed_batch = []
        for _ in range(500):
            try:
                status, rel_path, err_msg, audit = result_queue.get_nowait()
                if status == "OK":
                    ok += 1
                    if audit:
                        csv_batch.append(audit)
                else:
                    err += 1
                    failed_batch.append((rel_path, status, err_msg))
            except Empty:
                break
        if csv_batch:
            try:
                _append_csv(csv_path, csv_batch)
            except Exception as e:
                log.warning(f"[DRAIN_CSV_FAIL] {e}")
        if failed_batch:
            try:
                _append_failed_csv(failed_csv, failed_batch)
            except Exception as e:
                log.warning(f"[DRAIN_FAILED_CSV_FAIL] {e}")
        if ok or err:
            log.debug(f"[DRAIN] 扫描间隙消费: OK={ok}, ERR={err}")

    # ── 扫描并入队 ──
    scanned = skipped = 0
    src_path = Path(src_dir)
    last_log = time.time()
    try:
        for fp in src_path.rglob('*'):
            if interrupted['flag']:
                break
            if not fp.is_file():
                continue
            # 跳过日志目录内的文件
            try:
                rel = str(fp.relative_to(src_path))
            except ValueError:
                continue
            if rel.startswith("_desensitize_log"):
                continue
            if rel.endswith('.desens_tmp'):
                continue
            scanned += 1
            if rel in processed_set:
                skipped += 1
                continue
            while not interrupted['flag']:
                try:
                    task_queue.put(rel, timeout=0.5)
                    break
                except Full:
                    _drain_quick()
            if scanned % 500 == 0:
                _drain_quick()
                elapsed = time.time() - start_time
                log.info(f"[SCAN] 已扫描 {scanned}, 跳过 {skipped}, "
                         f"队列深度 {task_queue.qsize()}, 耗时 {elapsed:.1f}s")
                last_log = time.time()
            elif time.time() - last_log > 3:
                _drain_quick()
                log.info(f"[SCAN] 已扫描 {scanned}, 跳过 {skipped}, "
                         f"队列深度 {task_queue.qsize()}")
                last_log = time.time()
    except Exception as e:
        log.error(f"[SCAN] 扫描异常: {e}")

    scan_cost = time.time() - start_time
    log.info(f"[SCAN_DONE] 扫描完成! 总计 {scanned}, 跳过 {skipped}, 耗时 {scan_cost:.1f}s")

    # 毒丸
    pills_sent = 0
    pill_retries = 0
    while pills_sent < num_workers and pill_retries < num_workers * 3:
        if interrupted['flag']:
            break
        try:
            task_queue.put(None, timeout=2)
            pills_sent += 1
            pill_retries = 0
        except Full:
            pill_retries += 1
            _drain_quick()
            if pill_retries >= num_workers * 3:
                log.error(f"[PILL] 毒丸发送失败 {pill_retries} 次，可能 worker 卡死")

    if pills_sent < num_workers:
        log.warning(f"[PILL] 仅发出 {pills_sent}/{num_workers} 个毒丸，剩余 worker 将被 terminate")

    # ── 结果收集 ──
    ok_count = err_count = skip_count = skip_invalid_count = 0
    pending_progress = []
    pending_csv = []
    pending_failed = []
    total_tasks = scanned - skipped
    completed = 0
    last_prog_log = time.time()

    while True:
        got_any = False
        for _ in range(50):
            try:
                status, rel_path, err_msg, audit = result_queue.get(timeout=1 if not pending_progress else 0.1)
                got_any = True
            except Empty:
                break

            if status == "OK":
                ok_count += 1
                pending_progress.append(rel_path)
                if audit:
                    pending_csv.append(audit)
            elif status == "SKIP":
                skip_count += 1
                pending_failed.append((rel_path, status, err_msg))
            elif status == "SKIP_INVALID":
                skip_invalid_count += 1
                pending_failed.append((rel_path, status, err_msg))
            else:
                err_count += 1
                pending_failed.append((rel_path, status, err_msg))
                log.error(f"[ERROR] {rel_path}: {err_msg}")
            completed += 1

            if len(pending_progress) >= 50 or \
               (pending_progress and time.time() - last_prog_log > 5):
                try:
                    with open(progress_log, 'a', encoding='utf-8') as f:
                        f.write('\n'.join(pending_progress) + '\n')
                except Exception as e:
                    log.error(f"[PROGRESS_WRITE_FAIL] {e}")
                pending_progress.clear()
                if pending_csv:
                    try:
                        _append_csv(csv_path, pending_csv)
                    except Exception as e:
                        log.error(f"[CSV_WRITE_FAIL] {e}")
                    pending_csv.clear()
                if pending_failed:
                    try:
                        _append_failed_csv(failed_csv, pending_failed)
                    except Exception as e:
                        log.error(f"[FAILED_CSV_WRITE_FAIL] {e}")
                    pending_failed.clear()
                last_prog_log = time.time()

            if completed % 100 == 0 or completed == total_tasks:
                elapsed = time.time() - start_time
                speed = completed / elapsed if elapsed > 0 else 0
                remaining = total_tasks - completed
                eta_h = remaining / speed / 3600 if speed > 0 else float('inf')
                log.info(
                    f"[PROG] OK={ok_count} ERR={err_count} SKIP={skip_count} "
                    f"INVALID={skip_invalid_count} | {speed:.1f} f/s | ETA: {eta_h:.1f}h"
                )

        if not got_any:
            alive_count = sum(1 for w in workers if w.is_alive())
            if alive_count == 0 and result_queue.empty():
                break
            if interrupted['flag'] and (time.time() - start_time) > (scan_cost + 60):
                log.warning("[WAIT] 中断超时，强制结束 worker")
                for w in workers:
                    if w.is_alive():
                        w.terminate()
                break
            time.sleep(0.3)

    # 落盘剩余
    if pending_progress:
        try:
            with open(progress_log, 'a', encoding='utf-8') as f:
                f.write('\n'.join(pending_progress) + '\n')
        except Exception as e:
            log.error(f"[PROGRESS_WRITE_FAIL] {e}")
        pending_progress.clear()
    if pending_csv:
        try:
            _append_csv(csv_path, pending_csv)
        except Exception as e:
            log.error(f"[CSV_WRITE_FAIL] {e}")
        pending_csv.clear()
    if pending_failed:
        try:
            _append_failed_csv(failed_csv, pending_failed)
        except Exception as e:
            log.error(f"[FAILED_CSV_WRITE_FAIL] {e}")
        pending_failed.clear()

    try:
        task_queue.close()
        task_queue.join_thread()
    except Exception:
        pass
    try:
        result_queue.close()
        result_queue.join_thread()
    except Exception:
        pass

    for w in workers:
        w.join(timeout=3)
        if w.is_alive():
            log.warning(f"[CLEANUP] worker PID={w.pid} 仍在运行，强制终止")
            w.terminate()
            w.join(timeout=1)

    elapsed = time.time() - start_time
    log.info("=" * 50)
    log.info(f"[DONE] 任务结束! 耗时: {elapsed/3600:.2f}h ({elapsed/60:.1f} min)")
    log.info(f"      成功: {ok_count} | 失败: {err_count} | "
             f"跳过(非DICOM): {skip_count} | 无效文件: {skip_invalid_count}")
    unhandled = err_count + skip_count + skip_invalid_count
    log.info(f"      未成功处理合计: {unhandled}（失败+跳过+无效）")
    log.info(f"[CSV] 审计表: {csv_path}")
    log.info(f"[CSV] 未处理清单({unhandled} 条): {failed_csv}")
    if interrupted['flag']:
        log.info("      [状态] 被中断，可再次运行同命令断点续传")
    log.info("=" * 50)

    return 2 if interrupted['flag'] else (1 if err_count > 0 else 0)


# ============================================================
# 命令行入口
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="DICOM 原地脱敏引擎 v8.0-local (直接覆盖原文件)"
    )
    parser.add_argument("--src", required=True, help="源 DICOM 数据目录（原地修改）")
    parser.add_argument("--workers", type=int,
                        default=max(1, mp.cpu_count() - 1),
                        help="并发工作进程数（默认: CPU核心数-1）")
    parser.add_argument("--force", action="store_true",
                        help="跳过安全检查")
    parser.add_argument("--no-purge-private", action="store_true",
                        help="不清空私有标签（默认会清空）")
    parser.add_argument("--purge-nested-ids", action="store_true",
                        help="递归清空嵌套字段中的类影像号数字（4-12位纯数字，默认关闭）")
    parser.add_argument("--datetime-empty", action="store_true",
                        help="DA/DT 字段直接置空而非保留年-月（默认保留年-月）")
    return parser.parse_args()


if __name__ == '__main__':
    mp.freeze_support()
    _ensure_console()

    _is_worker = False
    try:
        _is_worker = mp.current_process().name.startswith('Worker')
    except Exception:
        pass

    if _is_worker:
        pass
    else:
        args = parse_args()

        if not os.path.isdir(args.src):
            print(f"错误: 源目录不存在: {args.src}")
            sys.exit(2)

        purge_private = not args.no_purge_private
        purge_nested_ids = args.purge_nested_ids
        datetime_empty = args.datetime_empty

        exit_code = run_engine(
            args.src, args.workers,
            purge_private=purge_private,
            purge_nested_ids=purge_nested_ids,
            datetime_empty=datetime_empty
        )
        sys.exit(exit_code)
