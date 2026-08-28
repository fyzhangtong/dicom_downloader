# -*- coding: utf-8 -*-
"""
DICOM 影像批量下载工具（可视化版）

功能：
    1. 读取 Excel 表格中的「影像号」(StudyInstanceUID) 列
    2. 通过 DICOM Q/R 协议（C-Move）从医院 PACS 批量拉取影像到本地
    3. 所有连接参数（PACS IP/端口/AE、本机 AE Title/接收端口）可视化配置

协议依据：众阳云开放服务 API「2.8.1.3 获取检查图像」——DICOM 3.0 Q/R
    C-Move 入参：StudyInstanceUID (0020,000D)

运行环境：Windows（打包为 exe 后双击运行即可，无需安装 Python）
"""

import json
import os
import queue
import socket
import sys
import threading
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from pynetdicom import AE, evt
from pynetdicom.presentation import StoragePresentationContexts, VerificationPresentationContexts
from pynetdicom.sop_class import (
    StudyRootQueryRetrieveInformationModelFind,
    StudyRootQueryRetrieveInformationModelMove,
    Verification,
)
from pydicom.dataset import Dataset


# ---------------------------------------------------------------------------
# 全局状态
# ---------------------------------------------------------------------------
OUTPUT_ROOT = ""
_lock = threading.Lock()
_current_study_dir = ""          # 当前正在下载的 Study 的落地子目录
_store_server = None            # Store SCP 服务实例
_store_started = False
_stop_event = threading.Event()
_rate_limit_kbps = 0           # 传输限速（KB/s），0 = 不限；由批量下载开始前设置
_rate_lock = threading.Lock()
_rate_last_time = 0.0          # 令牌桶：上次补充令牌的时间
_rate_tokens = 0.0             # 令牌桶：当前可用令牌（字节）

# 线程间通信队列（工作线程 -> GUI 主线程）
_ui_queue = queue.Queue()


def get_app_dir():
    """返回程序所在目录（兼容 PyInstaller 打包后的 exe）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(get_app_dir(), "config.json")


# ---------------------------------------------------------------------------
# 1) Store SCP：接收 PACS 通过 C-Move 推送的 DICOM 文件
# ---------------------------------------------------------------------------
def _safe_name(s):
    return "".join(c if (c.isalnum() or c in "._-") else "_" for c in str(s))


def handle_store(event):
    """处理 C-STORE：把收到的 DICOM 文件保存到对应 Study 子目录（并发安全）。"""
    ds = event.dataset
    ds.file_meta = event.file_meta
    sop_uid = getattr(ds, "SOPInstanceUID", None) or ("unknown-%d" % int(time.time()))
    study_uid = getattr(ds, "StudyInstanceUID", None)
    if study_uid:
        # 优先按图像自带的 StudyInstanceUID 定位目录（并发下载时不会串目录）
        d = os.path.join(OUTPUT_ROOT, _safe_name(study_uid))
    else:
        with _lock:
            d = _current_study_dir
    if not d:
        d = OUTPUT_ROOT
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, _safe_name(sop_uid) + ".dcm")
    ds.save_as(path, enforce_file_format=True)
    # 限速：按实际落盘文件大小节流，压低下行速率
    try:
        _throttle(os.path.getsize(path))
    except Exception:
        pass
    return 0x0000  # Success


def _throttle(size_bytes):
    """全局限速（并发安全，令牌桶）：把多个 C-STORE 的总下行速率压到目标值。"""
    global _rate_last_time, _rate_tokens
    limit = _rate_limit_kbps
    if limit <= 0 or size_bytes <= 0:
        return
    rate = limit * 1024.0  # 字节/秒
    while True:
        with _rate_lock:
            now = time.time()
            if _rate_last_time == 0:
                _rate_last_time = now
                _rate_tokens = rate  # 初始满桶（约 1 秒额度）
            else:
                elapsed = now - _rate_last_time
                _rate_tokens = min(rate, _rate_tokens + elapsed * rate)
                _rate_last_time = now
            if _rate_tokens >= size_bytes:
                _rate_tokens -= size_bytes
                return
            need = (size_bytes - _rate_tokens) / rate
        # 令牌不足：在锁外等待，再回来重试
        time.sleep(min(need, 5.0))


def start_store_scp(local_aet, local_port):
    """启动本机 Store SCP（后台线程，只启动一次）。"""
    global _store_server, _store_started
    if _store_started:
        _ui_queue.put(("log", "[StoreSCP] 接收服务已运行（AE=%s 端口=%d）" % (local_aet, local_port)))
        return True
    ae = AE()
    ae.add_supported_contexts(StoragePresentationContexts)
    ae.add_supported_contexts(VerificationPresentationContexts[0])
    handlers = [(evt.EVT_C_STORE, handle_store)]
    try:
        _store_server = ae.start_server(
            ("0.0.0.0", local_port), block=False, ae_title=local_aet, evt_handlers=handlers
        )
        _store_started = True
        _ui_queue.put(("log", "[StoreSCP] 接收服务已启动：AE=%s 端口=%d" % (local_aet, local_port)))
        return True
    except Exception as e:
        _ui_queue.put(("error", "Store SCP 启动失败：%s" % e))
        return False


# ---------------------------------------------------------------------------
# 2) 网络可达性检测
# ---------------------------------------------------------------------------
def test_connectivity(host, port, aet=None):
    """
    检测 PACS 前置机可达性。
    返回 (ok, messages)：
        1. TCP 端口连通性检查（判断网络/端口是否可达）
        2. DICOM C-Echo 检查（判断 DICOM 服务与 AE Title 是否有效，需要提供 aet）
    """
    msgs = []
    tcp_ok = False
    echo_ok = False

    # 1) TCP 连通性
    try:
        sock = socket.create_connection((host, port), timeout=5)
        sock.close()
        tcp_ok = True
        msgs.append("TCP 连接成功：%s:%d 可达" % (host, port))
    except Exception as e:
        msgs.append("TCP 连接失败：%s:%d 不可达（%s）" % (host, port, e))
        msgs.append("请检查：IP/端口是否正确、网络是否连通、是否需 VPN/白名单")
        return False, msgs

    # 2) DICOM C-Echo（验证 AE Title 与 DICOM 服务）
    if not aet:
        msgs.append("未填写 AE Title，跳过 DICOM C-Echo 验证")
        return tcp_ok, msgs

    try:
        ae = AE()
        ae.add_requested_context(Verification)
        assoc = ae.associate(host, port, ae_title=aet)
        if assoc.is_established:
            status = assoc.send_c_echo()
            if status and getattr(status, "Status", None) == 0x0000:
                echo_ok = True
                msgs.append("DICOM C-Echo 成功：AE Title=%s 有效" % aet)
            else:
                msgs.append("DICOM C-Echo 未返回成功状态")
            assoc.release()
        else:
            msgs.append("DICOM 关联建立失败：请检查 AE Title=%s 是否正确" % aet)
    except Exception as e:
        msgs.append("DICOM C-Echo 异常：%s" % e)

    return (tcp_ok and (echo_ok if aet else True)), msgs


# ---------------------------------------------------------------------------
# 3) 读取 Excel 中的「影像号」(StudyInstanceUID) 列
# ---------------------------------------------------------------------------
def read_study_uids(excel_path, column, sheet_name=None):
    """
    读取 Excel 指定列，返回去重后的非空 StudyInstanceUID 列表。
    column 可以是：
        - 列名（如 "影像号"，按表头匹配）
        - 字母（如 "A"）
        - 数字（如 "1"，1-based）
    """
    import openpyxl

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    all_rows = list(rows)

    def col_letter_to_idx(letter):
        idx = 0
        for ch in letter.upper():
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        return idx - 1

    col_idx = None
    start_row = 0  # 数据起始行（0 = 第 1 行）

    c = str(column).strip()
    if c.isdigit():
        col_idx = int(c) - 1
    elif c and all(ch.isalpha() for ch in c) and len(c) <= 3:
        col_idx = col_letter_to_idx(c)
    elif header:
        # 按表头匹配列名
        for i, h in enumerate(header):
            if str(h).strip() == c:
                col_idx = i
                break
        if col_idx is None:
            # 表头没匹配到，退回第一列，并把表头行也当作数据
            col_idx = 0
            start_row = 0
            all_rows = [header] + all_rows
    else:
        col_idx = 0

    if col_idx is None or col_idx < 0:
        col_idx = 0

    uids = []
    seen = set()
    for r in all_rows:
        if start_row == 0 and r is header:
            pass
        val = r[col_idx] if (r and len(r) > col_idx) else None
        if val is None:
            continue
        s = str(val).strip()
        if not s:
            continue
        if s not in seen:
            seen.add(s)
            uids.append(s)

    wb.close()
    return uids


# ---------------------------------------------------------------------------
# 3) C-Move 拉取单个 Study
# ---------------------------------------------------------------------------
_PENDING = {0xFF00, 0xFF01}


def _status_text(code):
    """把 DICOM C-Find/C-Move 状态码转成可读文本。"""
    m = {
        0x0000: "成功(Success)",
        0xFF00: "进行中(Pending)",
        0xFF01: "进行中(Pending, 有警告)",
        0xA700: "拒绝:资源不足(Out of Resources)",
        0xA701: "拒绝:无法计算匹配数(Unable to calculate matches)",
        0xA702: "拒绝:资源不足-无法计算匹配数",
        0xA801: "拒绝:目标节点未知(Move Destination unknown)",
        0xA900: "失败:标识符与SOP类不匹配(Identifier does not match SOP Class)",
        0xC000: "失败:无法处理(Unable to process)",
        0xC001: "失败:无法处理-部分键(Unable to process, some keys)",
        0xFE00: "取消(Cancel)",
        0xB000: "警告:子操作完成但有失败(Warning, some sub-ops failed)",
    }
    return m.get(code, "未知(0x%04X)" % code)


def find_studies_by_patient(assoc, patient_id):
    """C-Find 查询某个病人的所有 Study，返回 StudyInstanceUID 列表。"""
    ds = Dataset()
    ds.QueryRetrieveLevel = "STUDY"
    ds.PatientID = patient_id
    ds.StudyInstanceUID = ""
    ds.StudyDate = ""
    ds.Modality = ""
    ds.StudyDescription = ""
    ds.AccessionNumber = ""
    ds.StudyID = ""

    _ui_queue.put(("log", "      [C-Find] 查询 patientId=%s (QueryRetrieveLevel=STUDY)" % patient_id))
    uids = []
    last_status = None
    try:
        for status, identifier in assoc.send_c_find(
            ds, query_model=StudyRootQueryRetrieveInformationModelFind
        ):
            if status:
                last_status = status.Status
                _ui_queue.put(("log", "      [C-Find] 响应状态 0x%04X %s" % (status.Status, _status_text(status.Status))))
            if identifier:
                suid = getattr(identifier, "StudyInstanceUID", None)
                acc = getattr(identifier, "AccessionNumber", None)
                sid = getattr(identifier, "StudyID", None)
                sdate = getattr(identifier, "StudyDate", None)
                mod = getattr(identifier, "Modality", None)
                desc = getattr(identifier, "StudyDescription", None)
                _ui_queue.put(("log", "          -> StudyInstanceUID=%s 日期=%s 模态=%s 描述=%s 申请号=%s StudyID=%s" % (
                    suid, sdate, mod, desc, acc, sid)))
                if suid and str(suid) not in [str(x) for x in uids]:
                    uids.append(suid)
    except Exception as e:
        _ui_queue.put(("log", "      [C-Find 异常] %s" % e))
    _ui_queue.put(("log", "      [C-Find] 结束，最终状态 0x%04X %s，命中 %d 个 Study" % (
        last_status or 0, _status_text(last_status) if last_status is not None else "无响应", len(uids))))
    return uids


def pull_one_study(assoc, study_uid, local_aet, out_root):
    """拉取一个 Study，影像落到 out_root/<study_uid>/ 目录。"""
    global _current_study_dir
    subdir = os.path.join(out_root, _safe_name(study_uid))
    with _lock:
        _current_study_dir = subdir
    os.makedirs(subdir, exist_ok=True)

    ds = Dataset()
    ds.QueryRetrieveLevel = "STUDY"
    ds.StudyInstanceUID = study_uid

    _ui_queue.put(("log", "      [C-Move] 请求拉取 StudyInstanceUID=%s -> 目标AE=%s (QueryRetrieveLevel=STUDY)" % (study_uid, local_aet)))

    final_code = 0x0000
    has_error = False
    try:
        for status, _ in assoc.send_c_move(
            ds, local_aet, query_model=StudyRootQueryRetrieveInformationModelMove
        ):
            if status:
                code = status.Status
                # 子操作计数（C-Move 的 Pending 阶段会带这些值，可用来评估进度）
                remaining = getattr(status, "NumberOfRemainingSuboperations", None)
                completed = getattr(status, "NumberOfCompletedSuboperations", None)
                failed = getattr(status, "NumberOfFailedSuboperations", None)
                warned = getattr(status, "NumberOfWarningSuboperations", None)
                extra = ""
                if remaining is not None or completed is not None:
                    extra = " (剩余=%s 已完成=%s 失败=%s 警告=%s)" % (remaining, completed, failed, warned)
                _ui_queue.put(("log", "      [C-Move] 响应状态 0x%04X %s%s" % (code, _status_text(code), extra)))
                if code not in _PENDING:
                    if code != 0x0000:
                        final_code = code
                        has_error = True
    except Exception as e:
        has_error = True
        _ui_queue.put(("log", "      [C-Move 异常] %s" % e))

    n_files = len([f for f in os.listdir(subdir) if f.endswith(".dcm")]) if os.path.isdir(subdir) else 0
    _ui_queue.put(("log", "      [C-Move] 结束，最终状态 0x%04X %s，本地落盘 %d 个文件" % (final_code, _status_text(final_code), n_files)))
    return has_error, final_code, n_files, subdir


# ---------------------------------------------------------------------------
# 4) 批量下载（后台线程执行）
# ---------------------------------------------------------------------------
class DowloadConfig:
    def __init__(self):
        self.pacs_host = ""
        self.pacs_port = 104
        self.pacs_aet = ""
        self.local_aet = "MYAET"
        self.local_port = 11112
        self.excel_path = ""
        self.sheet_name = ""
        self.column = "影像号"
        self.key_type = "study_uid"   # study_uid(StudyInstanceUID) / patient_id(病人ID)
        self.out_dir = ""
        self.rate_limit_kbps = 0      # 限速（KB/s），0 = 不限
        self.pause_every = 0          # 每下载 N 个检查后暂停（仅串行模式），0 = 不启用
        self.pause_seconds = 30       # 暂停秒数
        self.concurrent_move = 1      # 并发下载的 Study 数（建议 3~5）
        self.cmove_timeout = 300      # C-MOVE 单次最大等待（秒），超时后重试
        self.cfind_timeout = 60       # C-FIND 单次最大等待（秒）
        self.cmove_retry = 1          # C-MOVE 失败/超时后重试次数
        self.cmove_retry_delay = 3    # C-MOVE 重试间隔（秒）


def _make_find_assoc(cfg):
    """建立 C-Find 用的 association（按 cfind_timeout 设置超时）。"""
    ae = AE()
    ae.add_requested_context(StudyRootQueryRetrieveInformationModelFind)
    ae.dimse_timeout = int(getattr(cfg, "cfind_timeout", 60) or 60)
    ae.network_timeout = None
    return ae.associate(cfg.pacs_host, cfg.pacs_port, ae_title=cfg.pacs_aet)


def _make_move_assoc(cfg):
    """建立 C-Move 用的 association（按 cmove_timeout 设置超时）。"""
    ae = AE()
    ae.add_requested_context(StudyRootQueryRetrieveInformationModelMove)
    ae.dimse_timeout = int(getattr(cfg, "cmove_timeout", 300) or 300)
    ae.network_timeout = None
    return ae.associate(cfg.pacs_host, cfg.pacs_port, ae_title=cfg.pacs_aet)


def _download_one(cfg, label, uid, idx, total):
    """下载单个 Study（失败/超时后自动重试）。返回 (idx, ok, fail_msg)。"""
    retry = int(getattr(cfg, "cmove_retry", 1) or 1)
    delay = int(getattr(cfg, "cmove_retry_delay", 3) or 3)
    fail_msg = ""
    n_files = 0
    subdir = ""
    for attempt in range(retry + 1):
        if _stop_event.is_set():
            _ui_queue.put(("log", "[%d/%d] [跳过] %s（手动停止）" % (idx, total, label)))
            return idx, False, "已停止"
        assoc = _make_move_assoc(cfg)
        if assoc.is_established:
            has_error, code, n_files, subdir = pull_one_study(assoc, uid, cfg.local_aet, OUTPUT_ROOT)
            try:
                assoc.release()
            except Exception:
                pass
            if not has_error:
                _ui_queue.put(("log", "[%d/%d] [完成] %s -> %d 个文件 %s" % (idx, total, label, n_files, subdir)))
                return idx, True, ""
            fail_msg = "状态码 0x%04X" % code
        else:
            fail_msg = "连接 PACS 失败"
        if attempt < retry:
            _ui_queue.put(("log", "[%d/%d] [重试 %d/%d] %s：%s，%d 秒后重试" % (
                idx, total, attempt + 1, retry, label, fail_msg, delay)))
            t_end = time.time() + delay
            while time.time() < t_end and not _stop_event.is_set():
                time.sleep(0.5)
    _ui_queue.put(("log", "[%d/%d] [失败] %s：%s（本地文件 %d 个）" % (idx, total, label, fail_msg, n_files)))
    return idx, False, fail_msg


def batch_download(cfg):
    """批量下载主流程（在线程中运行）。"""
    global OUTPUT_ROOT, _current_study_dir, _rate_limit_kbps
    global _rate_last_time, _rate_tokens
    _stop_event.clear()

    OUTPUT_ROOT = cfg.out_dir
    os.makedirs(OUTPUT_ROOT, exist_ok=True)
    _rate_limit_kbps = int(getattr(cfg, "rate_limit_kbps", 0) or 0)
    _rate_last_time = 0.0
    _rate_tokens = 0.0
    if _rate_limit_kbps > 0:
        _ui_queue.put(("log", "已启用限速：%d KB/s" % _rate_limit_kbps))

    # 1) 读 Excel（得到 keys：StudyInstanceUID 或 patientId）
    _ui_queue.put(("log", "正在读取 Excel：%s" % cfg.excel_path))
    try:
        keys = read_study_uids(cfg.excel_path, cfg.column, cfg.sheet_name or None)
    except Exception as e:
        _ui_queue.put(("error", "读取 Excel 失败：%s\n%s" % (e, traceback.format_exc())))
        return
    n_keys = len(keys)
    key_label = "StudyInstanceUID" if cfg.key_type == "study_uid" else "病人ID(patientId)"
    _ui_queue.put(("log", "共读取到 %d 个 %s" % (n_keys, key_label)))
    if n_keys == 0:
        _ui_queue.put(("error", "未读取到任何数据，请检查列名/Sheet 配置"))
        return

    # 2) 启动 Store SCP
    if not start_store_scp(cfg.local_aet, cfg.local_port):
        return

    # 3) 展开任务：把 keys 转成 (显示标签, StudyInstanceUID) 列表
    tasks = []
    if cfg.key_type == "patient_id":
        cfind_timeout = int(getattr(cfg, "cfind_timeout", 60) or 60)
        _ui_queue.put(("log", "模式：按病人ID查询，先 C-Find 找出每个病人的所有检查（超时 %d 秒）..." % cfind_timeout))
        assoc_find = _make_find_assoc(cfg)
        if not assoc_find.is_established:
            _ui_queue.put(("error", "连接 PACS 失败（C-Find），请检查 IP/端口/AE 及网络是否可达"))
            return
        for pi, pid in enumerate(keys, 1):
            if _stop_event.is_set():
                break
            uids = find_studies_by_patient(assoc_find, pid)
            if not uids:
                _ui_queue.put(("log", "  [%d/%d] patientId=%s 未查到任何检查" % (pi, n_keys, pid)))
                continue
            _ui_queue.put(("log", "  [%d/%d] patientId=%s -> 找到 %d 个检查" % (pi, n_keys, pid, len(uids))))
            for u in uids:
                tasks.append(("%s/%s" % (pid, u), u))
        try:
            assoc_find.release()
        except Exception:
            pass
    else:
        for u in keys:
            tasks.append((str(u), u))

    total = len(tasks)
    _ui_queue.put(("log", "共需下载 %d 个检查（Study）" % total))
    if total == 0:
        _ui_queue.put(("error", "没有可下载的检查"))
        return

    # 4) 下载（串行或并发）
    concurrent = int(getattr(cfg, "concurrent_move", 1) or 1)
    if concurrent < 1:
        concurrent = 1
    cmove_timeout = int(getattr(cfg, "cmove_timeout", 300) or 300)
    cmove_retry = int(getattr(cfg, "cmove_retry", 1) or 1)
    _ui_queue.put(("log", "开始下载：并发数=%d，C-Move 超时=%d 秒，重试=%d 次" % (concurrent, cmove_timeout, cmove_retry)))

    ok = 0
    fail = 0

    if concurrent == 1:
        # 串行（支持间隙暂停）
        pause_every = int(getattr(cfg, "pause_every", 0) or 0)
        pause_seconds = int(getattr(cfg, "pause_seconds", 0) or 0)
        for i, (label, uid) in enumerate(tasks, 1):
            if _stop_event.is_set():
                _ui_queue.put(("log", "已手动停止，剩余 %d 个未处理" % (total - i + 1)))
                break
            _ui_queue.put(("progress", (i, total)))
            _ui_queue.put(("log", "[%d/%d] 拉取 %s" % (i, total, label)))
            _idx, is_ok, msg = _download_one(cfg, label, uid, i, total)
            if is_ok:
                ok += 1
            elif msg == "已停止":
                break
            else:
                fail += 1

            # 间隙：每下载 N 个检查后暂停 M 秒（最后一个不暂停）
            if pause_every > 0 and pause_seconds > 0 and i % pause_every == 0 and i < total:
                if _stop_event.is_set():
                    break
                _ui_queue.put(("status", "下载暂停中（%d 秒）..." % pause_seconds))
                _ui_queue.put(("log", "  [间隙] 已下载 %d 个，暂停 %d 秒..." % (i, pause_seconds)))
                t_end = time.time() + pause_seconds
                while time.time() < t_end and not _stop_event.is_set():
                    time.sleep(0.5)
                _ui_queue.put(("status", "继续下载..."))
                _ui_queue.put(("log", "  [间隙] 暂停结束，继续下载"))
    else:
        # 并发：每个任务独立 association，线程池调度
        lock = threading.Lock()
        done = 0
        with ThreadPoolExecutor(max_workers=concurrent) as pool:
            futures = [pool.submit(_download_one, cfg, label, uid, i, total) for i, (label, uid) in enumerate(tasks, 1)]
            for fut in as_completed(futures):
                _idx, is_ok, msg = fut.result()
                with lock:
                    done += 1
                    if is_ok:
                        ok += 1
                    else:
                        fail += 1
                _ui_queue.put(("progress", (done, total)))
                _ui_queue.put(("log", "[汇总] %d/%d（成功 %d / 失败 %d）" % (done, total, ok, fail)))

    summary = "下载完成：成功 %d / 失败 %d / 共 %d" % (ok, fail, total)
    _ui_queue.put(("done", summary))


# ---------------------------------------------------------------------------
# 5) GUI 界面
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("DICOM 影像批量下载工具")
        root.geometry("760x720")
        root.minsize(720, 650)

        self.cfg = DowloadConfig()
        self.thread = None

        self._build_widgets()
        self._load_config()
        self._poll_queue()

    # ----- 布局 -----
    def _build_widgets(self):
        pad = dict(padx=6, pady=3)

        # PACS 配置
        f1 = ttk.LabelFrame(self.root, text="PACS 前置机配置（由医院实施工程师提供）")
        f1.pack(fill="x", padx=10, pady=(10, 4))
        self._entry_pair(f1, "PACS IP", "pacs_host", 0, default="")
        self._entry_pair(f1, "PACS 端口", "pacs_port", 1, default="104")
        self._entry_pair(f1, "PACS AE Title", "pacs_aet", 2, default="")

        # 本机配置
        f2 = ttk.LabelFrame(self.root, text="本机接收节点配置（需注册到 PACS 前置机）")
        f2.pack(fill="x", padx=10, pady=4)
        self._entry_pair(f2, "本机 AE Title", "local_aet", 0, default="MYAET")
        self._entry_pair(f2, "本机接收端口", "local_port", 1, default="11112")

        # 下载配置
        f3 = ttk.LabelFrame(self.root, text="下载配置")
        f3.pack(fill="x", padx=10, pady=4)

        row0 = ttk.Frame(f3); row0.pack(fill="x", **pad)
        ttk.Label(row0, text="Excel 文件:").pack(side="left")
        self.var_excel = tk.StringVar()
        ttk.Entry(row0, textvariable=self.var_excel, width=50).pack(side="left", padx=4)
        ttk.Button(row0, text="浏览...", command=self._browse_excel).pack(side="left")

        row_kt = ttk.Frame(f3); row_kt.pack(fill="x", **pad)
        ttk.Label(row_kt, text="查询键类型:").pack(side="left")
        self.var_key_type = tk.StringVar(value="StudyInstanceUID(影像号UID)")
        ttk.Combobox(
            row_kt, textvariable=self.var_key_type, state="readonly", width=28,
            values=["StudyInstanceUID(影像号UID)", "病人ID(patientId)"],
        ).pack(side="left", padx=4)
        ttk.Label(row_kt, text="(选病人ID时，会先查该病人的全部检查再拉取)").pack(side="left")

        row1 = ttk.Frame(f3); row1.pack(fill="x", **pad)
        ttk.Label(row1, text="关键列名:").pack(side="left")
        self.var_column = tk.StringVar(value="影像号")
        ttk.Entry(row1, textvariable=self.var_column, width=16).pack(side="left", padx=4)
        ttk.Label(row1, text="(列名/字母A/序号1)").pack(side="left")
        ttk.Label(row1, text="  Sheet(可空):").pack(side="left", padx=(16, 0))
        self.var_sheet = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_sheet, width=12).pack(side="left", padx=4)

        row2 = ttk.Frame(f3); row2.pack(fill="x", **pad)
        ttk.Label(row2, text="输出目录:").pack(side="left")
        self.var_out = tk.StringVar()
        ttk.Entry(row2, textvariable=self.var_out, width=50).pack(side="left", padx=4)
        ttk.Button(row2, text="浏览...", command=self._browse_out).pack(side="left")

        row3 = ttk.Frame(f3); row3.pack(fill="x", **pad)
        ttk.Label(row3, text="限速(KB/s):").pack(side="left")
        self.var_rate_limit = tk.StringVar(value="0")
        ttk.Entry(row3, textvariable=self.var_rate_limit, width=9).pack(side="left", padx=4)
        ttk.Label(row3, text="(0=不限速, 如1024=1MB/s)").pack(side="left")

        ttk.Label(row3, text="  每下载").pack(side="left", padx=(16, 0))
        self.var_pause_every = tk.StringVar(value="0")
        ttk.Entry(row3, textvariable=self.var_pause_every, width=6).pack(side="left", padx=4)
        ttk.Label(row3, text="个检查后暂停").pack(side="left")
        self.var_pause_seconds = tk.StringVar(value="30")
        ttk.Entry(row3, textvariable=self.var_pause_seconds, width=6).pack(side="left", padx=4)
        ttk.Label(row3, text="秒(0=不暂停)").pack(side="left")

        row4 = ttk.Frame(f3); row4.pack(fill="x", **pad)
        ttk.Label(row4, text="并发下载数:").pack(side="left")
        self.var_concurrent = tk.StringVar(value="1")
        ttk.Entry(row4, textvariable=self.var_concurrent, width=6).pack(side="left", padx=4)
        ttk.Label(row4, text="(建议3~5)").pack(side="left")

        ttk.Label(row4, text="  C-Move超时(秒):").pack(side="left", padx=(16, 0))
        self.var_cmove_timeout = tk.StringVar(value="300")
        ttk.Entry(row4, textvariable=self.var_cmove_timeout, width=6).pack(side="left", padx=4)
        ttk.Label(row4, text="  C-Find超时(秒):").pack(side="left", padx=(12, 0))
        self.var_cfind_timeout = tk.StringVar(value="60")
        ttk.Entry(row4, textvariable=self.var_cfind_timeout, width=6).pack(side="left", padx=4)

        row5 = ttk.Frame(f3); row5.pack(fill="x", **pad)
        ttk.Label(row5, text="重试次数:").pack(side="left")
        self.var_cmove_retry = tk.StringVar(value="1")
        ttk.Entry(row5, textvariable=self.var_cmove_retry, width=6).pack(side="left", padx=4)
        ttk.Label(row5, text="(失败/超时后)").pack(side="left")
        ttk.Label(row5, text="  重试间隔(秒):").pack(side="left", padx=(16, 0))
        self.var_cmove_retry_delay = tk.StringVar(value="3")
        ttk.Entry(row5, textvariable=self.var_cmove_retry_delay, width=6).pack(side="left", padx=4)

        # 按钮区
        fbtn = ttk.Frame(self.root)
        fbtn.pack(fill="x", padx=10, pady=6)
        self.btn_test = ttk.Button(fbtn, text="检测连通", command=self._test_connectivity)
        self.btn_test.pack(side="left", padx=4)
        self.btn_start = ttk.Button(fbtn, text="开始下载", command=self._start)
        self.btn_start.pack(side="left", padx=4)
        self.btn_stop = ttk.Button(fbtn, text="停止", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        ttk.Button(fbtn, text="保存配置", command=self._save_config).pack(side="left", padx=4)
        ttk.Button(fbtn, text="加载配置", command=self._load_config).pack(side="left", padx=4)

        # 进度条
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=4)
        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.var_status).pack(anchor="w", padx=12)

        # 日志
        self.log = scrolledtext.ScrolledText(self.root, height=16, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))

    def _entry_pair(self, parent, label, attr, row, default=""):
        r = ttk.Frame(parent); r.pack(fill="x", padx=6, pady=3)
        ttk.Label(r, text=label + ":", width=16).pack(side="left")
        var = tk.StringVar(value=default)
        ttk.Entry(r, textvariable=var, width=40).pack(side="left", padx=4)
        setattr(self, "var_" + attr, var)

    # ----- 事件 -----
    def _browse_excel(self):
        p = filedialog.askopenfilename(
            title="选择影像号 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if p:
            self.var_excel.set(p)

    def _browse_out(self):
        p = filedialog.askdirectory(title="选择输出目录")
        if p:
            self.var_out.set(p)

    def _collect_cfg(self):
        c = DowloadConfig()
        c.pacs_host = self.var_pacs_host.get().strip()
        c.pacs_port = int(self.var_pacs_port.get().strip() or 104)
        c.pacs_aet = self.var_pacs_aet.get().strip()
        c.local_aet = self.var_local_aet.get().strip() or "MYAET"
        c.local_port = int(self.var_local_port.get().strip() or 11112)
        c.excel_path = self.var_excel.get().strip()
        c.sheet_name = self.var_sheet.get().strip()
        c.column = self.var_column.get().strip() or "影像号"
        c.key_type = "patient_id" if "病人" in self.var_key_type.get() else "study_uid"
        c.out_dir = self.var_out.get().strip()
        try:
            c.rate_limit_kbps = int(self.var_rate_limit.get().strip() or 0)
        except ValueError:
            c.rate_limit_kbps = 0
        try:
            c.pause_every = int(self.var_pause_every.get().strip() or 0)
        except ValueError:
            c.pause_every = 0
        try:
            c.pause_seconds = int(self.var_pause_seconds.get().strip() or 30)
        except ValueError:
            c.pause_seconds = 30
        try:
            c.concurrent_move = int(self.var_concurrent.get().strip() or 1)
        except ValueError:
            c.concurrent_move = 1
        if c.concurrent_move < 1:
            c.concurrent_move = 1
        try:
            c.cmove_timeout = int(self.var_cmove_timeout.get().strip() or 300)
        except ValueError:
            c.cmove_timeout = 300
        try:
            c.cfind_timeout = int(self.var_cfind_timeout.get().strip() or 60)
        except ValueError:
            c.cfind_timeout = 60
        try:
            c.cmove_retry = int(self.var_cmove_retry.get().strip() or 1)
        except ValueError:
            c.cmove_retry = 1
        try:
            c.cmove_retry_delay = int(self.var_cmove_retry_delay.get().strip() or 3)
        except ValueError:
            c.cmove_retry_delay = 3
        return c

    def _test_connectivity(self):
        host = self.var_pacs_host.get().strip()
        if not host:
            messagebox.showwarning("提示", "请先填写 PACS IP")
            return
        try:
            port = int(self.var_pacs_port.get().strip() or 104)
        except ValueError:
            messagebox.showwarning("提示", "PACS 端口必须是数字")
            return
        aet = self.var_pacs_aet.get().strip() or None

        self._append_log("正在检测 %s:%d 连通性 ..." % (host, port))
        self.btn_test.config(state="disabled")

        def worker():
            ok, msgs = test_connectivity(host, port, aet)
            _ui_queue.put(("conn_result", (ok, msgs)))

        threading.Thread(target=worker, daemon=True).start()

    def _start(self):
        if self.thread and self.thread.is_alive():
            messagebox.showinfo("提示", "已有下载任务在运行中")
            return
        cfg = self._collect_cfg()
        if not cfg.pacs_host or not cfg.pacs_aet:
            messagebox.showwarning("提示", "请填写 PACS IP 和 AE Title")
            return
        if not cfg.excel_path or not os.path.exists(cfg.excel_path):
            messagebox.showwarning("提示", "请选择有效的影像号 Excel 文件")
            return
        if not cfg.out_dir:
            messagebox.showwarning("提示", "请选择输出目录")
            return

        self.progress["maximum"] = 100
        self.progress["value"] = 0
        self._clear_log()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.var_status.set("下载中...")

        self.thread = threading.Thread(target=batch_download, args=(cfg,), daemon=True)
        self.thread.start()

    def _stop(self):
        _stop_event.set()
        self._append_log("正在停止（当前 Study 完成后生效）...")

    def _save_config(self):
        cfg = self._collect_cfg()
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(cfg.__dict__, f, ensure_ascii=False, indent=2)
            self._append_log("配置已保存：%s" % CONFIG_PATH)
        except Exception as e:
            messagebox.showerror("错误", "保存配置失败：%s" % e)

    def _load_config(self):
        if not os.path.exists(CONFIG_PATH):
            return
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                d = json.load(f)
            for k, v in d.items():
                if hasattr(self.cfg, k):
                    setattr(self.cfg, k, v)
            self.var_pacs_host.set(str(self.cfg.pacs_host or ""))
            self.var_pacs_port.set(str(self.cfg.pacs_port))
            self.var_pacs_aet.set(str(self.cfg.pacs_aet or ""))
            self.var_local_aet.set(str(self.cfg.local_aet or ""))
            self.var_local_port.set(str(self.cfg.local_port))
            self.var_out.set(str(self.cfg.out_dir or ""))
            self.var_column.set(str(self.cfg.column or "影像号"))
            self.var_sheet.set(str(self.cfg.sheet_name or ""))
            if getattr(self.cfg, "key_type", "study_uid") == "patient_id":
                self.var_key_type.set("病人ID(patientId)")
            else:
                self.var_key_type.set("StudyInstanceUID(影像号UID)")
            self.var_rate_limit.set(str(getattr(self.cfg, "rate_limit_kbps", 0)))
            self.var_pause_every.set(str(getattr(self.cfg, "pause_every", 0)))
            self.var_pause_seconds.set(str(getattr(self.cfg, "pause_seconds", 30)))
            self.var_concurrent.set(str(getattr(self.cfg, "concurrent_move", 1)))
            self.var_cmove_timeout.set(str(getattr(self.cfg, "cmove_timeout", 300)))
            self.var_cfind_timeout.set(str(getattr(self.cfg, "cfind_timeout", 60)))
            self.var_cmove_retry.set(str(getattr(self.cfg, "cmove_retry", 1)))
            self.var_cmove_retry_delay.set(str(getattr(self.cfg, "cmove_retry_delay", 3)))
            self._append_log("已自动加载配置：%s" % CONFIG_PATH)
        except Exception as e:
            print("load config error:", e)

    # ----- 日志/进度 -----
    def _append_log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                kind, payload = _ui_queue.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "error":
                    self._append_log("[错误] " + payload)
                    messagebox.showerror("错误", payload[:500])
                elif kind == "conn_result":
                    ok, msgs = payload
                    for m in msgs:
                        self._append_log(m)
                    self.btn_test.config(state="normal")
                    if ok:
                        messagebox.showinfo("检测结果", "\n".join(msgs))
                    else:
                        messagebox.showwarning("检测结果", "\n".join(msgs))
                elif kind == "progress":
                    done, total = payload
                    if total > 0:
                        self.progress["value"] = done * 100.0 / total
                        self.var_status.set("进度 %d / %d" % (done, total))
                elif kind == "status":
                    self.var_status.set(payload)
                elif kind == "done":
                    self._append_log(payload)
                    self.var_status.set(payload)
                    self.btn_start.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    messagebox.showinfo("完成", payload)
        except queue.Empty:
            pass
        self.root.after(200, self._poll_queue)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()