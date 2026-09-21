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

import csv
import hashlib
import json
import os
import queue
import re
import socket
import sys
import threading
import time
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from pynetdicom import AE, evt
from pynetdicom.presentation import StoragePresentationContexts
try:
    # 更全的存储 SOP 类（含各类压缩格式），能接收更多推送，减少前置机"子操作失败"
    from pynetdicom.presentation import AllStoragePresentationContexts as _ALL_STORAGE_CTX
except Exception:
    _ALL_STORAGE_CTX = None
from pynetdicom.sop_class import (
    StudyRootQueryRetrieveInformationModelFind,
    StudyRootQueryRetrieveInformationModelMove,
    Verification,
)
from pydicom.dataset import Dataset

import pydicom
# 宽容解析（重要）：少数厂商会在私有标签里写入与 VR 不匹配的长度，pydicom 查不到私有标签
# 的 VR 时会按 "0 bytes per value" 直接抛错，严格模式下整份文件会被判失败（旧脱敏工具正是
# 在这类数据上全量失败）。打开该开关后按 UN 原始字节保留，随后由模块①的「清空私有标签」
# 规则清掉，落盘仍是脱敏后的影像；「读取校验降级为警告」可避免同类不规范的标签中断流程。
try:
    pydicom.config.convert_wrong_length_to_UN = True
    pydicom.config.settings.reading_validation_mode = pydicom.config.WARN
except Exception:
    pass


# ---------------------------------------------------------------------------
# 全局状态
# ---------------------------------------------------------------------------
OUTPUT_ROOT = ""
_store_server = None            # Store SCP 服务实例
_store_started = False
_store_aet = ""                 # Store SCP 当前监听的 AE Title
_store_port = 0                 # Store SCP 当前监听的端口
_stop_event = threading.Event()
_rate_limit_kbps = 0           # 传输限速（KB/s），0 = 不限；由批量下载开始前设置
_rate_lock = threading.Lock()
_rate_last_time = 0.0          # 令牌桶：上次补充令牌的时间
_rate_tokens = 0.0             # 令牌桶：当前可用令牌（字节）
_throttle_cap_logged = False   # 限速封顶提示是否已输出（避免刷屏）
_active_assocs = []            # 当前活跃的 DICOM association，停止时强制中断
_assoc_lock = threading.Lock()
# 当前正在下载的检查信息（供 handle_store 通知 GUI 进度用）
_current_label = ""
_current_study_uid = ""        # 当前正在下载的 StudyInstanceUID（handle_store 据此判断是否推送进度）
_current_expected_images = 0   # 前置机报告的预期影像数（用于显示"X / N"）
# 按 Study 精确统计本批次落盘情况：同一病人多个检查共用一个 PatientID 目录时，
# 不能按目录计数（会把上一个检查的文件算进来），必须按影像自带的 StudyInstanceUID 计数
_store_lock = threading.Lock()
_store_received = {}           # study_uid -> 本批次已成功处理（落盘或按规则丢弃）的文件数
_store_counts = {}             # study_uid -> 本批次实际落盘文件数（含过滤命中的子目录文件）
_store_dirs = {}               # study_uid -> 正常影像保存目录（原始 PatientID 命名）
_store_filtered_dirs = {}      # study_uid -> 过滤命中影像保存目录（_剂量报告/原始 PatientID）
_store_filtered = {}           # study_uid -> 其中命中过滤规则的文件数
_store_discarded = {}          # study_uid -> 其中按「不保存」开关被丢弃的文件数
_store_failed = {}             # study_uid -> 其中处理失败（解析/脱敏异常）而拒绝落盘的文件数
_store_failed_files = []       # 未落盘文件明细：study_uid / SOPInstanceUID / 原因 / 时间
_download_start_time = 0.0     # 本次批量下载的开始时间戳（用于"已耗时"）
_diag_log = False              # 诊断日志开关：PDU/协商/Pending 细节默认隐藏，排查时打开
_desens_enabled = False        # 模块① 标签脱敏开关（下载开始时由配置设置）
_desens_rules = []             # 模块① 已解析的标签脱敏规则（tag 为 (group, element)）
_desens_opts = {"purge_private": True, "date_month": True, "time_clear": True}
_filter_enabled = False        # 模块② 剂量报告/截屏过滤开关（下载开始时由配置设置）
_filter_rules = []             # 模块② 已解析的过滤规则
_filter_save_hit = True        # 命中的过滤影像是否保存（False = 直接丢弃不落盘）
# 模块① 脱敏审计日志：**一个影像一行**，每个配置的规则字段各占一列，单元格为「原值 → 新值」，
# 主键为影像保存的文件名。全局的日期压缩/时间清空与私有标签清空作用于任意数量的标签、
# 无法预先确定列，统一汇总进「其它变更」列。
# 注意：审计表含脱敏前的原值（PatientID 除外），属敏感文件。
_desens_audit_rows = []        # 待落盘的审计行缓冲
_desens_audit_csv = ""         # 本次下载的审计明细临时 CSV 路径
_desens_audit_header = False   # 临时 CSV 是否已写过表头（每次下载重置）
_desens_audit_fields = []      # 本次下载的字段列（"GGGG,EEEE" 列表，取自配置规则）
_desens_audit_lock = threading.Lock()
_DESENS_AUDIT_FLUSH = 2000     # 缓冲阈值：超过即刷盘
_DESENS_AUDIT_FIXED_HEAD = ["影像文件名", "StudyInstanceUID", "脱敏后PatientID"]
_DESENS_AUDIT_TAIL_HEAD = ["其它变更", "时间"]
# 反向探测状态：检测连通时发空 C-Move，观察前置机是否主动连入本机接收服务
_probe_active = False
_probe_conn_event = threading.Event()
_probe_conn_addr = ""

# 下载结果记录（每次下载一条，最终写入 CSV 报告并在 GUI 历史列表展示）
_download_records = []
_records_lock = threading.Lock()

# 线程间通信队列（工作线程 -> GUI 主线程）
_ui_queue = queue.Queue()


def _diag(msg):
    """诊断日志：仅在打开「诊断日志」开关时输出（PDU/协商/Pending 等细节）。"""
    if _diag_log:
        _ui_queue.put(("log", msg))


def get_app_dir():
    """返回程序所在目录（兼容 PyInstaller 打包后的 exe）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(get_app_dir(), "config.json")


# ---------------------------------------------------------------------------
# 0) 脱敏 / 过滤规则（可在「脱敏配置」界面维护，随 config.json 保存/加载）
#    在 C-STORE 接收时直接处理内存中的 Dataset，写出的就是脱敏后的影像，
#    省掉了"下载后再回读 + 重写"那轮磁盘 I/O。
#    执行顺序固定为「先过滤、后脱敏」：过滤可能依赖 SeriesDescription 这类
#    会被脱敏清空的字段，顺序反了就永远匹配不到。
# ---------------------------------------------------------------------------
_FILTERED_SUBDIR = "_剂量报告"   # 模块②命中的影像落到输出目录下该子目录

# 标签脱敏方法：(内部键, 界面显示名, 参数说明)
_DESENS_METHODS = [
    ("hash",       "哈希（SHA512 截断）",     "截断位数，默认 16"),
    ("replace",    "固定值替换",              "替换为（留空 = 置空）"),
    ("clear",      "清空该标签",              "（无参数）"),
    ("remove",     "删除该标签",              "（无参数）"),
    ("birth_year", "出生日期 → 年 + 0101",    "（无参数）"),
    ("date_month", "日期 → 年-月-01 (DA/DT)", "（无参数）"),
    ("time_clear", "时间清空 (TM)",           "（无参数）"),
]
_DESENS_METHOD_LABEL = {k: v for k, v, _h in _DESENS_METHODS}
_DESENS_METHOD_HINT = {k: h for k, _v, h in _DESENS_METHODS}
_DESENS_PARAM_METHODS = {"hash", "replace"}   # 需要填写参数的方法

# 过滤匹配模式
_FILTER_MODES = [
    ("exact",    "精确匹配"),
    ("contains", "包含匹配"),
    ("regex",    "正则匹配"),
]
_FILTER_MODE_LABEL = {k: v for k, v in _FILTER_MODES}

# 常用标签预设（下拉可直接选，也可手动输入关键字或 Tag 号）
_COMMON_TAGS = [
    "PatientName", "PatientID", "PatientBirthDate", "PatientSex", "PatientAge",
    "PatientAddress", "PatientTelephoneNumbers", "OtherPatientIDs",
    "AccessionNumber", "StudyID", "StudyDescription", "SeriesDescription",
    "InstitutionName", "InstitutionAddress", "InstitutionalDepartmentName",
    "StationName", "ReferringPhysicianName", "OperatorsName", "RequestingPhysician",
    "ScheduledProcedureStepID", "PerformedProcedureStepID",
    "ImageType", "Modality", "BodyPartExamined",
    "StudyDate", "SeriesDate", "AcquisitionDate",
    "StudyTime", "SeriesTime", "AcquisitionTime",
]

# 默认规则：与既有脱敏程序的默认行为一致（不改配置时行为不变）
_DEFAULT_DESENS_RULES = [
    {"tag": "0008,0050", "method": "hash",       "param": "16", "recursive": False},  # AccessionNumber
    {"tag": "0010,0020", "method": "hash",       "param": "16", "recursive": True},   # PatientID
    {"tag": "0040,0009", "method": "hash",       "param": "16", "recursive": True},   # ScheduledProcedureStepID
    {"tag": "0040,0253", "method": "hash",       "param": "16", "recursive": True},   # PerformedProcedureStepID
    {"tag": "0020,0010", "method": "hash",       "param": "16", "recursive": True},   # StudyID
    {"tag": "0010,0030", "method": "birth_year", "param": "",   "recursive": False},  # PatientBirthDate
    {"tag": "0010,0010", "method": "replace",    "param": "ANONYMOUS",                  "recursive": False},
    {"tag": "0008,0080", "method": "replace",    "param": "Shandong_Class3B_2_Hospital", "recursive": False},
    {"tag": "0008,0081", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0008,0090", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0008,1070", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0032,1032", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0010,1000", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0010,1040", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0010,2154", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0008,1040", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0008,1010", "method": "replace",    "param": "",   "recursive": False},
    {"tag": "0008,1030", "method": "replace",    "param": "",   "recursive": False},  # StudyDescription
]

_DEFAULT_FILTER_RULES = [
    {"tag": "0008,0008", "mode": "exact",
     "values": ["SCREEN SAVE", "SCREENSAVE"]},                                       # ImageType
    {"tag": "0008,103E", "mode": "exact",
     "values": ["DoseReport", "Dose Report", "BBS_Display", "BPM_Display"]},         # SeriesDescription
]


def tag_to_str(tag):
    """把 (group, element) 转成 "GGGG,EEEE" 字符串。"""
    try:
        return "%04X,%04X" % (int(tag[0]), int(tag[1]))
    except Exception:
        return ""


def parse_tag(text):
    """解析用户输入的标签为 (group, element)，供脱敏/过滤规则使用。

    支持：关键字（PatientName）、Tag 号（"0010,0020" / "(0010,0020)" / "0010 0020"）。
    无法识别时抛 ValueError，由界面提示用户。
    """
    s = str(text or "").strip()
    if not s:
        raise ValueError("标签不能为空")
    try:
        from pydicom.datadict import tag_for_keyword
        t = tag_for_keyword(s)
    except Exception:
        t = None
    if t is not None:
        return (int(t) >> 16, int(t) & 0xFFFF)
    m = re.match(r"^\(?\s*([0-9A-Fa-f]{4})\s*[,;\s]\s*([0-9A-Fa-f]{4})\s*\)?$", s)
    if m:
        return (int(m.group(1), 16), int(m.group(2), 16))
    raise ValueError("无法识别的标签：%s\n可输入关键字（如 PatientName）或 Tag 号（如 0010,0020）" % s)


def tag_display(text):
    """把标签渲染成 "PatientID (0010,0020)" 便于阅读。"""
    try:
        g, e = parse_tag(text)
    except Exception:
        return str(text or "")
    code = "%04X,%04X" % (g, e)
    try:
        from pydicom.datadict import keyword_for_tag
        kw = keyword_for_tag((g << 16) | e) or ""
    except Exception:
        kw = ""
    return ("%s (%s)" % (kw, code)) if kw else code


def prepare_desens_rules(raw_rules):
    """把配置里的标签脱敏规则解析成可执行规则（附加 _tag 元组）。"""
    out = []
    seen = set()
    for r in raw_rules or []:
        try:
            tag = parse_tag(r.get("tag"))
        except Exception:
            continue
        if tag in seen:
            continue          # 同一标签只取第一条，避免两条规则互相覆盖
        seen.add(tag)
        out.append({
            "tag": tag,
            "method": r.get("method", "clear"),
            "param": r.get("param", ""),
            "recursive": bool(r.get("recursive")),
        })
    return out


def prepare_filter_rules(raw_rules):
    """把配置里的过滤规则解析成可执行规则（附加 _tag 元组、values 列表）。"""
    out = []
    for r in raw_rules or []:
        try:
            tag = parse_tag(r.get("tag"))
        except Exception:
            continue
        vals = r.get("values") or []
        if isinstance(vals, str):
            vals = [v for v in re.split(r"[;\n]", vals) if v.strip()]
        vals = [str(v) for v in vals if str(v).strip()]
        if not vals:
            continue
        out.append({"tag": tag, "mode": r.get("mode", "exact"), "values": vals})
    return out


def _sha512_trunc(value, n=16):
    """SHA512 截断哈希；空值返回空串。"""
    if value is None:
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    return hashlib.sha512(raw.encode("utf-8")).hexdigest()[:n]


def _mk_ym01(y, m):
    """把年/月拼成 YYYYMM01；非法日期返回空串。"""
    if not (y and m):
        return ""
    try:
        datetime(int(y), int(m), 1)
    except ValueError:
        return ""
    return y + m + "01"


def _trunc_date(val):
    """DA 压缩为「年-月-01」。"""
    s = str(val or "").strip()
    if len(s) >= 8 and s[:8].isdigit():
        return _mk_ym01(s[:4], s[4:6])
    if len(s) >= 6 and s[:6].isdigit():
        return _mk_ym01(s[:6][:4], s[:6][4:])
    return ""


def _trunc_datetime(val):
    """DT 压缩为「年-月-01」。"""
    s = str(val or "").strip()
    if len(s) >= 8 and s[:8].isdigit():
        return _mk_ym01(s[:4], s[4:6])
    return ""


def _is_sequence(v):
    """判断元素值是否为序列（pydicom 的 Sequence/MultiValue 不一定继承 list）。"""
    if isinstance(v, (list, tuple)):
        return True
    if type(v).__name__ == "Sequence":
        return True
    try:
        if len(v) == 0:
            return True
        return hasattr(next(iter(v)), "data_element")
    except Exception:
        return False


def _audit_add(audit, elem, method, before, after):
    """把一条「实际发生变化」的字段脱敏记录追加到 audit（audit 为 None 时忽略）。

    PatientID 是主标识：审计表只保留脱敏后的值，原始 PatientID 不写入日志。
    """
    if audit is None:
        return
    tag = (elem.tag.group, elem.tag.element)
    if tag == (0x0010, 0x0020):
        before = "(原始值不记录)"
    # 清空类操作的「脱敏后」写成空单元格在 Excel 里不易分辨，统一标注
    after = "(空)" if after in (None, "") else str(after)
    try:
        from pydicom.datadict import keyword_for_tag
        kw = keyword_for_tag((tag[0] << 16) | tag[1]) or ""
    except Exception:
        kw = ""
    audit.append({
        "tag": "%04X,%04X" % tag,
        "keyword": kw,
        "vr": elem.VR or "",
        "method": _DESENS_METHOD_LABEL.get(method, method),
        "before": before,
        "after": after,
    })


def _apply_one_rule(elem, rule, audit=None):
    """对单个元素套用一条标签脱敏规则。

    返回 True 表示该标签应被删除。传入 audit 列表时，把实际发生变化的字段
    追加进去（脱敏前后对比），供脱敏审计日志使用；未变化的字段不记录。
    """
    method = rule.get("method")
    if method == "remove":
        _audit_add(audit, elem, "remove", str(elem.value), "(已删除)")
        return True
    if method == "hash":
        try:
            n = int(rule.get("param") or 16)
        except Exception:
            n = 16
        n = max(1, min(n, 128))
        old = elem.value
        new = _sha512_trunc(old, n)
        if new and new != str(old).strip():
            elem.value = new
            _audit_add(audit, elem, "hash", str(old), new)
    elif method == "replace":
        rep = rule.get("param", "")
        rep = "" if rep is None else str(rep)
        if str(elem.value).strip() != rep:
            old = str(elem.value)
            elem.value = rep
            _audit_add(audit, elem, "replace", old, rep)
    elif method == "clear":
        if str(elem.value or "").strip() != "":
            old = str(elem.value)
            elem.value = ""
            _audit_add(audit, elem, "clear", old, "")
    elif method == "birth_year":
        raw = str(elem.value or "").strip()
        if len(raw) >= 4:
            elem.value = raw[:4] + "0101"
            _audit_add(audit, elem, "birth_year", raw, elem.value)
    elif method == "date_month":
        vr = elem.VR
        new = _trunc_datetime(elem.value) if vr == "DT" else _trunc_date(elem.value)
        if new != elem.value:
            old = str(elem.value)
            elem.value = new
            _audit_add(audit, elem, "date_month", old, new)
    elif method == "time_clear":
        if elem.VR == "TM" and elem.value not in (None, ""):
            old = str(elem.value)
            elem.value = ""
            _audit_add(audit, elem, "time_clear", old, "")
    return False


def _deep_walk(ds, rules_map, recursive_map, opts, audit=None):
    """递归套用脱敏规则。

    - rules_map：标签 -> 规则（顶层与递归规则都在内），命中即套用并跳过 VR 通用处理；
    - recursive_map：只含「作用于序列内」的规则，递归进序列时改用它；
    - opts：全局开关（清空私有标签 / 日期压缩到月 / 时间清空），对所有层级生效；
    - audit：可选，收集「实际发生的字段变更」供脱敏审计日志使用。

    返回本层及所有嵌套层被删除的私有标签数量（供审计里汇总成一条记录）。
    """
    to_delete = []
    n_private = 0
    for elem in list(ds):
        tag = (elem.tag.group, elem.tag.element)
        vr = elem.VR
        is_seq = (vr == "SQ") or _is_sequence(elem.value)

        rule = rules_map.get(tag)
        if rule is not None and not is_seq:
            if _apply_one_rule(elem, rule, audit):
                to_delete.append(elem.tag)
            continue

        if not is_seq:
            if vr == "TM":
                if opts["time_clear"] and elem.value not in (None, ""):
                    old = str(elem.value)
                    elem.value = ""
                    _audit_add(audit, elem, "time_clear", old, "")
            elif vr in ("DA", "DT"):
                if opts["date_month"]:
                    new = _trunc_datetime(elem.value) if vr == "DT" else _trunc_date(elem.value)
                    if new != elem.value:
                        old = str(elem.value)
                        elem.value = new
                        _audit_add(audit, elem, "date_month", old, new)

        if is_seq:
            try:
                items = list(elem.value)
            except Exception:
                items = []
            for item in items:
                if hasattr(item, "data_element"):
                    n_private += _deep_walk(item, recursive_map, recursive_map, opts, audit)
            continue

        if opts["purge_private"] and elem.is_private:
            to_delete.append(elem.tag)
            n_private += 1

    for t in to_delete:
        try:
            del ds[t]
        except KeyError:
            pass
    return n_private


def desensitize_dataset(ds, rules, opts, audit=None):
    """就地脱敏一个 DICOM Dataset（模块①，按配置的规则执行）。

    保留 StudyInstanceUID / SeriesInstanceUID / SOPInstanceUID 等关联 UID 不变
    （除非用户显式对它们配置了规则），确保脱敏后仍是一套完整、可正常浏览的检查。

    传入 audit 列表时，逐字段记录脱敏前后对比（供脱敏审计日志）。
    """
    rules_map = {}
    recursive_map = {}
    for r in rules or []:
        tag = r.get("tag")
        if tag is None:
            continue
        if tag not in rules_map:
            rules_map[tag] = r
        if r.get("recursive") and tag not in recursive_map:
            recursive_map[tag] = r
    n_private = _deep_walk(ds, rules_map, recursive_map, opts, audit)
    # 私有标签数量多且无业务含义，逐个记录会让日志爆掉，这里在每个影像上汇总成一条
    if audit is not None and n_private:
        audit.append({
            "tag": "(私有标签)", "keyword": "", "vr": "",
            "method": _DESENS_METHOD_LABEL.get("clear_private", "清空私有标签"),
            "before": "(已批量删除)", "after": "共 %d 个" % n_private,
        })
    return ds


def _match_one_token(tok, values, mode):
    """单个取值 token 与规则命中值集合比对（不区分大小写）。"""
    t = str(tok or "").strip()
    if not t:
        return False
    tl = t.lower()
    for v in values:
        vs = str(v).strip()
        if not vs:
            continue
        if mode == "regex":
            try:
                if re.search(vs, t, re.IGNORECASE):
                    return True
            except re.error:
                # 正则非法时退化为包含匹配，避免静默漏判
                if vs.lower() in tl:
                    return True
        elif mode == "contains":
            if vs.lower() in tl:
                return True
        else:  # exact
            if tl == vs.lower():
                return True
    return False


def _matches_filter_rules(ds, rules):
    """判断影像是否命中「剂量报告/截屏」过滤规则（模块②，任一规则命中即算命中）。

    多值字段（如 ImageType = "ORIGINAL\\PRIMARY\\SCREEN SAVE"）按 '\\' 切分后逐段比对。
    """
    for rule in rules or []:
        tag = rule.get("tag")
        if tag is None:
            continue
        try:
            if tag not in ds:
                continue
            val = ds[tag].value
        except Exception:
            continue
        if isinstance(val, (str, bytes)):
            parts = [val]
        elif hasattr(val, "__iter__"):
            parts = list(val)
        else:
            parts = [val]
        mode = rule.get("mode", "exact")
        values = rule.get("values") or []
        for p in parts:
            for tok in str(p).split("\\"):
                if _match_one_token(tok, values, mode):
                    return True
    return False


# ---------------------------------------------------------------------------
# 1) Store SCP：接收 PACS 通过 C-Move 推送的 DICOM 文件
# ---------------------------------------------------------------------------
def _safe_name(s):
    if s is None:
        return "unnamed"
    s = str(s)
    safe = "".join(c if (c.isalnum() or c in "._-") else "_" for c in s)
    # 极端兜底：纯特殊字符（如 NULL 字节）会被全部替换为 _，但仍然是非空
    return safe if safe.strip("_") else "unnamed"


def _on_conn_open(event):
    """Store SCP 收到入站连接时记录来源，便于判断 PACS 是否真的在往本机推影像。"""
    global _probe_conn_addr
    try:
        addr = "%s:%s" % (event.address[0], event.address[1])
    except Exception:
        addr = "未知"
    _ui_queue.put(("log", "[StoreSCP] 收到入站连接：%s" % addr))
    if _probe_active:
        _probe_conn_addr = addr
        _probe_conn_event.set()


def _on_conn_close(event):
    """Store SCP 连接关闭时记录来源，配合「入站连接」判断对方连上后做了什么。"""
    try:
        addr = "%s:%s" % (event.address[0], event.address[1])
    except Exception:
        addr = "未知"
    _ui_queue.put(("log", "[StoreSCP] 连接关闭：%s" % addr))


def _on_echo(event):
    """Store SCP 支持 C-ECHO：部分前置机推影像前会先对目标 AE 做 C-ECHO 校验，
    不支持会导致对方中止推送（表现为 C-Move 成功但 0 文件落盘）。"""
    try:
        addr = "%s:%s" % (event.assoc.remote_address[0], event.assoc.remote_address[1])
    except Exception:
        addr = "未知"
    _ui_queue.put(("log", "[StoreSCP] 收到 C-ECHO 校验：%s（已应答成功）" % addr))
    return 0x0000


def _ae_str(v):
    """AE Title 可能是 bytes，统一转成可读字符串。"""
    try:
        if isinstance(v, bytes):
            return v.decode("ascii", "replace").strip()
        return str(v)
    except Exception:
        return "?"


def _cx_abstract_name(cx):
    try:
        return cx.abstract_syntax.name
    except Exception:
        try:
            return str(cx.abstract_syntax)
        except Exception:
            return "?"


def _cx_transfer_name(cx):
    try:
        ts = cx.transfer_syntax
        if isinstance(ts, (list, tuple)):
            ts = ts[0] if ts else ""
        return _ae_str(ts)
    except Exception:
        return "?"


def _pdu_summary(pdu):
    """PDU 一行摘要。P-DATA-TF（影像数据块）只输出字节长度，避免刷屏。"""
    name = type(pdu).__name__
    if name == "P_DATA_TF":
        try:
            n = pdu.pdu_length
        except Exception:
            n = -1
        return "P-DATA-TF（数据块 %d 字节）" % n
    return name


def _make_pdu_handlers(side):
    """带侧别前缀的 PDU 捕获处理器（诊断级，默认不输出），避免 SCU/SCP 两侧日志混淆。"""

    def _recv(event):
        try:
            _diag("[%s][PDU←] %s" % (side, _pdu_summary(event.pdu)))
        except Exception:
            pass

    def _sent(event):
        try:
            _diag("[%s][PDU→] %s" % (side, _pdu_summary(event.pdu)))
        except Exception:
            pass

    return _recv, _sent


def _log_assoc_negotiation(assoc, tag):
    """输出关联协商结果（双方 AE、同意/拒绝的呈现上下文）。诊断级，默认不输出。"""
    try:
        remote = "%s:%s" % (assoc.remote_address[0], assoc.remote_address[1])
    except Exception:
        remote = "未知"
    try:
        calling = _ae_str(assoc.requestor.primitive.calling_ae_title)
        called = _ae_str(assoc.requestor.primitive.called_ae_title)
    except Exception:
        calling, called = "?", "?"
    _diag("[%s] 关联协商完成：对端=%s 调用方AE=%s 被叫方AE=%s" % (tag, remote, calling, called))
    try:
        for cx in assoc.accepted_contexts:
            _diag("[%s]   [同意] %s / %s" % (tag, _cx_abstract_name(cx), _cx_transfer_name(cx)))
        for cx in assoc.rejected_contexts:
            _diag("[%s]   [拒绝] %s" % (tag, _cx_abstract_name(cx)))
    except Exception:
        pass


def _on_accepted(event):
    """Store SCP 接受关联时输出协商结果：排查"对方连上却不发 C-STORE"。"""
    _log_assoc_negotiation(event.assoc, "StoreSCP协商")


def _on_aborted(event):
    """关联被 ABORT 时记录来源：区分"对方主动中止"与"正常释放"。"""
    try:
        addr = "%s:%s" % (event.assoc.remote_address[0], event.assoc.remote_address[1])
    except Exception:
        addr = "未知"
    _ui_queue.put(("log", "[StoreSCP] 关联被中止（A-ABORT）：%s" % addr))


def _local_ips():
    """本机非回环 IPv4 地址列表，用于核对 PACS 侧注册的 IP。"""
    ips = []
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if ip not in ips:
                ips.append(ip)
    except Exception:
        pass
    return ips


def _record_failed_file(study_key, sop_uid, reason):
    """记录一个「处理失败、未落盘」的文件（并发安全），供统计与失败清单 CSV 使用。

    只在解析/脱敏等异常导致文件被拒绝落盘时调用；命中过滤主动丢弃的不算失败。
    """
    with _store_lock:
        _store_failed[study_key] = _store_failed.get(study_key, 0) + 1
        _store_failed_files.append({
            "study_uid": study_key,
            "sop_uid": sop_uid,
            "reason": reason,
            "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        })


def handle_store(event):
    """处理 C-STORE：可选「过滤 + 脱敏」后保存到对应子目录（并发安全）。

    落盘前即完成过滤与脱敏，写出的就是脱敏后的影像（比"下载后再单独跑脱敏程序"
    少一轮磁盘回读+重写）。脱敏失败时拒绝落盘，避免写入未脱敏影像。

    目录名固定用「影像自带的原始 PatientID」（缺失时回落到 StudyInstanceUID），
    与是否开启脱敏无关：即使开启了模块①，文件夹也不会变成哈希值。

    顺序固定「先过滤、后脱敏」：过滤可能依赖会被脱敏清空的字段。
    命中过滤且关闭「保存命中的过滤影像」时，文件直接丢弃不落盘（仍回成功状态，
    避免前置机把该子操作记为失败）。
    """
    # 预置失败标识：异常若发生在解析早期（取不到 SOP/Study 之前），仍能记录来源
    sop_uid = "?"
    suid_key = "_unknown"
    try:
        ds = event.dataset
        ds.file_meta = event.file_meta
        sop_uid = getattr(ds, "SOPInstanceUID", None) or ("unknown-%d" % int(time.time()))
        study_uid = getattr(ds, "StudyInstanceUID", None)
        suid_key = str(study_uid) if study_uid else "_unknown"
        # 关键防御：OUTPUT_ROOT 尚未设置或已被清空时（如程序关闭中/已停止）拒绝写入
        if not OUTPUT_ROOT:
            _ui_queue.put(("log", "[StoreSCP] 收到 DICOM 但输出目录未就绪，已拒绝：SOP=%s" % sop_uid))
            return 0xC000

        # 模块②：剂量报告/截屏判断
        # 注意：必须在脱敏之前判断——SeriesDescription 会被模块①清空，脱敏后就匹配不到了
        is_filtered = _filter_enabled and _matches_filter_rules(ds, _filter_rules)

        # 命中过滤 + 未开启「保存命中的过滤影像」→ 丢弃（不回写磁盘，但仍回成功状态）
        if is_filtered and not _filter_save_hit:
            with _store_lock:
                _store_received[suid_key] = _store_received.get(suid_key, 0) + 1
                _store_filtered[suid_key] = _store_filtered.get(suid_key, 0) + 1
                _store_discarded[suid_key] = _store_discarded.get(suid_key, 0) + 1
                n_recv = _store_received[suid_key]
            try:
                if suid_key == _current_study_uid:
                    n_total = max(n_recv, _current_expected_images)
                    _ui_queue.put(("image_progress", (_current_label, n_recv, n_total)))
            except Exception:
                pass
            return 0x0000  # Success：已按规则丢弃，不算失败

        # 目录名固定用「影像自带的原始 PatientID」（不论是否开启模块①，都不用脱敏后的哈希值），
        # 因此必须在脱敏前先取出原始值
        raw_patient_id = getattr(ds, "PatientID", None)

        # 模块①：标签脱敏（就地修改，落盘即为脱敏后影像）
        audit_rows = []
        if _desens_enabled:
            try:
                desensitize_dataset(ds, _desens_rules, _desens_opts, audit_rows)
            except Exception as e:
                reason = "脱敏失败：%s" % (str(e) or type(e).__name__)
                _record_failed_file(suid_key, sop_uid, reason)
                _ui_queue.put(("log", "[脱敏] 处理失败，已拒绝落盘以免泄露原始影像：SOP=%s %s" % (sop_uid, e)))
                return 0xC000

        # 统一用影像自带的原始 PatientID 作目录名（缺失时回落到 StudyInstanceUID）
        folder_key = raw_patient_id or study_uid
        base = os.path.join(OUTPUT_ROOT, _FILTERED_SUBDIR) if is_filtered else OUTPUT_ROOT
        if folder_key:
            d = os.path.join(base, _safe_name(folder_key))
        else:
            d = os.path.join(base, "_unknown_study")
            _ui_queue.put(("log", "[StoreSCP] 收到缺少定位字段的文件，落入 _unknown_study/"))
        os.makedirs(d, exist_ok=True)
        path = os.path.join(d, _safe_name(sop_uid) + ".dcm")
        # 极少数情况下（pynetdicom 内部并发回调）save_as 可能撞到 Windows 文件锁
        # 一次重试即可恢复
        try:
            ds.save_as(path, enforce_file_format=True)
        except Exception as e:
            _ui_queue.put(("log", "[StoreSCP] 首次保存失败，重试一次：%s" % e))
            time.sleep(0.05)
            ds.save_as(path, enforce_file_format=True)
        # 保存成功后才计数（并发安全）：按 Study 精确统计本批次落盘数与目录，
        # 同一病人多个检查共用 PatientID 目录时，按目录计数会把其它检查的文件算进来
        with _store_lock:
            _store_received[suid_key] = _store_received.get(suid_key, 0) + 1
            _store_counts[suid_key] = _store_counts.get(suid_key, 0) + 1
            if is_filtered:
                _store_filtered[suid_key] = _store_filtered.get(suid_key, 0) + 1
                _store_filtered_dirs[suid_key] = d
            else:
                _store_dirs[suid_key] = d
            n_recv = _store_received[suid_key]
        # 脱敏审计：仅在文件真正落盘后记录（未落盘的文件不入审计表），
        # 主键为影像保存的文件名，并带上脱敏后的 PatientID
        if audit_rows:
            _append_desens_audit(os.path.basename(path), suid_key,
                                 getattr(ds, "PatientID", None), audit_rows)
        # 限速：按实际落盘文件大小节流，压低下行速率
        try:
            _throttle(os.path.getsize(path))
        except Exception:
            pass
        # 通知 GUI：仅当前正在下载的检查推送进度（避免上一个检查的尾部推送干扰当前进度条）
        try:
            if suid_key == _current_study_uid:
                n_total = max(n_recv, _current_expected_images)
                _ui_queue.put(("image_progress", (_current_label, n_recv, n_total)))
        except Exception:
            pass
        return 0x0000  # Success
    except Exception as e:
        reason = "%s：%s" % (type(e).__name__, str(e) or "(无详细信息)")
        _record_failed_file(suid_key, sop_uid, reason)
        _ui_queue.put(("log", "[StoreSCP] 保存 DICOM 失败，已拒绝落盘：SOP=%s %s" % (sop_uid, reason)))
        return 0xC000  # Unable to process，避免单文件失败中断接收


def _throttle(size_bytes):
    """全局限速（并发安全，令牌桶）：把多个 C-STORE 的总下行速率压到目标值。

    注意：限速通过"延迟 C-STORE 应答"实现，前置机若长时间收不到应答会把该子操作
    记为失败（表现为 failed=N、最终 0xA702）。因此单文件等待封顶 2 秒：
    限速值过低时实际速率会高于设定值，但保证下载不被前置机超时打断。
    """
    global _rate_last_time, _rate_tokens, _throttle_cap_logged
    limit = _rate_limit_kbps
    if limit <= 0 or size_bytes <= 0:
        return
    rate = limit * 1024.0  # 字节/秒
    waited = 0.0
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
        # 单文件等待封顶 2 秒：避免 C-STORE 应答过慢被前置机判超时（failed/0xA702）
        if waited >= 2.0:
            if not _throttle_cap_logged:
                _throttle_cap_logged = True
                _ui_queue.put(("log", "[限速] 限速值过低，单文件等待已封顶 2 秒，实际速率将高于设定值（不影响下载成功）"))
            return
        step = min(need, 5.0, 2.0 - waited)
        time.sleep(step)
        waited += step


def _try_bind_port(port):
    """尝试绑定本机端口（绑定后立即释放），用于检测端口是否可用。返回 (ok, err)。"""
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        s.bind(("0.0.0.0", port))
        return True, ""
    except OSError as e:
        return False, str(e)
    finally:
        try:
            s.close()
        except Exception:
            pass


def _port_bind_hint(port):
    """本机端口无法绑定时的可操作建议。"""
    return ("端口 %d 无法绑定的可能原因：\n"
            "1) 本工具的其它实例或其它程序正占用该端口，请关闭后重试；\n"
            "2) 端口落在 Windows Hyper-V/WSL 保留的排除范围内，可在 cmd 运行 "
            "netsh interface ipv4 show excludedportrange protocol=tcp 查看，"
            "改用一个不在范围内的端口（需重新向医院前置机注册）或重启电脑后重试；\n"
            "3) 被杀毒软件/防火墙拦截，请将本程序加入白名单。" % port)


def start_store_scp(local_aet, local_port):
    """启动本机 Store SCP（后台线程）。AE/端口变化时自动重启，保证与 C-Move 目标一致。"""
    global _store_server, _store_started, _store_aet, _store_port
    if _store_started and _store_server is not None:
        if _store_aet == local_aet and _store_port == local_port:
            _ui_queue.put(("log", "[StoreSCP] 接收服务已运行（AE=%s 端口=%d）" % (local_aet, local_port)))
            return True
        # 监听配置变化：关闭旧服务，重新启动
        _ui_queue.put(("log", "[StoreSCP] 本地接收参数变化，正在重启（AE=%s 端口=%d）..." % (local_aet, local_port)))
        try:
            _store_server.shutdown()
        except Exception as e:
            _ui_queue.put(("log", "[StoreSCP] 关闭旧接收服务异常（忽略）：%s" % e))
        _store_server = None
        _store_started = False
        # 端口释放后等待再 bind，避免 TIME_WAIT 导致新服务启动失败
        time.sleep(0.3)

    ae = AE()
    # 优先用更全的存储上下文（含压缩格式），减少前置机因"格式不被接收"而子操作失败
    storage_ctxs = _ALL_STORAGE_CTX or StoragePresentationContexts
    # StoragePresentationContexts 是列表，需要逐个 add_supported_context
    for ctx in storage_ctxs:
        ae.add_supported_context(ctx.abstract_syntax, ctx.transfer_syntax)
    ae.add_requested_context(Verification)
    # 作为 SCP 也要支持 C-ECHO：部分前置机推影像前先对目标 AE 做 C-ECHO 校验
    ae.add_supported_context(Verification)
    pdu_recv, pdu_sent = _make_pdu_handlers("SCP")
    handlers = [
        (evt.EVT_C_STORE, handle_store),
        (evt.EVT_C_ECHO, _on_echo),
        (evt.EVT_CONN_OPEN, _on_conn_open),
        (evt.EVT_CONN_CLOSE, _on_conn_close),
        (evt.EVT_ACCEPTED, _on_accepted),
        (evt.EVT_ABORTED, _on_aborted),
        (evt.EVT_PDU_RECV, pdu_recv),
        (evt.EVT_PDU_SENT, pdu_sent),
    ]
    try:
        ae.ae_title = local_aet
    except Exception:
        pass
    try:
        _store_server = ae.start_server(
            ("0.0.0.0", local_port), block=False, evt_handlers=handlers, ae_title=local_aet,
        )
    except Exception as e:
        _ui_queue.put(("error", "Store SCP 启动失败：%s\n%s" % (e, _port_bind_hint(local_port))))
        return False
    _store_started = True
    _store_aet = local_aet
    _store_port = local_port
    ips = "、".join(_local_ips()) or "未知"
    _ui_queue.put(("log", "[StoreSCP] 接收服务已启动：AE=%s 端口=%d（本机 IP：%s，请确认与前置机注册一致）" % (local_aet, local_port, ips)))
    return True


# ---------------------------------------------------------------------------
# 2) 网络可达性检测
# ---------------------------------------------------------------------------
def test_connectivity(host, port, aet=None, local_port=None, local_aet=None):
    """
    双向检测 PACS 前置机与本机的连通性。
    返回 (ok, messages)：
        正向（本机 -> 前置机）：
            1. TCP 端口连通性检查
            2. DICOM C-Echo 检查（验证前置机服务与 AE Title，需要提供 aet）
        反向（前置机 -> 本机）：
            3. 本机接收服务自检：经本机外部 IP 连接收服务并 C-Echo（需要 local_port/local_aet）
            4. 前置机反向探测：发一个不存在 UID 的 C-Move，观察前置机是否主动连入本机
               接收服务（仅作参考提示，不影响 ok 判定，因为部分前置机对 0 匹配不连目标）
        本机接收端口自检：启动 Store SCP 本身即端口绑定测试
    """
    global _probe_active, _probe_conn_addr
    msgs = []
    tcp_ok = False
    echo_ok = False
    local_ok = True
    self_ok = True

    # 0) 确保本机接收服务在运行（反向检测依赖它；启动本身即端口绑定自检）
    store_ready = False
    if local_port:
        if local_aet:
            store_ready = start_store_scp(local_aet, local_port)
            if store_ready:
                msgs.append("本机接收服务就绪：AE=%s 端口=%d" % (local_aet, local_port))
            else:
                local_ok = False
                msgs.append("本机接收服务启动失败：端口 %d 无法绑定" % local_port)
                msgs.append(_port_bind_hint(local_port))
        elif not (_store_started and _store_port == local_port):
            ok, err = _try_bind_port(local_port)
            if ok:
                msgs.append("本机接收端口自检：端口 %d 可正常绑定" % local_port)
            else:
                local_ok = False
                msgs.append("本机接收端口自检失败：端口 %d 无法绑定（%s）" % (local_port, err))
                msgs.append(_port_bind_hint(local_port))

    # 1) TCP 连通性（正向）
    try:
        sock = socket.create_connection((host, port), timeout=5)
        sock.close()
        tcp_ok = True
        msgs.append("TCP 连接成功：%s:%d 可达" % (host, port))
    except Exception as e:
        msgs.append("TCP 连接失败：%s:%d 不可达（%s）" % (host, port, e))
        msgs.append("请检查：IP/端口是否正确、网络是否连通、是否需 VPN/白名单")
        return False, msgs

    # 2) DICOM C-Echo（正向：验证前置机服务与 AE Title）
    if not aet:
        msgs.append("未填写 PACS AE Title，跳过正向 DICOM C-Echo 验证")
    else:
        try:
            ae = AE()
            ae.add_requested_context(Verification)
            ae.acse_timeout = 10
            ae.network_timeout = 10
            assoc = ae.associate(host, port, ae_title=aet)
            if assoc.is_established:
                try:
                    status = assoc.send_c_echo()
                    if status and getattr(status, "Status", None) == 0x0000:
                        echo_ok = True
                        msgs.append("正向 DICOM C-Echo 成功：AE Title=%s 有效" % aet)
                    else:
                        msgs.append("正向 DICOM C-Echo 未返回成功状态")
                finally:
                    try:
                        assoc.release()
                    except Exception:
                        pass
            else:
                msgs.append("正向 DICOM 关联建立失败：请检查 AE Title=%s 是否正确" % aet)
                # 未建立时也要尝试 release 释放 socket 资源
                try:
                    assoc.release()
                except Exception:
                    pass
        except Exception as e:
            msgs.append("正向 DICOM C-Echo 异常：%s" % e)

    # 3) 反向自检：经本机外部 IP 连自己的接收服务并 C-Echo（验证监听/AE/应答全链路）
    # 逐个外部 IP 尝试（VPN/虚拟网卡可能不可路由，任一成功即视为通过）
    if store_ready or (_store_started and local_port and _store_port == local_port):
        self_ok = False
        self_err = ""
        for self_ip in (_local_ips() or ["127.0.0.1"]):
            try:
                ae = AE()
                ae.add_requested_context(Verification)
                ae.acse_timeout = 5
                ae.network_timeout = 5
                assoc = ae.associate(self_ip, local_port, ae_title=local_aet or "SELFECHO")
                if assoc.is_established:
                    try:
                        st = assoc.send_c_echo()
                        if st and getattr(st, "Status", None) == 0x0000:
                            msgs.append("反向自检成功：经 %s:%d 可达本机接收服务且 C-Echo 应答正常" % (self_ip, local_port))
                            self_ok = True
                        else:
                            self_err = "C-Echo 应答异常"
                    finally:
                        try:
                            assoc.release()
                        except Exception:
                            pass
                else:
                    self_err = "无法建立关联"
                    try:
                        assoc.release()
                    except Exception:
                        pass
            except Exception as e:
                self_err = str(e)
            if self_ok:
                break
        if not self_ok:
            msgs.append("反向自检失败（%s）：请检查 Windows 防火墙是否放行端口 %d 入站" % (self_err, local_port))

    # 4) 前置机反向探测：发不存在 UID 的 C-Move，观察前置机是否主动连入本机（仅提示）
    if (store_ready or (_store_started and local_port and _store_port == local_port)) and aet and local_aet:
        _probe_conn_event.clear()
        _probe_conn_addr = ""
        _probe_active = True
        try:
            ae = AE()
            ae.add_requested_context(StudyRootQueryRetrieveInformationModelMove)
            ae.acse_timeout = 10
            ae.dimse_timeout = 5
            ae.network_timeout = 10
            assoc = ae.associate(host, port, ae_title=aet)
            if assoc.is_established:
                try:
                    ds = Dataset()
                    ds.QueryRetrieveLevel = "STUDY"
                    ds.StudyInstanceUID = "1.2.3.999.999999.999999999"  # 不存在的 UID
                    for _st, _id in assoc.send_c_move(ds, local_aet, StudyRootQueryRetrieveInformationModelMove):
                        break  # 只需触发前置机动作，取首个响应即退出
                except Exception:
                    pass
                finally:
                    try:
                        assoc.release()
                    except Exception:
                        pass
                if _probe_conn_event.wait(8):
                    msgs.append("反向探测成功：前置机已主动连入本机接收服务（%s），反向通道正常" % _probe_conn_addr)
                else:
                    ips = "、".join(_local_ips()) or "未知"
                    msgs.append("反向探测：8 秒内前置机未连入本机接收服务。"
                                "若下载仍 0 文件，请核对前置机上「%s」注册的 IP/端口是否为本机（%s / %d）"
                                % (local_aet, ips, local_port))
            else:
                msgs.append("反向探测跳过：无法与 PACS 建立关联")
                try:
                    assoc.release()
                except Exception:
                    pass
        except Exception as e:
            msgs.append("反向探测异常：%s" % e)
        finally:
            _probe_active = False

    return (tcp_ok and local_ok and self_ok and (echo_ok if aet else True)), msgs


# ---------------------------------------------------------------------------
# 3) 读取 Excel 中的「影像号」(StudyInstanceUID) 列
# ---------------------------------------------------------------------------
def read_study_uids(excel_path, column, sheet_name=None, fallback_to_first_col=True):
    """
    读取 Excel 指定列，返回去重后的非空 StudyInstanceUID 列表。
    column 可以是：
        - 列名（如 "影像号"，按表头匹配）
        - 字母（如 "A"）
        - 数字（如 "1"，1-based）
    fallback_to_first_col=False 时，按列名匹配失败会直接抛 ValueError（不再静默读第一列）。
    """
    import openpyxl
    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    except PermissionError:
        raise PermissionError(
            "无法读取 Excel（文件被占用）：%s\n请先关闭 Excel 后再试。" % excel_path
        )
    except Exception as e:
        raise Exception("加载 Excel 失败：%s（路径：%s）" % (e, excel_path))
    try:
        ws = wb[sheet_name] if sheet_name else wb.active
    except KeyError:
        names = "、".join(wb.sheetnames)
        wb.close()
        raise ValueError("Sheet「%s」不存在，当前工作簿的 Sheet 为：%s" % (sheet_name, names))

    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    all_rows = list(rows)

    def col_letter_to_idx(letter):
        idx = 0
        for ch in letter.upper():
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        return idx - 1

    col_idx = None

    c = str(column).strip()
    if c.isdigit():
        col_idx = int(c) - 1
    elif c and c.isascii() and c.isalpha() and len(c) <= 3:
        # 仅纯英文字母（A-Z/a-z）才按“列字母”解析，避免中文列名被误判
        col_idx = col_letter_to_idx(c)
    elif header:
        # 按表头匹配列名
        for i, h in enumerate(header):
            if str(h).strip() == c:
                col_idx = i
                break
        if col_idx is None:
            if not fallback_to_first_col:
                headers = [str(h).strip() for h in header if h is not None and str(h).strip()]
                raise ValueError(
                    "在表头中找不到列「%s」。当前表头为：%s。请修改「关键列名」，或改用列字母（如 A）/序号（如 1）"
                    % (c, "、".join(headers[:15]) if headers else "(空)")
                )
            # 表头没匹配到，退回第一列，并把表头行也当作数据
            col_idx = 0
            all_rows = [header] + all_rows
    else:
        col_idx = 0

    if col_idx is None or col_idx < 0:
        col_idx = 0

    uids = []
    seen = set()
    for r in all_rows:
        val = r[col_idx] if (r and len(r) > col_idx) else None
        if val is None:
            continue
        if isinstance(val, float) and val.is_integer():
            # 长数字（影像号/病人ID）被 Excel 存成浮点数时，避免 str() 变成科学计数法
            s = str(int(val))
        else:
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
    seen_uids = set()  # 去重 set，避免 O(n²) 遍历
    last_status = None
    try:
        # send_c_find(dataset, query_model) - query_model 必须位置参数
        for status, identifier in assoc.send_c_find(ds, StudyRootQueryRetrieveInformationModelFind):
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
                if suid:
                    suid_str = str(suid)
                    if suid_str not in seen_uids:
                        seen_uids.add(suid_str)
                        uids.append(suid_str)
    except Exception as e:
        _ui_queue.put(("log", "      [C-Find 异常] %s" % e))
    _ui_queue.put(("log", "      [C-Find] 结束，最终状态 0x%04X %s，命中 %d 个 Study" % (
        last_status or 0, _status_text(last_status) if last_status is not None else "无响应", len(uids))))
    return uids


def pull_one_study(assoc, study_uid, local_aet):
    """拉取一个 Study。影像保存目录由 handle_store 决定（统一按原始 PatientID 命名）。

    本函数不再预建任何目录：文件数统计/校验按 Study 精确计数（handle_store 记录到
    _store_received / _store_counts），实际落盘目录从 _store_dirs 读取，避免产生以查询键
    命名的空文件夹，也避免同一病人多个检查共用目录时互相串数。

    计数分三类，避免互相干扰：
    - 接收数 _store_received：成功处理（落盘 + 按规则丢弃）——用于进度与成功判定；
    - 落盘数 _store_counts：真正写到磁盘的文件数——用于报告与目录展示；
    - 丢弃数 _store_discarded：命中过滤且关闭「保存命中的过滤影像」的文件数。
    - 失败数 _store_failed：解析/脱敏异常被拒绝落盘的文件数（不计入接收数，明细写失败清单）。
    若某检查全部命中过滤且被丢弃，落盘数为 0 但接收数 > 0，仍判定为成功（不误报失败）。
    失败数单独统计：个别文件失败不影响该检查判成功，但会在日志/报告里注明。

    C-Move 进行中（Pending 状态）会实时用「已完成 + 失败 + 警告」估算前置机已确定的子操作数，
    累计取最大值（避免某些前置机实现不规范导致分母倒退），并立即更新全局 _current_expected_images，
    让 handle_store 后续发的 image_progress 也能用上新分母；前置机不报子操作计数时退化为 C-Move
    完成后再用 接收数 + n_failed + n_warned 兜底作为分母。

    注意：NumberOfRemainingSuboperations 在该前置机上不可信（剩余=0 但已完成还在涨），
    所以不参与估算，仅在日志中展示。
    """
    global _current_expected_images

    def _n_received_now():
        """本批次该 Study 已成功处理的文件数（落盘 + 按规则丢弃）。"""
        with _store_lock:
            return _store_received.get(study_uid, 0)

    ds = Dataset()
    ds.QueryRetrieveLevel = "STUDY"
    ds.StudyInstanceUID = study_uid

    _ui_queue.put(("log", "      [C-Move] 请求拉取 StudyInstanceUID=%s -> 目标AE=%s (QueryRetrieveLevel=STUDY)" % (study_uid, local_aet)))

    final_code = 0x0000
    has_error = False
    n_failed = 0    # 子操作失败数（前置机推送失败的影像数）
    n_warned = 0    # 子操作警告数
    expected_total_max = 0  # 累计最大的"前置机应推总数"估算（取最大值避免分母倒退）
    try:
        # send_c_move(dataset, move_aet, query_model) - query_model 必须位置参数
        for status, _ in assoc.send_c_move(ds, local_aet, StudyRootQueryRetrieveInformationModelMove):
            if status:
                code = status.Status
                # 子操作计数（C-Move 的 Pending 阶段会带这些值，可用来评估进度）
                remaining = getattr(status, "NumberOfRemainingSuboperations", None)
                completed = getattr(status, "NumberOfCompletedSuboperations", None)
                failed = getattr(status, "NumberOfFailedSuboperations", None)
                warned = getattr(status, "NumberOfWarningSuboperations", None)
                if failed is not None:
                    n_failed = int(failed)
                if warned is not None:
                    n_warned = int(warned)
                # 估算前置机已确定要推的总数：只用 已完成 + 失败 + 警告
                # 注意：NumberOfRemainingSuboperations 在该前置机上不可信（剩余=0 但已完成还在涨，
                # 说明它只在前几个 Pending 报一次或干脆不更新；强行参与估算会导致分母小于分子）
                if completed is not None or failed is not None or warned is not None:
                    c_val = int(completed) if completed is not None else 0
                    f_val = int(failed) if failed is not None else 0
                    w_val = int(warned) if warned is not None else 0
                    est_from_subops = c_val + f_val + w_val
                    # 取「子操作估算值」与「当前已落盘数」的较大者，防止前置机的"已完成"上报有延迟
                    # 时显示「分子 > 分母」；同时永不倒退
                    n_recv_now = _n_received_now()
                    est = max(est_from_subops, n_recv_now, expected_total_max)
                    if est > expected_total_max:
                        expected_total_max = est
                        # 实时更新全局分母：handle_store 后续发的 image_progress 会立即用上新分母
                        _current_expected_images = est
                        # 主动推一次进度，让 UI 立即从「X / ?」变成「X / N」
                        _ui_queue.put(("image_progress", (_current_label, n_recv_now, est)))
                extra = ""
                if remaining is not None or completed is not None:
                    extra = " (剩余=%s 已完成=%s 失败=%s 警告=%s)" % (remaining, completed, failed, warned)
                msg = "      [C-Move] 响应状态 0x%04X %s%s" % (code, _status_text(code), extra)
                if code in _PENDING:
                    _diag(msg)  # Pending 过程日志默认隐藏，避免刷屏
                else:
                    _ui_queue.put(("log", msg))
                if code not in _PENDING:
                    final_code = code
                    # 0xB000=警告:子操作完成但有失败，属于“部分成功”，不当作硬错误；
                    # 由 _download_one 按本地实际落盘文件数记为成功，接受个别失败
                    if code not in (0x0000, 0xB000):
                        has_error = True
    except Exception as e:
        has_error = True
        _ui_queue.put(("log", "      [C-Move 异常] %s" % e))

    n_received = _n_received_now()
    # 部分医院前置机是"先返回 C-Move 成功、再异步推送影像"，
    # 状态成功但暂时 0 文件时，最多等 30 秒观察文件是否陆续落盘
    if not has_error and n_received == 0:
        _ui_queue.put(("log", "      [C-Move] 状态成功但暂无文件落盘，等待前置机异步推送（最多 30 秒）..."))
        for _ in range(60):
            time.sleep(0.5)
            if _stop_event.is_set():
                break
            n_received = _n_received_now()
            if n_received > 0:
                break
    # 落盘数 / 丢弃数：接收数 = 落盘数 + 丢弃数
    with _store_lock:
        real_dir = _store_dirs.get(study_uid) or _store_filtered_dirs.get(study_uid)
        n_saved = _store_counts.get(study_uid, 0)
        n_filtered = _store_filtered.get(study_uid, 0)
        n_discarded = _store_discarded.get(study_uid, 0)
        n_unparsed = _store_failed.get(study_uid, 0)
    _ui_queue.put(("log", "      [C-Move] 结束，最终状态 0x%04X %s，本地落盘 %d 个文件%s" % (
        final_code, _status_text(final_code), n_saved,
        ("（另有 %d 个命中过滤已丢弃）" % n_discarded) if n_discarded else "")))
    if n_unparsed:
        _ui_queue.put(("log", "      [异常] 本检查有 %d 个文件处理失败未落盘（已记入失败清单）" % n_unparsed))
    # n_total 兜底：取 (接收数 + 已失败 + 已警告)、(累计最大预期值)、(接收数) 三者最大
    # 防止前置机"剩余=0 但还在异步推"导致最终分母小于实际收到的影像数
    n_total = max(n_received + n_failed + n_warned, expected_total_max, n_received)
    if n_filtered:
        if n_discarded:
            _ui_queue.put(("log", "      [过滤] 本检查命中剂量报告/截屏规则 %d 个文件，其中 %d 个已按设置丢弃" % (n_filtered, n_discarded)))
        else:
            _ui_queue.put(("log", "      [过滤] 本检查命中剂量报告/截屏规则 %d 个文件，已存入 %s/" % (n_filtered, _FILTERED_SUBDIR)))
    return has_error, final_code, n_saved, n_discarded, real_dir, n_failed, n_warned, n_total, n_filtered, n_unparsed


# ---------------------------------------------------------------------------
# 4) 批量下载（后台线程执行）
# ---------------------------------------------------------------------------
class DownloadConfig:
    def __init__(self):
        self.pacs_host = ""
        self.pacs_port = 104
        self.pacs_aet = ""
        self.local_aet = "MYAET"
        self.local_port = 11112
        self.excel_path = ""
        self.sheet_name = ""
        self.column = "影像号"
        self.key_type = "patient_id"  # patient_id(病人ID) / study_uid(StudyInstanceUID)，默认按病人ID
        self.out_dir = ""
        self.rate_limit_kbps = 0      # 限速（KB/s），0 = 不限
        self.pause_every = 0          # 每下载 N 个检查后暂停（仅串行模式），0 = 不启用
        self.pause_seconds = 30       # 暂停秒数
        self.cmove_timeout = 300      # C-MOVE 超时（秒）
        self.cfind_timeout = 60       # C-FIND 超时（秒）
        self.desens_enabled = False   # 模块① 标签脱敏开关：落盘即为脱敏后影像
        self.filter_enabled = False   # 模块② 剂量报告/截屏过滤开关
        # 模块① 规则与全局选项（在「脱敏配置」中维护）
        self.desens_rules = [dict(r) for r in _DEFAULT_DESENS_RULES]
        self.desens_purge_private = True   # 清空所有私有标签
        self.desens_date_month = True      # 所有 DA/DT 统一压缩为「年-月-01」
        self.desens_time_clear = True      # 所有 TM 统一清空
        # 模块② 规则与命中处理方式
        self.filter_rules = [dict(r) for r in _DEFAULT_FILTER_RULES]
        self.filter_save_hit = True        # 命中过滤的影像是否保存（False = 丢弃不落盘）


def _register_assoc(assoc):
    with _assoc_lock:
        _active_assocs.append(assoc)


def _unregister_assoc(assoc):
    with _assoc_lock:
        try:
            _active_assocs.remove(assoc)
        except ValueError:
            pass


def _abort_active_assocs():
    """停止时强制中断所有活跃的 association，让卡在网络等待上的线程尽快退出。"""
    with _assoc_lock:
        for a in list(_active_assocs):
            try:
                a.abort()
            except Exception:
                pass


def _make_assoc(cfg, sop_class, label, timeout_attr, timeout_default):
    """建立 DICOM association（C-Find / C-Move 共用）。连接异常返回 None。"""
    _ui_queue.put(("log", "正在连接 PACS %s:%d（%s，AE=%s）..." % (cfg.pacs_host, cfg.pacs_port, label, cfg.pacs_aet)))
    ae = AE()
    ae.add_requested_context(sop_class)
    ae.acse_timeout = 15
    ae.dimse_timeout = int(getattr(cfg, timeout_attr, timeout_default) or timeout_default)
    ae.network_timeout = 30
    # 诊断：SCU 侧关联协商结果 + 底层 PDU 流捕获
    pdu_recv, pdu_sent = _make_pdu_handlers("SCU")
    handlers = [
        (evt.EVT_ACCEPTED, _on_accepted_scu),
        (evt.EVT_ABORTED, _on_aborted_scu),
        (evt.EVT_PDU_RECV, pdu_recv),
        (evt.EVT_PDU_SENT, pdu_sent),
    ]
    try:
        assoc = ae.associate(cfg.pacs_host, cfg.pacs_port, ae_title=cfg.pacs_aet, evt_handlers=handlers)
    except Exception as e:
        _ui_queue.put(("log", "[%s] 连接 PACS 异常：%s" % (label, e)))
        return None
    if assoc.is_established:
        _register_assoc(assoc)
    else:
        # 关联未建立（被 PACS 拒、网络异常等），确保释放对象内部 socket 资源
        try:
            assoc.release()
        except Exception:
            pass
    return assoc


def _on_accepted_scu(event):
    _log_assoc_negotiation(event.assoc, "SCU协商")


def _on_aborted_scu(event):
    try:
        addr = "%s:%s" % (event.assoc.remote_address[0], event.assoc.remote_address[1])
    except Exception:
        addr = "未知"
    _ui_queue.put(("log", "[SCU] 关联被中止（A-ABORT）：%s" % addr))


def _record_download(idx, key, uid, label, status, n_files, message=""):
    """追加一条下载结果记录（线程安全）。"""
    rec = {
        "idx": idx,
        "key": key,
        "study_uid": uid,
        "label": label,
        "status": status,
        "n_files": n_files,
        "message": message,
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    with _records_lock:
        _download_records.append(rec)


def _clear_records():
    with _records_lock:
        _download_records.clear()


def _snapshot_records():
    with _records_lock:
        return list(_download_records)


_REPORT_HEADER = ["序号", "查询键", "StudyInstanceUID", "结果", "文件数", "说明", "时间"]


def _write_report_csv(records, out_root):
    """把下载记录写入输出目录下的 CSV 报告，返回文件路径；失败返回 None。"""
    if not out_root:
        return None
    try:
        os.makedirs(out_root, exist_ok=True)
    except Exception:
        return None
    csv_path = os.path.join(out_root, "下载报告_%s.csv" % time.strftime("%Y%m%d_%H%M%S"))
    try:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(_REPORT_HEADER)
            for r in records:
                w.writerow([r["idx"], r["key"], r["study_uid"], r["status"],
                            r["n_files"], r["message"], r["time"]])
        return csv_path
    except Exception as e:
        _ui_queue.put(("log", "[报告] 写入下载报告失败：%s" % e))
        return None


_FAILED_CSV_HEADER = ["StudyInstanceUID", "SOPInstanceUID", "失败原因", "时间"]


def _snapshot_failed_files():
    """取「处理失败、未落盘」文件明细的快照（线程安全）。"""
    with _store_lock:
        return list(_store_failed_files)


def _write_failed_csv(items, out_root):
    """把「处理失败、未落盘」的文件清单写入输出目录下的 CSV，返回路径；无内容时返回 None。

    这些文件在接收时就因解析/脱敏异常被拒绝落盘（不会写出未处理的原始影像），
    单独留一份清单便于事后核对与补拉。
    """
    if not out_root or not items:
        return None
    try:
        os.makedirs(out_root, exist_ok=True)
    except Exception:
        return None
    csv_path = os.path.join(out_root, "未落盘文件清单_%s.csv" % time.strftime("%Y%m%d_%H%M%S"))
    try:
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(_FAILED_CSV_HEADER)
            for it in items:
                w.writerow([it["study_uid"], it["sop_uid"], it["reason"], it["time"]])
        return csv_path
    except Exception as e:
        _ui_queue.put(("log", "[报告] 写入未落盘文件清单失败：%s" % e))
        return None


def _audit_header():
    """审计表的表头：固定列 + 每个配置规则字段一列 + 其它变更 + 时间。"""
    return (_DESENS_AUDIT_FIXED_HEAD
            + [tag_display(t) for t in _desens_audit_fields]
            + _DESENS_AUDIT_TAIL_HEAD)


def _format_other_changes(items):
    """把「非规则字段」的变更（全局日期/时间、私有标签）汇总成一个单元格。"""
    parts = []
    for c in items:
        if c["tag"] == "(私有标签)":
            parts.append("私有标签：%s" % c["after"])
        else:
            parts.append("%s：%s → %s" % (c["keyword"] or c["tag"], c["before"], c["after"]))
    return "；".join(parts)


def _build_audit_row(file_name, study_uid, pid_after, changes, now):
    """把一次脱敏的字段变更整理成「一个影像一行」的审计记录。

    配置了规则的字段按列对齐（未变更则该格为空），其余变更汇总进「其它变更」列。
    """
    field_set = set(_desens_audit_fields)
    by_tag = {}
    others = []
    for c in changes or []:
        if c["tag"] in field_set and c["tag"] not in by_tag:
            by_tag[c["tag"]] = "%s → %s" % (c["before"], c["after"])
        else:
            others.append(c)
    row = [file_name, study_uid, "" if pid_after is None else str(pid_after)]
    row += [by_tag.get(t, "") for t in _desens_audit_fields]
    row += [_format_other_changes(others), now]
    return row


def _append_desens_audit(file_name, study_uid, patient_id_after, changes):
    """把一个影像的审计记录（一行）追加到缓冲，攒够阈值即刷进临时 CSV。"""
    row = _build_audit_row(file_name, study_uid, patient_id_after, changes,
                           time.strftime("%Y-%m-%d %H:%M:%S"))
    with _desens_audit_lock:
        _desens_audit_rows.append(row)
        if len(_desens_audit_rows) >= _DESENS_AUDIT_FLUSH:
            _flush_desens_audit()


def _flush_desens_audit():
    """把审计缓冲刷入临时 CSV（调用方需已持有 _desens_audit_lock）。"""
    global _desens_audit_rows, _desens_audit_header
    if not _desens_audit_rows or not _desens_audit_csv:
        _desens_audit_rows = []
        return
    try:
        with open(_desens_audit_csv, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if not _desens_audit_header:
                w.writerow(_audit_header())
                _desens_audit_header = True
            w.writerows(_desens_audit_rows)
    except Exception as e:
        _ui_queue.put(("log", "[审计] 写入脱敏审计明细失败：%s" % e))
    _desens_audit_rows = []


def _finalize_desens_audit(out_root):
    """把脱敏审计明细转成 Excel，返回 xlsx 路径；无内容时返回 None。

    几十万行的大表用 write_only 模式流式写入，避免一次性占满内存；
    转换失败时保留临时 CSV，确保审计数据不丢。
    """
    global _desens_audit_csv, _desens_audit_header
    with _desens_audit_lock:
        _flush_desens_audit()
        csv_path = _desens_audit_csv
        _desens_audit_csv = ""
        _desens_audit_header = False
    if not csv_path or not os.path.exists(csv_path):
        return None
    xlsx_path = os.path.join(out_root, "脱敏审计_%s.xlsx" % time.strftime("%Y%m%d_%H%M%S"))
    n_rows = 0
    _ui_queue.put(("log", "[审计] 正在生成脱敏审计 Excel（大表可能需要几秒）..."))
    try:
        from openpyxl import Workbook
        wb = Workbook(write_only=True)
        ws = wb.create_sheet("脱敏审计")
        with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
            for row in csv.reader(f):
                ws.append(row)
                n_rows += 1
        wb.save(xlsx_path)
        wb.close()
    except Exception as e:
        _ui_queue.put(("log", "[审计] 生成 Excel 失败，明细保留为 CSV：%s（%s）" % (csv_path, e)))
        return None
    try:
        os.remove(csv_path)   # 转换成功即删除临时 CSV，避免同一份敏感明细留两份
    except Exception:
        pass
    _ui_queue.put(("log", "[审计] 脱敏审计 Excel 已生成：%s（%d 行字段变更）" % (xlsx_path, max(0, n_rows - 1))))
    return xlsx_path


def _download_one(cfg, key, label, uid, idx, total):
    """下载单个 Study（不重试）。返回 (idx, status, 落盘文件数, message)。"""
    global _current_label, _current_study_uid, _current_expected_images
    _current_label = label
    _current_study_uid = str(uid)  # handle_store 据此只推当前检查的进度
    _current_expected_images = 0   # 拉取前未知，由 pull_one_study 返回值/异常时回填
    # 清空本批次计数/目录表：上一个检查的残留推送不计入当前检查
    with _store_lock:
        _store_received.clear()
        _store_counts.clear()
        _store_discarded.clear()
        _store_dirs.clear()
        _store_filtered_dirs.clear()
        _store_filtered.clear()
        _store_failed.clear()

    # 停止优先：停止后剩余任务统一记为“停止”
    if _stop_event.is_set():
        _ui_queue.put(("log", "[%d/%d] [跳过] %s（手动停止）" % (idx, total, label)))
        _record_download(idx, key, uid, label, "停止", 0)
        return idx, "stopped", 0, "已停止"

    def _success(n, note=""):
        """成功统一出口：记录 + 返回。"""
        _record_download(idx, key, uid, label, "成功", n, note)
        return idx, "success", n, note

    # 全量拉取该 Study（前置机不支持 IMAGE/SERIES 级查询，无法精准补拉，接受个别失败）
    assoc = _make_assoc(cfg, StudyRootQueryRetrieveInformationModelMove, "C-Move", "cmove_timeout", 300)
    if assoc and assoc.is_established:
        try:
            has_error, code, n_saved, n_discarded, subdir, n_failed, n_warned, n_total, n_filtered, n_unparsed = pull_one_study(assoc, uid, cfg.local_aet)
            n_recv = n_saved + n_discarded
            _current_expected_images = n_total
            _ui_queue.put(("image_progress", (label, n_recv, n_total)))
        finally:
            _unregister_assoc(assoc)
            try:
                assoc.release()
            except Exception:
                pass
        # subdir 为实际落盘目录（原始 PatientID 命名）；无文件落盘时为 None
        dir_hint = subdir or "（无文件落盘）"
        if n_filtered:
            if n_discarded:
                filter_hint = "，命中过滤 %d 个（丢弃 %d 个）" % (n_filtered, n_discarded)
            else:
                filter_hint = "，其中 %d 个命中过滤存入 %s/" % (n_filtered, _FILTERED_SUBDIR)
        else:
            filter_hint = ""
        # 处理失败（解析/脱敏异常被拒绝落盘）单独提示；少数文件失败不影响该检查整体判成功
        fail_hint = "，%d 个文件处理失败未落盘" % n_unparsed if n_unparsed else ""
        hint = filter_hint + fail_hint
        # 成功判定用「接收数」：某检查全部命中过滤且被丢弃时落盘为 0，但仍算成功
        if not has_error and n_recv > 0:
            # 状态码 0x0000 且无子操作失败/警告：完全成功
            if n_failed == 0 and n_warned == 0 and code == 0x0000:
                _ui_queue.put(("log", "[%d/%d] [完成] %s -> 落盘 %d 个文件%s %s" % (idx, total, label, n_saved, hint, dir_hint)))
                return _success(n_saved, hint.lstrip("，"))
            # 子操作有失败/警告（0xB000）：按本地已落盘文件数视为成功，接受个别失败
            _ui_queue.put(("log", "[%d/%d] [完成] %s -> 落盘 %d 个文件%s（子操作 %d 失败/%d 警告，接受） %s" % (idx, total, label, n_saved, hint, n_failed, n_warned, dir_hint)))
            return _success(n_saved, "子操作 %d 失败/%d 警告%s" % (n_failed, n_warned, hint))
        elif has_error:
            fail_msg = "状态码 0x%04X（本地文件 %d 个）" % (code, n_saved)
        elif n_unparsed:
            # 有文件推来了但全部处理失败：直接说明原因，避免误导为「前置机没推」
            fail_msg = "本检查 %d 个文件全部处理失败、未落盘（解析或脱敏异常，详见失败清单 CSV）" % n_unparsed
        else:
            # 状态码成功但 0 文件：前置机没有把影像推到本机（注册信息不符/防火墙/异步未推）
            ips = "、".join(_local_ips()) or "未知"
            fail_msg = ("C-Move 状态成功但等待 30 秒后仍 0 个文件落盘。"
                        "请核对：1) 医院前置机上「%s」注册的 IP 是否为本机当前 IP（%s）；"
                        "2) 注册端口是否为 %d；3) Windows 防火墙是否放行 %d 入站。"
                        % (cfg.local_aet, ips, cfg.local_port, cfg.local_port))
    else:
        fail_msg = "连接 PACS 失败"
    _ui_queue.put(("log", "[%d/%d] [失败] %s：%s" % (idx, total, label, fail_msg)))
    _record_download(idx, key, uid, label, "失败", n_saved, fail_msg)
    return idx, "failed", n_saved, fail_msg


def batch_download(cfg):
    """批量下载主流程（在线程中运行）。任何退出路径（成功/报错/异常）都会复位 GUI 状态并写报告。"""
    try:
        _batch_download_inner(cfg)
    except Exception as e:
        # 提取出错位置、错误类型与原因：弹窗给简洁信息，完整堆栈写日志便于排查
        exc_type = type(e).__name__
        reason = str(e) or "(无具体错误信息)"
        tb = traceback.extract_tb(sys.exc_info()[2])
        loc = "未知位置"
        if tb:
            loc = "%s 第 %d 行" % (os.path.basename(tb[-1].filename), tb[-1].lineno)
        _ui_queue.put(("log", "[错误] 出错位置：%s\n%s" % (loc, traceback.format_exc())))
        _ui_queue.put(("error", "出错位置：%s\n错误类型：%s\n错误原因：%s" % (loc, exc_type, reason)))
    finally:
        _abort_active_assocs()
        records = _snapshot_records()
        csv_path = _write_report_csv(records, cfg.out_dir) if records else None
        # 处理失败（未落盘）的文件单独出一份清单，便于核对与补拉
        failed_items = _snapshot_failed_files()
        failed_csv = _write_failed_csv(failed_items, cfg.out_dir)
        if failed_items:
            _ui_queue.put(("log", "[报告] 本次共 %d 个文件处理失败未落盘，清单：%s" % (len(failed_items), failed_csv)))
        # 脱敏审计明细 → Excel（仅开启模块①时才有内容；无内容则不生成文件）
        _finalize_desens_audit(cfg.out_dir)
        _ui_queue.put(("report", (records, csv_path)))
        _ui_queue.put(("reset", None))


def _batch_download_inner(cfg):
    global OUTPUT_ROOT, _rate_limit_kbps
    global _rate_last_time, _rate_tokens, _download_start_time, _throttle_cap_logged
    global _desens_enabled, _desens_rules, _desens_opts
    global _filter_enabled, _filter_rules, _filter_save_hit
    global _desens_audit_rows, _desens_audit_csv, _desens_audit_header
    # 关键：重置停止标志，避免上一次“停止”被传染到本次下载
    _stop_event.clear()
    _clear_records()
    with _store_lock:
        _store_failed.clear()
        _store_failed_files.clear()   # 失败明细按本次下载重新累积
    # 脱敏审计按本次下载重新累积（临时 CSV 与表头状态一并复位）
    with _desens_audit_lock:
        _desens_audit_rows = []
        _desens_audit_header = False
        _desens_audit_csv = ""

    OUTPUT_ROOT = cfg.out_dir
    _download_start_time = time.time()  # 记录本次下载开始时间
    os.makedirs(OUTPUT_ROOT, exist_ok=True)
    _rate_limit_kbps = int(getattr(cfg, "rate_limit_kbps", 0) or 0)
    _rate_last_time = 0.0
    _rate_tokens = 0.0
    _throttle_cap_logged = False  # 每次下载重置，保证限速封顶提示能在每次下载时生效
    if _rate_limit_kbps > 0:
        _ui_queue.put(("log", "已启用限速：%d KB/s" % _rate_limit_kbps))

    # 脱敏/过滤开关（下载开始时一次性生效，中途修改配置不影响本次任务）
    _desens_enabled = bool(getattr(cfg, "desens_enabled", False))
    _filter_enabled = bool(getattr(cfg, "filter_enabled", False))
    _desens_rules = prepare_desens_rules(getattr(cfg, "desens_rules", None))
    _desens_opts = {
        "purge_private": bool(getattr(cfg, "desens_purge_private", True)),
        "date_month": bool(getattr(cfg, "desens_date_month", True)),
        "time_clear": bool(getattr(cfg, "desens_time_clear", True)),
    }
    _filter_rules = prepare_filter_rules(getattr(cfg, "filter_rules", None))
    _filter_save_hit = bool(getattr(cfg, "filter_save_hit", True))
    if _desens_enabled:
        _ui_queue.put(("log", "已启用模块①：标签脱敏（%d 条规则；清空私有标签=%s，日期压缩到月=%s，时间清空=%s）" % (
            len(_desens_rules), _desens_opts["purge_private"], _desens_opts["date_month"], _desens_opts["time_clear"])))
        for r in _desens_rules:
            _diag("      [脱敏规则] %s -> %s %s%s" % (
                tag_display(tag_to_str(r["tag"])), _DESENS_METHOD_LABEL.get(r["method"], r["method"]),
                ("参数=%s" % r["param"]) if r["param"] else "",
                "（含序列内）" if r["recursive"] else ""))
        if not _desens_rules:
            _ui_queue.put(("log", "[脱敏] 未配置任何标签规则，仅按全局选项处理"))
        # 脱敏审计：开启模块①即自动生成（明细先落临时 CSV，下载结束转成 Excel）
        with _desens_audit_lock:
            _desens_audit_csv = os.path.join(cfg.out_dir, "_脱敏审计明细.tmp.csv")
        _ui_queue.put(("log", "[审计] 已开启脱敏审计：本次将逐字段记录脱敏前后对比，结束时生成 Excel"))
    if _filter_enabled:
        if _filter_save_hit:
            _ui_queue.put(("log", "已启用模块②：剂量报告/截屏过滤（%d 条规则；命中项存入 %s/ 子目录）" % (
                len(_filter_rules), _FILTERED_SUBDIR)))
        else:
            _ui_queue.put(("log", "已启用模块②：剂量报告/截屏过滤（%d 条规则；命中项将直接丢弃，不落盘）" % len(_filter_rules)))
        for r in _filter_rules:
            _diag("      [过滤规则] %s %s %s" % (
                tag_display(tag_to_str(r["tag"])), _FILTER_MODE_LABEL.get(r["mode"], r["mode"]), r["values"]))
        if not _filter_rules:
            _ui_queue.put(("log", "[过滤] 未配置任何过滤规则，本次不会过滤任何影像"))

    # 1) 读 Excel（得到 keys：StudyInstanceUID 或 patientId）
    _ui_queue.put(("log", "正在读取 Excel：%s" % cfg.excel_path))
    try:
        keys = read_study_uids(cfg.excel_path, cfg.column, cfg.sheet_name or None, fallback_to_first_col=False)
    except Exception as e:
        _ui_queue.put(("error", "读取 Excel 失败：%s" % e))
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
        assoc_find = _make_assoc(cfg, StudyRootQueryRetrieveInformationModelFind, "C-Find", "cfind_timeout", 60)
        if not assoc_find or not assoc_find.is_established:
            _ui_queue.put(("error", "连接 PACS 失败（C-Find），请检查 IP/端口/AE 及网络是否可达"))
            return
        try:
            for pi, pid in enumerate(keys, 1):
                if _stop_event.is_set():
                    break
                uids = find_studies_by_patient(assoc_find, pid)
                if not uids:
                    _ui_queue.put(("log", "  [%d/%d] patientId=%s 未查到任何检查" % (pi, n_keys, pid)))
                    continue
                _ui_queue.put(("log", "  [%d/%d] patientId=%s -> 找到 %d 个检查" % (pi, n_keys, pid, len(uids))))
                for u in uids:
                    tasks.append((str(pid), "%s/%s" % (pid, u), u))
        finally:
            _unregister_assoc(assoc_find)
            try:
                assoc_find.release()
            except Exception:
                pass
    else:
        for u in keys:
            tasks.append((str(u), str(u), u))

    total = len(tasks)
    _ui_queue.put(("log", "共需下载 %d 个检查（Study）" % total))
    if total == 0:
        _ui_queue.put(("error", "没有可下载的检查"))
        return

    # 4) 串行下载（稳态，支持间隙暂停、停止中断）
    cmove_timeout = int(getattr(cfg, "cmove_timeout", 300) or 300)
    _ui_queue.put(("log", "开始下载：串行（稳态），C-Move 超时=%d 秒" % cmove_timeout))

    ok = 0
    fail = 0
    pause_every = int(getattr(cfg, "pause_every", 0) or 0)
    pause_seconds = int(getattr(cfg, "pause_seconds", 0) or 0)

    for i, (key, label, uid) in enumerate(tasks, 1):
        if _stop_event.is_set():
            _ui_queue.put(("log", "已手动停止，剩余 %d 个未处理" % (total - i + 1)))
            break
        _ui_queue.put(("progress", (i, total)))
        _ui_queue.put(("log", "[%d/%d] 拉取 %s" % (i, total, label)))
        _idx, status, _, _ = _download_one(cfg, key, label, uid, i, total)
        if status == "success":
            ok += 1
        elif status == "stopped":
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

    # 处理失败（未落盘）的影像总数：不影响检查成功判定，但要在汇总里明确告知
    n_failed_files = len(_snapshot_failed_files())
    fail_note = "（另有 %d 个影像处理失败未落盘，详见失败清单 CSV）" % n_failed_files if n_failed_files else ""
    if _stop_event.is_set():
        _ui_queue.put(("done", "已停止：成功 %d / 失败 %d / 共 %d%s（剩余任务未处理）" % (ok, fail, total, fail_note)))
        _ui_queue.put(("log", "已停止：成功 %d / 失败 %d / 共 %d%s" % (ok, fail, total, fail_note)))
    else:
        _ui_queue.put(("done", "下载完成：成功 %d / 失败 %d / 共 %d%s" % (ok, fail, total, fail_note)))


# ---------------------------------------------------------------------------
# 5) GUI 界面
# ---------------------------------------------------------------------------
def _center_on_parent(win, parent):
    """把对话框大致居中到主窗口上方（不同平台都能落在屏幕内）。"""
    try:
        win.update_idletasks()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = win.winfo_width(), win.winfo_height()
        x = px + max(0, (pw - w) // 2)
        y = py + max(0, (ph - h) // 3)
        win.geometry("+%d+%d" % (x, y))
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("DICOM 影像批量下载工具")
        root.geometry("880x900")
        root.minsize(760, 620)

        self.cfg = DownloadConfig()
        self.thread = None
        self.records = []       # 最近一次下载的结果记录
        self.report_csv = None  # 最近一次下载报告 CSV 路径
        # 规则表（界面为唯一真源；新增/编辑/删除后立即重建表格）
        self.desens_rules = [dict(r) for r in _DEFAULT_DESENS_RULES]
        self.filter_rules = [dict(r) for r in _DEFAULT_FILTER_RULES]

        self._build_widgets()
        self._rebuild_desens_table()
        self._rebuild_filter_table()
        self._load_config()
        self._poll_queue()

    # ----- 布局 -----
    def _build_widgets(self):
        pad = dict(padx=6, pady=3)

        # 配置区（上下罗列，可滚动）：高度随内容动态变化，优先完整显示配置
        host = ttk.Frame(self.root)
        host.pack(fill="x", expand=False)
        self._cfg_host = host
        self._cfg_canvas = tk.Canvas(host, highlightthickness=0, height=200)
        vsb = ttk.Scrollbar(host, orient="vertical", command=self._cfg_canvas.yview)
        self._cfg_canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._cfg_canvas.pack(side="left", fill="both", expand=True)
        holder = ttk.Frame(self._cfg_canvas)
        self._cfg_holder = holder
        self._cfg_window = self._cfg_canvas.create_window((0, 0), window=holder, anchor="nw")
        holder.bind("<Configure>", self._on_cfg_content_change)
        self._cfg_canvas.bind("<Configure>",
                              lambda e: self._cfg_canvas.itemconfigure(self._cfg_window, width=e.width))
        # 鼠标滚轮仅在指针位于配置区时生效（避免抢走日志框的滚动）
        self._cfg_canvas.bind("<Enter>", self._bind_wheel)
        self._cfg_canvas.bind("<Leave>", self._unbind_wheel)
        # 窗口尺寸变化时重新分配「配置区 / 日志」的高度
        self.root.bind("<Configure>", self._sync_cfg_height)

        # PACS 配置
        f1 = ttk.LabelFrame(holder, text="PACS 前置机配置（由医院实施工程师提供）")
        f1.pack(fill="x", padx=10, pady=(10, 4))
        self._entry_pair(f1, "PACS IP", "pacs_host", default="")
        self._entry_pair(f1, "PACS 端口", "pacs_port", default="104")
        self._entry_pair(f1, "PACS AE Title", "pacs_aet", default="")

        # 本机配置
        f2 = ttk.LabelFrame(holder, text="本机接收节点配置（需注册到 PACS 前置机）")
        f2.pack(fill="x", padx=10, pady=4)
        self._entry_pair(f2, "本机 AE Title", "local_aet", default="MYAET")
        self._entry_pair(f2, "本机接收端口", "local_port", default="11112")

        # 下载配置
        f3 = ttk.LabelFrame(holder, text="下载配置")
        f3.pack(fill="x", padx=10, pady=4)

        row0 = ttk.Frame(f3); row0.pack(fill="x", **pad)
        ttk.Label(row0, text="Excel 文件:").pack(side="left")
        self.var_excel = tk.StringVar()
        ttk.Entry(row0, textvariable=self.var_excel, width=50).pack(side="left", padx=4)
        ttk.Button(row0, text="浏览...", command=self._browse_excel).pack(side="left")

        row_kt = ttk.Frame(f3); row_kt.pack(fill="x", **pad)
        ttk.Label(row_kt, text="查询键类型:").pack(side="left")
        self.var_key_type = tk.StringVar(value="病人ID(patientId)")
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
        ttk.Label(row3, text="(0=不限速, 如1024=1MB/s; 过低会触发封顶保护)").pack(side="left")

        ttk.Label(row3, text="  每下载").pack(side="left", padx=(16, 0))
        self.var_pause_every = tk.StringVar(value="0")
        ttk.Entry(row3, textvariable=self.var_pause_every, width=6).pack(side="left", padx=4)
        ttk.Label(row3, text="个检查后暂停").pack(side="left")
        self.var_pause_seconds = tk.StringVar(value="30")
        ttk.Entry(row3, textvariable=self.var_pause_seconds, width=6).pack(side="left", padx=4)
        ttk.Label(row3, text="秒(0=不暂停)").pack(side="left")

        row4 = ttk.Frame(f3); row4.pack(fill="x", **pad)
        ttk.Label(row4, text="C-Move超时(秒):").pack(side="left")
        self.var_cmove_timeout = tk.StringVar(value="300")
        ttk.Entry(row4, textvariable=self.var_cmove_timeout, width=8).pack(side="left", padx=4)
        ttk.Label(row4, text="C-Find超时(秒):").pack(side="left", padx=(12, 0))
        self.var_cfind_timeout = tk.StringVar(value="60")
        ttk.Entry(row4, textvariable=self.var_cfind_timeout, width=6).pack(side="left", padx=4)
        self.var_diag_log = tk.BooleanVar(value=False)
        ttk.Checkbutton(row4, text="诊断日志", variable=self.var_diag_log,
                        command=self._sync_diag).pack(side="left", padx=(12, 0))

        # 脱敏配置（独立模块，与下载配置分开）
        f4 = ttk.LabelFrame(holder, text="脱敏配置（独立模块，两个开关相互独立；开启后落盘即为处理后的影像）")
        f4.pack(fill="x", padx=10, pady=(4, 10))
        self._build_filter_section(f4)   # 过滤在上（执行顺序也是先过滤后脱敏）
        self._build_desens_section(f4)

        # ---- 底部固定区（按钮 / 进度 / 日志，不随配置区滚动）----
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
        self.btn_history = ttk.Button(fbtn, text="下载记录", command=self._show_history, state="disabled")
        self.btn_history.pack(side="left", padx=4)

        # 进度条
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=4)
        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.var_status).pack(anchor="w", padx=12)
        # 第二行：当前检查的影像张数 + 总张数 + 已耗时
        info_row = ttk.Frame(self.root)
        info_row.pack(fill="x", padx=12, pady=(0, 4))
        self.var_image_count = tk.StringVar(value="影像数: 0 / ?")
        ttk.Label(info_row, textvariable=self.var_image_count).pack(side="left")
        ttk.Label(info_row, text="    ").pack(side="left")
        self.var_elapsed = tk.StringVar(value="已耗时: 00:00:00")
        ttk.Label(info_row, textvariable=self.var_elapsed).pack(side="left")

        # 日志：占配置区之外的剩余空间（配置区越高，日志越小，但不小于 3 行）
        self._log_min_height = 60
        self.log = scrolledtext.ScrolledText(self.root, height=3, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.root.update_idletasks()
        self._log_min_height = self.log.winfo_reqheight()
        self._sync_cfg_height()

    # ----- 配置区 / 日志 的动态高度分配 -----
    def _on_cfg_content_change(self, _event=None):
        """配置区内容变化（勾选展开/收起、增删规则等）时更新滚动区域并重新分配高度。"""
        self._cfg_canvas.configure(scrollregion=self._cfg_canvas.bbox("all"))
        self._sync_cfg_height()

    @staticmethod
    def _pack_pady(widget):
        """返回控件的 pack 上下 pady 之和。"""
        try:
            p = widget.pack_info().get("pady", 0)
        except Exception:
            return 0
        if isinstance(p, (tuple, list)):
            try:
                return int(p[0]) + int(p[-1])
            except Exception:
                return 0
        try:
            return int(p) * 2
        except Exception:
            return 0

    def _sync_cfg_height(self, _event=None):
        """动态分配高度：以完整显示配置区为主，剩余空间给日志。

        - 配置内容短 → 画布收缩到内容高度，日志占据剩余空间（变高）；
        - 配置内容长 → 画布尽量占满可用高度，日志收缩到最小（3 行）；
        - 仍放不下 → 配置区内部滚动（右侧滚动条可用）。
        """
        holder = getattr(self, "_cfg_holder", None)
        host = getattr(self, "_cfg_host", None)
        log = getattr(self, "log", None)
        if holder is None or host is None or log is None:
            return
        try:
            content_h = holder.winfo_reqheight()
        except Exception:
            return
        # 除配置区、日志外的固定区域（按钮/进度/状态等）高度，含上下 pady
        reserved = 0
        try:
            for w in self.root.pack_slaves():
                if w is host:
                    continue
                if w is log:
                    reserved += self._pack_pady(w)   # 日志只计上下留白，高度另用 _log_min_height
                else:
                    reserved += w.winfo_reqheight() + self._pack_pady(w)
        except Exception:
            pass
        avail = self.root.winfo_height() - reserved - getattr(self, "_log_min_height", 60)
        avail = max(120, avail)
        target = content_h if content_h < avail else avail
        try:
            self._cfg_canvas.configure(height=int(target))
        except Exception:
            pass

    # ----- 脱敏配置区（模块①）-----
    def _build_desens_section(self, parent):
        f = ttk.LabelFrame(parent, text="模块① 标签脱敏（按下列规则逐条处理；未配置的标签不受影响）")
        f.pack(fill="x", padx=6, pady=(0, 4))

        top = ttk.Frame(f); top.pack(fill="x", padx=6, pady=3)
        self.var_desens = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="启用标签脱敏", variable=self.var_desens,
                        command=self._toggle_desens_detail).pack(side="left")
        ttk.Label(top, text="（勾选后显示详细规则配置）", foreground="#666").pack(side="left", padx=6)

        # 详细配置：仅勾选「启用标签脱敏」时显示
        detail = ttk.Frame(f)
        self.desens_detail = detail

        opt = ttk.Frame(detail); opt.pack(fill="x", padx=6, pady=3)
        ttk.Label(opt, text="全局选项:").pack(side="left")
        self.var_purge_private = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="清空私有标签", variable=self.var_purge_private).pack(side="left", padx=(6, 0))
        self.var_date_month = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="所有日期压缩到「年-月-01」", variable=self.var_date_month).pack(side="left", padx=(10, 0))
        self.var_time_clear = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="所有时间清空", variable=self.var_time_clear).pack(side="left", padx=(10, 0))

        tbl = ttk.Frame(detail); tbl.pack(fill="x", padx=6, pady=(0, 3))
        cols = ("tag", "method", "param", "scope")
        self.tv_desens = ttk.Treeview(tbl, columns=cols, show="headings", height=6)
        for cid, text, w in (("tag", "DICOM 标签", 280), ("method", "脱敏方法", 180),
                             ("param", "参数", 150), ("scope", "作用范围", 110)):
            self.tv_desens.heading(cid, text=text)
            self.tv_desens.column(cid, width=w, anchor="w", stretch=(cid == "param"))
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tv_desens.yview)
        self.tv_desens.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tv_desens.pack(side="left", fill="x", expand=True)
        self.tv_desens.bind("<Double-1>", lambda e: self._edit_desens_rule())

        bar = ttk.Frame(detail); bar.pack(fill="x", padx=6, pady=(0, 6))
        self.btn_desens_add = ttk.Button(bar, text="添加规则", command=self._add_desens_rule)
        self.btn_desens_add.pack(side="left")
        self.btn_desens_edit = ttk.Button(bar, text="编辑", command=self._edit_desens_rule)
        self.btn_desens_edit.pack(side="left", padx=4)
        self.btn_desens_del = ttk.Button(bar, text="删除", command=self._del_desens_rule)
        self.btn_desens_del.pack(side="left")
        self.btn_desens_reset = ttk.Button(bar, text="恢复默认规则", command=self._reset_desens_rules)
        self.btn_desens_reset.pack(side="left", padx=4)
        ttk.Label(bar, text="（双击行可编辑；同一标签只生效第一条规则）",
                  foreground="#666").pack(side="left", padx=8)
        self._desens_buttons = [self.btn_desens_add, self.btn_desens_edit,
                                self.btn_desens_del, self.btn_desens_reset]
        self._toggle_desens_detail()

    def _toggle_desens_detail(self):
        """勾选「启用标签脱敏」才展开详细规则配置。"""
        if self.var_desens.get():
            self.desens_detail.pack(fill="x")
        else:
            self.desens_detail.pack_forget()

    # ----- 脱敏配置区（模块②）-----
    def _build_filter_section(self, parent):
        f = ttk.LabelFrame(parent, text="模块② 剂量报告/截屏过滤（任一规则命中即算命中）")
        f.pack(fill="x", padx=6, pady=(0, 6))

        top = ttk.Frame(f); top.pack(fill="x", padx=6, pady=3)
        self.var_filter = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="启用剂量报告/截屏过滤", variable=self.var_filter,
                        command=self._toggle_filter_detail).pack(side="left")
        ttk.Label(top, text="（勾选后显示详细规则配置）", foreground="#666").pack(side="left", padx=6)

        # 详细配置：仅勾选「启用剂量报告/截屏过滤」时显示
        detail = ttk.Frame(f)
        self.filter_detail = detail

        opt = ttk.Frame(detail); opt.pack(fill="x", padx=6, pady=3)
        self.var_filter_save_hit = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="保存命中的过滤影像", variable=self.var_filter_save_hit).pack(side="left")
        ttk.Label(opt, text="（勾选：存入 输出目录/%s/ 子目录；不勾选：直接丢弃不落盘）" % _FILTERED_SUBDIR,
                  foreground="#666").pack(side="left", padx=6)

        tbl = ttk.Frame(detail); tbl.pack(fill="x", padx=6, pady=(0, 3))
        cols = ("tag", "mode", "values")
        self.tv_filter = ttk.Treeview(tbl, columns=cols, show="headings", height=4)
        for cid, text, w in (("tag", "DICOM 标签", 280), ("mode", "匹配模式", 110),
                             ("values", "命中值（多个值任一命中即可）", 330)):
            self.tv_filter.heading(cid, text=text)
            self.tv_filter.column(cid, width=w, anchor="w", stretch=(cid == "values"))
        sb = ttk.Scrollbar(tbl, orient="vertical", command=self.tv_filter.yview)
        self.tv_filter.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tv_filter.pack(side="left", fill="x", expand=True)
        self.tv_filter.bind("<Double-1>", lambda e: self._edit_filter_rule())

        bar = ttk.Frame(detail); bar.pack(fill="x", padx=6, pady=(0, 6))
        self.btn_filter_add = ttk.Button(bar, text="添加规则", command=self._add_filter_rule)
        self.btn_filter_add.pack(side="left")
        self.btn_filter_edit = ttk.Button(bar, text="编辑", command=self._edit_filter_rule)
        self.btn_filter_edit.pack(side="left", padx=4)
        self.btn_filter_del = ttk.Button(bar, text="删除", command=self._del_filter_rule)
        self.btn_filter_del.pack(side="left")
        self.btn_filter_reset = ttk.Button(bar, text="恢复默认规则", command=self._reset_filter_rules)
        self.btn_filter_reset.pack(side="left", padx=4)
        ttk.Label(bar, text="（双击行可编辑）", foreground="#666").pack(side="left", padx=8)
        self._filter_buttons = [self.btn_filter_add, self.btn_filter_edit,
                                self.btn_filter_del, self.btn_filter_reset]
        self._toggle_filter_detail()

    def _toggle_filter_detail(self):
        """勾选「启用剂量报告/截屏过滤」才展开详细规则配置。"""
        if self.var_filter.get():
            self.filter_detail.pack(fill="x")
        else:
            self.filter_detail.pack_forget()

    def _entry_pair(self, parent, label, attr, default=""):
        r = ttk.Frame(parent); r.pack(fill="x", padx=6, pady=3)
        ttk.Label(r, text=label + ":", width=16).pack(side="left")
        var = tk.StringVar(value=default)
        ttk.Entry(r, textvariable=var, width=40).pack(side="left", padx=4)
        setattr(self, "var_" + attr, var)

    def _iter_editable_widgets(self):
        """遍历主窗口下所有 Entry/Combobox，用于下载期间统一禁用/恢复。
        （原实现用 nametowidget(StringVar._name) 是无效的：StringVar 的 _name 是
        Tcl 变量名而非 widget 路径，会静默失败，导致下载时输入框实际未被禁用。）"""

        def walk(w):
            for ch in w.winfo_children():
                if isinstance(ch, (ttk.Entry, ttk.Combobox)):
                    yield ch
                yield from walk(ch)

        yield from walk(self.root)

    # ----- 配置区滚动 -----
    def _bind_wheel(self, _event=None):
        self._cfg_canvas.bind_all("<MouseWheel>", self._on_wheel)
        self._cfg_canvas.bind_all("<Button-4>", self._on_wheel)
        self._cfg_canvas.bind_all("<Button-5>", self._on_wheel)

    def _unbind_wheel(self, _event=None):
        self._cfg_canvas.unbind_all("<MouseWheel>")
        self._cfg_canvas.unbind_all("<Button-4>")
        self._cfg_canvas.unbind_all("<Button-5>")

    def _on_wheel(self, event):
        """滚轮滚动配置区：兼容 Windows/macOS（delta）与 X11（Button-4/5）。"""
        try:
            if getattr(event, "num", None) == 4:
                step = -1
            elif getattr(event, "num", None) == 5:
                step = 1
            else:
                d = getattr(event, "delta", 0)
                if abs(d) >= 120:
                    step = int(-d / 120) or (-1 if d > 0 else 1)
                else:
                    step = -1 if d > 0 else 1
            self._cfg_canvas.yview_scroll(step, "units")
        except Exception:
            pass

    # ----- 规则表：渲染 -----
    def _rebuild_desens_table(self):
        tv = self.tv_desens
        tv.delete(*tv.get_children())
        for i, r in enumerate(self.desens_rules):
            tv.insert("", "end", iid=str(i), values=(
                tag_display(r.get("tag", "")),
                _DESENS_METHOD_LABEL.get(r.get("method"), r.get("method", "")),
                r.get("param", ""),
                "顶层 + 序列内" if r.get("recursive") else "仅顶层",
            ))

    def _rebuild_filter_table(self):
        tv = self.tv_filter
        tv.delete(*tv.get_children())
        for i, r in enumerate(self.filter_rules):
            vals = r.get("values") or []
            if isinstance(vals, str):
                vals = [vals]
            tv.insert("", "end", iid=str(i), values=(
                tag_display(r.get("tag", "")),
                _FILTER_MODE_LABEL.get(r.get("mode"), r.get("mode", "")),
                " ； ".join(str(v) for v in vals),
            ))

    @staticmethod
    def _selected_index(tv):
        sel = tv.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    @staticmethod
    def _select_row(tv, idx):
        try:
            iid = str(idx)
            tv.selection_set(iid)
            tv.see(iid)
        except Exception:
            pass

    # ----- 规则表：增删改（模块①）-----
    def _add_desens_rule(self):
        r = self._desens_dialog()
        if not r:
            return
        # 同一标签只保留一条：新加的覆盖旧的
        self.desens_rules = [x for x in self.desens_rules if x.get("tag") != r["tag"]]
        self.desens_rules.append(r)
        self._rebuild_desens_table()
        self._select_row(self.tv_desens, len(self.desens_rules) - 1)

    def _edit_desens_rule(self):
        i = self._selected_index(self.tv_desens)
        if i is None or not (0 <= i < len(self.desens_rules)):
            messagebox.showinfo("提示", "请先在表格中选择一条规则", parent=self.root)
            return
        r = self._desens_dialog(self.desens_rules[i])
        if not r:
            return
        new_list = []
        for n, x in enumerate(self.desens_rules):
            if n != i and x.get("tag") == r["tag"]:
                continue          # 与其它规则标签冲突：丢弃冲突的那条
            new_list.append(r if n == i else x)
        self.desens_rules = new_list
        self._rebuild_desens_table()
        self._select_row(self.tv_desens, new_list.index(r))

    def _del_desens_rule(self):
        i = self._selected_index(self.tv_desens)
        if i is None or not (0 <= i < len(self.desens_rules)):
            messagebox.showinfo("提示", "请先在表格中选择一条规则", parent=self.root)
            return
        if not messagebox.askyesno("确认删除", "确定删除规则：%s ？" % tag_display(self.desens_rules[i].get("tag", "")),
                                   parent=self.root):
            return
        del self.desens_rules[i]
        self._rebuild_desens_table()

    def _reset_desens_rules(self):
        if not messagebox.askyesno("恢复默认规则", "将丢弃当前规则表并恢复为默认规则，确定？", parent=self.root):
            return
        self.desens_rules = [dict(r) for r in _DEFAULT_DESENS_RULES]
        self._rebuild_desens_table()

    def _desens_dialog(self, rule=None):
        """新增/编辑一条标签脱敏规则。确定返回规则 dict，取消返回 None。"""
        rule = dict(rule or {})
        dlg = tk.Toplevel(self.root)
        dlg.title("标签脱敏规则")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        result = {"value": None}

        body = ttk.Frame(dlg); body.pack(fill="both", expand=True, padx=14, pady=12)

        r1 = ttk.Frame(body); r1.pack(fill="x", pady=3)
        ttk.Label(r1, text="DICOM 标签:", width=12, anchor="w").pack(side="left")
        var_tag = tk.StringVar(value=str(rule.get("tag", "")))
        ttk.Combobox(r1, textvariable=var_tag, width=44,
                     values=_COMMON_TAGS).pack(side="left", padx=4)
        ttk.Label(body, text="可从下拉选择常用标签；也可手输关键字（PatientName）或 Tag 号（0010,0020）",
                  foreground="#666").pack(anchor="w", padx=(96, 0))

        r2 = ttk.Frame(body); r2.pack(fill="x", pady=3)
        ttk.Label(r2, text="脱敏方法:", width=12, anchor="w").pack(side="left")
        var_method = tk.StringVar(
            value=_DESENS_METHOD_LABEL.get(rule.get("method", "hash"), _DESENS_METHOD_LABEL["hash"]))
        ttk.Combobox(r2, textvariable=var_method, state="readonly", width=30,
                     values=[lab for _k, lab, _h in _DESENS_METHODS]).pack(side="left", padx=4)

        r3 = ttk.Frame(body); r3.pack(fill="x", pady=3)
        ttk.Label(r3, text="参数:", width=12, anchor="w").pack(side="left")
        var_param = tk.StringVar(value=str(rule.get("param", "") or ""))
        ent_param = ttk.Entry(r3, textvariable=var_param, width=30)
        ent_param.pack(side="left", padx=4)
        lbl_hint = ttk.Label(body, text="", foreground="#666")
        lbl_hint.pack(anchor="w", padx=(96, 0))

        r4 = ttk.Frame(body); r4.pack(fill="x", pady=3)
        ttk.Label(r4, text="作用范围:", width=12, anchor="w").pack(side="left")
        var_rec = tk.BooleanVar(value=bool(rule.get("recursive")))
        ttk.Checkbutton(r4, text="同时作用于序列（SQ）内的同名标签", variable=var_rec).pack(side="left")

        def _method_key():
            lab = var_method.get()
            for k, v, _h in _DESENS_METHODS:
                if v == lab:
                    return k
            return "hash"

        def _sync_param(*_a):
            key = _method_key()
            lbl_hint.config(text=_DESENS_METHOD_HINT.get(key, ""))
            ent_param.config(state="normal" if key in _DESENS_PARAM_METHODS else "disabled")

        var_method.trace_add("write", lambda *a: _sync_param())
        _sync_param()

        btns = ttk.Frame(body); btns.pack(fill="x", pady=(12, 0))

        def _ok():
            try:
                tag = parse_tag(var_tag.get())
            except ValueError as e:
                messagebox.showwarning("标签有误", str(e), parent=dlg)
                return
            method = _method_key()
            param = var_param.get().strip() if method in _DESENS_PARAM_METHODS else ""
            if method == "hash":
                try:
                    n = int(param or 16)
                except ValueError:
                    messagebox.showwarning("参数有误", "哈希截断位数必须是整数（默认 16）", parent=dlg)
                    return
                if not (1 <= n <= 128):
                    messagebox.showwarning("参数有误", "哈希截断位数需在 1~128 之间", parent=dlg)
                    return
                param = str(n)
            result["value"] = {"tag": tag_to_str(tag), "method": method,
                               "param": param, "recursive": bool(var_rec.get())}
            dlg.destroy()

        ttk.Button(btns, text="确定", command=_ok).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side="right")
        dlg.bind("<Return>", lambda e: _ok())
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        _center_on_parent(dlg, self.root)
        dlg.grab_set()
        self.root.wait_window(dlg)
        return result["value"]

    # ----- 规则表：增删改（模块②）-----
    def _add_filter_rule(self):
        r = self._filter_dialog()
        if r:
            self.filter_rules.append(r)
            self._rebuild_filter_table()
            self._select_row(self.tv_filter, len(self.filter_rules) - 1)

    def _edit_filter_rule(self):
        i = self._selected_index(self.tv_filter)
        if i is None or not (0 <= i < len(self.filter_rules)):
            messagebox.showinfo("提示", "请先在表格中选择一条规则", parent=self.root)
            return
        r = self._filter_dialog(self.filter_rules[i])
        if not r:
            return
        self.filter_rules[i] = r
        self._rebuild_filter_table()
        self._select_row(self.tv_filter, i)

    def _del_filter_rule(self):
        i = self._selected_index(self.tv_filter)
        if i is None or not (0 <= i < len(self.filter_rules)):
            messagebox.showinfo("提示", "请先在表格中选择一条规则", parent=self.root)
            return
        if not messagebox.askyesno("确认删除", "确定删除规则：%s ？" % tag_display(self.filter_rules[i].get("tag", "")),
                                   parent=self.root):
            return
        del self.filter_rules[i]
        self._rebuild_filter_table()

    def _reset_filter_rules(self):
        if not messagebox.askyesno("恢复默认规则", "将丢弃当前规则表并恢复为默认规则，确定？", parent=self.root):
            return
        self.filter_rules = [dict(r) for r in _DEFAULT_FILTER_RULES]
        self._rebuild_filter_table()

    def _filter_dialog(self, rule=None):
        """新增/编辑一条过滤规则。确定返回规则 dict，取消返回 None。"""
        rule = dict(rule or {})
        dlg = tk.Toplevel(self.root)
        dlg.title("剂量报告/截屏过滤规则")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        result = {"value": None}

        body = ttk.Frame(dlg); body.pack(fill="both", expand=True, padx=14, pady=12)

        r1 = ttk.Frame(body); r1.pack(fill="x", pady=3)
        ttk.Label(r1, text="DICOM 标签:", width=12, anchor="w").pack(side="left")
        var_tag = tk.StringVar(value=str(rule.get("tag", "")))
        ttk.Combobox(r1, textvariable=var_tag, width=44,
                     values=_COMMON_TAGS).pack(side="left", padx=4)
        ttk.Label(body, text="可从下拉选择常用标签；也可手输关键字（ImageType）或 Tag 号（0008,0008）",
                  foreground="#666").pack(anchor="w", padx=(96, 0))

        r2 = ttk.Frame(body); r2.pack(fill="x", pady=3)
        ttk.Label(r2, text="匹配模式:", width=12, anchor="w").pack(side="left")
        var_mode = tk.StringVar(
            value=_FILTER_MODE_LABEL.get(rule.get("mode", "exact"), _FILTER_MODE_LABEL["exact"]))
        ttk.Combobox(r2, textvariable=var_mode, state="readonly", width=30,
                     values=[lab for _k, lab in _FILTER_MODES]).pack(side="left", padx=4)

        r3 = ttk.Frame(body); r3.pack(fill="x", pady=3)
        ttk.Label(r3, text="命中值:", width=12, anchor="w").pack(side="left")
        cur_vals = rule.get("values") or []
        if isinstance(cur_vals, str):
            cur_vals = [cur_vals]
        var_vals = tk.StringVar(value=";".join(str(v) for v in cur_vals))
        ttk.Entry(r3, textvariable=var_vals, width=44).pack(side="left", padx=4)
        ttk.Label(body, text="多个值用「;」分隔，任一命中即算命中；匹配不区分大小写；"
                             "多值字段（如 ImageType）会按「\\」切分后逐段比对",
                  foreground="#666").pack(anchor="w", padx=(96, 0))

        btns = ttk.Frame(body); btns.pack(fill="x", pady=(12, 0))

        def _mode_key():
            lab = var_mode.get()
            for k, v in _FILTER_MODES:
                if v == lab:
                    return k
            return "exact"

        def _ok():
            try:
                tag = parse_tag(var_tag.get())
            except ValueError as e:
                messagebox.showwarning("标签有误", str(e), parent=dlg)
                return
            mode = _mode_key()
            vals = [v.strip() for v in re.split(r"[;\n]", var_vals.get()) if v.strip()]
            if not vals:
                messagebox.showwarning("命中值有误", "请至少填写一个命中值", parent=dlg)
                return
            result["value"] = {"tag": tag_to_str(tag), "mode": mode, "values": vals}
            dlg.destroy()

        ttk.Button(btns, text="确定", command=_ok).pack(side="right", padx=4)
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side="right")
        dlg.bind("<Return>", lambda e: _ok())
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        _center_on_parent(dlg, self.root)
        dlg.grab_set()
        self.root.wait_window(dlg)
        return result["value"]

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
        c = DownloadConfig()
        c.pacs_host = self.var_pacs_host.get().strip()
        try:
            c.pacs_port = int(self.var_pacs_port.get().strip() or 104)
        except ValueError:
            c.pacs_port = 104
        c.pacs_aet = self.var_pacs_aet.get().strip()
        c.local_aet = self.var_local_aet.get().strip() or "MYAET"
        try:
            c.local_port = int(self.var_local_port.get().strip() or 11112)
        except ValueError:
            c.local_port = 11112
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
            c.cmove_timeout = int(self.var_cmove_timeout.get().strip() or 300)
        except ValueError:
            c.cmove_timeout = 300
        try:
            c.cfind_timeout = int(self.var_cfind_timeout.get().strip() or 60)
        except ValueError:
            c.cfind_timeout = 60
        c.desens_enabled = bool(self.var_desens.get())
        c.desens_rules = [dict(r) for r in self.desens_rules]
        c.desens_purge_private = bool(self.var_purge_private.get())
        c.desens_date_month = bool(self.var_date_month.get())
        c.desens_time_clear = bool(self.var_time_clear.get())
        c.filter_enabled = bool(self.var_filter.get())
        c.filter_rules = [dict(r) for r in self.filter_rules]
        c.filter_save_hit = bool(self.var_filter_save_hit.get())
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
        local_aet = self.var_local_aet.get().strip() or None
        try:
            local_port = int(self.var_local_port.get().strip() or 11112)
        except ValueError:
            local_port = None

        self._append_log("正在双向检测 %s:%d 连通性（正向 C-Echo + 反向接收探测）..." % (host, port))
        self.btn_test.config(state="disabled")

        def worker():
            ok, msgs = test_connectivity(host, port, aet, local_port=local_port, local_aet=local_aet)
            _ui_queue.put(("conn_result", (ok, msgs)))

        threading.Thread(target=worker, daemon=True).start()

    def _sync_diag(self):
        """勾选/取消「诊断日志」立即生效（不必等下次开始下载）。"""
        global _diag_log
        _diag_log = bool(self.var_diag_log.get())
        self._append_log("诊断日志已%s" % ("开启" if _diag_log else "关闭"))

    def _start(self):
        global _diag_log
        if self.thread and self.thread.is_alive():
            messagebox.showinfo("提示", "已有下载任务在运行中")
            return
        _diag_log = bool(self.var_diag_log.get())
        try:
            int(self.var_pacs_port.get().strip() or 104)
            int(self.var_local_port.get().strip() or 11112)
        except ValueError:
            messagebox.showwarning("提示", "PACS端口和本机接收端口必须是数字")
            return
        cfg = self._collect_cfg()
        if not cfg.pacs_host or not cfg.pacs_aet:
            messagebox.showwarning("提示", "请填写 PACS IP 和 AE Title")
            return
        # AE Title 校验：1~16 字符，不能含空格/中文/特殊字符（pynetdicom 编码会失败或被 PACS 拒）
        for nm, v in (("PACS AE Title", cfg.pacs_aet), ("本机 AE Title", cfg.local_aet)):
            if not v:
                messagebox.showwarning("提示", "%s 不能为空" % nm)
                return
            if len(v) > 16 or any(ch.isspace() for ch in v) or not v.isascii() or not v.isprintable():
                messagebox.showwarning("提示", "%s 不合法：必须是 1~16 位 ASCII 可显示字符，且不能含空格（当前值：%r）" % (nm, v))
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
        # 启动后禁用关键输入框，防止中途修改导致与后台线程不一致
        for w in self._iter_editable_widgets():
            try:
                w.config(state="disabled")
            except Exception:
                pass
        self.btn_test.config(state="disabled")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_history.config(state="disabled")
        for b in self._desens_buttons + self._filter_buttons:
            try:
                b.config(state="disabled")
            except Exception:
                pass
        self.var_status.set("下载中...")
        self.var_image_count.set("影像数: 0 / ?")
        self.var_elapsed.set("已耗时: 00:00:00")
        # 启动 1Hz 定时器刷新"已耗时"显示
        self._download_start_ts = time.time()
        self._tick_after_id = None
        self._tick_running = True
        self._tick_update()

        self.thread = threading.Thread(target=batch_download, args=(cfg,), daemon=True)
        self.thread.start()

    def _stop(self):
        _stop_event.set()
        _abort_active_assocs()
        self._append_log("正在停止（中断当前连接，稍候即生效）...")

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
            self.var_excel.set(str(self.cfg.excel_path or ""))
            self.var_column.set(str(self.cfg.column or "影像号"))
            self.var_sheet.set(str(self.cfg.sheet_name or ""))
            if getattr(self.cfg, "key_type", "patient_id") == "patient_id":
                self.var_key_type.set("病人ID(patientId)")
            else:
                self.var_key_type.set("StudyInstanceUID(影像号UID)")
            self.var_rate_limit.set(str(getattr(self.cfg, "rate_limit_kbps", 0)))
            self.var_pause_every.set(str(getattr(self.cfg, "pause_every", 0)))
            self.var_pause_seconds.set(str(getattr(self.cfg, "pause_seconds", 30)))
            self.var_cmove_timeout.set(str(getattr(self.cfg, "cmove_timeout", 300)))
            self.var_cfind_timeout.set(str(getattr(self.cfg, "cfind_timeout", 60)))
            self.var_desens.set(bool(getattr(self.cfg, "desens_enabled", False)))
            self.var_filter.set(bool(getattr(self.cfg, "filter_enabled", False)))
            self.var_purge_private.set(bool(getattr(self.cfg, "desens_purge_private", True)))
            self.var_date_month.set(bool(getattr(self.cfg, "desens_date_month", True)))
            self.var_time_clear.set(bool(getattr(self.cfg, "desens_time_clear", True)))
            self.var_filter_save_hit.set(bool(getattr(self.cfg, "filter_save_hit", True)))
            # 按加载后的开关状态同步展开/收起详细配置
            self._toggle_desens_detail()
            self._toggle_filter_detail()
            # 规则表：配置里没有（旧版 config.json）时保留界面上的默认规则
            dr = getattr(self.cfg, "desens_rules", None)
            if isinstance(dr, list) and dr:
                self.desens_rules = [dict(r) for r in dr if isinstance(r, dict)]
                self._rebuild_desens_table()
            fr = getattr(self.cfg, "filter_rules", None)
            if isinstance(fr, list) and fr:
                self.filter_rules = [dict(r) for r in fr if isinstance(r, dict)]
                self._rebuild_filter_table()
            self._append_log("已自动加载配置：%s（脱敏规则 %d 条，过滤规则 %d 条）" % (
                CONFIG_PATH, len(self.desens_rules), len(self.filter_rules)))
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

    def _reset_ui(self, status_text=None):
        """复位按钮/状态（下载线程结束或出错时调用），保证界面不会卡死。"""
        # 停止 1Hz 定时器
        self._tick_running = False
        if getattr(self, "_tick_after_id", None) is not None:
            try:
                self.root.after_cancel(self._tick_after_id)
            except Exception:
                pass
            self._tick_after_id = None
        # 恢复所有输入框为可编辑（下载结束/异常退出时）；
        # Combobox 需恢复为 readonly（原本只读），避免被输入任意文本
        for w in self._iter_editable_widgets():
            try:
                w.config(state="readonly" if isinstance(w, ttk.Combobox) else "normal")
            except Exception:
                pass
        self.btn_test.config(state="normal")
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        for b in self._desens_buttons + self._filter_buttons:
            try:
                b.config(state="normal")
            except Exception:
                pass
        if status_text:
            self.var_status.set(status_text)

    def _tick_update(self):
        """每秒刷新"已耗时"，下载结束后自动停止。"""
        try:
            if not getattr(self, "_tick_running", False):
                return
            start = getattr(self, "_download_start_ts", None)
            if start is not None:
                elapsed = max(0, int(time.time() - start))
                self.var_elapsed.set("已耗时: " + self._format_hms(elapsed))
            self._tick_after_id = self.root.after(1000, self._tick_update)
        except Exception:
            self._tick_after_id = None

    @staticmethod
    def _format_hms(seconds):
        s = int(seconds)
        h = s // 3600
        m = (s % 3600) // 60
        sec = s % 60
        return "%02d:%02d:%02d" % (h, m, sec)

    def _show_history(self):
        """弹出下载记录列表窗口。"""
        if not self.records:
            messagebox.showinfo("提示", "暂无下载记录，请先执行一次下载")
            return
        win = tk.Toplevel(self.root)
        win.title("下载记录")
        win.geometry("1060x520")
        win.transient(self.root)

        cols = ("idx", "key", "uid", "status", "n_files", "message", "time")
        headers = tuple(_REPORT_HEADER)  # 复用报告表头，避免重复定义
        widths = (60, 150, 240, 100, 70, 260, 150)

        main = ttk.Frame(win)
        main.pack(fill="both", expand=True, padx=8, pady=8)
        tree = ttk.Treeview(main, columns=cols, show="headings")
        for c, h, w in zip(cols, headers, widths):
            tree.heading(c, text=h)
            tree.column(c, width=w, anchor="w")
        for r in self.records:
            tree.insert("", "end", values=(
                r["idx"], r["key"], r["study_uid"], r["status"],
                r["n_files"], r["message"], r["time"],
            ))

        vsb = ttk.Scrollbar(main, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(main, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(bar, text="导出 CSV", command=lambda: self._export_history(win)).pack(side="left", padx=4)
        ttk.Button(bar, text="关闭", command=win.destroy).pack(side="left", padx=4)

    def _export_history(self, parent=None):
        """把当前下载记录导出为用户指定位置的 CSV。"""
        if not self.records:
            return
        p = filedialog.asksaveasfilename(
            parent=parent, title="导出下载记录",
            defaultextension=".csv", initialfile="下载记录.csv",
            filetypes=[("CSV 文件", "*.csv"), ("所有文件", "*.*")],
        )
        if not p:
            return
        try:
            with open(p, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(_REPORT_HEADER)
                for r in self.records:
                    w.writerow([r["idx"], r["key"], r["study_uid"], r["status"],
                                r["n_files"], r["message"], r["time"]])
            messagebox.showinfo("提示", "已导出：%s" % p)
        except Exception as e:
            messagebox.showerror("错误", "导出失败：%s" % e)

    def _poll_queue(self):
        try:
            while True:
                kind, payload = _ui_queue.get_nowait()
                try:
                    if kind == "log":
                        self._append_log(payload)
                    elif kind == "error":
                        self._append_log("[错误] " + payload)
                        self._reset_ui("出错，已停止（可修改配置后重新开始）")
                        # 延后弹窗：错误信息较短（位置/类型/原因），完整堆栈已写日志
                        self.root.after(50, lambda p=payload: messagebox.showerror("错误", p[:2000]))
                    elif kind == "conn_result":
                        ok, msgs = payload
                        for m in msgs:
                            self._append_log(m)
                        # 仅在无下载任务运行时才恢复检测按钮，避免与下载期间的禁用冲突
                        if not (self.thread and self.thread.is_alive()):
                            self.btn_test.config(state="normal")
                        if ok:
                            self.root.after(50, lambda m=msgs: messagebox.showinfo("检测结果", "\n".join(m)))
                        else:
                            self.root.after(50, lambda m=msgs: messagebox.showwarning("检测结果", "\n".join(m)))
                    elif kind == "progress":
                        done, total = payload
                        if total > 0:
                            self.progress["value"] = done * 100.0 / total
                            self.var_status.set("进度 %d / %d" % (done, total))
                    elif kind == "image_progress":
                        # (label, n_files, n_total) - 当前检查的影像张数进度
                        try:
                            _lbl, n_done, n_total = payload
                            total_str = str(n_total) if n_total > 0 else "?"
                            self.var_image_count.set("当前[%s] 影像数: %d / %s" % (_lbl, n_done, total_str))
                        except Exception:
                            pass
                    elif kind == "status":
                        self.var_status.set(payload)
                    elif kind == "done":
                        self._append_log(payload)
                        self.var_status.set(payload)
                        self._reset_ui()
                        self.root.after(50, lambda p=payload: messagebox.showinfo("完成", p))
                    elif kind == "report":
                        records, csv_path = payload
                        self.records = records or []
                        self.report_csv = csv_path
                        self.btn_history.config(state="normal" if self.records else "disabled")
                        if csv_path:
                            self._append_log("[报告] 下载报告已保存：%s" % csv_path)
                    elif kind == "reset":
                        # 下载线程任何退出路径都会发送，兜底复位界面
                        self._reset_ui()
                except Exception:
                    # 单条消息处理失败不影响后续轮询
                    traceback.print_exc()
        except queue.Empty:
            pass
        except Exception:
            traceback.print_exc()
        self.root.after(200, self._poll_queue)

    def _on_close(self):
        """主窗口关闭事件：优雅停止 Store SCP + 关联释放，避免 Windows 端口 TIME_WAIT。"""
        # 标记停止，避免正在跑的下载继续
        try:
            _stop_event.set()
        except Exception:
            pass
        # 主动 abort 所有活跃 DICOM association（复用统一实现）
        _abort_active_assocs()
        # 优雅关闭 Store SCP
        global _store_server, _store_started
        if _store_started and _store_server is not None:
            try:
                _store_server.shutdown()
            except Exception:
                pass
            _store_server = None
            _store_started = False
        self.root.destroy()


def main():
    root = tk.Tk()
    app = App(root)
    # 关键：拦截窗口关闭事件，优雅停 Store SCP，避免 Windows 端口 TIME_WAIT
    root.protocol("WM_DELETE_WINDOW", app._on_close)
    root.mainloop()


if __name__ == "__main__":
    main()