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
import json
import os
import queue
import socket
import sys
import threading
import time
import traceback

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
_active_assocs = []            # 当前活跃的 DICOM association，停止时强制中断
_assoc_lock = threading.Lock()

# 下载结果记录（每次下载一条，最终写入 CSV 报告并在 GUI 历史列表展示）
_download_records = []
_records_lock = threading.Lock()

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
    if s is None:
        return "unnamed"
    s = str(s)
    safe = "".join(c if (c.isalnum() or c in "._-") else "_" for c in s)
    # 极端兜底：纯特殊字符（如 NULL 字节）会被全部替换为 _，但仍然是非空
    return safe if safe.strip("_") else "unnamed"


def _on_conn_open(event):
    """Store SCP 收到入站连接时记录来源，便于判断 PACS 是否真的在往本机推影像。"""
    try:
        addr = "%s:%s" % (event.address[0], event.address[1])
    except Exception:
        addr = "未知"
    _ui_queue.put(("log", "[StoreSCP] 收到入站连接：%s" % addr))


def _count_dcm(subdir):
    try:
        return len([f for f in os.listdir(subdir) if f.endswith(".dcm")])
    except Exception:
        return 0


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


def handle_store(event):
    """处理 C-STORE：把收到的 DICOM 文件保存到对应 Study 子目录（并发安全）。"""
    try:
        ds = event.dataset
        ds.file_meta = event.file_meta
        sop_uid = getattr(ds, "SOPInstanceUID", None) or ("unknown-%d" % int(time.time()))
        study_uid = getattr(ds, "StudyInstanceUID", None)
        # 关键防御：OUTPUT_ROOT 尚未设置或已被清空时（如程序关闭中/已停止）拒绝写入
        if not OUTPUT_ROOT:
            _ui_queue.put(("log", "[StoreSCP] 收到 DICOM 但输出目录未就绪，已拒绝：SOP=%s" % sop_uid))
            return 0xC000
        if study_uid:
            # 优先按图像自带的 StudyInstanceUID 定位目录（并发下载时不会串目录）
            d = os.path.join(OUTPUT_ROOT, _safe_name(study_uid))
        else:
            # 缺失 StudyInstanceUID 的异常文件统一兜底，避免并发时串到别的目录
            d = os.path.join(OUTPUT_ROOT, "_unknown_study")
            _ui_queue.put(("log", "[StoreSCP] 收到缺少 StudyInstanceUID 的文件，落入 _unknown_study/"))
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
        # 限速：按实际落盘文件大小节流，压低下行速率
        try:
            _throttle(os.path.getsize(path))
        except Exception:
            pass
        return 0x0000  # Success
    except Exception as e:
        _ui_queue.put(("log", "[StoreSCP] 保存 DICOM 失败：%s" % e))
        return 0xC000  # Unable to process，避免单文件失败中断接收


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
    # StoragePresentationContexts 是列表，需要逐个 add_supported_context（pynetdicom 3.x 没有复数 API）
    for ctx in storage_ctxs:
        ae.add_supported_context(ctx.abstract_syntax, ctx.transfer_syntax)
    ae.add_requested_context(Verification)
    handlers = [(evt.EVT_C_STORE, handle_store), (evt.EVT_CONN_OPEN, _on_conn_open)]
    # 优先用属性设置 AE Title（所有 pynetdicom 版本通用，2.x 之前 start_server 不支持 ae_title 关键字）
    try:
        ae.ae_title = local_aet
    except Exception:
        pass
    try:
        # 跨 2.x/3.x 兼容：先 try 关键字参数，失败回落位置参数
        try:
            _store_server = ae.start_server(
                ("0.0.0.0", local_port), block=False, evt_handlers=handlers, ae_title=local_aet,
            )
        except TypeError:
            # 2.x 早期：ae_title / evt_handlers 是位置参数
            _store_server = ae.start_server(
                ("0.0.0.0", local_port), False, handlers, local_aet,
            )
    except Exception as e:
        _ui_queue.put(("error", "Store SCP 启动失败：%s\n%s" % (e, _port_bind_hint(local_port))))
        return False
    # 成功启动后统一更新状态（之前在 3.x 路径下会漏设置 _store_started）
    _store_started = True
    _store_aet = local_aet
    _store_port = local_port
    ips = "、".join(_local_ips()) or "未知"
    _ui_queue.put(("log", "[StoreSCP] 接收服务已启动：AE=%s 端口=%d（本机 IP：%s，请确认与前置机注册一致）" % (local_aet, local_port, ips)))
    return True


# ---------------------------------------------------------------------------
# 2) 网络可达性检测
# ---------------------------------------------------------------------------
def test_connectivity(host, port, aet=None, local_port=None):
    """
    检测 PACS 前置机可达性。
    返回 (ok, messages)：
        1. TCP 端口连通性检查（判断网络/端口是否可达）
        2. DICOM C-Echo 检查（判断 DICOM 服务与 AE Title 是否有效，需要提供 aet）
        3. 本机接收端口自检（判断下载时能否绑定本地端口，需要提供 local_port）
    """
    msgs = []
    tcp_ok = False
    echo_ok = False
    local_ok = True

    # 0) 本机接收端口自检（与 PACS 无关，先查，避免"测试成功但下载绑定失败"）
    # 注意：本工具的 Store SCP 若已在监听该端口，属于正常占用，跳过自检
    if local_port and not (_store_started and _store_port == local_port):
        ok, err = _try_bind_port(local_port)
        if ok:
            msgs.append("本机接收端口自检：端口 %d 可正常绑定" % local_port)
        else:
            local_ok = False
            msgs.append("本机接收端口自检失败：端口 %d 无法绑定（%s）" % (local_port, err))
            msgs.append(_port_bind_hint(local_port))

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
        return (tcp_ok and local_ok), msgs

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
                    msgs.append("DICOM C-Echo 成功：AE Title=%s 有效" % aet)
                else:
                    msgs.append("DICOM C-Echo 未返回成功状态")
            finally:
                try:
                    assoc.release()
                except Exception:
                    pass
        else:
            msgs.append("DICOM 关联建立失败：请检查 AE Title=%s 是否正确" % aet)
            # 未建立时也要尝试 release 释放 socket 资源
            try:
                assoc.release()
            except Exception:
                pass
    except Exception as e:
        msgs.append("DICOM C-Echo 异常：%s" % e)

    return (tcp_ok and local_ok and (echo_ok if aet else True)), msgs


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
        # 注意：pynetdicom 2.x/3.x 签名兼容写法 - query_model 位置参数（3.x 不支持 query_model= 关键字）
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


def pull_one_study(assoc, study_uid, local_aet, out_root):
    """拉取一个 Study，影像落到 out_root/<study_uid>/ 目录。"""
    subdir = os.path.join(out_root, _safe_name(study_uid))
    os.makedirs(subdir, exist_ok=True)

    ds = Dataset()
    ds.QueryRetrieveLevel = "STUDY"
    ds.StudyInstanceUID = study_uid

    _ui_queue.put(("log", "      [C-Move] 请求拉取 StudyInstanceUID=%s -> 目标AE=%s (QueryRetrieveLevel=STUDY)" % (study_uid, local_aet)))

    final_code = 0x0000
    has_error = False
    n_failed = 0    # 子操作失败数（前置机推送失败的影像数）
    n_warned = 0    # 子操作警告数
    try:
        # 3.x 签名：send_c_move(dataset, move_aet, query_model) - query_model 必须位置参数
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
                extra = ""
                if remaining is not None or completed is not None:
                    extra = " (剩余=%s 已完成=%s 失败=%s 警告=%s)" % (remaining, completed, failed, warned)
                _ui_queue.put(("log", "      [C-Move] 响应状态 0x%04X %s%s" % (code, _status_text(code), extra)))
                if code not in _PENDING:
                    final_code = code
                    # 0xB000=警告:子操作完成但有失败，属于“部分成功”，不当作硬错误；
                    # 交由 _download_one 依据 n_failed/n_warned 走增量补拉，避免整 Study 全量重拉
                    if code not in (0x0000, 0xB000):
                        has_error = True
    except Exception as e:
        has_error = True
        _ui_queue.put(("log", "      [C-Move 异常] %s" % e))

    n_files = _count_dcm(subdir)
    # 部分医院前置机是"先返回 C-Move 成功、再异步推送影像"，
    # 状态成功但暂时 0 文件时，最多等 30 秒观察文件是否陆续落盘
    if not has_error and n_files == 0:
        _ui_queue.put(("log", "      [C-Move] 状态成功但暂无文件落盘，等待前置机异步推送（最多 30 秒）..."))
        for _ in range(60):
            time.sleep(0.5)
            if _stop_event.is_set():
                break
            n_files = _count_dcm(subdir)
            if n_files > 0:
                break
    _ui_queue.put(("log", "      [C-Move] 结束，最终状态 0x%04X %s，本地落盘 %d 个文件" % (final_code, _status_text(final_code), n_files)))
    return has_error, final_code, n_files, subdir, n_failed, n_warned


def find_missing_sop_uids(cfg, study_uid, subdir):
    """IMAGE 级 C-Find 列出该 Study 应有的全部 SOPInstanceUID，与本地已落盘文件对比，
    返回缺失的 UID 列表。查询失败/无结果时返回 None（表示无法增量核对，需整 Study 重拉）。"""
    assoc = _make_assoc(cfg, StudyRootQueryRetrieveInformationModelFind, "C-Find", "cfind_timeout", 60)
    if not (assoc and assoc.is_established):
        return None
    try:
        ds = Dataset()
        ds.QueryRetrieveLevel = "IMAGE"
        ds.StudyInstanceUID = study_uid
        ds.SOPInstanceUID = ""
        expected = []
        seen = set()
        try:
            for status, identifier in assoc.send_c_find(ds, StudyRootQueryRetrieveInformationModelFind):
                if identifier:
                    suid = getattr(identifier, "SOPInstanceUID", None)
                    if suid:
                        s = str(suid)
                        if s not in seen:
                            seen.add(s)
                            expected.append(s)
        except Exception as e:
            _ui_queue.put(("log", "      [增量核对] IMAGE 级 C-Find 异常：%s" % e))
            return None
        if not expected:
            # 前置机可能不支持 IMAGE 级查询
            return None
        local = set()
        try:
            for f in os.listdir(subdir):
                if f.endswith(".dcm"):
                    local.add(f[:-4])  # 文件名即 _safe_name(SOPInstanceUID)
        except Exception:
            pass
        missing = [u for u in expected if _safe_name(u) not in local]
        _ui_queue.put(("log", "      [增量核对] 前置机应有 %d 个影像，本地已有 %d 个，缺失 %d 个" % (
            len(expected), len(local), len(missing))))
        return missing
    finally:
        _unregister_assoc(assoc)
        try:
            assoc.release()
        except Exception:
            pass


def pull_images(cfg, study_uid, sop_uids, local_aet):
    """IMAGE 级 C-Move 只补拉指定影像。返回 (has_error, final_code)。"""
    assoc = _make_assoc(cfg, StudyRootQueryRetrieveInformationModelMove, "C-Move", "cmove_timeout", 300)
    if not (assoc and assoc.is_established):
        return True, 0xA700
    try:
        ds = Dataset()
        ds.QueryRetrieveLevel = "IMAGE"
        ds.StudyInstanceUID = study_uid
        ds.SOPInstanceUID = sop_uids
        final_code = 0x0000
        has_error = False
        try:
            for status, _ in assoc.send_c_move(ds, local_aet, StudyRootQueryRetrieveInformationModelMove):
                if status:
                    code = status.Status
                    if code not in _PENDING:
                        _ui_queue.put(("log", "      [补拉] 响应状态 0x%04X %s" % (code, _status_text(code))))
                        final_code = code
                        # 0xB000 同样视为“部分成功”，回到调用处做二次增量核对，
                        # 而不是直接判定失败进入重试
                        if code not in (0x0000, 0xB000):
                            has_error = True
        except Exception as e:
            has_error = True
            _ui_queue.put(("log", "      [补拉 异常] %s" % e))
        return has_error, final_code
    finally:
        _unregister_assoc(assoc)
        try:
            assoc.release()
        except Exception:
            pass


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
        self.cmove_timeout = 300      # C-MOVE 单次最大等待（秒），超时后重试
        self.cfind_timeout = 60       # C-FIND 单次最大等待（秒）
        self.cmove_retry = 1          # C-MOVE 失败/超时后重试次数
        self.cmove_retry_delay = 3    # C-MOVE 重试间隔（秒）


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
    try:
        assoc = ae.associate(cfg.pacs_host, cfg.pacs_port, ae_title=cfg.pacs_aet)
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


_MANIFEST_LOCK = threading.Lock()


def _manifest_path():
    """成功清单文件路径（与下载输出同目录）。"""
    if not OUTPUT_ROOT:
        return None
    return os.path.join(OUTPUT_ROOT, "成功清单.csv")


def _load_manifest_uids():
    """读取成功清单，返回 set。文件不存在或异常时返回空集。"""
    p = _manifest_path()
    if not p or not os.path.isfile(p):
        return set()
    s = set()
    try:
        with _MANIFEST_LOCK:
            with open(p, "r", encoding="utf-8-sig", newline="") as f:
                r = csv.reader(f)
                for row in r:
                    if not row:
                        continue
                    cell = row[0].strip()
                    if cell and cell.lower() != "studyinstanceuid":
                        s.add(cell)
    except Exception as e:
        _ui_queue.put(("log", "[清单] 读取成功清单失败：%s" % e))
    return s


def _append_manifest_uid(uid):
    """下载成功后追加一条 UID 到成功清单。"""
    p = _manifest_path()
    if not p:
        return
    try:
        with _MANIFEST_LOCK:
            os.makedirs(OUTPUT_ROOT, exist_ok=True)
            new_file = not os.path.isfile(p)
            with open(p, "a", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                if new_file:
                    w.writerow(["StudyInstanceUID"])
                w.writerow([uid])
    except Exception as e:
        _ui_queue.put(("log", "[清单] 写入成功清单失败：%s" % e))


def _already_downloaded(uid, manifest_set):
    """断点续传判断：以“成功清单”为准，清单里有的视为已成功。"""
    if uid in manifest_set:
        # 顺便统计本地落盘数（仅供日志显示，不影响跳过判断）
        d = os.path.join(OUTPUT_ROOT, _safe_name(uid))
        n = 0
        if os.path.isdir(d):
            try:
                n = len([f for f in os.listdir(d) if f.endswith(".dcm")])
            except Exception:
                n = 0
        return True, n
    return False, 0


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


def _download_one(cfg, key, label, uid, idx, total, manifest_set):
    """下载单个 Study（失败/超时自动重试，已存在则跳过）。返回 (idx, status, n_files, message)。"""
    retry = int(getattr(cfg, "cmove_retry", 1) or 1)
    delay = int(getattr(cfg, "cmove_retry_delay", 3) or 3)
    fail_msg = ""
    n_files = 0

    # 停止优先于断点续传：停止后剩余任务统一记为“停止”
    if _stop_event.is_set():
        _ui_queue.put(("log", "[%d/%d] [跳过] %s（手动停止）" % (idx, total, label)))
        _record_download(idx, key, uid, label, "停止", 0)
        return idx, "stopped", 0, "已停止"

    # 断点续传：成功清单里有该 UID 则跳过（精准判断，不受半成品影响）
    exists, n_exist = _already_downloaded(uid, manifest_set)
    if exists:
        _ui_queue.put(("log", "[%d/%d] [跳过] %s（成功清单已存在，本地 %d 个文件，不重复下载）" % (idx, total, label, n_exist)))
        _record_download(idx, key, uid, label, "已存在(跳过)", n_exist)
        return idx, "skipped", n_exist, ""

    def _success(n, note=""):
        """成功统一出口：写成功清单 + 记录 + 返回。"""
        _append_manifest_uid(uid)
        manifest_set.add(uid)
        _record_download(idx, key, uid, label, "成功", n, note)
        return idx, "success", n, note

    for attempt in range(retry + 1):
        if _stop_event.is_set():
            _ui_queue.put(("log", "[%d/%d] [跳过] %s（手动停止）" % (idx, total, label)))
            _record_download(idx, key, uid, label, "停止", 0)
            return idx, "stopped", 0, "已停止"
        subdir = os.path.join(OUTPUT_ROOT, _safe_name(uid))
        n_files = _count_dcm(subdir)

        # 增量核对：本地已有文件时，先用 IMAGE 级 C-Find 列出缺失影像
        missing = None
        if n_files > 0:
            missing = find_missing_sop_uids(cfg, uid, subdir)
            if missing is not None and not missing:
                # 本地已齐全（如之前下载过但清单丢失）：直接算成功，不重拉
                _ui_queue.put(("log", "[%d/%d] [完成] %s 本地已齐全（%d 个文件，增量核对通过）" % (idx, total, label, n_files)))
                return _success(n_files, "增量核对齐全")

        if missing:
            # 增量补拉：只重拉缺失的影像，不整 Study 重拉
            _ui_queue.put(("log", "[%d/%d] [增量补拉] %s：缺失 %d 个，仅补拉缺失部分" % (idx, total, label, len(missing))))
            has_error, code = pull_images(cfg, uid, missing, cfg.local_aet)
            n_files = _count_dcm(subdir)
            if not has_error:
                missing2 = find_missing_sop_uids(cfg, uid, subdir)
                if not missing2:  # [] 齐全 或 None 无法核对，均视为完成
                    _ui_queue.put(("log", "[%d/%d] [完成] %s -> %d 个文件（增量补拉完成） %s" % (idx, total, label, n_files, subdir)))
                    return _success(n_files, "增量补拉完成")
                fail_msg = "增量补拉后仍缺 %d 个（可能是前置机侧该影像本身有问题）" % len(missing2)
            else:
                fail_msg = "增量补拉状态码 0x%04X（本地文件 %d 个）" % (code, n_files)
        else:
            # 首次拉取 或 无法增量核对：整 Study 全量拉取
            assoc = _make_assoc(cfg, StudyRootQueryRetrieveInformationModelMove, "C-Move", "cmove_timeout", 300)
            if assoc and assoc.is_established:
                try:
                    has_error, code, n_files, subdir, n_failed, n_warned = pull_one_study(assoc, uid, cfg.local_aet, OUTPUT_ROOT)
                finally:
                    _unregister_assoc(assoc)
                    try:
                        assoc.release()
                    except Exception:
                        pass
                # 双重判定：状态码 0x0000 且本地至少落盘 1 个 .dcm 才视为成功
                if not has_error and n_files > 0:
                    # 仅当状态码为 0x0000 且无子操作失败/警告时才算完全成功；
                    # 若状态码是 0xB000（即使前置机没回填失败计数），也要走增量补拉
                    if n_failed == 0 and n_warned == 0 and code == 0x0000:
                        _ui_queue.put(("log", "[%d/%d] [完成] %s -> %d 个文件 %s" % (idx, total, label, n_files, subdir)))
                        return _success(n_files)
                    # 子操作有失败/警告：增量补拉缺失影像一次
                    _ui_queue.put(("log", "[%d/%d] [补拉] 前置机报告子操作 %d 失败/%d 警告，增量补拉缺失影像" % (idx, total, n_failed, n_warned)))
                    miss = find_missing_sop_uids(cfg, uid, subdir)
                    if miss is None:
                        # 前置机不支持 IMAGE 级查询，无法核对：有文件即视为成功（带提醒）
                        _ui_queue.put(("log", "[%d/%d] [提醒] %s 无法增量核对（前置机可能不支持 IMAGE 级查询），按文件数视为成功" % (idx, total, label)))
                        return _success(n_files, "子操作 %d 失败/%d 警告，无法增量核对" % (n_failed, n_warned))
                    if miss:
                        pull_images(cfg, uid, miss, cfg.local_aet)
                        n_files = _count_dcm(subdir)
                    miss2 = find_missing_sop_uids(cfg, uid, subdir)
                    if not miss2:
                        _ui_queue.put(("log", "[%d/%d] [完成] %s -> %d 个文件（补拉完成） %s" % (idx, total, label, n_files, subdir)))
                        return _success(n_files, "补拉完成")
                    fail_msg = "子操作 %d 失败/%d 警告，补拉后仍缺 %d 个" % (n_failed, n_warned, len(miss2))
                elif has_error:
                    fail_msg = "状态码 0x%04X（本地文件 %d 个）" % (code, n_files)
                else:
                    # 状态码成功但 0 文件：前置机没有把影像推到本机（注册信息不符/防火墙/异步未推）
                    ips = "、".join(_local_ips()) or "未知"
                    fail_msg = ("C-Move 状态成功但等待 30 秒后仍 0 个文件落盘。"
                                "请核对：1) 医院前置机上「%s」注册的 IP 是否为本机当前 IP（%s）；"
                                "2) 注册端口是否为 %d；3) Windows 防火墙是否放行 %d 入站。"
                                % (cfg.local_aet, ips, cfg.local_port, cfg.local_port))
            else:
                fail_msg = "连接 PACS 失败"
        if attempt < retry:
            _ui_queue.put(("log", "[%d/%d] [重试 %d/%d] %s：%s，%d 秒后重试" % (
                idx, total, attempt + 1, retry, label, fail_msg, delay)))
            t_end = time.time() + delay
            while time.time() < t_end and not _stop_event.is_set():
                time.sleep(0.5)
    _ui_queue.put(("log", "[%d/%d] [失败] %s：%s" % (idx, total, label, fail_msg)))
    _record_download(idx, key, uid, label, "失败", n_files, fail_msg)
    return idx, "failed", n_files, fail_msg


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
        _ui_queue.put(("report", (records, csv_path)))
        _ui_queue.put(("reset", None))


def _batch_download_inner(cfg):
    global OUTPUT_ROOT, _rate_limit_kbps
    global _rate_last_time, _rate_tokens
    # 关键：重置停止标志，避免上一次“停止”被传染到本次下载
    _stop_event.clear()
    _clear_records()

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

    # 4) 串行下载（稳态，支持间隙暂停、停止中断、重试、断点续传）
    cmove_timeout = int(getattr(cfg, "cmove_timeout", 300) or 300)
    cmove_retry = int(getattr(cfg, "cmove_retry", 1) or 1)
    _ui_queue.put(("log", "开始下载：串行（稳态），C-Move 超时=%d 秒，重试=%d 次" % (cmove_timeout, cmove_retry)))

    # 加载成功清单（精准断点续传依据）
    manifest_set = _load_manifest_uids()
    if manifest_set:
        _ui_queue.put(("log", "成功清单已加载：%d 个 UID 标记为已下载，将自动跳过" % len(manifest_set)))

    ok = 0
    skip = 0
    fail = 0
    pause_every = int(getattr(cfg, "pause_every", 0) or 0)
    pause_seconds = int(getattr(cfg, "pause_seconds", 0) or 0)

    for i, (key, label, uid) in enumerate(tasks, 1):
        if _stop_event.is_set():
            _ui_queue.put(("log", "已手动停止，剩余 %d 个未处理" % (total - i + 1)))
            break
        _ui_queue.put(("progress", (i, total)))
        _ui_queue.put(("log", "[%d/%d] 拉取 %s" % (i, total, label)))
        _idx, status, n_files, msg = _download_one(cfg, key, label, uid, i, total, manifest_set)
        if status == "success":
            ok += 1
        elif status == "skipped":
            skip += 1
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

    if _stop_event.is_set():
        _ui_queue.put(("done", "已停止：成功 %d / 跳过(已存在) %d / 失败 %d / 共 %d（剩余任务未处理）" % (ok, skip, fail, total)))
        _ui_queue.put(("log", "已停止：成功 %d / 跳过 %d / 失败 %d / 共 %d" % (ok, skip, fail, total)))
    else:
        _ui_queue.put(("done", "下载完成：成功 %d / 跳过(已存在) %d / 失败 %d / 共 %d" % (ok, skip, fail, total)))


# ---------------------------------------------------------------------------
# 5) GUI 界面
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("DICOM 影像批量下载工具")
        root.geometry("760x720")
        root.minsize(720, 650)

        self.cfg = DownloadConfig()
        self.thread = None
        self.records = []       # 最近一次下载的结果记录
        self.report_csv = None  # 最近一次下载报告 CSV 路径

        self._build_widgets()
        self._load_config()
        self._poll_queue()

    # ----- 布局 -----
    def _build_widgets(self):
        pad = dict(padx=6, pady=3)

        # PACS 配置
        f1 = ttk.LabelFrame(self.root, text="PACS 前置机配置（由医院实施工程师提供）")
        f1.pack(fill="x", padx=10, pady=(10, 4))
        self._entry_pair(f1, "PACS IP", "pacs_host", default="")
        self._entry_pair(f1, "PACS 端口", "pacs_port", default="104")
        self._entry_pair(f1, "PACS AE Title", "pacs_aet", default="")

        # 本机配置
        f2 = ttk.LabelFrame(self.root, text="本机接收节点配置（需注册到 PACS 前置机）")
        f2.pack(fill="x", padx=10, pady=4)
        self._entry_pair(f2, "本机 AE Title", "local_aet", default="MYAET")
        self._entry_pair(f2, "本机接收端口", "local_port", default="11112")

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
        ttk.Label(row3, text="(0=不限速, 如1024=1MB/s)").pack(side="left")

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
        self.btn_history = ttk.Button(fbtn, text="下载记录", command=self._show_history, state="disabled")
        self.btn_history.pack(side="left", padx=4)

        # 进度条
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=4)
        self.var_status = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.var_status).pack(anchor="w", padx=12)

        # 日志
        self.log = scrolledtext.ScrolledText(self.root, height=16, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))

    def _entry_pair(self, parent, label, attr, default=""):
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
        try:
            c.cmove_retry = int(self.var_cmove_retry.get().strip() or 1)
        except ValueError:
            c.cmove_retry = 1
        # 夹逼：0=不重试；超过 5 没意义
        if c.cmove_retry < 0:
            c.cmove_retry = 0
        elif c.cmove_retry > 5:
            c.cmove_retry = 5
        try:
            c.cmove_retry_delay = int(self.var_cmove_retry_delay.get().strip() or 3)
        except ValueError:
            c.cmove_retry_delay = 3
        if c.cmove_retry_delay < 1:
            c.cmove_retry_delay = 1
        elif c.cmove_retry_delay > 60:
            c.cmove_retry_delay = 60
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
        try:
            local_port = int(self.var_local_port.get().strip() or 11112)
        except ValueError:
            local_port = None

        self._append_log("正在检测 %s:%d 连通性（含本机接收端口自检）..." % (host, port))
        self.btn_test.config(state="disabled")

        def worker():
            ok, msgs = test_connectivity(host, port, aet, local_port=local_port)
            _ui_queue.put(("conn_result", (ok, msgs)))

        threading.Thread(target=worker, daemon=True).start()

    def _start(self):
        if self.thread and self.thread.is_alive():
            messagebox.showinfo("提示", "已有下载任务在运行中")
            return
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
        for attr in ("pacs_host", "pacs_port", "pacs_aet", "local_aet", "local_port",
                     "excel", "out", "column", "sheet", "key_type",
                     "rate_limit", "pause_every", "pause_seconds",
                     "cmove_timeout", "cfind_timeout", "cmove_retry", "cmove_retry_delay"):
            v = getattr(self, "var_" + attr, None)
            if v is not None:
                try:
                    self.root.nametowidget(v._name).config(state="disabled")
                except Exception:
                    pass
        self.btn_test.config(state="disabled")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_history.config(state="disabled")
        self.var_status.set("下载中...")

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

    def _reset_ui(self, status_text=None):
        """复位按钮/状态（下载线程结束或出错时调用），保证界面不会卡死。"""
        # 恢复所有输入框为可编辑（下载结束/异常退出时）
        for attr in ("pacs_host", "pacs_port", "pacs_aet", "local_aet", "local_port",
                     "excel", "out", "column", "sheet", "key_type",
                     "rate_limit", "pause_every", "pause_seconds",
                     "cmove_timeout", "cfind_timeout", "cmove_retry", "cmove_retry_delay"):
            v = getattr(self, "var_" + attr, None)
            if v is not None:
                try:
                    self.root.nametowidget(v._name).config(state="normal")
                except Exception:
                    pass
        self.btn_test.config(state="normal")
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        if status_text:
            self.var_status.set(status_text)

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