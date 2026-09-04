# DICOM 影像批量下载工具（Windows 可视化版）

读取 Excel 表格中的一列标识（StudyInstanceUID 或 病人ID），通过 **DICOM Q/R 协议**（C-Find + C-Move）
从医院 PACS 批量拉取影像到本地目录，所有连接参数可视化配置。

> 协议依据：众阳云开放服务 API「2.8.1.3 获取检查图像」—— DICOM 3.0 Q/R
> C-Move 入参：StudyInstanceUID (0020,000D)

---

## 一、⚠️ 关键：查询键必须是下面两种之一

众阳 PACS 的 C-Find / C-Move **只支持**两个查询键：

| 可用的键 | DICOM 标签 | 你 Excel 里对应的列 | 是否可用 |
|---------|-----------|------------------|---------|
| **StudyInstanceUID** | (0020,000D) | 形如 `1.2.840.xxx...` 的列 | ✅ 精确到一次检查 |
| **patientId / 病人ID** | (0010,0020) | 「病人ID」列（如 `1554033`） | ✅ 一病人可能对应多次检查 |

> ⚠️ **「影像号」这一列不能直接当作查询键！** 医院导出的「影像号」是纯数字流水号（如 `1006567`），
> 它在 DICOM 里是「检查号 AccessionNumber」，**不在众阳支持的查询键里**。
>
> 本工具支持「影像号」列名仅作为默认占位列名，**实际查询必须用 StudyInstanceUID 或病人ID**。
> 推荐让医院/众阳工程师在导数据时补一列 **StudyInstanceUID**（或提供「影像号 → StudyInstanceUID」映射表），再按 UID 精确下载。

### 两种查询模式

| 模式 | 输入列 | 下载结果 |
|------|--------|---------|
| `StudyInstanceUID(影像号UID)` | UID 所在列（默认列名「影像号」可改） | 每个 UID 精确拉一次检查 |
| `病人ID(patientId)` | 病人ID 列 | 先 C-Find 查该病人**全部检查**，再逐个拉取（可能多拉） |

---

## 二、主要功能

- 可视化配置 PACS IP / 端口 / AE Title、本机 AE Title / 接收端口
- **检测连通（双向）**：正向 TCP + DICOM C-Echo；反向自检（经本机外部 IP 连自己的接收服务并 C-Echo）+ 前置机反向探测（发空 C-Move 观察前置机是否主动连入）
- **查询键两种模式**：StudyInstanceUID 精确拉取 / 病人ID 展开后逐个拉取
- **串行下载**（稳态，避免给前置机造成并发压力）
- **间隙暂停**：每下载 N 个检查后暂停 M 秒，给 PACS 喘息
- **限速**：按 KB/s 节流 C-STORE 落盘速率（限速通过延迟应答实现，过低会触发单文件 2 秒封顶保护，避免前置机判超时）
- **进度显示**：总进度条 + 当前检查的「影像数 X / N」+ 已耗时
- **下载报告**：每次下载生成 `下载报告_YYYYMMDD_HHMMSS.csv`（UTF-8 with BOM）
- **手动停止**：中断当前 association，剩余任务不再处理
- **优雅退出**：关闭主窗口时主动 shutdown Store SCP，避免 Windows 端口 TIME_WAIT

---

## 三、交付文件

| 文件 | 作用 |
|------|------|
| `dicom_batch_gui.py` | 程序源码（Python 3.9+） |
| `dicom_batch_gui.spec` | PyInstaller 打包配置 |
| `build_exe.bat` | Windows 一键打包脚本（双击运行） |
| `requirements.txt` | Python 依赖清单 |
| `config.json` | 运行后自动保存/加载的界面配置（首次运行后生成） |
| `dist/DICOMBatchDownloader.exe` | 打包产物（运行 `build_exe.bat` 后生成） |

---

## 四、生成 exe（仅打包者需要）

1. 安装 **Python 3.9+**（https://www.python.org/downloads/ ，安装时勾选 **Add python.exe to PATH**）
2. 把整个目录拷贝到 Windows 电脑
3. 双击 `build_exe.bat`，等待打包完成
4. 到 `dist` 目录双击 `DICOMBatchDownloader.exe` 即可运行

---

## 五、运行前需向医院（众阳）确认/索取

| 项 | 填写位置 | 说明 |
|----|----------|------|
| PACS 前置机 IP | 「PACS IP」 | 医院影像归档前置机地址 |
| PACS 端口 | 「PACS 端口」 | 通常 104 或自定义 |
| PACS AE Title | 「PACS AE Title」 | 前置机的 DICOM AE Title |
| 本机 AE Title | 「本机 AE Title」 | 自定义，如 `CLIENT_AET`（1~16 位 ASCII，不能含空格/中文） |
| 本机接收端口 | 「本机接收端口」 | 自定义，如 `11112`（**不能落在 Hyper-V/WSL 排除范围**） |
| **节点注册** | —— | 把「本机 AE Title + 本机 IP + 接收端口」交给工程师，注册进前置机并开通 Q/R 权限 |
| **网络** | —— | 运行电脑需能访问前置机 IP:端口（内网 / VPN / 白名单） |
| **StudyInstanceUID 列** | —— | 想精确下载必须由工程师在导出数据时补这一列 |

> ℹ️ 本机 IP 会显示在「检测连通」结果中，请把它（而非 127.0.0.1）报给工程师注册。

---

## 六、使用步骤

1. **填好 PACS 与本机参数**（首次可点「保存配置」记到 `config.json`）
2. 点「**检测连通**」—— 双向检测：正向 TCP + C-Echo；反向自检 + 前置机反向探测，任一失败都会给出可操作建议
3. **「浏览...」选择 Excel**（`.xlsx` / `.xlsm`；Excel 需先关闭，否则会报「文件被占用」）
4. **选「查询键类型」**，并在「关键列名」填写对应列名（默认「影像号」，可改为 `StudyInstanceUID` / `病人ID` / `A` / `1` 等）
5. **选择输出目录**（下载报告 CSV 会写入该目录）
6. 高级参数可保持默认：
   - C-Move 超时 300 秒 / C-Find 超时 60 秒
   - 限速 0 = 不限；每 0 个检查暂停 = 不暂停
   - 「诊断日志」默认关闭：只输出关键日志；勾选后输出 PDU/协商/Pending 细节，排查问题时使用
7. 点「**开始下载**」；日志实时输出每个 Study 的 C-Move 最终状态与结果
8. 下载结束自动弹出汇总（成功 / 失败 / 共），并生成报告 CSV

> 「**下载记录**」按钮可查看本次明细，「**导出 CSV**」可保存到任意位置。

---

## 七、下载报告

- 每次下载结束会在输出目录生成 `下载报告_YYYYMMDD_HHMMSS.csv`（UTF-8 with BOM），包含序号、查询键、StudyInstanceUID、结果、文件数、说明、时间。
- 失败/停止的 Study 不会自动跳过，重新运行即会再次拉取。

---

## 八、常见问题

| 现象 | 原因 / 处理 |
|------|-----------|
| 「检测连通」TCP 失败 | IP/端口错误、网络不通、需 VPN/白名单 |
| TCP 成功但 C-Echo 失败 | AE Title 不正确，或前置机未放行本机 IP |
| 报 `Move Destination unknown` (0xA801) | 本机 AE Title 未注册到前置机 |
| 报 `Out of Resources` (0xA700) | 查询条件不匹配 / 该 UID 不存在 |
| 「检测连通」提示「端口无法绑定」 | 端口被占用 / 落在 Hyper-V/WSL 排除范围 / 被防火墙拦截。提示中给出 `netsh` 命令与处理建议 |
| C-Move 状态成功但 0 个文件 | 前置机是「先返回成功、异步推影像」特性，工具会多等 30 秒；若仍 0 文件，大概率是本机注册 IP/端口与前置机侧不一致 |
| C-Move 状态为 0xB000（部分失败） | 少数影像子操作失败，工具按**本地已落盘文件数**记为成功；如有个别缺失，重跑该 Study 即可全量补拉 |
| 用「影像号」查几乎全失败 | 影像号不是有效查询键，改用 StudyInstanceUID 或病人ID |
| Excel 报「文件被占用」 | 关闭 Excel 后再试 |
| exe 被 Windows 拦截 | 右键属性 → 勾选「解除锁定」，或 SmartScreen 选「仍要运行」 |
| 关闭主窗口后端口残留 TIME_WAIT | 工具已主动 shutdown Store SCP；若仍有残留，等 30~60 秒或换个端口 |

---

## 九、配置项默认值（界面可改，保存到 `config.json`）

| 项 | 默认 | 说明 |
|----|------|------|
| PACS 端口 | 104 | 常用 DICOM 端口 |
| 本机 AE Title | `MYAET` | 1~16 位 ASCII |
| 本机接收端口 | 11112 | 必须提前向医院注册 |
| 关键列名 | `影像号` | 支持列名 / 列字母 / 序号 |
| 限速 (KB/s) | 0 | 0 = 不限速 |
| 每 N 个检查暂停 | 0 | 0 = 不暂停 |
| 暂停秒数 | 30 | 仅在「每 N 个」> 0 时生效 |
| C-Move 超时(秒) | 300 | 单次 C-Move 最大等待 |
| C-Find 超时(秒) | 60 | 仅在「按病人ID」模式生效 |
| 诊断日志 | 关 | 勾选后输出 PDU/协商/Pending 细节 |

---

## 十、技术栈与依赖

- Python 3.9+
- pynetdicom ≥ 3.0（DICOM 协议栈；使用 `send_c_move` / `send_c_find` 位置参数）
- pydicom ≥ 3.0
- openpyxl ≥ 3.1（读取 .xlsx / .xlsm）
- PyInstaller ≥ 6.0（仅打包时需要）
- tkinter（Python 内置，GUI）

依赖见 `requirements.txt`。
