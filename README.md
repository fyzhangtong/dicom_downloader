# DICOM 影像批量下载工具（Windows 版）使用说明

## 一、工具能做什么

读取一个 Excel 表格中的一列标识，通过 **DICOM Q/R 协议** 从医院 PACS 批量拉取影像到本地目录。
- 所有连接参数可视化配置
- 带「检测连通」功能（TCP 连通 + DICOM C-Echo 验证）
- 支持两种查询键模式（见下）

## 二、⚠️ 关键：查询键必须是下面两种之一

众阳 PACS 的 C-Find/C-Move **只支持**两个查询键：

| 可用的键 | 你 Excel 里对应的列 | 是否可用 |
|---------|-------------------|---------|
| **StudyInstanceUID** (0020,000D) | 形如 `1.2.840.xxx...` 的列 | ✅ 精确到一次检查 |
| **patientId / 病人ID** (0010,0020) | 「病人ID」列（`1554033`） | ✅ 可用，但一病人可能对应多次检查 |

> ⚠️ **「影像号」这一列不能用！** 医院导出的「影像号」是纯数字流水号（如 `1006567`），
> 它在 DICOM 里是「检查号 AccessionNumber」，**不在众阳支持的查询键里**，
> 直接用它查会查不到/报错。

**两种模式：**

| 模式 | 输入列 | 下载结果 |
|------|--------|---------|
| `StudyInstanceUID` | 影像号对应的 UID 列 | 每个 UID 精确拉一次检查 |
| `病人ID(patientId)` | 病人ID 列 | 先查该病人的**全部检查**，再逐个拉取（可能多拉） |

**推荐**：让医院/众阳工程师在导数据时补一列 **StudyInstanceUID**（或提供「影像号→StudyInstanceUID」映射表），再按 StudyInstanceUID 精确下载。

## 三、交付文件说明

| 文件 | 作用 |
|------|------|
| `dicom_batch_gui.py` | 程序源码（Python） |
| `dicom_batch_gui.spec` | PyInstaller 打包配置 |
| `build_exe.bat` | 一键打包脚本（双击运行） |
| `requirements.txt` | 依赖清单 |
| `dist/DICOMBatchDownloader.exe` | 打包产物（运行脚本后生成） |

## 四、生成 exe（只需做一次）

1. 安装 **Python 3.9+**（官网 https://www.python.org/downloads/ ，安装时务必勾选 **Add python.exe to PATH**）。
2. 把本目录整个拷贝到 Windows 电脑。
3. 双击 `build_exe.bat`，等待打包完成。
4. 到 `dist` 目录双击 `DICOMBatchDownloader.exe` 即可运行。

## 五、运行前需要向医院（众阳）工程师确认/索取

| 项 | 填写位置 | 说明 |
|----|----------|------|
| PACS 前置机 IP | 「PACS IP」 | 医院影像归档前置机地址 |
| PACS 端口 | 「PACS 端口」 | 通常 104 或自定义 |
| PACS AE Title | 「PACS AE Title」 | 前置机的 DICOM AE Title |
| 本机 AE Title | 「本机 AE Title」 | 你自己定，如 `CLIENT_AET` |
| 本机接收端口 | 「本机接收端口」 | 你自己定，如 `11112` |
| **节点注册** | —— | 把「本机 AE Title + 本机 IP + 接收端口」交给工程师，注册进前置机并开通 Q/R 权限 |
| **网络** | —— | 运行电脑需能访问前置机 IP:端口（内网 / VPN / 白名单） |
| **StudyInstanceUID 列** | —— | 若想精确下载，需工程师在导数据时补这一列 |

## 六、使用步骤

1. 填好连接参数，点「**检测连通**」—— 先 TCP 连通检查，再 DICOM C-Echo 验证 AE Title 是否有效。
2. 「**浏览...**」选择 Excel 文件。
3. 选「**查询键类型**」：
   - `StudyInstanceUID`：在「关键列名」填 UID 所在列名
   - `病人ID(patientId)`：在「关键列名」填「病人ID」
4. 选择输出目录。
5. 点「**开始下载**」；影像保存到 `<输出目录>/<StudyInstanceUID>/<...>.dcm`。

## 七、常见问题

| 现象 | 原因/处理 |
|------|-----------|
| 「检测连通」TCP 失败 | IP/端口错误、网络不通、需 VPN/白名单 |
| TCP 成功但 C-Echo 失败 | AE Title 不正确，或前置机未放行你的 IP |
| 下载报 `Move Destination unknown` | 本机 AE Title 未注册到前置机 |
| 报 `Out of Resources` | 查询条件不匹配 / 该 UID 不存在 |
| 用「影像号」查几乎全失败 | 影像号不是有效查询键，改用 StudyInstanceUID 或病人ID |
| exe 被 Windows 拦截 | 右键属性勾选"解除锁定"，或允许"仍要运行" |