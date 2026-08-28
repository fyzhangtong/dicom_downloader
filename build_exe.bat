@echo off
setlocal
cd /d %~dp0

echo ==============================================
echo   DICOM 影像批量下载工具 - Windows 打包脚本
echo ==============================================
echo.

rem ---- 1) 检测 Python（优先 python，其次 py 启动器）----
set "PY="
python --version >nul 2>&1
if not errorlevel 1 set "PY=python"
if not defined PY (
    py --version >nul 2>&1
    if not errorlevel 1 set "PY=py"
)

if not defined PY (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+ 并勾选 "Add python.exe to PATH"
    echo        下载地址：https://www.python.org/downloads/windows/
    echo        安装时务必勾选底部的 "Add python to PATH" 复选框！
    pause
    exit /b 1
)

echo [1/3] 检测到 Python：
%PY% --version

echo.
echo [2/3] 安装依赖（首次可能需要几分钟）...
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络后重试
    pause
    exit /b 1
)

echo.
echo [3/3] 打包为 exe（PyInstaller）...
%PY% -m PyInstaller --clean --noconfirm dicom_batch_gui.spec
if errorlevel 1 (
    echo [错误] 打包失败，请把本窗口截图发给技术人员
    pause
    exit /b 1
)

echo.
echo ==============================================
echo   打包完成！可执行文件位于 dist 目录：
echo ==============================================
dir /b dist\*.exe
echo.
echo 双击 dist\DICOMBatchDownloader.exe 即可运行。
echo 注意：exe 首次运行如需联网，请在 Windows 防火墙弹窗中允许访问。
pause
