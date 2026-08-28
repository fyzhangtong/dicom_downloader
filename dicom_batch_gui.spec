# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置：DICOM 影像批量下载工具（Windows exe）

block_cipher = None

a = Analysis(
    ['dicom_batch_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pynetdicom',
        'pynetdicom.presentation',
        'pynetdicom.sop_class',
        'pydicom',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DICOMBatchDownloader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,            # GUI 程序，不显示黑色控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                # 如需图标，可改为 'icon.ico'
)