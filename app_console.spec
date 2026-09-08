# -*- mode: python ; coding: utf-8 -*-
"""
调试版打包配置（带黑色控制台窗口）。
当 dist\\WordFormat.exe 双击后无任何反应 / 一闪而过时，用这个打包一份能看到报错的版本：

    pyinstaller app_console.spec --noconfirm --clean

产物：dist\\WordFormat_Debug.exe
"""

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules("docx")
hiddenimports += ["lxml._elementpath", "profile"]

block_cipher = None

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
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
    name="WordFormat_Debug",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,           # <-- 保留控制台，报错看得见
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
