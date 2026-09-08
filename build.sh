#!/usr/bin/env bash
# Linux / macOS 打包（生成的是 ELF/Mach-O 可执行文件，不是 exe；exe 请在 Windows 上跑 build.bat）
set -e
python3 -m pip install -r requirements.txt
pyinstaller app.spec --noconfirm --clean
echo "打包完成：dist/WordFormat"
