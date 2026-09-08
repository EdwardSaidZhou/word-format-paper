@echo off
setlocal
REM ============================================================
REM  WordFormat - Build EXE (Windows)
REM  NOTE: keep this file pure-ASCII with CRLF line endings.
REM        Do NOT save as UTF-8, do NOT add "chcp 65001".
REM        cmd.exe reads .bat byte-by-byte and will abort.
REM ============================================================

cd /d "%~dp0"

echo ============================================================
echo   WordFormat - Build EXE
echo ============================================================
echo.

REM ---------- 1. locate Python (py -3 / python / python3) ----------
set "PY=py -3"
%PY% -c "import sys" >nul 2>nul || set "PY=python"
%PY% -c "import sys" >nul 2>nul || set "PY=python3"
%PY% -c "import sys" >nul 2>nul || goto NOPYTHON

echo [1/4] Python found:
%PY% -c "import sys,platform;print('      '+sys.executable+'  ('+platform.python_version()+')')"
echo.

REM ---------- 2. version check (need 3.8+) ----------
%PY% -c "import sys;sys.exit(0 if sys.version_info>=(3,8) else 1)" >nul 2>nul
if errorlevel 1 goto OLDVER

REM ---------- 3. install dependencies ----------
echo [2/4] Upgrading pip ...
%PY% -m pip install --upgrade pip
echo.

echo [3/4] Installing requirements (python-docx, pyinstaller) ...
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto ERR
echo.

REM ---------- 4. build ----------
echo [4/4] Building with PyInstaller, please wait ...
%PY% -m PyInstaller app.spec --noconfirm --clean
if errorlevel 1 goto ERR

echo.
echo ============================================================
echo   BUILD OK  -^>  dist\WordFormat.exe
echo   Copy that single file anywhere, no Python needed.
echo ============================================================
echo.
pause
goto :eof

:NOPYTHON
echo.
echo [ERROR] Python was not found.
echo   - Install Python 3.8+ from https://www.python.org/downloads/
echo   - IMPORTANT: tick "Add python.exe to PATH" during install.
echo   - If you prefer "py -3", make sure the Python Launcher is installed.
echo.
pause
goto :eof

:OLDVER
echo.
echo [ERROR] Python 3.8 or newer is required.
%PY% -c "import sys;print('   Current version:',sys.version)"
echo.
pause
goto :eof

:ERR
echo.
echo [ERROR] Build failed. Scroll up for the red error message.
echo   Common causes:
echo     1. No internet / pip blocked  -> use a mirror:
echo        python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
echo     2. Antivirus quarantined PyInstaller files -> add folder to whitelist
echo     3. Missing files -> run this .bat inside the app.py folder
echo.
pause
goto :eof
