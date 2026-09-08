@echo off
setlocal
REM ============================================================
REM  WordFormat - run the GUI without building an exe
REM  NOTE: keep this file pure-ASCII with CRLF line endings.
REM ============================================================

cd /d "%~dp0"

set "PY=py -3"
%PY% -c "import sys" >nul 2>nul || set "PY=python"
%PY% -c "import sys" >nul 2>nul || set "PY=python3"
%PY% -c "import sys" >nul 2>nul || goto NOPYTHON

%PY% -c "import docx" >nul 2>nul
if errorlevel 1 (
    echo python-docx not installed, installing now ...
    %PY% -m pip install -r requirements.txt
    if errorlevel 1 goto ERR
)

%PY% app.py --gui
if errorlevel 1 goto ERR
goto :eof

:NOPYTHON
echo.
echo [ERROR] Python was not found. Install Python 3.8+ from
echo         https://www.python.org/downloads/
echo         and tick "Add python.exe to PATH".
echo.
pause
goto :eof

:ERR
echo.
echo [ERROR] Failed to start. Scroll up for details.
echo         Crash details (if any) are written to crash.log
echo.
pause
goto :eof
