@echo off
setlocal
cd /d "%~dp0"
py -3.12 -m venv .venv-build
if errorlevel 1 goto failed
.venv-build\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto failed
.venv-build\Scripts\python.exe scripts\build_windows.py
if errorlevel 1 goto failed
.venv-build\Scripts\python.exe scripts\verify_windows.py
if errorlevel 1 goto failed
echo Build complete: dist\ReadQRToxlsx.exe
pause
exit /b 0
:failed
echo Build failed. See the error above. Python 3.12 x64 is required.
pause
exit /b 1
