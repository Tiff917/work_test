@echo off
cd /d "%~dp0"
set "RUNTIME=C:\Users\Xiao\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%RUNTIME%" (
  "%RUNTIME%" desktop_app.py
) else (
  python desktop_app.py
)
if errorlevel 1 pause
