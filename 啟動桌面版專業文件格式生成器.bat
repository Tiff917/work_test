@echo off
cd /d "%~dp0"
set "RUNTIME=C:\Users\Xiao\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "RUNTIMEW=C:\Users\Xiao\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if exist "%RUNTIMEW%" (
  start "" "%RUNTIMEW%" desktop_app.py
) else if exist "%RUNTIME%" (
  start "" "%RUNTIME%" desktop_app.py
) else (
  start "" python desktop_app.py
)
