@echo off
cd /d "%~dp0"
set "RUNTIME=C:\Users\Xiao\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "RUNTIMEW=C:\Users\Xiao\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if exist "%RUNTIMEW%" (
  start "" "%RUNTIMEW%" app.py
) else if exist "%RUNTIME%" (
  start "" "%RUNTIME%" app.py
) else (
  start "" python app.py
)
