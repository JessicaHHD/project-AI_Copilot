@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0src
set "PYTHON_EXE=C:\Users\hansiying.1\AppData\Local\Programs\Python\Python312\python.exe"
if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" -m newcomer_tool.main
) else (
    python -m newcomer_tool.main
)
pause
