@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src"
set PYTHONDONTWRITEBYTECODE=1
set "PYTHON_EXE=C:\Users\hansiying.1\AppData\Local\Programs\Python\Python312\python.exe"
echo 启动新人价自动化工作台只读 API...
echo 接口地址：http://127.0.0.1:8765
echo 当前 API 只读取配置、输出文件和日志，不会执行查价、提报或下载动作。
if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" -m newcomer_tool.gui_api --host 127.0.0.1 --port 8765
) else (
    python -m newcomer_tool.gui_api --host 127.0.0.1 --port 8765
)
pause
