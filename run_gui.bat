@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONPATH=%~dp0src
set PYTHONDONTWRITEBYTECODE=1
set NEWCOMER_TOOL_CONFIG=%~dp0config.test.yaml
set "PYTHON_EXE=C:\Users\hansiying.1\AppData\Local\Programs\Python\Python312\python.exe"
if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" -m streamlit run src\newcomer_tool\gui_mvp.py
) else (
    python -m streamlit run src\newcomer_tool\gui_mvp.py
)
pause
