@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "NODE_HOME=D:\install\node-v24.19.0-win-x64"
set "FRONTEND_DIR=%~dp0gui_frontend"

echo 新人价自动化工作台总入口
echo.
echo 本入口只启动只读 API 和 React 前端，不执行筛品、查价、Outlook 下载、后台提报或审核回填。
echo.
if exist "%NODE_HOME%\node.exe" (
    set "PATH=%NODE_HOME%;%PATH%"
) else (
    echo [错误] 未找到 Node.js：%NODE_HOME%\node.exe
    echo 请修改本脚本中的 NODE_HOME，指向你的 Node.js 解压目录。
    pause
    exit /b 1
)

if not exist "%FRONTEND_DIR%\node_modules" (
    echo [错误] 前端依赖尚未安装：%FRONTEND_DIR%\node_modules
    echo 请先进入 gui_frontend 运行 npm install。
    pause
    exit /b 1
)

echo [1/2] 启动只读 API 窗口...
start "新人价工作台 API" cmd /k call "%~dp0run_gui_api.bat"

echo [2/2] 启动 React 前端窗口...
start "新人价工作台前端" cmd /k call "%~dp0run_frontend.bat"

echo.
echo 工作台正在启动，请稍等几秒后打开：
echo 前端页面：http://localhost:8443/
echo API 健康检查：http://127.0.0.1:8765/api/health
echo.
echo 如果端口已被占用，请先关闭之前打开的 API 或前端窗口。
pause
