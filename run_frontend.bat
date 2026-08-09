@echo off
chcp 65001 >nul
cd /d "%~dp0gui_frontend"
set "NODE_HOME=D:\install\node-v24.19.0-win-x64"
if exist "%NODE_HOME%\node.exe" (
    set "PATH=%NODE_HOME%;%PATH%"
)
set "npm_config_cache=%CD%\.npm-cache"
echo 启动新人价自动化工作台前端...
echo 页面地址：http://localhost:8443
echo 如提示 npm 不存在，请先安装或解压 Node.js，并修改本脚本中的 NODE_HOME。
npm run dev
pause
