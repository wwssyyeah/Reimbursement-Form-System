@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM 优先用 node 启动本地服务（完全离线可用 lib/ 引擎）
where node >nul 2>nul
if %errorlevel%==0 (
  node server.js
  goto :done
)

REM 退回 python
where python >nul 2>nul
if %errorlevel%==0 (
  echo 未检测到 node，使用 python 启动本地服务（端口 8080）...
  start "" "http://localhost:8080/index.html"
  python -m http.server 8080
  goto :done
)

echo 未检测到 node 或 python，无法启动本地服务。
echo 请安装其一后重试；或直接在浏览器中双击 index.html（仅在线 CDN 模式）。
pause
:done
