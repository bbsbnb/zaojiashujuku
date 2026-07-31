@echo off
chcp 65001 >nul
cd /d %~dp0

echo 正在启动 AI 人材机价格助手 MVP...
echo.
echo 浏览器地址：http://127.0.0.1:8765/
echo.
echo 如果页面打不开，请确认没有旧服务占用 8765 端口。
echo 关闭本窗口即可停止本地服务。
echo.

python -m cost_ai_mvp.web_app

echo.
echo 服务已停止。
pause
