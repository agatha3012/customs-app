@echo off
cd /d %~dp0
echo ========================================
echo   报关辅助系统 Customs App v2.0
echo   启动中...
echo ========================================
echo.
echo 正在启动服务器，浏览器将自动打开...
echo 如未自动打开，请手动访问: http://localhost:3000
echo.
node server.js
pause
