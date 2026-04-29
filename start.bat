@echo off
chcp 65001 >nul
echo ==========================================
echo   Excel 图片提取工具
echo ==========================================
echo.
echo 正在启动服务...
echo 访问地址: http://localhost:5001
echo.
echo 按 Ctrl+C 停止服务
echo.

REM 尝试自动打开浏览器
timeout /t 1 >nul
start http://localhost:5001

REM 启动 Flask
python app.py
