@echo off
chcp 65001 >nul
echo ==========================================
echo   重启 Excel 图片提取服务
echo ==========================================
echo.

REM 查找并终止所有占用 5000 端口的进程
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :5001') do (
    echo 正在终止进程 PID: %%a
    taskkill /PID %%a /F >nul 2>&1
)

timeout /t 2 >nul

echo.
echo 正在启动服务...
start http://127.0.0.1:5001
python app.py
