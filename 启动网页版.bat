@echo off
echo ========================================
echo   排班自动替换工具 - 网页版
echo ========================================
echo.
echo 正在启动，请稍候...
echo.

pip install flask openpyxl -q

echo 启动成功！
echo 请在浏览器中打开: http://localhost:5000
echo 按 Ctrl+C 可以停止服务
echo.

python app.py
pause
