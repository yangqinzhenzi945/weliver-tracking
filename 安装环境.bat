@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m pip install -r requirements.txt
if errorlevel 1 (echo 安装依赖失败 & pause & exit /b 1)
py -m playwright install chromium
if errorlevel 1 (echo 安装浏览器失败 & pause & exit /b 1)
echo 安装完成
pause
