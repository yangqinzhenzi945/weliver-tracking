@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 Weliver 网页查询工具...
py -m streamlit run web_app.py --server.address 127.0.0.1
if errorlevel 1 (
  echo 启动失败。请先双击「安装环境.bat」安装依赖。
  pause
)
