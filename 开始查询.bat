@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Weliver 出库查询
echo.
echo 请将 Excel 放在此文件夹中。
set "EXCEL="
set /p "EXCEL=请输入 Excel 文件名（例如 cargoboard待查询出库.xlsx）: "
set "EXCEL=%EXCEL:"=%"
if not defined EXCEL (echo 未输入文件名 & pause & exit /b 1)
if not exist "%EXCEL%" (echo 文件不存在：%EXCEL% & pause & exit /b 1)
py "weliver_tracking.py" "%EXCEL%"
echo.
pause
