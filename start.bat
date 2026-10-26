@echo off
chcp 65001 >nul 2>&1
title 规检 - Word 文档格式检查
cd /d "%~dp0"

:: 检测 Python
where python >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    goto :found
)
where python3 >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python3
    goto :found
)
where py >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py -3
    goto :found
)

echo ❌ 未找到 Python，请先安装：
echo    打开 Microsoft Store 搜索 "Python 3.12" 安装即可
echo    或去 https://www.python.org/downloads/ 下载
echo.
pause
exit /b 1

:found
:: 创建虚拟环境
if not exist "venv" (
    echo 📦 首次运行，正在安装依赖...
    %PYTHON% -m venv venv
    call venv\Scripts\activate.bat
    pip install -r requirements.txt -q
) else (
    call venv\Scripts\activate.bat
)

:: 从 5001 开始寻找可用端口，避免误打开其他程序
set PORT=5001
:check_port
netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul
if %errorlevel% equ 0 (
    set /a PORT+=1
    goto :check_port
)

:: 启动
echo.
echo ✅ 规检启动成功！浏览器将自动打开...
echo    如果没有自动打开，请访问: http://127.0.0.1:%PORT%
echo    关闭此窗口即可停止程序
echo.

:: 延迟打开浏览器
start "" /b cmd /c "timeout /t 2 >nul & start http://127.0.0.1:%PORT%"

:: 运行 Flask
python app.py --port %PORT%
pause
