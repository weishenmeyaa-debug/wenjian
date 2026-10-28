@echo off
setlocal
title Word Document Format Checker
cd /d "%~dp0"

set "PYTHON="
where py >nul 2>&1
if not errorlevel 1 goto :use_py
where python >nul 2>&1
if not errorlevel 1 goto :use_python
where python3 >nul 2>&1
if not errorlevel 1 goto :use_python3
goto :no_python

:use_py
set "PYTHON=py -3"
goto :found

:use_python
set "PYTHON=python"
goto :found

:use_python3
set "PYTHON=python3"
goto :found

:no_python
echo [ERROR] Python not found. Please install Python 3 first.
echo        Download: https://www.python.org/downloads/
echo.
pause
exit /b 1

:found
if exist "venv\Scripts\python.exe" goto :have_venv
echo [SETUP] First run: creating virtual environment...
%PYTHON% -m venv venv
if errorlevel 1 goto :venv_fail

:have_venv
call venv\Scripts\activate.bat

python -c "import flask, docx" >nul 2>&1
if not errorlevel 1 goto :deps_ok

echo [SETUP] Installing dependencies, needs internet, about 1 minute...
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if not errorlevel 1 goto :deps_check
echo [INFO] Mirror failed, retrying with default index...
python -m pip install -r requirements.txt

:deps_check
python -c "import flask, docx" >nul 2>&1
if not errorlevel 1 goto :deps_ok
echo [ERROR] Dependency installation failed. Check your network and retry.
pause
exit /b 1

:venv_fail
echo [ERROR] Failed to create virtual environment.
pause
exit /b 1

:deps_ok
set "PORT=5001"
:check_port
netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul
if errorlevel 1 goto :port_ok
set /a PORT+=1
goto :check_port

:port_ok
echo.
echo [OK] Format Checker is starting. Browser will open automatically.
echo     If not, open manually: http://127.0.0.1:%PORT%
echo     Close this window to stop the service.
echo.

start "" /b cmd /c "timeout /t 2 >nul & start http://127.0.0.1:%PORT%"

python app.py --port %PORT%
pause
