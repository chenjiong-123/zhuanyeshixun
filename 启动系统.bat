@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 goto use_py
where python >nul 2>nul
if not errorlevel 1 goto use_python
echo Python 3.9 or newer is required.
echo Install Python and enable Add Python to PATH.
pause
exit /b 1

:use_py
py -3 app.py
goto finished

:use_python
python app.py

:finished
if errorlevel 1 (
  echo.
  echo The application stopped because an error occurred.
  pause
)
endlocal
