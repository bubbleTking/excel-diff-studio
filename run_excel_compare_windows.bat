@echo off
setlocal

cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%~dp0excel_compare_portable.py"
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    python "%~dp0excel_compare_portable.py"
  ) else (
    echo Python was not found.
    echo Please install Python 3 from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
  )
)

echo.
pause
