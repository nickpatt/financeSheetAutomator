@echo off
echo Daily Summary Generator - Build Script
echo ======================================
echo.

REM Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    py --version >nul 2>&1
    if errorlevel 1 (
        echo Error: Python is not installed!
        pause
        exit /b 1
    ) else (
        set PY_CMD=py
    )
) else (
    set PY_CMD=python
)

echo Using Python: %PY_CMD%
echo.

REM Install/upgrade pip
echo Installing/upgrading pip...
%PY_CMD% -m pip install --upgrade pip

REM Install requirements
echo Installing requirements...
%PY_CMD% -m pip install -r requirements.txt

REM Run the build script
echo Building executable...
%PY_CMD% build_exe.py

echo.
echo Build process complete!
echo If successful, you should now have DailySummaryGenerator.exe
echo.
pause 