@echo off

setlocal
set PY_CMD=python

REM Check if python is available
python --version >nul 2>&1
if errorlevel 1 (
    REM Try py launcher
    py --version >nul 2>&1
    if errorlevel 1 (
        echo Python is not installed. Attempting to install Python using winget...
        winget --version >nul 2>&1
        if errorlevel 1 (
            echo Error: Python is not installed and winget is not available to install it.
            echo Please install Python manually and try again.
            pause
            exit /b 1
        )
        winget install -e --id Python.Python.3
        if errorlevel 1 (
            echo Error: Failed to install Python using winget.
            pause
            exit /b 1
        )
        echo Python installation complete. Checking again...
        python --version >nul 2>&1
        if errorlevel 1 (
            py --version >nul 2>&1
            if errorlevel 1 (
                echo Error: Python installation did not succeed. Please install Python manually.
                pause
                exit /b 1
            ) else (
                set PY_CMD=py
            )
        ) else (
            set PY_CMD=python
        )
    ) else (
        set PY_CMD=py
    )
) else (
    set PY_CMD=python
)

REM Check if required packages are installed
echo Checking dependencies...
%PY_CMD% -c "import pandas, docx, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    %PY_CMD% -m pip install pandas python-docx openpyxl
    if errorlevel 1 (
        echo Error: Failed to install required packages
        pause
        exit /b 1
    )
)

REM Run the daily summary script in interactive mode
echo Starting Daily Summary Generator...
echo.
%PY_CMD% daily_summary_generator.py --interactive

echo.
echo Press any key to exit...
pause >nul 
endlocal 