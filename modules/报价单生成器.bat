@echo off

REM Check if 'pip' is available
where pip > nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo 'pip' command not found. Please install Python with pip.
    pause
    exit /b
)

REM Install required dependencies if not already installed
pip install -r requirements.txt

REM Check if the installation was successful
IF %ERRORLEVEL% NEQ 0 (
    echo Failed to install required dependencies.
    pause
    exit /b
)

REM Run your Python script
python ".\quotation_selector_gui.py"
pause