@echo off
REM Check if Python is installed
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed. Please install Python and try again.
    exit /b 1
)

REM Create a virtual environment
python -m venv .venv

REM Install dependencies from requirements.txt
IF EXIST requirements.txt (
    .venv\Scripts\activate
    pip install -r requirements.txt
    echo Virtual environment created and dependencies installed successfully.
) ELSE (
    echo requirements.txt not found. Virtual environment created without installing dependencies.
)

REM Deactivate the virtual environment
deactivate

echo Done.
