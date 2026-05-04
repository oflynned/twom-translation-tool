@echo off
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
    echo Python 3.9+ required. Download from https://www.python.org/
    echo Tick "Add Python to PATH" during install.
    pause & exit /b 1
)
python twom_translator.py
if errorlevel 1 ( echo. & echo Error - see above & pause )
