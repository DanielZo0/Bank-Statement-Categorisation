@echo off
REM Bank Statement Categorization Tool Launcher
REM Double-click this file to start the categorization tool

title Bank Statement Categorization Tool

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from python.org
    echo.
    pause
    exit /b 1
)

REM Run the categorization script
python categorize_statement.py

REM Keep window open on error
if errorlevel 1 (
    echo.
    echo An error occurred. Press any key to exit...
    pause >nul
)

