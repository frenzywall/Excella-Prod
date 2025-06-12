@echo off
setlocal enabledelayedexpansion

echo ================================================
echo         Building Excella Executable
echo ================================================
echo.

:: Check if build mode is passed
if "%~1"=="" (
    echo [ERROR] Build mode not specified.
    echo Usage: build_exe.bat [onedir | onefile]
    exit /b 1
)

set BUILD_TYPE=%~1

if /I NOT "%BUILD_TYPE%"=="onedir" if /I NOT "%BUILD_TYPE%"=="onefile" (
    echo [ERROR] Invalid build mode: %BUILD_TYPE%
    echo Valid options: onedir or onefile
    exit /b 1
)

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.8 or higher.
    exit /b 1
)

:: Install dependencies
echo [INFO] Installing required Python packages...
pip install -r requirements.txt >nul 2>&1
pip install pyinstaller >nul 2>&1

:: Clean previous build
echo [INFO] Cleaning previous build directories...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
if exist Excella.spec del /q Excella.spec

:: Set PyInstaller flags
set BUILD_FLAGS=--noconfirm --windowed --icon=icon.ico --name="Excella" --add-data="icon.ico;."

if /I "%BUILD_TYPE%"=="onefile" (
    set BUILD_FLAGS=!BUILD_FLAGS! --onefile
)

:: Build with PyInstaller
echo [INFO] Building executable in %BUILD_TYPE% mode...
python -m PyInstaller !BUILD_FLAGS! app.py

echo.
echo ================================================
echo [SUCCESS] Build complete!
if /I "%BUILD_TYPE%"=="onedir" (
    echo Check dist\Excella\ for the final executable.
) else (
    echo Check dist\Excella.exe for the final executable.
)
echo ================================================
pause
