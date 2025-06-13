@echo off
setlocal enabledelayedexpansion

echo ================================================
echo         Building Excella Executable
echo ================================================
echo.

:: Check if build mode is passed
if "%~1"=="" (
    echo [ERROR] Build mode not specified.
    echo Usage: build_exe.bat [onedir ^| onefile]
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

:: Check if dependencies are already installed to avoid reinstalling
echo [INFO] Checking dependencies...
python -c "import PyInstaller" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [INFO] Installing PyInstaller...
    pip install pyinstaller --quiet --disable-pip-version-check
) else (
    echo [INFO] PyInstaller already installed, skipping...
)

:: Check requirements.txt dependencies
if exist requirements.txt (
    echo [INFO] Installing requirements...
    pip install -r requirements.txt --quiet --disable-pip-version-check --upgrade-strategy only-if-needed
) else (
    echo [INFO] No requirements.txt found, skipping...
)

:: Clean previous build (only if they exist)
echo [INFO] Cleaning previous build directories...
if exist build (
    echo [INFO] Removing build directory...
    rmdir /s /q build
)
if exist dist (
    echo [INFO] Removing dist directory...
    rmdir /s /q dist
)
for /d %%i in (*__pycache__*) do rmdir /s /q "%%i" 2>nul
if exist Excella.spec del /q Excella.spec

:: Check if icon exists before adding it
set ICON_FLAG=
if exist icon.ico (
    set ICON_FLAG=--icon=icon.ico --add-data="icon.ico;."
    echo [INFO] Using icon.ico
) else (
    echo [INFO] No icon.ico found, building without icon
)

:: Set PyInstaller flags with performance optimizations
set BUILD_FLAGS=--noconfirm --windowed --name="Excella" !ICON_FLAG! --noupx --exclude-module=tkinter --exclude-module=matplotlib --clean

:: Add onefile flag if specified
if /I "%BUILD_TYPE%"=="onefile" (
    set BUILD_FLAGS=!BUILD_FLAGS! --onefile
)

:: Check if spec file exists from previous run to enable incremental builds
if exist Excella.spec (
    echo [INFO] Using existing spec file for faster incremental build...
    python -m PyInstaller Excella.spec --noconfirm
) else (
    echo [INFO] Building executable in %BUILD_TYPE% mode...
    python -m PyInstaller !BUILD_FLAGS! app.py
)

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Build failed!
    exit /b 1
)

echo.
echo ================================================
echo [SUCCESS] Build complete!
if /I "%BUILD_TYPE%"=="onedir" (
    echo Check dist\Excella\ for the final executable.
) else (
    echo Check dist\Excella.exe for the final executable.
)
echo ================================================
echo [INFO] Build completed at %DATE% %TIME%
pause