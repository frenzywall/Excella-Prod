@echo off
echo Building Excella Installer...
echo.

REM Check if Inno Setup is installed
if not exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
    if not exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" (
        echo Inno Setup 6 is not installed. Please install it from https://jrsoftware.org/isdl.php
        echo.
        pause
        exit /b 1
    )
)

REM Create output directory
if not exist installer mkdir installer

REM Compile the installer
echo Compiling installer...
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
    "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" excella_setup.iss
) else (
    "%ProgramFiles%\Inno Setup 6\ISCC.exe" excella_setup.iss
)

echo.
echo Installer build complete! The setup file is in the installer folder.
echo.
