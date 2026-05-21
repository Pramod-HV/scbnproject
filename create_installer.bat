@echo off
REM Script to create Windows installer using Inno Setup
REM This creates a single .exe installer file

echo ==========================================
echo Sterling Map Migrator - Installer Creator
echo ==========================================
echo.

REM Check if Inno Setup is installed
set INNO_PATH="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %INNO_PATH% (
    echo.
    echo ERROR: Inno Setup not found!
    echo.
    echo Please install Inno Setup from: https://jrsoftware.org/isdl.php
    echo.
    echo After installation, run this script again.
    echo.
    pause
    exit /b 1
)

REM Check if the application has been built
if not exist "dist\SterlingMapMigrator" (
    echo.
    echo ERROR: Application not built yet!
    echo.
    echo Please run build.bat first to build the application.
    echo.
    pause
    exit /b 1
)

REM Create output directory for installer
if not exist "installer_output" mkdir installer_output

REM Build the installer
echo.
echo Building Windows installer...
echo.
%INNO_PATH% installer_windows.iss

if errorlevel 0 (
    echo.
    echo ==========================================
    echo Installer created successfully!
    echo ==========================================
    echo.
    echo Installer location:
    echo    installer_output\SterlingMapMigrator_Setup_v1.0.0.exe
    echo.
    echo File size: ~50-80 MB
    echo.
    echo You can now distribute this single .exe file!
    echo.
    echo When users run it:
    echo  1. Installation wizard will guide them
    echo  2. Application will be installed to Program Files
    echo  3. Desktop shortcut will be created (optional)
    echo  4. Start Menu entry will be created
    echo  5. Application can be launched immediately
    echo.
) else (
    echo.
    echo ERROR: Failed to create installer
    echo Please check the error messages above.
    echo.
)

pause

@REM Made with Bob
