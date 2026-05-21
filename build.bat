@echo off
REM Build script for Sterling B2Bi Map Migration Accelerator (Windows)
REM This script creates a standalone desktop application

echo ==========================================
echo Sterling Map Migrator - Build Script
echo ==========================================
echo.

REM Check if PyInstaller is installed
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo Failed to install PyInstaller
        pause
        exit /b 1
    )
)

REM Check if required dependencies are installed
echo Checking dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install dependencies
    pause
    exit /b 1
)

REM Clean previous builds
echo.
echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
del /q *.pyc 2>nul

REM Build the application
echo.
echo Building application...
pyinstaller build_app.spec --clean --noconfirm

if errorlevel 0 (
    echo.
    echo ==========================================
    echo Build completed successfully!
    echo ==========================================
    echo.
    echo Application location:
    echo    dist\SterlingMapMigrator\
    echo.
    echo To run the application:
    echo    dist\SterlingMapMigrator\SterlingMapMigrator.exe
    echo.
    echo You can distribute the entire 'dist\SterlingMapMigrator' folder
    echo.
) else (
    echo.
    echo Build failed. Please check the error messages above.
    pause
    exit /b 1
)

pause

@REM Made with Bob
