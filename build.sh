#!/bin/bash
# Build script for Sterling B2Bi Map Migration Accelerator
# This script creates a standalone desktop application

echo "=========================================="
echo "Sterling Map Migrator - Build Script"
echo "=========================================="
echo ""

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "❌ PyInstaller not found. Installing..."
    pip install pyinstaller
    if [ $? -ne 0 ]; then
        echo "❌ Failed to install PyInstaller"
        exit 1
    fi
fi

# Check if required dependencies are installed
echo "📦 Checking dependencies..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "❌ Failed to install dependencies"
    exit 1
fi

# Clean previous builds
echo ""
echo "🧹 Cleaning previous builds..."
rm -rf build/
rm -rf dist/
rm -rf __pycache__/
rm -f *.pyc

# Build the application
echo ""
echo "🔨 Building application..."
pyinstaller build_app.spec --clean --noconfirm

if [ $? -eq 0 ]; then
    echo ""
    echo "=========================================="
    echo "✅ Build completed successfully!"
    echo "=========================================="
    echo ""
    echo "📁 Application location:"
    
    if [[ "$OSTYPE" == "darwin"* ]]; then
        # macOS
        echo "   dist/SterlingMapMigrator.app"
        echo ""
        echo "🚀 To run the application:"
        echo "   open dist/SterlingMapMigrator.app"
    else
        # Linux/Windows
        echo "   dist/SterlingMapMigrator/"
        echo ""
        echo "🚀 To run the application:"
        echo "   ./dist/SterlingMapMigrator/SterlingMapMigrator"
    fi
    echo ""
    echo "📦 You can distribute the entire 'dist/SterlingMapMigrator' folder"
    echo ""
else
    echo ""
    echo "❌ Build failed. Please check the error messages above."
    exit 1
fi

# Made with Bob
