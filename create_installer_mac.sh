#!/bin/bash
# Script to create macOS installer (.dmg)
# This creates a disk image that users can mount and drag-drop to install

echo "=========================================="
echo "Sterling Map Migrator - macOS Installer"
echo "=========================================="
echo ""

# Check if the application has been built
if [ ! -d "dist/SterlingMapMigrator.app" ]; then
    echo "❌ ERROR: Application not built yet!"
    echo ""
    echo "Please run ./build.sh first to build the application."
    echo ""
    exit 1
fi

# Check if create-dmg is installed
if ! command -v create-dmg &> /dev/null; then
    echo "📦 Installing create-dmg..."
    if command -v brew &> /dev/null; then
        brew install create-dmg
    else
        echo "❌ ERROR: Homebrew not found!"
        echo ""
        echo "Please install Homebrew first:"
        echo "  /bin/bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\""
        echo ""
        echo "Then run this script again."
        exit 1
    fi
fi

# Create output directory
mkdir -p installer_output

# Create the DMG
echo ""
echo "🔨 Creating macOS installer (.dmg)..."
echo ""

create-dmg \
  --volname "Sterling Map Migrator" \
  --volicon "dist/SterlingMapMigrator.app/Contents/Resources/icon.icns" \
  --window-pos 200 120 \
  --window-size 800 400 \
  --icon-size 100 \
  --icon "SterlingMapMigrator.app" 200 190 \
  --hide-extension "SterlingMapMigrator.app" \
  --app-drop-link 600 185 \
  --no-internet-enable \
  "installer_output/SterlingMapMigrator_v1.0.0.dmg" \
  "dist/SterlingMapMigrator.app"

if [ $? -eq 0 ]; then
    echo ""
    echo "=========================================="
    echo "✅ Installer created successfully!"
    echo "=========================================="
    echo ""
    echo "📁 Installer location:"
    echo "   installer_output/SterlingMapMigrator_v1.0.0.dmg"
    echo ""
    echo "📊 File size: ~60-90 MB"
    echo ""
    echo "🚀 You can now distribute this single .dmg file!"
    echo ""
    echo "When users open it:"
    echo "  1. DMG will mount automatically"
    echo "  2. Window shows app and Applications folder"
    echo "  3. User drags app to Applications"
    echo "  4. App is installed and ready to use"
    echo "  5. Can be launched from Applications or Launchpad"
    echo ""
else
    echo ""
    echo "❌ Failed to create installer"
    echo "Please check the error messages above."
    echo ""
    exit 1
fi

# Made with Bob
