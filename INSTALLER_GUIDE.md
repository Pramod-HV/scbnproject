# Installer Creation Guide - Sterling B2Bi Map Migration Accelerator

## 🎯 Overview

This guide shows you how to create a **single-file installer** that users can download and run. The installer will:
- Install the application automatically
- Create desktop shortcuts
- Add to Start Menu (Windows) or Applications (macOS)
- Allow users to launch the app anytime like any other desktop application

---

## 🪟 Windows Installer (.exe)

### What You Get:
A single `.exe` file (~50-80 MB) that users can download and run. It will:
1. Show an installation wizard
2. Install to Program Files
3. Create desktop shortcut (optional)
4. Add to Start Menu
5. Launch the app after installation

### Prerequisites:

**Install Inno Setup** (free):
1. Download from: https://jrsoftware.org/isdl.php
2. Run the installer
3. Install with default settings

### Steps to Create Installer:

#### Step 1: Build the Application
```cmd
cd Integrated_Folder
build.bat
```

#### Step 2: Create the Installer
```cmd
create_installer.bat
```

#### Step 3: Get Your Installer
The installer will be created at:
```
installer_output/SterlingMapMigrator_Setup_v1.0.0.exe
```

### What Users Do:

1. **Download** `SterlingMapMigrator_Setup_v1.0.0.exe`
2. **Double-click** the installer
3. **Follow the wizard**:
   - Click "Next"
   - Choose installation location (default: Program Files)
   - Select "Create desktop shortcut" (optional)
   - Click "Install"
4. **Launch** the application (checkbox at end)
5. **Done!** App is now installed and can be launched anytime from:
   - Desktop shortcut
   - Start Menu → Sterling Map Migrator
   - Windows Search → "Sterling Map Migrator"

### Uninstall:
- Start Menu → Sterling Map Migrator → Uninstall
- Or: Settings → Apps → Sterling Map Migrator → Uninstall

---

## 🍎 macOS Installer (.dmg)

### What You Get:
A single `.dmg` file (~60-90 MB) that users can download and open. It will:
1. Mount as a disk image
2. Show drag-and-drop installation window
3. User drags app to Applications folder
4. App is installed and ready to use

### Prerequisites:

**Install create-dmg** (free):
```bash
# Install Homebrew if not already installed
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install create-dmg
brew install create-dmg
```

### Steps to Create Installer:

#### Step 1: Build the Application
```bash
cd Integrated_Folder
chmod +x build.sh
./build.sh
```

#### Step 2: Create the Installer
```bash
chmod +x create_installer_mac.sh
./create_installer_mac.sh
```

#### Step 3: Get Your Installer
The installer will be created at:
```
installer_output/SterlingMapMigrator_v1.0.0.dmg
```

### What Users Do:

1. **Download** `SterlingMapMigrator_v1.0.0.dmg`
2. **Double-click** the DMG file
3. **Drag** the app icon to the Applications folder
4. **Eject** the disk image
5. **Launch** from:
   - Applications folder
   - Launchpad
   - Spotlight search

### First Launch (macOS Security):
If macOS shows "App is damaged" or security warning:
```bash
xattr -cr /Applications/SterlingMapMigrator.app
```
Or: System Preferences → Security & Privacy → "Open Anyway"

### Uninstall:
- Drag app from Applications to Trash
- Empty Trash

---

## 🐧 Linux (AppImage - Optional)

For Linux, you can create an AppImage (single executable file):

### Install AppImage Tools:
```bash
# Download appimagetool
wget https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage
chmod +x appimagetool-x86_64.AppImage
```

### Create AppImage:
```bash
# Build the app first
./build.sh

# Create AppImage structure
mkdir -p SterlingMapMigrator.AppDir/usr/bin
cp -r dist/SterlingMapMigrator/* SterlingMapMigrator.AppDir/usr/bin/

# Create desktop entry
cat > SterlingMapMigrator.AppDir/SterlingMapMigrator.desktop << EOF
[Desktop Entry]
Name=Sterling Map Migrator
Exec=SterlingMapMigrator
Icon=sterling
Type=Application
Categories=Utility;
EOF

# Build AppImage
./appimagetool-x86_64.AppImage SterlingMapMigrator.AppDir
```

---

## 📊 Comparison

| Platform | Format | Size | Installation | Uninstall |
|----------|--------|------|--------------|-----------|
| Windows | `.exe` | 50-80 MB | Wizard | Control Panel |
| macOS | `.dmg` | 60-90 MB | Drag & Drop | Trash |
| Linux | `.AppImage` | 50-80 MB | Make executable | Delete file |

---

## 🚀 Complete Workflow

### For Windows:

```cmd
REM 1. Build the application
cd Integrated_Folder
build.bat

REM 2. Create installer
create_installer.bat

REM 3. Distribute
REM Upload installer_output/SterlingMapMigrator_Setup_v1.0.0.exe
```

### For macOS:

```bash
# 1. Build the application
cd Integrated_Folder
./build.sh

# 2. Create installer
./create_installer_mac.sh

# 3. Distribute
# Upload installer_output/SterlingMapMigrator_v1.0.0.dmg
```

---

## 📤 Distribution

### Upload to Cloud Storage:

**Google Drive:**
1. Upload the installer file
2. Right-click → "Get link"
3. Set to "Anyone with link can view"
4. Share the link

**Dropbox:**
1. Upload the installer
2. Click "Share" → "Create link"
3. Share the link

**Direct Download:**
- Host on your website
- Use file hosting service
- Email (if under 25 MB)

### Send to Users:

**Email Template:**
```
Subject: Sterling Map Migrator - Installation

Hi,

Please install the Sterling Map Migrator application:

Download: [Your Link Here]

Installation Instructions:

Windows:
1. Download the .exe file
2. Double-click to run the installer
3. Follow the installation wizard
4. Launch from desktop shortcut or Start Menu

macOS:
1. Download the .dmg file
2. Open the DMG
3. Drag the app to Applications folder
4. Launch from Applications or Launchpad

The application will appear as a regular desktop application and can be launched anytime.

If you encounter any issues, please let me know.

Thanks!
```

---

## 🔧 Customization

### Add Custom Icon:

**Windows (Inno Setup):**
1. Create/get an `.ico` file
2. Edit `installer_windows.iss`:
   ```
   SetupIconFile=path\to\your\icon.ico
   ```

**macOS:**
1. Create/get an `.icns` file
2. Place in `dist/SterlingMapMigrator.app/Contents/Resources/`
3. Update `build_app.spec`:
   ```python
   icon='path/to/your/icon.icns'
   ```

### Change Application Name:

Edit `installer_windows.iss` or `create_installer_mac.sh` and change:
- `MyAppName`
- `MyAppVersion`
- Output filename

---

## ✅ Testing Checklist

Before distributing:

- [ ] Build application successfully
- [ ] Create installer successfully
- [ ] Test installer on clean machine
- [ ] Verify desktop shortcut works
- [ ] Verify Start Menu/Applications entry
- [ ] Test launching the application
- [ ] Test all application features
- [ ] Test uninstallation
- [ ] Check file size is reasonable
- [ ] Verify no antivirus false positives

---

## 🆘 Troubleshooting

### Windows Installer Issues:

**"Inno Setup not found"**
- Install from: https://jrsoftware.org/isdl.php
- Ensure installed to default location

**"Application not built"**
- Run `build.bat` first
- Check `dist/SterlingMapMigrator` exists

**Installer won't run**
- Right-click → "Run as administrator"
- Check antivirus isn't blocking

### macOS Installer Issues:

**"create-dmg not found"**
- Install Homebrew first
- Run: `brew install create-dmg`

**"App is damaged"**
- Run: `xattr -cr /Applications/SterlingMapMigrator.app`
- Or: System Preferences → Security → "Open Anyway"

**DMG won't mount**
- Check file downloaded completely
- Try downloading again

---

## 🎉 You're Ready!

You now have everything needed to create professional installers for your application. Users can simply download and install like any other desktop application!

### Quick Summary:

1. **Build**: `build.bat` or `./build.sh`
2. **Create Installer**: `create_installer.bat` or `./create_installer_mac.sh`
3. **Distribute**: Upload installer file and share link
4. **Users**: Download → Install → Launch → Use!

Happy distributing! 🚀