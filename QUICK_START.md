# Quick Start Guide - Sterling B2Bi Map Migration Accelerator

## 🚀 Building the Desktop Application

### Option 1: Automated Build (Recommended)

#### macOS/Linux:
```bash
cd Integrated_Folder
chmod +x build.sh
./build.sh
```

#### Windows:
```cmd
cd Integrated_Folder
build.bat
```

### Option 2: Manual Build

```bash
cd Integrated_Folder
pip install pyinstaller
pip install -r requirements.txt
pyinstaller build_app.spec --clean --noconfirm
```

## 📦 What You Get

After building, you'll have a **standalone desktop application** that:

✅ **Requires NO Python installation** on target machines
✅ **Requires NO dependency installation**
✅ **Works as a native desktop app**
✅ **Includes all necessary files** (Excel templates, scripts, etc.)

### Output Location:

- **macOS**: `dist/SterlingMapMigrator.app`
- **Windows**: `dist/SterlingMapMigrator/SterlingMapMigrator.exe`
- **Linux**: `dist/SterlingMapMigrator/SterlingMapMigrator`

## 🎯 Running the Application

### From Source (Development):
```bash
cd Integrated_Folder
python tkinter_app.py
```

### From Built Package:

#### macOS:
```bash
open dist/SterlingMapMigrator.app
```
Or double-click the app in Finder

#### Windows:
Double-click `SterlingMapMigrator.exe` in the `dist/SterlingMapMigrator` folder

#### Linux:
```bash
./dist/SterlingMapMigrator/SterlingMapMigrator
```

## 📋 Application Features

### Phase 0: MXL File Parsing
- Extract and organize MXL files from ZIP archives
- Create folder structure: `zipfiles/`, `old_mxlFiles/`, `mxl_files/`
- Support for nested ZIP files

### Phase 1: Codelist Extraction
- Extract codelists from MXL files
- Generate codelist reports
- Rename codelists based on rules

### Phase 2: Rules Extraction & Application
- Extract rules from MXL files
- Apply rules to target maps
- Support for inbound/outbound map features

### Phase 3: Process All Maps
- Batch process multiple maps
- Apply character encoding modifications
- Update syntax tokens
- Remove namespace prefixes

### Phase 4: Process Data Updates
- Update process data based on checklist rules
- Insert/update ExplicitRule tags
- Match with Generic_checklistMain.xlsm

## 📁 Required Files

Ensure these files are in the same directory as the application:

1. **Generic_checklistMain.xlsm** - Main checklist template
2. **mapping_results.xlsx** - Mapping results
3. **process_data_rules.xlsx** - Process data rules

These files are automatically included when you build the application.

## 🔧 Troubleshooting

### Build Issues

**PyInstaller not found:**
```bash
pip install pyinstaller
```

**Dependencies missing:**
```bash
pip install -r requirements.txt
```

**Permission denied (macOS/Linux):**
```bash
chmod +x build.sh
```

### Runtime Issues

**macOS: "App is damaged"**
```bash
xattr -cr dist/SterlingMapMigrator.app
```

**Windows: SmartScreen warning**
Click "More info" → "Run anyway"

**Missing Excel files:**
Ensure all `.xlsm` and `.xlsx` files are in the application directory

## 📊 Workflow Example

1. **Start Application** → Launch the desktop app
2. **Select Phase** → Choose Phase 0, 1, 2, 3, or 4
3. **Browse Input** → Select ZIP file, folder, or MXL files
4. **Configure Options** → Set parameters as needed
5. **Run Phase** → Click "Run Phase X" button
6. **View Logs** → Monitor progress in the log viewer
7. **Check Output** → Review results in output folders

## 🎨 Application Interface

The application features:
- **Modern dark theme** with color-coded logs
- **Tab-based navigation** for each phase
- **Real-time log viewer** with timestamps
- **Progress indicators** for long-running tasks
- **File browser dialogs** for easy file selection
- **Cross-phase map sharing** via "Use Maps from Phase 0" button

## 📦 Distribution

To distribute the application:

1. **Copy the entire folder**:
   - macOS: Copy `SterlingMapMigrator.app`
   - Windows/Linux: Copy the entire `SterlingMapMigrator` folder

2. **Include Excel templates** (if not already bundled)

3. **No installation required** on target machines

4. **Just run the executable**

## 💡 Tips

- **First run may be slow** - The app needs to initialize
- **Keep Excel files updated** - Use latest checklist versions
- **Check logs regularly** - Monitor for errors or warnings
- **Use Phase 0 first** - Establish proper folder structure
- **Backup original files** - Files are saved to `old_mxlFiles/`

## 📖 Full Documentation

For detailed information, see:
- `BUILD_INSTRUCTIONS.md` - Complete build guide
- `README.md` - Project overview
- `FRONTEND_README.md` - Frontend documentation

## 🆘 Support

If you encounter issues:
1. Check the log viewer in the application
2. Review error messages carefully
3. Ensure all required files are present
4. Verify Python version (3.8+ for development)
5. Check file permissions

## 🎉 Ready to Test!

Your application is now ready for testing. Build it and try it out!

```bash
# Build
./build.sh  # or build.bat on Windows

# Run
open dist/SterlingMapMigrator.app  # macOS
# or
dist/SterlingMapMigrator/SterlingMapMigrator.exe  # Windows
```

Happy migrating! 🚀
