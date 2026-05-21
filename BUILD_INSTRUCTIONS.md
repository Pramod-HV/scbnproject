# Building Sterling B2Bi Map Migration Accelerator

This guide explains how to build the Sterling B2Bi Map Migration Accelerator as a standalone desktop application.

## Prerequisites

1. **Python 3.8 or higher** installed on your system
2. **pip** (Python package installer)
3. All project dependencies installed

## Quick Build

### For macOS/Linux:

```bash
chmod +x build.sh
./build.sh
```

### For Windows:

```cmd
build.bat
```

## Manual Build Steps

If you prefer to build manually or the automated scripts don't work:

### 1. Install PyInstaller

```bash
pip install pyinstaller
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Clean Previous Builds (Optional)

```bash
# macOS/Linux
rm -rf build/ dist/ __pycache__/

# Windows
rmdir /s /q build dist __pycache__
```

### 4. Build the Application

```bash
pyinstaller build_app.spec --clean --noconfirm
```

## Output

After a successful build, you'll find the application in:

### macOS:
- **Location**: `dist/SterlingMapMigrator.app`
- **Run**: Double-click the app or use `open dist/SterlingMapMigrator.app`
- **Distribute**: Share the entire `.app` bundle

### Windows:
- **Location**: `dist/SterlingMapMigrator/SterlingMapMigrator.exe`
- **Run**: Double-click the `.exe` file
- **Distribute**: Share the entire `SterlingMapMigrator` folder

### Linux:
- **Location**: `dist/SterlingMapMigrator/SterlingMapMigrator`
- **Run**: `./dist/SterlingMapMigrator/SterlingMapMigrator`
- **Distribute**: Share the entire `SterlingMapMigrator` folder

## What Gets Packaged

The build includes:

✅ **Application Files:**
- Main Tkinter GUI (`tkinter_app.py`)
- Backend API (`backend_api.py`)
- All processing scripts (MXL parser, processors, updaters, etc.)

✅ **Data Files:**
- `Generic_checklistMain.xlsm` (Checklist template)
- `mapping_results.xlsx` (Mapping results)
- `process_data_rules.xlsx` (Process data rules)
- `README.md` (Documentation)

✅ **Dependencies:**
- Python runtime
- pandas, openpyxl
- tkinter (GUI framework)
- All required Python standard libraries

## Distribution

### Single Folder Distribution (Recommended)

The `dist/SterlingMapMigrator` folder contains everything needed to run the application:

1. **Copy the entire folder** to the target machine
2. **No Python installation required** on the target machine
3. **No dependency installation required**
4. **Just run the executable**

### Important Notes

- **First Run**: The application may take a few seconds to start on first launch
- **Antivirus**: Some antivirus software may flag the executable. This is normal for PyInstaller apps
- **Permissions**: On macOS, you may need to allow the app in System Preferences > Security & Privacy
- **File Paths**: The application creates working directories (`zipfiles/`, `old_mxlFiles/`, `mxl_files/`) in the same location as the executable

## Troubleshooting

### Build Fails

1. **Check Python version**: `python --version` (should be 3.8+)
2. **Update pip**: `pip install --upgrade pip`
3. **Reinstall PyInstaller**: `pip uninstall pyinstaller && pip install pyinstaller`
4. **Check dependencies**: `pip install -r requirements.txt --upgrade`

### Application Won't Start

1. **Check console output**: Run from terminal to see error messages
2. **Missing files**: Ensure all Excel templates are in the same directory
3. **Permissions**: Ensure the executable has run permissions

### macOS "App is damaged" Error

```bash
xattr -cr dist/SterlingMapMigrator.app
```

### Windows SmartScreen Warning

Click "More info" → "Run anyway"

## Advanced Configuration

### Custom Icon

To add a custom icon, edit `build_app.spec`:

```python
icon='path/to/your/icon.ico'  # Windows
icon='path/to/your/icon.icns'  # macOS
```

### Console Window (Debug Mode)

To show console output for debugging, edit `build_app.spec`:

```python
console=True  # Change from False to True
```

### Exclude Unnecessary Modules

To reduce file size, add modules to the `excludes` list in `build_app.spec`:

```python
excludes=[
    'matplotlib',
    'numpy',
    # Add more modules here
]
```

## File Size

Expected application size:
- **Windows**: ~50-80 MB
- **macOS**: ~60-90 MB
- **Linux**: ~50-80 MB

Size varies based on Python version and included dependencies.

## Support

For issues or questions:
1. Check the error logs in the application
2. Review the console output when running from terminal
3. Ensure all required Excel files are present
4. Verify Python and dependency versions

## Version Information

- **Application**: Sterling B2Bi Map Migration Accelerator
- **Version**: 1.0.0
- **Build Tool**: PyInstaller
- **Python**: 3.8+
- **GUI Framework**: Tkinter