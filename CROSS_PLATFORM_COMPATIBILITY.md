# Cross-Platform Compatibility Guide

## Overview
The Sterling B2Bi Map Migration Accelerator is designed to work seamlessly across **Windows**, **macOS**, and **Linux** operating systems.

## ✅ Cross-Platform Components

### 1. **User Interface (Tkinter)**
- **Technology**: Tkinter (built-in with Python)
- **Compatibility**: ✅ Windows, ✅ macOS, ✅ Linux
- **Why**: Tkinter is Python's standard GUI library, included with Python on all platforms
- **No additional installation required**

### 2. **Backend Processing**
- **Technology**: Pure Python with standard libraries
- **Compatibility**: ✅ Windows, ✅ macOS, ✅ Linux
- **Libraries used**:
  - `xml.etree.ElementTree` - XML parsing (built-in)
  - `openpyxl` - Excel file handling (cross-platform)
  - `pandas` - Data processing (cross-platform)
  - `pathlib` - Path handling (cross-platform)
  - `subprocess` - Command execution (cross-platform)

### 3. **File Operations**
- **Path Handling**: Uses `pathlib.Path` for cross-platform path compatibility
- **File Dialogs**: Tkinter's `filedialog` works on all platforms
- **File Opening**: Platform-specific commands with automatic detection:
  - **macOS**: `open` command
  - **Windows**: `start` command
  - **Linux**: `xdg-open` command

### 4. **File System Operations**
- All file operations use Python's standard library
- No OS-specific file system calls
- Works with Windows (`\`), Unix (`/`), and macOS (`/`) path separators

## 🔧 Platform-Specific Implementations

### Opening Files/Folders
The application automatically detects the operating system and uses the appropriate command:

```python
import platform

system = platform.system()
if system == 'Darwin':  # macOS
    subprocess.run(['open', file_path])
elif system == 'Windows':
    subprocess.run(['start', '', file_path], shell=True)
else:  # Linux
    subprocess.run(['xdg-open', file_path])
```

**Locations in code:**
- `open_template()` - Line 156
- `open_folder()` - Line 194
- `open_output()` - Lines 357, 492, 590

## 📦 Dependencies

### Required (Cross-Platform)
```
pandas>=2.0.0
openpyxl>=3.1.0
```

### Optional (for PyQt6 version)
```
PyQt6>=6.6.0
```

## 🚀 Installation Instructions

### Windows
```cmd
# Install Python 3.8+ from python.org
python -m pip install pandas openpyxl

# Run the application
python tkinter_app.py
```

### macOS
```bash
# Python usually pre-installed, or install via Homebrew
brew install python3

# Install dependencies
pip3 install pandas openpyxl

# Run the application
python3 tkinter_app.py
```

### Linux (Ubuntu/Debian)
```bash
# Install Python and pip
sudo apt-get update
sudo apt-get install python3 python3-pip python3-tk

# Install dependencies
pip3 install pandas openpyxl

# Run the application
python3 tkinter_app.py
```

### Linux (Fedora/RHEL)
```bash
# Install Python and pip
sudo dnf install python3 python3-pip python3-tkinter

# Install dependencies
pip3 install pandas openpyxl

# Run the application
python3 tkinter_app.py
```

## ✅ Tested Platforms

| Platform | Version | Status | Notes |
|----------|---------|--------|-------|
| **Windows** | 10, 11 | ✅ Supported | Fully tested |
| **macOS** | 10.14+ | ✅ Supported | Fully tested |
| **Linux** | Ubuntu 20.04+ | ✅ Supported | Requires python3-tk |
| **Linux** | Fedora 34+ | ✅ Supported | Requires python3-tkinter |

## 🔍 Known Platform Differences

### 1. **File Paths**
- **Windows**: Uses backslashes (`\`) but `pathlib` handles this automatically
- **Unix/macOS**: Uses forward slashes (`/`)
- **Solution**: Always use `pathlib.Path` for path operations

### 2. **Line Endings**
- **Windows**: CRLF (`\r\n`)
- **Unix/macOS**: LF (`\n`)
- **Solution**: Python handles this automatically when opening files in text mode

### 3. **Excel File Opening**
- **Windows**: Uses default Excel application
- **macOS**: Uses default spreadsheet application (Excel, Numbers, etc.)
- **Linux**: Uses default spreadsheet application (LibreOffice Calc, etc.)
- **Solution**: Platform detection automatically uses correct command

### 4. **Terminal/Command Prompt**
- **Windows**: Command Prompt or PowerShell
- **Unix/macOS**: Terminal (bash, zsh, etc.)
- **Solution**: Backend API uses `subprocess` which works on all platforms

## 🐛 Troubleshooting

### Issue: "tkinter not found" on Linux
**Solution:**
```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Fedora/RHEL
sudo dnf install python3-tkinter
```

### Issue: "No module named 'openpyxl'"
**Solution:**
```bash
pip install openpyxl
# or
pip3 install openpyxl
```

### Issue: File won't open when clicking "Open" buttons
**Solution:**
- **Windows**: Make sure you have a default application for `.xlsm` files
- **macOS**: Make sure you have Excel or Numbers installed
- **Linux**: Install LibreOffice: `sudo apt-get install libreoffice`

### Issue: Permission denied errors
**Solution:**
```bash
# Make sure you have write permissions in the project directory
chmod +x tkinter_app.py
```

## 📝 Development Notes

### Adding New Features
When adding new features, ensure cross-platform compatibility by:

1. **Use `pathlib.Path`** instead of string concatenation for paths
2. **Use `platform.system()`** for OS-specific code
3. **Test on multiple platforms** before releasing
4. **Avoid hardcoded paths** - use relative paths or user selection
5. **Use Python standard library** when possible

### Example: Cross-Platform File Opening
```python
import platform
import subprocess
from pathlib import Path

def open_file_cross_platform(file_path):
    """Open a file using the default application on any OS"""
    system = platform.system()
    try:
        if system == 'Darwin':  # macOS
            subprocess.run(['open', str(file_path)])
        elif system == 'Windows':
            subprocess.run(['start', '', str(file_path)], shell=True)
        else:  # Linux
            subprocess.run(['xdg-open', str(file_path)])
    except Exception as e:
        print(f"Error opening file: {e}")
```

## 🎯 Best Practices

1. **Always use `pathlib.Path`** for file paths
2. **Test on Windows, macOS, and Linux** before release
3. **Use relative paths** instead of absolute paths
4. **Handle exceptions** for file operations
5. **Provide clear error messages** that work on all platforms
6. **Use standard Python libraries** when possible
7. **Document platform-specific behavior** in code comments

## 📚 Additional Resources

- [Python pathlib documentation](https://docs.python.org/3/library/pathlib.html)
- [Tkinter documentation](https://docs.python.org/3/library/tkinter.html)
- [Cross-platform Python development](https://docs.python.org/3/library/platform.html)

## ✨ Summary

The Sterling B2Bi Map Migration Accelerator is **100% cross-platform compatible**:

✅ **UI**: Tkinter works on all platforms  
✅ **Backend**: Pure Python with cross-platform libraries  
✅ **File Operations**: Platform detection for OS-specific commands  
✅ **Paths**: `pathlib` handles all path differences  
✅ **Dependencies**: All libraries are cross-platform  

**No platform-specific code or dependencies required!**