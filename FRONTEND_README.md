# Frontend Development Guide
## AI-Powered Map Migration Accelerator

This guide provides everything you need to build the desktop frontend application.

---

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements_frontend.txt
```

### 2. Run the Template Application

```bash
python pyqt6_app_template.py
```

---

## Project Structure

```
Integrated_Folder/
├── FRONTEND_SPECIFICATION.md    # Complete UI/UX specifications
├── backend_api.py                # Backend integration layer
├── pyqt6_app_template.py         # PyQt6 application template
├── requirements_frontend.txt     # Frontend dependencies
│
├── Backend Scripts (DO NOT MODIFY)
│   ├── mxl_parser.py
│   ├── process_data_manager.py
│   ├── process_data_updater_v3.py
│   ├── process_all_mxl_files.py
│   └── ... (other backend scripts)
```

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────┐
│                  PyQt6 Frontend                         │
│  ┌───────────────────────────────────────────────────┐ │
│  │  Main Window (pyqt6_app_template.py)             │ │
│  │  ├── Phase 0 Tab                                  │ │
│  │  ├── Phase 1 Tab                                  │ │
│  │  ├── Phase 2 Tab                                  │ │
│  │  ├── Phase 3 Tab                                  │ │
│  │  ├── Utilities Tab                                │ │
│  │  └── Log Viewer                                   │ │
│  └───────────────────────────────────────────────────┘ │
│                         ↓                               │
│  ┌───────────────────────────────────────────────────┐ │
│  │  Backend API (backend_api.py)                     │ │
│  │  - Handles all backend communication              │ │
│  │  - Provides thread-safe execution                 │ │
│  │  - Streams logs in real-time                      │ │
│  └───────────────────────────────────────────────────┘ │
│                         ↓                               │
│  ┌───────────────────────────────────────────────────┐ │
│  │  Backend Scripts (Python)                         │ │
│  │  - mxl_parser.py                                  │ │
│  │  - process_data_manager.py                        │ │
│  │  - process_data_updater_v3.py                     │ │
│  │  - etc.                                            │ │
│  └───────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
```

---

## Backend API Usage

### Initialization

```python
from backend_api import BackendAPI

def log_handler(level, message):
    """Handle log messages"""
    print(f"[{level}] {message}")

api = BackendAPI(log_callback=log_handler)
```

### Available Methods

#### Phase 0: Extract Map Details
```python
success = api.extract_map_details(
    input_source='zip',  # or 'mxl'
    input_folder='zipfiles',
    output_excel='mapping_results.xlsx'
)
```

#### Phase 1: Extract Rules
```python
success = api.extract_rules(
    mxl_folder='mxl_files',
    output_excel='process_data_rules.xlsx'
)
```

#### Phase 2: Apply Rules
```python
success = api.apply_rules(
    rules_excel='process_data_rules.xlsx',
    mxl_folder='mxl_files'
)
```

#### Phase 3: Process Maps
```python
# Process all files
success = api.process_all_maps(
    mapping_results='mapping_results.xlsx',
    checklist='Generic_checklistMain.xlsm',
    mxl_folder='mxl_files'
)

# Process selected files
success = api.process_all_maps(
    mapping_results='mapping_results.xlsx',
    checklist='Generic_checklistMain.xlsm',
    mxl_folder='mxl_files',
    selected_files=['file1.mxl', 'file2.mxl']
)
```

#### Utility Methods
```python
# Get list of MXL files
files = api.get_mxl_files('mxl_files')

# Validate required files exist
validation = api.validate_files()

# Backup files
success = api.backup_files('mxl_files', 'backup')

# Stop current process
api.stop_current_process()
```

---

## Threading Best Practices

### Always Run Backend Tasks in Worker Threads

```python
from PyQt6.QtCore import QThread, pyqtSignal

class WorkerThread(QThread):
    log_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal(bool)
    
    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
    
    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.finished_signal.emit(result)
        except Exception as e:
            self.log_signal.emit('ERROR', str(e))
            self.finished_signal.emit(False)

# Usage
worker = WorkerThread(api.extract_map_details, 'zip', 'zipfiles')
worker.log_signal.connect(handle_log)
worker.finished_signal.connect(task_finished)
worker.start()
```

---

## Log Viewer Implementation

### Color-Coded Logs

```python
from PyQt6.QtWidgets import QTextEdit
from PyQt6.QtGui import QColor

class LogViewer(QTextEdit):
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        
        self.colors = {
            'INFO': QColor(33, 150, 243),      # Blue
            'SUCCESS': QColor(76, 175, 80),    # Green
            'WARNING': QColor(255, 152, 0),    # Orange
            'ERROR': QColor(244, 67, 54)       # Red
        }
    
    def append_log(self, level, message):
        color = self.colors.get(level, QColor(0, 0, 0))
        self.setTextColor(color)
        self.append(message)
        
        # Auto-scroll
        self.verticalScrollBar().setValue(
            self.verticalScrollBar().maximum()
        )
```

---

## UI Components

### File Browser

```python
from PyQt6.QtWidgets import QFileDialog

# Browse for folder
folder = QFileDialog.getExistingDirectory(
    self, "Select Folder"
)

# Browse for file
file, _ = QFileDialog.getOpenFileName(
    self, "Select File", "", "Excel Files (*.xlsx *.xls)"
)
```

### Progress Bar

```python
from PyQt6.QtWidgets import QProgressBar

# Indeterminate progress
progress_bar = QProgressBar()
progress_bar.setRange(0, 0)  # Indeterminate

# Determinate progress
progress_bar.setRange(0, 100)
progress_bar.setValue(50)  # 50%
```

### Message Boxes

```python
from PyQt6.QtWidgets import QMessageBox

# Information
QMessageBox.information(self, "Success", "Task completed!")

# Warning
QMessageBox.warning(self, "Warning", "Please select a file")

# Error
QMessageBox.critical(self, "Error", "Task failed!")

# Question
reply = QMessageBox.question(
    self, "Confirm", "Are you sure?",
    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
)
if reply == QMessageBox.StandardButton.Yes:
    # User clicked Yes
    pass
```

---

## Styling

### Qt Stylesheets (QSS)

```python
# Apply stylesheet to entire application
app.setStyleSheet("""
    QMainWindow {
        background-color: #FAFAFA;
    }
    
    QPushButton {
        background-color: #2196F3;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        font-size: 10pt;
    }
    
    QPushButton:hover {
        background-color: #1976D2;
    }
    
    QPushButton:pressed {
        background-color: #0D47A1;
    }
    
    QGroupBox {
        border: 1px solid #CCCCCC;
        border-radius: 5px;
        margin-top: 10px;
        padding-top: 10px;
    }
    
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 5px;
    }
""")
```

---

## Testing

### Test Backend API

```python
# test_backend.py
from backend_api import BackendAPI

def test_log(level, message):
    print(f"[{level}] {message}")

api = BackendAPI(log_callback=test_log)

# Test Phase 0
print("Testing Phase 0...")
success = api.extract_map_details('mxl', 'old_mxlFiles')
print(f"Phase 0: {'SUCCESS' if success else 'FAILED'}")

# Test getting files
print("\nTesting file list...")
files = api.get_mxl_files()
print(f"Found {len(files)} files")
```

### Run Tests

```bash
python test_backend.py
```

---

## Packaging for Distribution

### Using PyInstaller

```bash
# Install PyInstaller
pip install pyinstaller

# Create executable
pyinstaller --onefile --windowed \
    --name "Map Migration Accelerator" \
    --icon=icon.ico \
    pyqt6_app_template.py

# Output will be in dist/ folder
```

### PyInstaller Spec File

```python
# map_migration.spec
a = Analysis(
    ['pyqt6_app_template.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('backend_api.py', '.'),
        ('*.py', '.'),  # Include all Python files
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Map Migration Accelerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'
)
```

---

## Troubleshooting

### Common Issues

#### 1. Backend Process Not Responding
- Check if process is running: `api.is_running`
- Stop process: `api.stop_current_process()`

#### 2. Logs Not Appearing
- Ensure log_callback is properly connected
- Check if worker thread is running

#### 3. UI Freezing
- Always run backend tasks in worker threads
- Never call backend API directly from UI thread

#### 4. File Not Found Errors
- Use absolute paths when possible
- Validate file existence before processing

---

## Support

For questions or issues:
1. Check FRONTEND_SPECIFICATION.md for detailed UI specs
2. Review backend_api.py for API documentation
3. Examine pyqt6_app_template.py for implementation examples

---

## Next Steps

1. **Customize UI**: Modify pyqt6_app_template.py to match your design
2. **Add Features**: Implement additional tabs or utilities
3. **Test Thoroughly**: Test all workflows with real data
4. **Package**: Create executable for distribution

Good luck with your frontend development!