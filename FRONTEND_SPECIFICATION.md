# AI-Powered Map Migration Accelerator - Frontend Specification

## Overview
Desktop application for Sterling B2Bi EDI map migration automation with independent function execution and real-time log viewing.

---

## Technology Stack Options

### **Option 1: PyQt6 / PySide6 (Recommended)**
- **Pros**: Native desktop feel, cross-platform (Windows/Mac/Linux), professional UI, excellent performance
- **Cons**: Steeper learning curve
- **Best for**: Professional enterprise applications

### **Option 2: Tkinter (Built-in)**
- **Pros**: No installation needed, simple, lightweight
- **Cons**: Less modern UI, limited styling
- **Best for**: Quick prototypes, simple interfaces

### **Option 3: CustomTkinter (Modern Tkinter)**
- **Pros**: Modern look, easy to use, built on Tkinter
- **Cons**: Additional dependency
- **Best for**: Modern UI with Tkinter simplicity

### **Recommended: PyQt6** for professional enterprise application

---

## Application Structure

```
┌─────────────────────────────────────────────────────────────┐
│  AI-Powered Map Migration Accelerator                       │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  Menu Bar: File | Tools | Help                      │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                              │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  Tab 1: Phase 0 - Extract Map Details               │   │
│  │  Tab 2: Phase 1 - Extract Rules                     │   │
│  │  Tab 3: Phase 2 - Apply Rules                       │   │
│  │  Tab 4: Phase 3 - Process Maps                      │   │
│  │  Tab 5: Utilities                                    │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                              │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  Log Output (Real-time)                             │   │
│  │  ┌───────────────────────────────────────────────┐ │   │
│  │  │ [INFO] Processing started...                  │ │   │
│  │  │ [INFO] Found 32 MXL files                     │ │   │
│  │  │ [SUCCESS] Processing complete                 │ │   │
│  │  └───────────────────────────────────────────────┘ │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                              │
│  [Progress Bar: ████████████░░░░░░░░ 60%]                  │
│  Status: Processing file 20 of 32...                        │
└─────────────────────────────────────────────────────────────┘
```

---

## Detailed Tab Specifications

### **Tab 1: Phase 0 - Extract Map Details**

```
┌─────────────────────────────────────────────────────────────┐
│  Phase 0: Extract Map Details                               │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  Input Source:                                               │
│  ○ ZIP Files    ○ MXL Files                                 │
│                                                              │
│  Input Folder:  [___________________________] [Browse...]   │
│                                                              │
│  Output Excel:  [mapping_results.xlsx      ] [Browse...]   │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Options:                                             │  │
│  │  ☑ Copy to mxl_files folder                          │  │
│  │  ☑ Skip duplicates                                    │  │
│  │  ☑ Save to old_mxlFiles (archive)                    │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  [Run Extraction]  [Clear]  [Download Results]              │
│                                                              │
│  Results:                                                    │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  ✓ Processed: 32 files                               │  │
│  │  ✓ Created: mapping_results.xlsx                     │  │
│  │  ✓ Copied: 32 files to mxl_files/                    │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### **Tab 2: Phase 1 - Extract Rules**

```
┌─────────────────────────────────────────────────────────────┐
│  Phase 1: Extract Process Data Rules                        │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  MXL Files Folder:  [mxl_files/            ] [Browse...]   │
│                                                              │
│  Output Excel:      [process_data_rules.xlsx] [Browse...]   │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Extract Options:                                     │  │
│  │  ☑ Include field names                               │  │
│  │  ☑ Include XPath information                         │  │
│  │  ☑ Auto-format Excel                                 │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  [Run Extraction]  [Clear]  [Download Excel]                │
│                                                              │
│  Instructions:                                               │
│  1. Click "Run Extraction" to generate process_data_rules   │
│  2. Download the Excel file                                  │
│  3. Open and add comments/uncomments as needed               │
│  4. Upload in Phase 2 to apply changes                       │
└─────────────────────────────────────────────────────────────┘
```

### **Tab 3: Phase 2 - Apply Rules**

```
┌─────────────────────────────────────────────────────────────┐
│  Phase 2: Apply Process Data Rules                          │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  Rules Excel:       [___________________________] [Upload]  │
│                     (Upload your modified process_data_rules)│
│                                                              │
│  MXL Files Folder:  [mxl_files/            ] [Browse...]   │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Apply Options:                                       │  │
│  │  ☑ Backup files before changes                       │  │
│  │  ☑ Validate rules before applying                    │  │
│  │  ☑ Generate change report                            │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  [Validate Rules]  [Apply Changes]  [Rollback]              │
│                                                              │
│  Summary:                                                    │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Files to modify: 15                                 │  │
│  │  Comments to add: 45                                  │  │
│  │  Uncomments to add: 12                                │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### **Tab 4: Phase 3 - Process Maps**

```
┌─────────────────────────────────────────────────────────────┐
│  Phase 3: Process All Maps                                  │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  Mapping Results:   [mapping_results.xlsx  ] [Browse...]   │
│  Generic Checklist: [Generic_checklistMain.xlsm] [Browse...]│
│  MXL Files Folder:  [mxl_files/            ] [Browse...]   │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Processing Options:                                  │  │
│  │  ☑ Run Pre-session Rules                             │  │
│  │  ☑ Modify X Syntax Token                             │  │
│  │  ☑ Modify Character Encoding                         │  │
│  │  ☑ Process Freeformat                                │  │
│  │  ☑ Process Inbound Features                          │  │
│  │  ☑ Generate Error Report                             │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  [Process All]  [Process Selected]  [Stop]                  │
│                                                              │
│  File Selection:                                             │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  ☑ CBI_XML_AMAZON_I_850_4010.mxl                     │  │
│  │  ☑ CBI_XML_WALMART_I_850_5010.mxl                    │  │
│  │  ☑ CBI_POS_TARGET_I_856_4010.mxl                     │  │
│  │  ... (32 files total)                                │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  [Download Error Report]  [Download All Results]            │
└─────────────────────────────────────────────────────────────┘
```

### **Tab 5: Utilities**

```
┌─────────────────────────────────────────────────────────────┐
│  Utilities & Individual Functions                           │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Individual Function Execution:                       │  │
│  │                                                        │  │
│  │  [Pre-session Rules Only]                             │  │
│  │  [Modify X Syntax Token Only]                         │  │
│  │  [Character Encoding Only]                            │  │
│  │  [Freeformat Processing Only]                         │  │
│  │  [Remove Namespace Prefixes]                          │  │
│  │  [Inbound Features Only]                              │  │
│  │  [Outbound Features Only]                             │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  File Management:                                     │  │
│  │                                                        │  │
│  │  [Clean Old Files]  [Backup Current State]            │  │
│  │  [Restore from Backup]  [Export Logs]                 │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                              │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  Settings:                                            │  │
│  │                                                        │  │
│  │  Default Folders:                                     │  │
│  │  ZIP Files:     [zipfiles/]                           │  │
│  │  MXL Files:     [mxl_files/]                          │  │
│  │  Archive:       [old_mxlFiles/]                       │  │
│  │                                                        │  │
│  │  [Save Settings]  [Reset to Defaults]                 │  │
│  └──────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

---

## Key Features

### 1. **Real-time Log Viewer**
- Color-coded logs (INFO=blue, WARNING=yellow, ERROR=red, SUCCESS=green)
- Auto-scroll to latest
- Search/filter logs
- Export logs to file
- Clear logs button

### 2. **Progress Tracking**
- Overall progress bar
- Current file being processed
- Estimated time remaining
- Files processed / Total files

### 3. **File Management**
- Browse buttons for all file/folder selections
- Drag-and-drop support for files
- Recent files list
- File validation before processing

### 4. **Download/Upload**
- Download generated Excel files
- Upload modified Excel files
- Download error reports
- Download all results as ZIP

### 5. **Error Handling**
- Clear error messages
- Suggested fixes
- Retry failed operations
- Rollback capability

---

## UI/UX Guidelines

### Colors
- **Primary**: #2196F3 (Blue)
- **Success**: #4CAF50 (Green)
- **Warning**: #FF9800 (Orange)
- **Error**: #F44336 (Red)
- **Background**: #FAFAFA (Light Gray)
- **Text**: #212121 (Dark Gray)

### Fonts
- **Headers**: 14pt Bold
- **Body**: 10pt Regular
- **Logs**: 9pt Monospace

### Spacing
- Padding: 10px
- Margin: 5px
- Button height: 35px
- Input height: 30px

---

## Backend Integration Points

### API Endpoints (to be created)

```python
# Phase 0
POST /api/extract-map-details
    - input_source: "zip" | "mxl"
    - input_folder: str
    - output_excel: str
    - options: dict

# Phase 1
POST /api/extract-rules
    - mxl_folder: str
    - output_excel: str

# Phase 2
POST /api/apply-rules
    - rules_excel: file
    - mxl_folder: str

# Phase 3
POST /api/process-maps
    - mapping_results: str
    - checklist: str
    - mxl_folder: str
    - options: dict
    - selected_files: list

# Utilities
POST /api/run-function
    - function_name: str
    - parameters: dict

# File Operations
GET /api/download/{filename}
POST /api/upload
GET /api/logs
```

---

## Installation & Setup

### Requirements
```
PyQt6>=6.6.0
pandas>=2.0.0
openpyxl>=3.1.0
```

### Launch Command
```bash
python frontend_app.py
```

---

## Next Steps

1. **Choose Technology**: PyQt6 recommended
2. **Create UI Mockups**: Use Qt Designer or Figma
3. **Implement Backend API**: Create Flask/FastAPI wrapper
4. **Build Frontend**: Implement UI with chosen framework
5. **Integration**: Connect frontend to backend
6. **Testing**: Test all workflows
7. **Packaging**: Create executable with PyInstaller

---

## File Structure

```
Integrated_Folder/
├── frontend_app.py           # Main application
├── backend_api.py            # Backend wrapper
├── ui/
│   ├── main_window.py        # Main window
│   ├── tab_extract.py        # Phase 0 tab
│   ├── tab_extract_rules.py  # Phase 1 tab
│   ├── tab_apply_rules.py    # Phase 2 tab
│   ├── tab_process.py        # Phase 3 tab
│   ├── tab_utilities.py      # Utilities tab
│   └── log_viewer.py         # Log viewer widget
├── utils/
│   ├── file_manager.py       # File operations
│   ├── logger.py             # Logging handler
│   └── config.py             # Configuration
└── resources/
    ├── icons/                # UI icons
    └── styles.qss            # Qt stylesheets
```

---

This specification provides everything needed for a frontend developer to create the desktop application!