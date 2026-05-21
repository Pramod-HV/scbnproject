#!/usr/bin/env python3
"""
Sterling B2Bi Map Migration Accelerator - Desktop Application (Tkinter)
A cross-platform desktop application for managing EDI map migration workflows.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
from pathlib import Path
import sys
from backend_api import BackendAPI

class LogViewer(scrolledtext.ScrolledText):
    """Custom text widget for displaying color-coded logs"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(state='disabled', wrap='word', font=('Courier', 9), bg='#1e1e1e', fg='#d4d4d4')
        
        # Configure tags for different log levels (terminal-like colors)
        self.tag_config('INFO', foreground='#4ec9b0')  # Cyan
        self.tag_config('SUCCESS', foreground='#4ec9b0')  # Green
        self.tag_config('WARNING', foreground='#dcdcaa')  # Yellow
        self.tag_config('ERROR', foreground='#f48771')  # Red
        self.tag_config('DEBUG', foreground='#808080')  # Gray
        self.tag_config('TIMESTAMP', foreground='#858585')  # Dark gray
        
    def append_log(self, level, message):
        """Append a log message with appropriate color and timestamp"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        self.configure(state='normal')
        
        # Add timestamp
        self.insert('end', f"[{timestamp}] ", 'TIMESTAMP')
        
        # Add level and message
        if level in ['INFO', 'SUCCESS', 'WARNING', 'ERROR', 'DEBUG']:
            self.insert('end', f"{message}\n", level)
        else:
            self.insert('end', f"{message}\n", 'INFO')
        
        self.see('end')
        self.configure(state='disabled')
        
    def clear_logs(self):
        """Clear all logs"""
        self.configure(state='normal')
        self.delete('1.0', 'end')
        self.configure(state='disabled')


class WorkerThread(threading.Thread):
    """Worker thread for running backend operations"""
    
    def __init__(self, func, callback, *args, **kwargs):
        super().__init__(daemon=True)
        self.func = func
        self.callback = callback
        self.args = args
        self.kwargs = kwargs
        
    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.callback(True, result)
        except Exception as e:
            self.callback(False, str(e))


class Phase0Tab(ttk.Frame):
    """Phase 0: Extract Map Details"""
    
    def __init__(self, parent, backend, log_callback, shared_folders):
        super().__init__(parent)
        self.backend = backend
        self.log_callback = log_callback
        self.shared_folders = shared_folders
        self.folder_browsed = False  # Track if user has browsed for a folder
        self.operation_completed = False  # Track if operation has been run successfully
        self.create_widgets()
        
    def create_widgets(self):
        # Title
        title = ttk.Label(self, text="📋 Phase 0: Map Inventory & Renaming", font=('Arial', 18, 'bold'))
        title.pack(pady=15)
        
        # Description
        desc = ttk.Label(self, text="Extract map details and rename maps from ZIP or MXL files\n(Automatically detects file type)",
                        wraplength=600, font=('Arial', 11))
        desc.pack(pady=8)
        
        # Folder selection (removed mode selection - auto-detection handles it)
        folder_frame = ttk.LabelFrame(self, text="Input Folder", padding=10)
        folder_frame.pack(fill='x', padx=20, pady=10)
        
        self.folder_var = tk.StringVar(value='')  # No default folder
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50)
        folder_entry.pack(side='left', padx=5)
        
        ttk.Button(folder_frame, text="Browse...", command=self.browse_folder).pack(side='left', padx=5)
        
        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=20)
        
        self.run_btn = ttk.Button(btn_frame, text="🚀 Run Operation", command=self.run_phase0)
        self.run_btn.pack(side='left', padx=8)
        
        ttk.Button(btn_frame, text="📂 Open Output", command=self.open_output).pack(side='left', padx=8)
    
    def browse_folder(self):
        """Allow user to select MXL/ZIP files or folder"""
        # Open file dialog that allows selecting files
        files = filedialog.askopenfilenames(
            title="Select MXL or ZIP File(s) - or Cancel to select a folder",
            filetypes=[
                ("Map Files", "*.mxl *.zip"),
                ("MXL Files", "*.mxl"),
                ("ZIP Files", "*.zip"),
                ("All Files", "*.*")
            ]
        )
        
        if files:
            # Files were selected
            folder = str(Path(files[0]).parent)
            if len(files) == 1:
                self.folder_var.set(str(Path(files[0])))
                self.log_callback('INFO', f"✅ Selected file: {Path(files[0]).name}")
            else:
                self.folder_var.set(f"{folder} ({len(files)} files)")
                self.log_callback('INFO', f"✅ Selected {len(files)} files from: {folder}")
            self.folder_browsed = True
            self.shared_folders['phase0'] = folder
        else:
            # No files selected, offer folder selection
            folder = filedialog.askdirectory(title="Select Folder with MXL/ZIP Files")
            if folder:
                self.folder_var.set(folder)
                self.folder_browsed = True
                self.shared_folders['phase0'] = folder
                self.log_callback('INFO', f"✅ Selected folder: {folder}")
        
    def run_phase0(self):
        # Check if user has browsed for a folder
        if not self.folder_browsed:
            messagebox.showwarning(
                "Folder Not Selected",
                "⚠️ Please browse and select a folder before running the operation.\n\n"
                "Click the 'Browse...' button to select your input folder."
            )
            return
        
        input_folder = self.folder_var.get()
        if not input_folder or not Path(input_folder).exists():
            messagebox.showerror("Error", "Please select a valid input folder")
            return
        
        self.run_btn.config(state='disabled')
        self.log_callback('INFO', "Starting Phase 0 with auto-detection...")
        self.log_callback('INFO', f"Input: {input_folder}")
        
        def task():
            # Auto-detection: backend will determine if it's ZIP or MXL
            return self.backend.extract_map_details(
                input_source='auto',  # Let mxl_parser.py auto-detect
                input_folder=input_folder
            )
        
        def callback(success, result):
            self.run_btn.config(state='normal')
            if success:
                self.operation_completed = True  # Mark operation as completed
                self.log_callback('SUCCESS', "✅ Phase 0 completed successfully!")
                messagebox.showinfo("Success", "Map details extracted successfully!\nCheck mapping_results.xlsx")
            else:
                self.log_callback('ERROR', f"❌ Phase 0 failed: {result}")
                messagebox.showerror("Error", f"Phase 0 failed: {result}")
        
        WorkerThread(task, callback).start()
        
    def open_output(self):
        """Open the output file directly from project folder"""
        # Check if operation has been completed
        if not self.operation_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run the operation first before opening the output.\n\n"
                "Click 'Run Operation' to process your files."
            )
            return
        
        output_file = Path("mapping_results.xlsx")
        
        # Check if output file exists
        if not output_file.exists():
            messagebox.showwarning(
                "Output File Not Found",
                "⚠️ No output file found.\n\n"
                "Please run the operation first to generate the output file."
            )
            return
        
        import subprocess
        import platform
        
        try:
            # Open the file from project folder
            system = platform.system()
            if system == 'Darwin':  # macOS
                subprocess.run(['open', str(output_file)])
            elif system == 'Windows':
                subprocess.run(['start', '', str(output_file)], shell=True)
            else:  # Linux
                subprocess.run(['xdg-open', str(output_file)])
            
            self.log_callback('INFO', "📊 Opening mapping_results.xlsx from project folder...")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")


class CodelistTab(ttk.Frame):
    """Phase 1: Codelist Management"""
    
    def __init__(self, parent, backend, log_callback, shared_folders):
        super().__init__(parent)
        self.backend = backend
        self.log_callback = log_callback
        self.shared_folders = shared_folders
        self.folder_browsed = False  # Track if user has browsed for a folder
        self.extract_completed = False  # Track if extract operation completed
        self.rename_completed = False  # Track if rename operation completed
        self.create_widgets()
        
    def create_widgets(self):
        # Title - Bigger and more impressive
        title = ttk.Label(self, text="Phase 1: Codelist Management", font=('Arial', 18, 'bold'))
        title.pack(pady=15)
        
        # Description - Larger font
        desc = ttk.Label(self, text="Extract and rename codelists from MXL or ZIP files (auto-detects)",
                        font=('Arial', 11), wraplength=600)
        desc.pack(pady=8)
        
        # Info box - More prominent
        info_frame = ttk.Frame(self)
        info_frame.pack(fill='x', padx=20, pady=10)
        ttk.Label(info_frame, text="💡 Tip: Place either .mxl files or .zip files in the folder. The tool will auto-detect the file type.",
                 foreground='blue', font=('Arial', 10, 'italic'), wraplength=600).pack()
        
        # Folder selection - Larger
        folder_frame = ttk.LabelFrame(self, text="Input Folder", padding=15)
        folder_frame.pack(fill='x', padx=20, pady=15)
        
        self.folder_var = tk.StringVar(value='')  # No default folder
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50, font=('Arial', 10))
        folder_entry.pack(side='left', padx=5)
        
        ttk.Button(folder_frame, text="Browse...", command=self.browse_folder).pack(side='left', padx=5)
        ttk.Button(folder_frame, text="📁 Use Maps from Phase 0", command=self.use_phase0_maps).pack(side='left', padx=5)
        
        # Step 1: Extract Codelists - Enhanced
        extract_frame = ttk.LabelFrame(self, text="Step 1: Extract Codelists", padding=15)
        extract_frame.pack(fill='x', padx=20, pady=15)
        
        ttk.Label(extract_frame, text="Extract codelist names from MXL files and generate Excel report",
                 font=('Arial', 10), wraplength=600).pack(pady=8)
        
        extract_btn_frame = ttk.Frame(extract_frame)
        extract_btn_frame.pack(pady=12)
        
        self.extract_btn = ttk.Button(extract_btn_frame, text="📊 Extract Codelists",
                                     command=self.extract_codelists)
        self.extract_btn.pack(side='left', padx=5)
        
        ttk.Button(extract_btn_frame, text="📂 Open Report",
                  command=self.open_report).pack(side='left', padx=5)
        
        # Step 2: Rename Codelists - Enhanced
        rename_frame = ttk.LabelFrame(self, text="Step 2: Rename Codelists", padding=15)
        rename_frame.pack(fill='x', padx=20, pady=15)
        
        ttk.Label(rename_frame, text="After filling 'New Codelist Name' column in Excel, click Rename to apply changes",
                 font=('Arial', 10), wraplength=600, foreground='white').pack(pady=8)
        
        rename_btn_frame = ttk.Frame(rename_frame)
        rename_btn_frame.pack(pady=12)
        
        self.rename_btn = ttk.Button(rename_btn_frame, text="🔄 Rename Codelists",
                                    command=self.rename_codelists)
        self.rename_btn.pack(side='left', padx=5)
        
        ttk.Button(rename_btn_frame, text="📦 Download Updated Files",
                  command=self.download_updated_files).pack(side='left', padx=5)
    
    def browse_folder(self):
        """Allow user to select MXL/ZIP files or folder"""
        # Open file dialog that allows selecting files
        files = filedialog.askopenfilenames(
            title="Select MXL or ZIP File(s) - or Cancel to select a folder",
            filetypes=[
                ("Map Files", "*.mxl *.zip"),
                ("MXL Files", "*.mxl"),
                ("ZIP Files", "*.zip"),
                ("All Files", "*.*")
            ]
        )
        
        if files:
            # Files were selected
            folder = str(Path(files[0]).parent)
            if len(files) == 1:
                self.folder_var.set(str(Path(files[0])))
                self.log_callback('INFO', f"✅ Selected file: {Path(files[0]).name}")
            else:
                self.folder_var.set(f"{folder} ({len(files)} files)")
                self.log_callback('INFO', f"✅ Selected {len(files)} files from: {folder}")
            self.folder_browsed = True
            self.shared_folders['phase1'] = folder
        else:
            # No files selected, offer folder selection
            folder = filedialog.askdirectory(title="Select Folder with MXL/ZIP Files")
            if folder:
                self.folder_var.set(folder)
                self.folder_browsed = True
                self.shared_folders['phase1'] = folder
                self.log_callback('INFO', f"✅ Selected folder: {folder}")
    
    def use_phase0_maps(self):
        """Use the folder from Phase 0"""
        phase0_folder = self.shared_folders.get('phase0', '')
        if not phase0_folder:
            messagebox.showwarning(
                "Phase 0 Not Run",
                "⚠️ No folder found from Phase 0.\n\n"
                "Please run Phase 0 first or browse to select a folder manually."
            )
            return
        
        self.folder_var.set(phase0_folder)
        self.folder_browsed = True
        self.shared_folders['phase1'] = phase0_folder
        self.log_callback('INFO', f"✅ Using maps from Phase 0: {phase0_folder}")
    
    def extract_codelists(self):
        # Check if user has browsed for a folder
        if not self.folder_browsed:
            messagebox.showwarning(
                "Folder Not Selected",
                "⚠️ Please browse and select a folder before extracting codelists.\n\n"
                "Click the 'Browse...' button to select your MXL files folder."
            )
            return
        
        input_folder = self.folder_var.get()
        if not input_folder or not Path(input_folder).exists():
            messagebox.showerror("Error", "Please select a valid input folder")
            return
        
        self.extract_btn.config(state='disabled')
        self.log_callback('INFO', f"🔍 Extracting codelists from {input_folder}...")
        self.log_callback('INFO', "Auto-detecting file type (ZIP or MXL)...")
        
        def task():
            return self.backend.extract_codelists(
                input_folder=input_folder,
                output_file='codelist_report.xlsx'
            )
        
        def callback(success, result):
            self.extract_btn.config(state='normal')
            if success:
                self.extract_completed = True  # Mark extract as completed
                self.log_callback('SUCCESS', "✅ Codelist extraction complete!")
                self.log_callback('INFO', "📊 Report: codelist_report.xlsx")
                messagebox.showinfo("Success",
                    "Codelist extraction complete!\n\n"
                    "Next steps:\n"
                    "1. Click 'Open Report' to view codelist_report.xlsx\n"
                    "2. Fill in the 'New Codelist Name' column\n"
                    "3. Save the file\n"
                    "4. Click 'Rename Codelists' to apply changes")
            else:
                self.log_callback('ERROR', f"❌ Extraction failed: {result}")
                messagebox.showerror("Error", f"Extraction failed: {result}")
        
        WorkerThread(task, callback).start()
    
    def rename_codelists(self):
        # Check if extract operation has been completed
        if not self.extract_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Extract Codelists' first before renaming.\n\n"
                "Click 'Extract Codelists' to process your files."
            )
            return
        
        # Check if folder has been selected
        input_folder = self.folder_var.get()
        if not input_folder or not Path(input_folder).exists():
            messagebox.showerror("Error",
                "No folder selected!\n\n"
                "Please select a folder with MXL files first.")
            return
        
        # Check if report exists
        if not Path('codelist_report.xlsx').exists():
            messagebox.showerror("Error",
                "codelist_report.xlsx not found!\n\n"
                "Please run 'Extract Codelists' first.")
            return
        
        self.rename_btn.config(state='disabled')
        self.log_callback('INFO', "🔄 Renaming codelists...")
        
        def task():
            return self.backend.rename_codelists(input_file='codelist_report.xlsx')
        
        def callback(success, result):
            self.rename_btn.config(state='normal')
            if success:
                self.rename_completed = True  # Mark rename as completed
                self.log_callback('SUCCESS', "✅ Codelist renaming complete!")
                self.log_callback('INFO', "📄 Check rename_summary.txt for details")
                messagebox.showinfo("Success",
                    "Codelist renaming complete!\n\n"
                    "Files have been updated in-place with .backup files created.\n"
                    "Check rename_summary.txt for detailed changes.")
            else:
                self.log_callback('ERROR', f"❌ Renaming failed: {result}")
                messagebox.showerror("Error", f"Renaming failed: {result}")
        
        WorkerThread(task, callback).start()
    
    def open_report(self):
        """Open the report file directly from project folder"""
        # Check if extract operation has been completed
        if not self.extract_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Extract Codelists' first before opening the report.\n\n"
                "Click 'Extract Codelists' to process your files."
            )
            return
        
        report_file = Path("codelist_report.xlsx")
        if not report_file.exists():
            messagebox.showwarning("Not Found", "codelist_report.xlsx not found. Run 'Extract Codelists' first.")
            return
        
        import subprocess
        import platform
        
        try:
            # Open the file from project folder
            system = platform.system()
            if system == 'Darwin':  # macOS
                subprocess.run(['open', str(report_file)])
            elif system == 'Windows':
                subprocess.run(['start', '', str(report_file)], shell=True)
            else:  # Linux
                subprocess.run(['xdg-open', str(report_file)])
            
            self.log_callback('INFO', "📊 Opening codelist_report.xlsx from project folder...")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")
    
    def download_updated_files(self):
        """Download all updated MXL files to Downloads folder"""
        # Check if rename operation has been completed
        if not self.rename_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Rename Codelists' first before downloading files.\n\n"
                "Click 'Rename Codelists' to update your files."
            )
            return
        
        input_folder = self.folder_var.get()
        if not input_folder:
            messagebox.showerror("Error", "No input folder selected")
            return
        
        mxl_folder = Path(input_folder)
        if not mxl_folder.exists():
            messagebox.showerror("Error", "Input folder does not exist")
            return
        
        import subprocess
        import platform
        import shutil
        from datetime import datetime
        import zipfile
        
        try:
            # Get Downloads folder
            home = Path.home()
            downloads_folder = home / "Downloads"
            
            # Create timestamped ZIP filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_file = downloads_folder / f"updated_mxl_files_{timestamp}.zip"
            
            # Create ZIP file with all MXL files
            with zipfile.ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for mxl_file in mxl_folder.glob('*.mxl'):
                    zipf.write(mxl_file, mxl_file.name)
            
            self.log_callback('SUCCESS', f"✅ Downloaded to: {zip_file}")
            messagebox.showinfo("Success", f"Updated MXL files downloaded to:\n{zip_file}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to download files: {str(e)}")


class Phase1Tab(ttk.Frame):
    """Phase 1: Process Data Rules - Extract & Apply"""
    
    def __init__(self, parent, backend, log_callback, shared_folders):
        super().__init__(parent)
        self.backend = backend
        self.log_callback = log_callback
        self.shared_folders = shared_folders
        self.folder_browsed = False  # Track if user has browsed for a folder
        self.extract_completed = False  # Track if extract operation completed
        self.apply_completed = False  # Track if apply operation completed
        self.create_widgets()
        
    def create_widgets(self):
        # Title with emoji
        title = ttk.Label(self, text="📊 Phase 1: Process Data Rules - Extract & Apply", font=('Arial', 18, 'bold'))
        title.pack(pady=15)
        
        # Description
        desc = ttk.Label(self, text="Extract rules from MXL files, modify them, and apply back to MXL files",
                        wraplength=600, font=('Arial', 11))
        desc.pack(pady=8)
        
        # Folder selection
        folder_frame = ttk.LabelFrame(self, text="MXL Files Folder", padding=10)
        folder_frame.pack(fill='x', padx=20, pady=10)
        
        self.folder_var = tk.StringVar(value='')  # No default folder
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50)
        folder_entry.pack(side='left', padx=5)
        
        ttk.Button(folder_frame, text="📁 Browse...", command=self.browse_folder).pack(side='left', padx=5)
        ttk.Button(folder_frame, text="📁 Use Maps from Phase 0", command=self.use_phase0_maps).pack(side='left', padx=5)
        ttk.Button(folder_frame, text="📁 Use Maps from Phase 1", command=self.use_phase1_maps).pack(side='left', padx=5)
        
        # Step 1: Extract Rules
        step1_frame = ttk.LabelFrame(self, text="Step 1: Extract Rules", padding=15)
        step1_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(step1_frame, text="Extract process data rules from MXL files into Excel",
                 font=('Arial', 10)).pack(anchor='w', pady=5)
        
        ttk.Label(step1_frame, text="✅ The file will open directly from the project folder - just edit and save!",
                 foreground='green', font=('Arial', 9, 'bold')).pack(anchor='w', pady=2)
        
        extract_btn_frame = ttk.Frame(step1_frame)
        extract_btn_frame.pack(pady=10)
        
        self.extract_btn = ttk.Button(extract_btn_frame, text="🔍 Extract Rules", command=self.extract_rules)
        self.extract_btn.pack(side='left', padx=8)
        
        ttk.Button(extract_btn_frame, text="📝 Open Rules File",
                  command=self.open_rules_file).pack(side='left', padx=8)
        
        # Step 2: Apply Rules
        step2_frame = ttk.LabelFrame(self, text="Step 2: Apply Modified Rules", padding=15)
        step2_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(step2_frame, text="⚠️ Make sure you've edited and saved process_data_rules.xlsx before applying",
                 foreground='orange', font=('Arial', 10, 'bold')).pack(anchor='w', pady=5)
        
        ttk.Label(step2_frame, text="Apply the modified rules back to MXL files",
                 font=('Arial', 10)).pack(anchor='w', pady=5)
        
        apply_btn_frame = ttk.Frame(step2_frame)
        apply_btn_frame.pack(pady=10)
        
        self.apply_btn = ttk.Button(apply_btn_frame, text="⚙️ Apply Rules to MXL Files",
                                    command=self.apply_rules)
        self.apply_btn.pack(side='left', padx=8)
        
        ttk.Button(apply_btn_frame, text="📦 Download Updated MXL Files",
                  command=self.download_updated_mxl_files).pack(side='left', padx=8)
    
    def browse_folder(self):
        """Allow user to select MXL/ZIP files or folder"""
        # Open file dialog that allows selecting files
        files = filedialog.askopenfilenames(
            title="Select MXL or ZIP File(s) - or Cancel to select a folder",
            initialdir="mxl_files",
            filetypes=[
                ("Map Files", "*.mxl *.zip"),
                ("MXL Files", "*.mxl"),
                ("ZIP Files", "*.zip"),
                ("All Files", "*.*")
            ]
        )
        
        if files:
            # Files were selected
            folder = str(Path(files[0]).parent)
            if len(files) == 1:
                self.folder_var.set(str(Path(files[0])))
                self.log_callback('INFO', f"✅ Selected file: {Path(files[0]).name}")
            else:
                self.folder_var.set(f"{folder} ({len(files)} files)")
                self.log_callback('INFO', f"✅ Selected {len(files)} files from: {folder}")
            self.shared_folders['phase2'] = folder
        else:
            # No files selected, offer folder selection
            folder = filedialog.askdirectory(title="Select Folder with MXL/ZIP Files", initialdir="mxl_files")
            if folder:
                self.folder_var.set(folder)
                self.shared_folders['phase2'] = folder
                self.log_callback('INFO', f"✅ Selected folder: {folder}")
    
    def use_phase0_maps(self):
        """Use the folder from Phase 0"""
        phase0_folder = self.shared_folders.get('phase0', '')
        if not phase0_folder:
            messagebox.showwarning(
                "Phase 0 Not Run",
                "⚠️ Please run Phase 0 first to extract details.\n\n"
                "Phase 0 must be completed before you can use its maps."
            )
            return
        
        self.folder_var.set(phase0_folder)
        self.log_callback('INFO', f"✅ Using maps from Phase 0: {phase0_folder}")
        messagebox.showinfo("Success", f"Now using maps from Phase 0:\n{phase0_folder}")
    
    def use_phase1_maps(self):
        """Use the folder from Phase 1"""
        phase1_folder = self.shared_folders.get('phase1', '')
        if not phase1_folder:
            messagebox.showwarning(
                "Phase 1 Not Run",
                "⚠️ Please run Phase 1 (Codelist Management) first.\n\n"
                "Phase 1 must be completed before you can use its maps."
            )
            return
        
        self.folder_var.set(phase1_folder)
        self.log_callback('INFO', f"✅ Using maps from Phase 1: {phase1_folder}")
        messagebox.showinfo("Success", f"Now using maps from Phase 1:\n{phase1_folder}")
        
    def extract_rules(self):
        mxl_folder = self.folder_var.get()
        if not mxl_folder or not Path(mxl_folder).exists():
            messagebox.showerror("Error", "Please select a valid MXL files folder")
            return
        
        self.extract_completed = False  # Reset flag
        self.extract_btn.config(state='disabled')
        self.log_callback('INFO', "🔍 Extracting process data rules...")
        
        def task():
            return self.backend.extract_rules(mxl_folder=mxl_folder)
        
        def callback(success, result):
            self.extract_btn.config(state='normal')
            if success:
                self.extract_completed = True  # Mark extract as completed
                self.log_callback('SUCCESS', "✅ Rules extracted successfully!")
                messagebox.showinfo("Success",
                    "Process data rules extracted successfully!\n\n"
                    "Next steps:\n"
                    "1. Click 'Open Rules File' to view the Excel file\n"
                    "2. Modify the rules as needed\n"
                    "3. Save the file\n"
                    "4. Click 'Apply Rules to MXL Files'")
            else:
                self.log_callback('ERROR', f"❌ Extraction failed: {result}")
                messagebox.showerror("Error", f"Extraction failed: {result}")
        
        WorkerThread(task, callback).start()
    
    def apply_rules(self):
        # Check if extract operation has been completed
        if not self.extract_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Extract Rules' first before applying rules.\n\n"
                "Click 'Extract Rules' to process your files."
            )
            return
        
        input_folder = self.folder_var.get()
        
        mxl_folder = input_folder
        if not Path(mxl_folder).exists():
            messagebox.showerror("Error", "Selected folder does not exist")
            return
        
        if not Path('process_data_rules.xlsx').exists():
            messagebox.showerror("Error",
                "process_data_rules.xlsx not found!\n\n"
                "Please extract rules first.")
            return
        
        response = messagebox.askyesno("Confirm",
            "Have you modified and saved process_data_rules.xlsx?\n\n"
            "This will update your MXL files with the modified rules.")
        if not response:
            return
            
        self.apply_btn.config(state='disabled')
        self.log_callback('INFO', "⚙️ Applying rules to MXL files...")
        
        def task():
            return self.backend.apply_rules(
                rules_excel='process_data_rules.xlsx',
                mxl_folder=mxl_folder
            )
        
        def callback(success, result):
            self.apply_btn.config(state='normal')
            if success:
                self.apply_completed = True  # Mark apply as completed
                self.log_callback('SUCCESS', "✅ Rules applied successfully!")
                messagebox.showinfo("Success",
                    "Rules applied successfully to MXL files!\n\n"
                    "You can now download the updated MXL files.")
            else:
                self.log_callback('ERROR', f"❌ Apply failed: {result}")
                messagebox.showerror("Error", f"Apply failed: {result}")
        
        WorkerThread(task, callback).start()
        
    def open_rules_file(self):
        """Open the rules file directly from the project folder"""
        # Check if extract operation has been completed
        if not self.extract_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Extract Rules' first before opening the rules file.\n\n"
                "Click 'Extract Rules' to process your files."
            )
            return
        
        rules_file = Path("process_data_rules.xlsx")
        
        # Check if rules file exists
        if not rules_file.exists():
            messagebox.showwarning("Not Found",
                "process_data_rules.xlsx not found.\n\n"
                "Please extract rules first.")
            return
        
        import subprocess
        import platform
        
        try:
            # Open the file directly from project folder
            system = platform.system()
            if system == 'Darwin':  # macOS
                subprocess.run(['open', str(rules_file)])
            elif system == 'Windows':
                subprocess.run(['start', '', str(rules_file)], shell=True)
            else:  # Linux
                subprocess.run(['xdg-open', str(rules_file)])
            
            self.log_callback('INFO', "📝 Opening process_data_rules.xlsx from project folder")
            messagebox.showinfo("Success",
                "Opening process_data_rules.xlsx\n\n"
                "✅ Edit and save the file directly\n"
                "✅ Changes will be applied when you click 'Apply Rules'")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")
    
    def download_updated_mxl_files(self):
        """Download all updated MXL files to Downloads folder"""
        # Check if apply operation has been completed
        if not self.apply_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Apply Rules to MXL Files' first before downloading.\n\n"
                "Click 'Apply Rules to MXL Files' to update your files."
            )
            return
        
        mxl_folder = self.folder_var.get()
        if not mxl_folder or not Path(mxl_folder).exists():
            messagebox.showerror("Error", "Please select a valid MXL files folder")
            return
        
        import subprocess
        import platform
        import shutil
        from datetime import datetime
        import zipfile
        
        try:
            # Get Downloads folder
            home = Path.home()
            downloads_folder = home / "Downloads"
            
            # Create timestamped zip filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"updated_mxl_files_{timestamp}.zip"
            zip_path = downloads_folder / zip_filename
            
            # Create zip file with all MXL files
            mxl_path = Path(mxl_folder)
            mxl_files = list(mxl_path.glob('*.mxl'))
            
            if not mxl_files:
                messagebox.showwarning("No Files", "No MXL files found in the selected folder")
                return
            
            self.log_callback('INFO', f"📦 Creating zip file with {len(mxl_files)} MXL files...")
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for mxl_file in mxl_files:
                    zipf.write(mxl_file, mxl_file.name)
            
            self.log_callback('SUCCESS', f"✅ Downloaded to: {zip_path}")
            
            # Ask if user wants to open the Downloads folder
            response = messagebox.askyesno("Success",
                f"Updated MXL files downloaded successfully!\n\n"
                f"File: {zip_filename}\n"
                f"Location: {downloads_folder}\n\n"
                f"Do you want to open the Downloads folder?")
            
            if response:
                system = platform.system()
                if system == 'Darwin':  # macOS
                    subprocess.run(['open', str(downloads_folder)])
                elif system == 'Windows':
                    subprocess.run(['explorer', str(downloads_folder)])
                else:  # Linux
                    subprocess.run(['xdg-open', str(downloads_folder)])
            
        except Exception as e:
            self.log_callback('ERROR', f"❌ Download failed: {str(e)}")
            messagebox.showerror("Error", f"Failed to download MXL files: {str(e)}")



class Phase3Tab(ttk.Frame):
    """Phase 3: Process All Maps"""
    
    def __init__(self, parent, backend, log_callback, shared_folders):
        super().__init__(parent)
        self.backend = backend
        self.log_callback = log_callback
        self.shared_folders = shared_folders
        self.folder_browsed = False  # Track if user has browsed for a folder
        self.operation_completed = False  # Track if operation completed
        self.create_widgets()
        
    def create_widgets(self):
        # Title
        title = ttk.Label(self, text="📋 Phase 3: Process All Maps", font=('Arial', 18, 'bold'))
        title.pack(pady=15)
        
        # Description
        desc = ttk.Label(self, text="Complete automation - Process all MXL files with all features",
                        wraplength=600, font=('Arial', 11))
        desc.pack(pady=8)
        
        # Download Checklist Section
        checklist_frame = ttk.LabelFrame(self, text="Step 1: Download & Configure Checklist", padding=15)
        checklist_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(checklist_frame, text="Download the Generic Checklist, configure it, and save before processing",
                 font=('Arial', 10)).pack(anchor='w', pady=5)
        
        checklist_btn_frame = ttk.Frame(checklist_frame)
        checklist_btn_frame.pack(pady=10)
        
        ttk.Button(checklist_btn_frame, text="📥 Download Checklist",
                  command=self.download_checklist).pack(side='left', padx=5)
        ttk.Button(checklist_btn_frame, text="📂 Open Checklist",
                  command=self.open_checklist).pack(side='left', padx=5)
        
        # Info box
        info_frame = ttk.Frame(self)
        info_frame.pack(fill='x', padx=20, pady=10)
        info_text = """💡 This phase will apply ALL transformations to your MXL files:
• Generic_Rule tab updates
• Process data tab updates
• Inbound Maps features
• Outbound Maps features"""
        ttk.Label(info_frame, text=info_text, foreground='blue', font=('Arial', 10, 'italic'), justify='left').pack(anchor='w')
        
        # Folder selection
        folder_frame = ttk.LabelFrame(self, text="Step 2: Select MXL Files Folder", padding=15)
        folder_frame.pack(fill='x', padx=20, pady=10)
        
        self.folder_var = tk.StringVar(value='')  # No default folder
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50)
        folder_entry.pack(side='left', padx=5)
        
        ttk.Button(folder_frame, text="📁 Browse...", command=self.browse_folder).pack(side='left', padx=5)
        ttk.Button(folder_frame, text="📁 Use Maps from Phase 0", command=self.use_phase0_maps).pack(side='left', padx=5)
        ttk.Button(folder_frame, text="📁 Use Maps from Phase 1", command=self.use_phase1_maps).pack(side='left', padx=5)
        ttk.Button(folder_frame, text="📁 Use Maps from Phase 2", command=self.use_phase2_maps).pack(side='left', padx=5)
        
        # Process Section
        process_frame = ttk.LabelFrame(self, text="Step 3: Process Maps", padding=15)
        process_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(process_frame, text="⚠️ Ensure you have configured and saved the checklist before processing",
                 foreground='orange', font=('Arial', 10, 'bold')).pack(anchor='w', pady=5)
        
        btn_frame = ttk.Frame(process_frame)
        btn_frame.pack(pady=15)
        
        self.run_btn = ttk.Button(btn_frame, text="🚀 Process All Maps", command=self.run_phase4)
        self.run_btn.pack(side='left', padx=8)
        
        ttk.Button(btn_frame, text="📦 Download Updated Files",
                  command=self.download_updated_files).pack(side='left', padx=8)
        
        ttk.Button(btn_frame, text="📄 Download Error Report",
                  command=self.download_error_report).pack(side='left', padx=8)
    
    def browse_folder(self):
        """Allow user to select MXL/ZIP files or folder"""
        # Open file dialog that allows selecting files
        files = filedialog.askopenfilenames(
            title="Select MXL or ZIP File(s) - or Cancel to select a folder",
            initialdir="mxl_files",
            filetypes=[
                ("Map Files", "*.mxl *.zip"),
                ("MXL Files", "*.mxl"),
                ("ZIP Files", "*.zip"),
                ("All Files", "*.*")
            ]
        )
        
        if files:
            # Files were selected
            folder = str(Path(files[0]).parent)
            if len(files) == 1:
                self.folder_var.set(str(Path(files[0])))
                self.log_callback('INFO', f"✅ Selected file: {Path(files[0]).name}")
            else:
                self.folder_var.set(f"{folder} ({len(files)} files)")
                self.log_callback('INFO', f"✅ Selected {len(files)} files from: {folder}")
            self.shared_folders['phase3'] = folder
            self.folder_browsed = True
        else:
            # No files selected, offer folder selection
            folder = filedialog.askdirectory(title="Select Folder with MXL Files", initialdir="mxl_files")
            if folder:
                self.folder_var.set(folder)
                self.shared_folders['phase3'] = folder
                self.folder_browsed = True
                self.log_callback('INFO', f"✅ Selected folder: {folder}")
    
    def download_checklist(self):
        """Download the Generic Checklist to Downloads folder"""
        checklist_file = Path("Generic_checklistMain.xlsm")
        
        if not checklist_file.exists():
            messagebox.showerror("Error",
                "Generic_checklistMain.xlsm not found in project folder.\n\n"
                "Please ensure the checklist file is in the project directory.")
            return
        
        import shutil
        from datetime import datetime
        
        try:
            # Get Downloads folder
            home = Path.home()
            downloads_folder = home / "Downloads"
            
            # Copy file to Downloads
            dest_file = downloads_folder / "Generic_checklistMain.xlsm"
            shutil.copy2(checklist_file, dest_file)
            
            self.log_callback('SUCCESS', f"✅ Checklist downloaded to: {dest_file}")
            messagebox.showinfo("Success",
                f"Generic Checklist downloaded successfully!\n\n"
                f"Location: {dest_file}\n\n"
                f"Next steps:\n"
                f"1. Click 'Open Checklist' to configure it\n"
                f"2. Make your changes and save\n"
                f"3. The saved file will be used for processing")
            
        except Exception as e:
            self.log_callback('ERROR', f"❌ Download failed: {str(e)}")
            messagebox.showerror("Error", f"Failed to download checklist: {str(e)}")
    
    def open_checklist(self):
        """Open the checklist file from project folder"""
        checklist_file = Path("Generic_checklistMain.xlsm")
        
        if not checklist_file.exists():
            messagebox.showwarning("Not Found",
                "Generic_checklistMain.xlsm not found.\n\n"
                "Please download the checklist first.")
            return
        
        import subprocess
        import platform
        
        try:
            system = platform.system()
            if system == 'Darwin':  # macOS
                subprocess.run(['open', str(checklist_file)])
            elif system == 'Windows':
                subprocess.run(['start', '', str(checklist_file)], shell=True)
            else:  # Linux
                subprocess.run(['xdg-open', str(checklist_file)])
            
            self.log_callback('INFO', "📂 Opening Generic_checklistMain.xlsm from project folder")
            messagebox.showinfo("Info",
                "Opening Generic_checklistMain.xlsm\n\n"
                "✅ Edit and save the file directly\n"
                "✅ Changes will be used when you click 'Process All Maps'")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")
    
    def use_phase0_maps(self):
        """Use the folder from Phase 0"""
        phase0_folder = self.shared_folders.get('phase0', '')
        if not phase0_folder:
            messagebox.showwarning(
                "Phase 0 Not Run",
                "⚠️ Please run Phase 0 first to extract details.\n\n"
                "Phase 0 must be completed before you can use its maps."
            )
            return
        
        self.folder_var.set(phase0_folder)
        self.folder_browsed = True
        self.log_callback('INFO', f"✅ Using maps from Phase 0: {phase0_folder}")
        messagebox.showinfo("Success", f"Now using maps from Phase 0:\n{phase0_folder}")
    
    def use_phase1_maps(self):
        """Use the folder from Phase 1"""
        phase1_folder = self.shared_folders.get('phase1', '')
        if not phase1_folder:
            messagebox.showwarning(
                "Phase 1 Not Run",
                "⚠️ Please run Phase 1 (Codelist Management) first.\n\n"
                "Phase 1 must be completed before you can use its maps."
            )
            return
        
        self.folder_var.set(phase1_folder)
        self.folder_browsed = True
        self.log_callback('INFO', f"✅ Using maps from Phase 1: {phase1_folder}")
        messagebox.showinfo("Success", f"Now using maps from Phase 1:\n{phase1_folder}")
    
    def use_phase2_maps(self):
        """Use the folder from Phase 2"""
        phase2_folder = self.shared_folders.get('phase2', '')
        if not phase2_folder:
            messagebox.showwarning(
                "Phase 2 Not Run",
                "⚠️ Please run Phase 2 (Process Data Rules) first.\n\n"
                "Phase 2 must be completed before you can use its maps."
            )
            return
        
        self.folder_var.set(phase2_folder)
        self.folder_browsed = True
        self.log_callback('INFO', f"✅ Using maps from Phase 2: {phase2_folder}")
        messagebox.showinfo("Success", f"Now using maps from Phase 2:\n{phase2_folder}")
        
    def run_phase4(self):
        # Check if user has browsed for a folder
        if not self.folder_browsed:
            messagebox.showwarning(
                "Folder Not Selected",
                "⚠️ Please browse and select a folder before processing maps.\n\n"
                "Click the 'Browse...' button to select your MXL files folder."
            )
            return
        
        mxl_folder = self.folder_var.get()
        if not mxl_folder or not Path(mxl_folder).exists():
            messagebox.showerror("Error", "Please select a valid MXL files folder")
            return
        
        # Check required files
        if not Path('mapping_results.xlsx').exists():
            messagebox.showerror("Error", "mapping_results.xlsx not found. Run Phase 0 first.")
            return
        if not Path('Generic_checklistMain.xlsm').exists():
            messagebox.showerror("Error", "Generic_checklistMain.xlsm not found.")
            return
        
        response = messagebox.askyesno("Confirm",
            f"This will process ALL MXL files in {mxl_folder} with ALL features.\n\nContinue?")
        if not response:
            return
        
        self.run_btn.config(state='disabled')
        self.log_callback('INFO', "🚀 Starting Phase 4 - Processing all maps...")
        
        def task():
            return self.backend.process_all_maps(
                mapping_results='mapping_results.xlsx',
                checklist='Generic_checklistMain.xlsm',
                mxl_folder=mxl_folder
            )
        
        def callback(success, result):
            self.run_btn.config(state='normal')
            if success:
                self.operation_completed = True  # Mark operation as completed
                self.log_callback('SUCCESS', "✅ Phase 3 completed successfully!")
                messagebox.showinfo("Success",
                    "All maps processed successfully!\n\n"
                    "You can now download the updated MXL files.")
            else:
                self.log_callback('ERROR', f"❌ Phase 4 failed: {result}")
                messagebox.showerror("Error", f"Phase 4 failed: {result}")
        
        WorkerThread(task, callback).start()
    
    def download_updated_files(self):
        """Download all updated MXL files to Downloads folder"""
        # Check if operation has been completed
        if not self.operation_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Process All Maps' first before downloading.\n\n"
                "Click 'Process All Maps' to update your files."
            )
            return
        
        mxl_folder = self.folder_var.get()
        if not mxl_folder or not Path(mxl_folder).exists():
            messagebox.showerror("Error", "Please select a valid MXL files folder")
            return
        
        import subprocess
        import platform
        import shutil
        from datetime import datetime
        import zipfile
        
        try:
            # Get Downloads folder
            home = Path.home()
            downloads_folder = home / "Downloads"
            
            # Create timestamped zip filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"phase3_updated_mxl_files_{timestamp}.zip"
            zip_path = downloads_folder / zip_filename
            
            # Create zip file with all MXL files
            mxl_path = Path(mxl_folder)
            mxl_files = list(mxl_path.glob('*.mxl'))
            
            if not mxl_files:
                messagebox.showwarning("No Files", "No MXL files found in the selected folder")
                return
            
            self.log_callback('INFO', f"📦 Creating zip file with {len(mxl_files)} MXL files...")
            
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for mxl_file in mxl_files:
                    zipf.write(mxl_file, mxl_file.name)
            
            self.log_callback('SUCCESS', f"✅ Downloaded to: {zip_path}")
            
            # Ask if user wants to open the Downloads folder
            response = messagebox.askyesno("Success",
                f"Updated MXL files downloaded successfully!\n\n"
                f"File: {zip_filename}\n"
                f"Location: {downloads_folder}\n\n"
                f"Do you want to open the Downloads folder?")
            
            if response:
                system = platform.system()
                if system == 'Darwin':  # macOS
                    subprocess.run(['open', str(downloads_folder)])
                elif system == 'Windows':
                    subprocess.run(['explorer', str(downloads_folder)])
                else:  # Linux
                    subprocess.run(['xdg-open', str(downloads_folder)])
            
        except Exception as e:
            self.log_callback('ERROR', f"❌ Download failed: {str(e)}")
            messagebox.showerror("Error", f"Failed to download MXL files: {str(e)}")
    
    def download_error_report(self):
        """Download the error report to Downloads folder"""
        # Check if operation has been completed
        if not self.operation_completed:
            messagebox.showwarning(
                "Operation Not Run",
                "⚠️ Please run 'Process All Maps' first before downloading error report.\n\n"
                "The error report is generated only after processing maps."
            )
            return
        
        # Check if error report exists
        error_report = Path("error_report.xlsx")
        if not error_report.exists():
            messagebox.showinfo("No Errors",
                "✅ No error report found!\n\n"
                "This means all maps were processed successfully without errors.")
            return
        
        import subprocess
        import platform
        import shutil
        from datetime import datetime
        
        try:
            # Get Downloads folder
            home = Path.home()
            downloads_folder = home / "Downloads"
            
            # Create timestamped filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest_file = downloads_folder / f"error_report_{timestamp}.xlsx"
            
            # Copy error report to Downloads
            shutil.copy2(error_report, dest_file)
            
            self.log_callback('SUCCESS', f"✅ Error report downloaded to: {dest_file}")
            
            # Ask if user wants to open the file
            response = messagebox.askyesno("Success",
                f"Error report downloaded successfully!\n\n"
                f"File: error_report_{timestamp}.xlsx\n"
                f"Location: {downloads_folder}\n\n"
                f"Do you want to open the error report?")
            
            if response:
                system = platform.system()
                if system == 'Darwin':  # macOS
                    subprocess.run(['open', str(dest_file)])
                elif system == 'Windows':
                    subprocess.run(['start', '', str(dest_file)], shell=True)
                else:  # Linux
                    subprocess.run(['xdg-open', str(dest_file)])
            
        except Exception as e:
            self.log_callback('ERROR', f"❌ Download failed: {str(e)}")
            messagebox.showerror("Error", f"Failed to download error report: {str(e)}")


        
class MainApplication(tk.Tk):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        
        self.title("Sterling B2Bi Map Migration Accelerator")
        self.geometry("1000x700")
        
        # Initialize backend
        self.backend = BackendAPI()
        
        # Shared folder tracking across phases
        self.shared_folders = {
            'phase0': '',  # Folder used in Phase 0
            'phase1': '',  # Folder used in Phase 1
            'phase2': ''   # Folder used in Phase 2
        }
        
        # Create main container
        main_container = ttk.PanedWindow(self, orient='vertical')
        main_container.pack(fill='both', expand=True)
        
        # Top section: Tabs
        tab_frame = ttk.Frame(main_container)
        main_container.add(tab_frame, weight=3)
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(tab_frame)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Add tabs (Setup tab removed - checklist download moved to Phase 3)
        self.phase0_tab = Phase0Tab(self.notebook, self.backend, self.log_message, self.shared_folders)
        self.notebook.add(self.phase0_tab, text="Phase 0: Extract Details")
        
        self.codelist_tab = CodelistTab(self.notebook, self.backend, self.log_message, self.shared_folders)
        self.notebook.add(self.codelist_tab, text="Phase 1: Codelists")
        
        self.phase2_tab = Phase1Tab(self.notebook, self.backend, self.log_message, self.shared_folders)
        self.notebook.add(self.phase2_tab, text="Phase 2: Process Data Rules")
        
        self.phase3_tab = Phase3Tab(self.notebook, self.backend, self.log_message, self.shared_folders)
        self.notebook.add(self.phase3_tab, text="Phase 3: Process All")
        
        # Bottom section: Log viewer
        log_frame = ttk.LabelFrame(main_container, text="Logs", padding=5)
        main_container.add(log_frame, weight=1)
        
        # Log viewer with scrollbar
        self.log_viewer = LogViewer(log_frame, height=10)
        self.log_viewer.pack(fill='both', expand=True)
        
        # Clear logs button
        btn_frame = ttk.Frame(log_frame)
        btn_frame.pack(fill='x', pady=5)
        ttk.Button(btn_frame, text="Clear Logs", command=self.clear_logs).pack(side='right')
        
        # Initial log message
        self.log_message('INFO', "Application started. Ready to process maps.")
        
    def log_message(self, level, message):
        """Add a log message to the log viewer"""
        self.log_viewer.append_log(level, message)
        
    def clear_logs(self):
        """Clear all logs"""
        self.log_viewer.clear_logs()
        self.log_message('INFO', "Logs cleared.")


def main():
    """Main entry point"""
    app = MainApplication()
    app.mainloop()


if __name__ == "__main__":
    main()

# Made with Bob
