"""
PyQt6 Desktop Application Template
AI-Powered Map Migration Accelerator

This is a complete template that frontend developers can use to build the UI.
All backend integration is handled through backend_api.py
"""

import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog,
    QRadioButton, QCheckBox, QProgressBar, QListWidget, QGroupBox,
    QMessageBox, QComboBox, QSplitter
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QTextCursor, QColor
from pathlib import Path
import os

# Import backend API
from backend_api import BackendAPI


class WorkerThread(QThread):
    """Worker thread for running backend operations without freezing UI"""
    log_signal = pyqtSignal(str, str)  # level, message
    finished_signal = pyqtSignal(bool)  # success
    
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
            self.log_signal.emit('ERROR', f'Exception: {str(e)}')
            self.finished_signal.emit(False)


class LogViewer(QTextEdit):
    """Custom log viewer with color-coded messages"""
    
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setFont(QFont('Courier', 9))
        
        # Color scheme
        self.colors = {
            'INFO': QColor(33, 150, 243),      # Blue
            'SUCCESS': QColor(76, 175, 80),    # Green
            'WARNING': QColor(255, 152, 0),    # Orange
            'ERROR': QColor(244, 67, 54)       # Red
        }
    
    def append_log(self, level: str, message: str):
        """Append a log message with appropriate color"""
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        
        # Set color based on level
        color = self.colors.get(level, QColor(0, 0, 0))
        self.setTextColor(color)
        
        # Append message
        self.append(message)
        
        # Auto-scroll to bottom
        self.verticalScrollBar().setValue(self.verticalScrollBar().maximum())


class Phase0Tab(QWidget):
    """Phase 0: Extract Map Details"""
    
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Phase 0: Extract Map Details")
        title.setFont(QFont('Arial', 14, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # Input Source
        source_group = QGroupBox("Input Source")
        source_layout = QHBoxLayout()
        self.radio_zip = QRadioButton("ZIP Files")
        self.radio_mxl = QRadioButton("MXL Files")
        self.radio_zip.setChecked(True)
        source_layout.addWidget(self.radio_zip)
        source_layout.addWidget(self.radio_mxl)
        source_group.setLayout(source_layout)
        layout.addWidget(source_group)
        
        # Input Folder
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("Input Folder:"))
        self.input_folder = QLineEdit()
        self.input_folder.setPlaceholderText("Select input folder...")
        folder_layout.addWidget(self.input_folder)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self.browse_input_folder)
        folder_layout.addWidget(btn_browse)
        layout.addLayout(folder_layout)
        
        # Output Excel
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Output Excel:"))
        self.output_excel = QLineEdit("mapping_results.xlsx")
        excel_layout.addWidget(self.output_excel)
        layout.addLayout(excel_layout)
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        self.chk_copy = QCheckBox("Copy to mxl_files folder")
        self.chk_copy.setChecked(True)
        self.chk_skip = QCheckBox("Skip duplicates")
        self.chk_skip.setChecked(True)
        self.chk_archive = QCheckBox("Save to old_mxlFiles (archive)")
        self.chk_archive.setChecked(True)
        options_layout.addWidget(self.chk_copy)
        options_layout.addWidget(self.chk_skip)
        options_layout.addWidget(self.chk_archive)
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_run = QPushButton("Run Extraction")
        self.btn_run.clicked.connect(self.run_extraction)
        self.btn_clear = QPushButton("Clear")
        self.btn_clear.clicked.connect(self.clear_fields)
        btn_layout.addWidget(self.btn_run)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def browse_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Input Folder")
        if folder:
            self.input_folder.setText(folder)
    
    def run_extraction(self):
        input_folder = self.input_folder.text()
        if not input_folder:
            QMessageBox.warning(self, "Error", "Please select an input folder")
            return
        
        input_source = 'zip' if self.radio_zip.isChecked() else 'mxl'
        output_excel = self.output_excel.text()
        
        # Run in worker thread
        self.parent.run_backend_task(
            self.parent.api.extract_map_details,
            input_source, input_folder, output_excel
        )
    
    def clear_fields(self):
        self.input_folder.clear()
        self.output_excel.setText("mapping_results.xlsx")


class Phase1Tab(QWidget):
    """Phase 1: Extract Rules"""
    
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Phase 1: Extract Process Data Rules")
        title.setFont(QFont('Arial', 14, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # MXL Folder
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("MXL Files Folder:"))
        self.mxl_folder = QLineEdit("mxl_files")
        folder_layout.addWidget(self.mxl_folder)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self.browse_mxl_folder)
        folder_layout.addWidget(btn_browse)
        layout.addLayout(folder_layout)
        
        # Output Excel
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Output Excel:"))
        self.output_excel = QLineEdit("process_data_rules.xlsx")
        excel_layout.addWidget(self.output_excel)
        layout.addLayout(excel_layout)
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_run = QPushButton("Run Extraction")
        self.btn_run.clicked.connect(self.run_extraction)
        btn_layout.addWidget(self.btn_run)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # Instructions
        instructions = QLabel(
            "Instructions:\n"
            "1. Click 'Run Extraction' to generate process_data_rules.xlsx\n"
            "2. Download and open the Excel file\n"
            "3. Add comments/uncomments as needed\n"
            "4. Upload in Phase 2 to apply changes"
        )
        instructions.setStyleSheet("background-color: #E3F2FD; padding: 10px; border-radius: 5px;")
        layout.addWidget(instructions)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def browse_mxl_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select MXL Folder")
        if folder:
            self.mxl_folder.setText(folder)
    
    def run_extraction(self):
        mxl_folder = self.mxl_folder.text()
        output_excel = self.output_excel.text()
        
        self.parent.run_backend_task(
            self.parent.api.extract_rules,
            mxl_folder, output_excel
        )


class Phase2Tab(QWidget):
    """Phase 2: Apply Rules"""
    
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Phase 2: Apply Process Data Rules")
        title.setFont(QFont('Arial', 14, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # Rules Excel
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(QLabel("Rules Excel:"))
        self.rules_excel = QLineEdit()
        self.rules_excel.setPlaceholderText("Upload your modified process_data_rules.xlsx")
        excel_layout.addWidget(self.rules_excel)
        btn_upload = QPushButton("Upload")
        btn_upload.clicked.connect(self.upload_rules)
        excel_layout.addWidget(btn_upload)
        layout.addLayout(excel_layout)
        
        # MXL Folder
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("MXL Files Folder:"))
        self.mxl_folder = QLineEdit("mxl_files")
        folder_layout.addWidget(self.mxl_folder)
        btn_browse = QPushButton("Browse...")
        btn_browse.clicked.connect(self.browse_mxl_folder)
        folder_layout.addWidget(btn_browse)
        layout.addLayout(folder_layout)
        
        # Options
        options_group = QGroupBox("Apply Options")
        options_layout = QVBoxLayout()
        self.chk_backup = QCheckBox("Backup files before changes")
        self.chk_backup.setChecked(True)
        self.chk_validate = QCheckBox("Validate rules before applying")
        self.chk_validate.setChecked(True)
        options_layout.addWidget(self.chk_backup)
        options_layout.addWidget(self.chk_validate)
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_apply = QPushButton("Apply Changes")
        self.btn_apply.clicked.connect(self.apply_changes)
        btn_layout.addWidget(self.btn_apply)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        layout.addStretch()
        self.setLayout(layout)
    
    def upload_rules(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Rules Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if file:
            self.rules_excel.setText(file)
    
    def browse_mxl_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select MXL Folder")
        if folder:
            self.mxl_folder.setText(folder)
    
    def apply_changes(self):
        rules_excel = self.rules_excel.text()
        mxl_folder = self.mxl_folder.text()
        
        if not rules_excel:
            QMessageBox.warning(self, "Error", "Please upload rules Excel file")
            return
        
        self.parent.run_backend_task(
            self.parent.api.apply_rules,
            rules_excel, mxl_folder
        )


class Phase3Tab(QWidget):
    """Phase 3: Process Maps"""
    
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        
        # Title
        title = QLabel("Phase 3: Process All Maps")
        title.setFont(QFont('Arial', 14, QFont.Weight.Bold))
        layout.addWidget(title)
        
        # File inputs
        self.add_file_input(layout, "Mapping Results:", "mapping_results.xlsx", "mapping_results")
        self.add_file_input(layout, "Generic Checklist:", "Generic_checklistMain.xlsm", "checklist")
        self.add_folder_input(layout, "MXL Files Folder:", "mxl_files", "mxl_folder")
        
        # File list
        list_group = QGroupBox("Select Files to Process")
        list_layout = QVBoxLayout()
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        list_layout.addWidget(self.file_list)
        
        btn_list_layout = QHBoxLayout()
        btn_refresh = QPushButton("Refresh List")
        btn_refresh.clicked.connect(self.refresh_file_list)
        btn_select_all = QPushButton("Select All")
        btn_select_all.clicked.connect(self.select_all_files)
        btn_list_layout.addWidget(btn_refresh)
        btn_list_layout.addWidget(btn_select_all)
        list_layout.addLayout(btn_list_layout)
        
        list_group.setLayout(list_layout)
        layout.addWidget(list_group)
        
        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_process_all = QPushButton("Process All")
        self.btn_process_all.clicked.connect(self.process_all)
        self.btn_process_selected = QPushButton("Process Selected")
        self.btn_process_selected.clicked.connect(self.process_selected)
        btn_layout.addWidget(self.btn_process_all)
        btn_layout.addWidget(self.btn_process_selected)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        
        # Load initial file list
        self.refresh_file_list()
    
    def add_file_input(self, layout, label, default, attr_name):
        row = QHBoxLayout()
        row.addWidget(QLabel(label))
        line_edit = QLineEdit(default)
        setattr(self, attr_name, line_edit)
        row.addWidget(line_edit)
        btn = QPushButton("Browse...")
        btn.clicked.connect(lambda: self.browse_file(line_edit))
        row.addWidget(btn)
        layout.addLayout(row)
    
    def add_folder_input(self, layout, label, default, attr_name):
        row = QHBoxLayout()
        row.addWidget(QLabel(label))
        line_edit = QLineEdit(default)
        setattr(self, attr_name, line_edit)
        row.addWidget(line_edit)
        btn = QPushButton("Browse...")
        btn.clicked.connect(lambda: self.browse_folder(line_edit))
        row.addWidget(btn)
        layout.addLayout(row)
    
    def browse_file(self, line_edit):
        file, _ = QFileDialog.getOpenFileName(self, "Select File")
        if file:
            line_edit.setText(file)
    
    def browse_folder(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            line_edit.setText(folder)
    
    def refresh_file_list(self):
        self.file_list.clear()
        mxl_folder = self.mxl_folder.text()
        files = self.parent.api.get_mxl_files(mxl_folder)
        self.file_list.addItems(files)
    
    def select_all_files(self):
        for i in range(self.file_list.count()):
            self.file_list.item(i).setSelected(True)
    
    def process_all(self):
        self.parent.run_backend_task(
            self.parent.api.process_all_maps,
            self.mapping_results.text(),
            self.checklist.text(),
            self.mxl_folder.text()
        )
    
    def process_selected(self):
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Please select files to process")
            return
        
        selected_files = [item.text() for item in selected_items]
        self.parent.run_backend_task(
            self.parent.api.process_all_maps,
            self.mapping_results.text(),
            self.checklist.text(),
            self.mxl_folder.text(),
            selected_files
        )


class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.api = BackendAPI(log_callback=self.handle_log)
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("AI-Powered Map Migration Accelerator")
        self.setGeometry(100, 100, 1200, 800)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout with splitter
        main_layout = QVBoxLayout()
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Top section: Tabs
        self.tabs = QTabWidget()
        self.tabs.addTab(Phase0Tab(self), "Phase 0: Extract Details")
        self.tabs.addTab(Phase1Tab(self), "Phase 1: Extract Rules")
        self.tabs.addTab(Phase2Tab(self), "Phase 2: Apply Rules")
        self.tabs.addTab(Phase3Tab(self), "Phase 3: Process Maps")
        splitter.addWidget(self.tabs)
        
        # Bottom section: Log viewer
        log_widget = QWidget()
        log_layout = QVBoxLayout()
        log_label = QLabel("Log Output")
        log_label.setFont(QFont('Arial', 10, QFont.Weight.Bold))
        log_layout.addWidget(log_label)
        
        self.log_viewer = LogViewer()
        log_layout.addWidget(self.log_viewer)
        
        # Log controls
        log_controls = QHBoxLayout()
        btn_clear_log = QPushButton("Clear Logs")
        btn_clear_log.clicked.connect(self.log_viewer.clear)
        btn_export_log = QPushButton("Export Logs")
        btn_export_log.clicked.connect(self.export_logs)
        log_controls.addWidget(btn_clear_log)
        log_controls.addWidget(btn_export_log)
        log_controls.addStretch()
        log_layout.addLayout(log_controls)
        
        log_widget.setLayout(log_layout)
        splitter.addWidget(log_widget)
        
        # Set splitter sizes (60% tabs, 40% logs)
        splitter.setSizes([480, 320])
        
        main_layout.addWidget(splitter)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Status bar
        self.status_label = QLabel("Ready")
        main_layout.addWidget(self.status_label)
        
        central_widget.setLayout(main_layout)
        
        # Initial log message
        self.log_viewer.append_log('INFO', 'Application started. Ready to process maps.')
    
    def handle_log(self, level, message):
        """Handle log messages from backend"""
        self.log_viewer.append_log(level, message)
    
    def run_backend_task(self, func, *args, **kwargs):
        """Run a backend task in a worker thread"""
        if self.worker and self.worker.isRunning():
            QMessageBox.warning(self, "Busy", "A task is already running. Please wait.")
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate
        self.status_label.setText("Processing...")
        
        self.worker = WorkerThread(func, *args, **kwargs)
        self.worker.log_signal.connect(self.handle_log)
        self.worker.finished_signal.connect(self.task_finished)
        self.worker.start()
    
    def task_finished(self, success):
        """Handle task completion"""
        self.progress_bar.setVisible(False)
        if success:
            self.status_label.setText("Task completed successfully")
            QMessageBox.information(self, "Success", "Task completed successfully!")
        else:
            self.status_label.setText("Task failed")
            QMessageBox.warning(self, "Error", "Task failed. Check logs for details.")
    
    def export_logs(self):
        """Export logs to file"""
        file, _ = QFileDialog.getSaveFileName(
            self, "Export Logs", "logs.txt", "Text Files (*.txt)"
        )
        if file:
            with open(file, 'w') as f:
                f.write(self.log_viewer.toPlainText())
            QMessageBox.information(self, "Success", f"Logs exported to {file}")


def main():
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()

# Made with Bob
