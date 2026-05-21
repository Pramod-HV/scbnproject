"""
Backend API Wrapper for AI-Powered Map Migration Accelerator
Provides a clean interface between frontend and backend processing scripts
"""

import sys
import os
import subprocess
import threading
from pathlib import Path
from typing import Callable, Optional, Dict, Any, List
import logging
from datetime import datetime


class BackendAPI:
    """
    Backend API wrapper for all processing functions.
    Provides thread-safe execution with real-time log streaming.
    """
    
    def __init__(self, log_callback: Optional[Callable[[str, str], None]] = None):
        """
        Initialize Backend API
        
        Args:
            log_callback: Function to call with (level, message) for each log entry
                         level can be: 'INFO', 'WARNING', 'ERROR', 'SUCCESS'
        """
        self.log_callback = log_callback
        self.current_process = None
        self.is_running = False
        
    def _log(self, level: str, message: str):
        """Internal logging method"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] [{level}] {message}"
        
        if self.log_callback:
            self.log_callback(level, log_message)
        else:
            print(log_message)
    
    def _run_command(self, command: List[str], cwd: Optional[str] = None) -> tuple[bool, str]:
        """
        Run a command and stream output in real-time
        
        Returns:
            (success, output)
        """
        try:
            self.is_running = True
            self._log('INFO', f"Running command: {' '.join(command)}")
            
            # Run process with real-time output
            process = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=cwd or os.getcwd(),
                bufsize=1,
                universal_newlines=True
            )
            
            self.current_process = process
            output_lines = []
            
            # Read output line by line in real-time
            for line in process.stdout:
                line = line.strip()
                if line:
                    output_lines.append(line)
                    # Determine log level from line content
                    if 'ERROR' in line.upper() or 'FAILED' in line.upper() or 'EXCEPTION' in line.upper():
                        self._log('ERROR', line)
                    elif 'WARNING' in line.upper() or 'WARN' in line.upper():
                        self._log('WARNING', line)
                    elif 'SUCCESS' in line.upper() or '✓' in line or '✅' in line or 'COMPLETED' in line.upper():
                        self._log('SUCCESS', line)
                    elif 'DEBUG' in line.upper():
                        self._log('DEBUG', line)
                    else:
                        self._log('INFO', line)
            
            # Read stderr as well
            stderr_lines = []
            if process.stderr:
                for line in process.stderr:
                    line = line.strip()
                    if line:
                        stderr_lines.append(line)
                        self._log('ERROR', line)
            
            # Wait for process to complete
            return_code = process.wait()
            self.is_running = False
            self.current_process = None
            
            if return_code == 0:
                self._log('SUCCESS', '✅ Command completed successfully')
                return True, '\n'.join(output_lines)
            else:
                self._log('ERROR', f'❌ Command failed with exit code {return_code}')
                return False, '\n'.join(stderr_lines) if stderr_lines else 'Command failed'
                
        except Exception as e:
            self.is_running = False
            self.current_process = None
            self._log('ERROR', f'Exception: {str(e)}')
            return False, str(e)
    
    def stop_current_process(self):
        """Stop the currently running process"""
        if self.current_process:
            self._log('WARNING', 'Stopping current process...')
            self.current_process.terminate()
            self.current_process = None
            self.is_running = False
    
    # ==================== Phase 0: Extract Map Details ====================
    
    def extract_map_details(
        self,
        input_source: str,  # 'auto', 'zip', or 'mxl' (auto-detection recommended)
        input_folder: str,
        output_excel: str = 'mapping_results.xlsx'
    ) -> bool:
        """
        Phase 0: Extract map details from ZIP or MXL files
        
        Args:
            input_source: 'auto' (recommended), 'zip', or 'mxl'
            input_folder: Path to input folder/file
            output_excel: Output Excel filename
        
        Returns:
            True if successful
        """
        self._log('INFO', '=' * 70)
        self._log('INFO', 'PHASE 0: EXTRACT MAP DETAILS')
        self._log('INFO', '=' * 70)
        self._log('INFO', f'Input source: {input_source}')
        self._log('INFO', f'Input folder: {input_folder}')
        
        command = ['python', 'mxl_parser.py', input_folder]
        
        if output_excel != 'mapping_results.xlsx':
            command.append(output_excel)
        
        # Only add --mxl flag if explicitly specified (not for 'auto')
        if input_source == 'mxl':
            command.append('--mxl')
        # For 'auto' or 'zip', let mxl_parser.py auto-detect
        
        self._log('INFO', f'Running command: {" ".join(command)}')
        success, output = self._run_command(command)
        return success
    
    # ==================== Helper: Process Input and Setup Folders ====================
    
    def _process_input_and_setup_folders(self, input_path: str) -> bool:
        """
        Helper method to process ZIP/folder/MXL input and set up folder structure.
        Creates: zipfiles/, old_mxlFiles/, mxl_files/
        
        Args:
            input_path: Path to ZIP file, folder, or MXL file(s)
            
        Returns:
            True if successful, False otherwise
        """
        self._log('INFO', '📁 Processing input and setting up folder structure...')
        self._log('INFO', f'Input: {input_path}')
        
        # Call mxl_parser.py to handle extraction and folder setup
        command = ['python', 'mxl_parser.py', input_path]
        
        self._log('INFO', f'Running: {" ".join(command)}')
        success, output = self._run_command(command)
        
        if success:
            self._log('SUCCESS', '✅ Folder structure created successfully')
            self._log('INFO', '  📁 zipfiles/ - Nested ZIP files')
            self._log('INFO', '  📁 old_mxlFiles/ - Immutable backups')
            self._log('INFO', '  📁 mxl_files/ - Working directory')
        else:
            self._log('ERROR', '❌ Failed to process input and setup folders')
        
        return success
    
    # ==================== Phase 1: Extract Rules ====================
    
    def extract_rules(
        self,
        mxl_folder: str,
        output_excel: str = 'process_data_rules.xlsx'
    ) -> bool:
        """
        Phase 2: Extract process data rules from MXL files
        
        Args:
            mxl_folder: Path to MXL files folder, ZIP file, or individual MXL files
            output_excel: Output Excel filename
        
        Returns:
            True if successful
        """
        self._log('INFO', '=' * 70)
        self._log('INFO', 'PHASE 2: EXTRACT PROCESS DATA RULES')
        self._log('INFO', '=' * 70)
        
        # Step 1: Process input and setup folder structure
        if not self._process_input_and_setup_folders(mxl_folder):
            return False
        
        # Step 2: Extract rules from mxl_files/ folder
        working_folder = 'mxl_files'
        command = [
            'python', 'process_data_manager.py',
            'extract', working_folder, output_excel
        ]
        
        success, output = self._run_command(command)
        return success
    
    # ==================== Phase 2: Apply Rules ====================
    
    def apply_rules(
        self,
        rules_excel: str,
        mxl_folder: str
    ) -> bool:
        """
        Phase 2: Apply process data rules to MXL files
        
        Args:
            rules_excel: Path to rules Excel file
            mxl_folder: Path to MXL files folder
        
        Returns:
            True if successful
        """
        self._log('INFO', '=' * 70)
        self._log('INFO', 'PHASE 2: APPLY PROCESS DATA RULES')
        self._log('INFO', '=' * 70)
        
        command = [
            'python', 'process_data_manager.py',
            'apply', rules_excel, mxl_folder
        ]
        
        success, output = self._run_command(command)
        return success
    
    # ==================== Phase 3: Process Maps ====================
    
    def process_all_maps(
        self,
        mapping_results: str,
        checklist: str,
        mxl_folder: str = 'mxl_files',
        selected_files: Optional[List[str]] = None
    ) -> bool:
        """
        Phase 3/4: Process all maps with full automation
        
        Args:
            mapping_results: Path to mapping_results.xlsx
            checklist: Path to Generic_checklistMain.xlsm
            mxl_folder: Path to MXL files folder, ZIP file, or individual MXL files
            selected_files: List of specific files to process (None = all)
        
        Returns:
            True if successful
        """
        self._log('INFO', '=' * 70)
        self._log('INFO', 'PHASE 3/4: PROCESS ALL MAPS')
        self._log('INFO', '=' * 70)
        
        # Step 1: Process input and setup folder structure (only if not already using mxl_files)
        if mxl_folder != 'mxl_files':
            if not self._process_input_and_setup_folders(mxl_folder):
                return False
            # After processing, use mxl_files as the working folder
            mxl_folder = 'mxl_files'
        
        # Step 2: Process maps from mxl_files/ folder
        if selected_files:
            # Process selected files one by one
            success_count = 0
            for mxl_file in selected_files:
                file_path = os.path.join(mxl_folder, mxl_file)
                command = [
                    'python', 'process_data_updater_v3.py',
                    mapping_results, checklist, file_path
                ]
                success, _ = self._run_command(command)
                if success:
                    success_count += 1
            
            self._log('INFO', f'Processed {success_count}/{len(selected_files)} files successfully')
            return success_count == len(selected_files)
        else:
            # Process all files
            command = [
                'python', 'process_all_mxl_files.py',
                mapping_results, checklist
            ]
            success, output = self._run_command(command)
            return success
    
    # ==================== Individual Functions ====================
    
    def run_presession_rules(
        self,
        checklist: str,
        mxl_file: str,
        mapping_results: str,
        map_name: Optional[str] = None
    ) -> bool:
        """Run pre-session rules only"""
        self._log('INFO', 'Running Pre-session Rules...')
        
        command = [
            'python', 'update_presession_rules.py',
            checklist, mxl_file, mapping_results
        ]
        if map_name:
            command.append(map_name)
        
        success, output = self._run_command(command)
        return success
    
    def modify_character_encoding(
        self,
        mxl_file: str,
        map_name: str,
        checklist: str,
        mapping_results: str
    ) -> bool:
        """Modify character encoding only"""
        self._log('INFO', 'Modifying Character Encoding...')
        
        # This is called internally by process_data_updater_v3.py
        # For standalone use, we need to import and call the function
        try:
            from modify_character_encoding import modify_character_encoding
            result = modify_character_encoding(mxl_file, map_name, checklist, mapping_results)
            if result:
                self._log('SUCCESS', 'Character encoding modified successfully')
            else:
                self._log('ERROR', 'Failed to modify character encoding')
            return result
        except Exception as e:
            self._log('ERROR', f'Error: {str(e)}')
            return False
    
    def remove_namespace_prefixes(
        self,
        mxl_file: str
    ) -> bool:
        """Remove namespace prefixes only"""
        self._log('INFO', 'Removing Namespace Prefixes...')
        
        command = ['python', 'remove_namespace_prefixes.py', mxl_file]
        success, output = self._run_command(command)
        return success
    
    def process_inbound_features(
        self,
        checklist: str,
        mxl_file: str
    ) -> bool:
        """Process inbound features only"""
        self._log('INFO', 'Processing Inbound Features...')
        
        try:
            from inbound_mapsFeatures import process_mxl_file
            result = process_mxl_file(checklist, mxl_file)
            if result:
                self._log('SUCCESS', 'Inbound features processed successfully')
            else:
                self._log('WARNING', 'Inbound features processing completed with warnings')
            return result
        except Exception as e:
            self._log('ERROR', f'Error: {str(e)}')
            return False
    
    # ==================== Utility Functions ====================
    
    def validate_files(self) -> Dict[str, bool]:
        """Validate that all required files exist"""
        required_files = {
            'Generic_checklistMain.xlsm': False,
            'mxl_files': False,
        }
        
        for file_name in required_files:
            required_files[file_name] = os.path.exists(file_name)
            if required_files[file_name]:
                self._log('SUCCESS', f'Found: {file_name}')
            else:
                self._log('WARNING', f'Missing: {file_name}')
        
        return required_files
    
    def get_mxl_files(self, folder: str = 'mxl_files') -> List[str]:
        """Get list of MXL files in folder"""
        try:
            mxl_files = [f for f in os.listdir(folder) if f.endswith('.mxl')]
            self._log('INFO', f'Found {len(mxl_files)} MXL files in {folder}')
            return sorted(mxl_files)
        except Exception as e:
            self._log('ERROR', f'Error reading folder: {str(e)}')
            return []
    
    def backup_files(self, source_folder: str, backup_folder: str) -> bool:
        """Backup files to a backup folder"""
        try:
            import shutil
            from datetime import datetime
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_path = f"{backup_folder}_{timestamp}"
            
            shutil.copytree(source_folder, backup_path)
            self._log('SUCCESS', f'Backup created: {backup_path}')
            return True
        except Exception as e:
            self._log('ERROR', f'Backup failed: {str(e)}')
            return False
    
    def extract_codelists(self, input_folder: str, output_file: str = 'codelist_report.xlsx') -> bool:
        """
        Extract codelists from MXL files or ZIP archives (auto-detects)
        
        Args:
            input_folder: Path to folder containing MXL or ZIP files, or single file
            output_file: Name of output Excel file
            
        Returns:
            True if successful, False otherwise
        """
        try:
            self._log('INFO', f'Starting codelist extraction from {input_folder}')
            
            # Step 1: Process input and setup folder structure
            if not self._process_input_and_setup_folders(input_folder):
                return False
            
            # Step 2: Extract codelists from mxl_files/ folder
            mxl_folder = Path('mxl_files')
            if not mxl_folder.exists():
                self._log('ERROR', 'mxl_files/ folder not found after processing input')
                return False
            
            mxl_files = list(mxl_folder.glob('*.mxl'))
            if not mxl_files:
                self._log('ERROR', 'No MXL files found in mxl_files/ folder')
                return False
            
            self._log('INFO', f'Found {len(mxl_files)} MXL file(s) in mxl_files/')
            mxl_paths = [str(mf) for mf in mxl_files]
            command = ['python3', 'codelist_extractor.py'] + mxl_paths + ['-o', output_file]
            
            success, output = self._run_command(command)
            
            if success:
                self._log('SUCCESS', f'Codelist extraction complete! Report: {output_file}')
            else:
                self._log('ERROR', 'Codelist extraction failed')
            
            return success
            
        except Exception as e:
            self._log('ERROR', f'Error during codelist extraction: {str(e)}')
            return False
    
    def rename_codelists(self, input_file: str = 'codelist_report.xlsx') -> bool:
        """
        Rename codelists in MXL files based on Excel mappings
        
        Args:
            input_file: Path to Excel file with rename mappings
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Check if Excel file exists
            if not Path(input_file).exists():
                self._log('ERROR', f'Excel file not found: {input_file}')
                self._log('ERROR', 'Please run Extract Codelists first and fill in the "New Codelist Name" column')
                return False
            
            self._log('INFO', f'Starting codelist renaming from {input_file}')
            
            command = ['python3', 'codelist_renamer.py', '-i', input_file]
            success, output = self._run_command(command)
            
            if success:
                self._log('SUCCESS', 'Codelist renaming complete!')
                self._log('INFO', 'Check rename_summary.txt for details')
            else:
                self._log('ERROR', 'Codelist renaming failed')
            
            return success
            
        except Exception as e:
            self._log('ERROR', f'Error during codelist renaming: {str(e)}')
            return False


# Example usage
if __name__ == '__main__':
    def log_handler(level, message):
        print(f"{level}: {message}")
    
    api = BackendAPI(log_callback=log_handler)
    
    # Example: Extract map details
    # api.extract_map_details('mxl', 'old_mxlFiles')
    
    # Example: Get MXL files
    files = api.get_mxl_files()
    print(f"Found {len(files)} files")

# Made with Bob
