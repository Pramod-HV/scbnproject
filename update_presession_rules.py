import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime
import os
import re
from pathlib import Path
import zipfile
import tempfile
import shutil
import sys


class PreSessionRuleUpdater:
    """
    Updates presessionrule tags in XML mapping files based on Excel configuration.
    """
    
    def __init__(self, excel_file_path, mapping_results_path=None, map_name=None):
        """
        Initialize the updater with Excel file path.
        
        Args:
            excel_file_path (str): Path to the Excel file containing configuration
            mapping_results_path (str, optional): Path to mapping_results.xlsx file
            map_name (str, optional): Name of the map to look up Base Name
        """
        self.excel_file_path = excel_file_path
        self.mapping_results_path = mapping_results_path
        self.map_name = map_name
        self.decision_flag = None
        self.reviewer_name = None
        self.base_name = None
        
    def read_excel_config(self):
        """
        Read configuration from Generic_rule sheet in Excel file.
        Also reads Base Name from mapping_results.xlsx if provided.
        
        Returns:
            bool: True if processing should continue (flag is 'Yes'), False otherwise
        """
        try:
            # Read the Generic_rule sheet
            df = pd.read_excel(self.excel_file_path, sheet_name='Generic_rule', header=None)
            
            # Read first row only
            # Column B (index 1) → decision flag (Yes/No)
            # Column C (index 2) → reviewer name (only used if Column B is 'Yes')
            self.decision_flag = str(df.iloc[0, 1]).strip()
            
            print(f"Decision Flag: {self.decision_flag}")
            
            # Check if flag is 'Yes'
            if self.decision_flag.lower() != 'yes':
                print("Decision flag is not 'Yes'. Exiting without processing.")
                self.reviewer_name = None
                return False
            
            # Only read Column C if decision flag is 'Yes'
            col_c_value = str(df.iloc[0, 2]).strip()
            
            # Column C should contain the actual reviewer name
            # Check if it's a valid name (not a prompt or NaN)
            if col_c_value and col_c_value.lower() not in ['nan', 'none', ''] and not col_c_value.lower().startswith('enter'):
                self.reviewer_name = col_c_value
            else:
                # If Column C doesn't have a valid name, use "Unknown"
                self.reviewer_name = "Unknown"
                print(f"Warning: Column C should contain the reviewer name when Column B is 'Yes'")
                print(f"Current value in Column C: '{col_c_value}'")
            
            print(f"Reviewer Name: {self.reviewer_name}")
            
            # Read Base Name from mapping_results.xlsx if provided
            if self.mapping_results_path and self.map_name:
                try:
                    mapping_df = pd.read_excel(self.mapping_results_path)
                    # Find the row with matching Map Name
                    map_row = mapping_df[mapping_df['Map Name'] == self.map_name]
                    if not map_row.empty:
                        # Column E is index 4 (0-based: A=0, B=1, C=2, D=3, E=4)
                        base_name_value = map_row.iloc[0, 4]
                        if pd.notna(base_name_value):
                            self.base_name = str(base_name_value).strip()
                            print(f"Base Name: {self.base_name}")
                        else:
                            print(f"Warning: Base Name (Column E) is empty for map '{self.map_name}'")
                    else:
                        print(f"Warning: Map '{self.map_name}' not found in mapping_results.xlsx")
                except Exception as e:
                    print(f"Warning: Could not read Base Name from mapping_results.xlsx: {e}")
            
            return True
            
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return False
    
    def process_presessionrule_content(self, content):
        """
        Process the content inside presessionrule tag:
        - Remove ONLY old standardized header comment lines (//Created, //Reviewed By, //Base Name)
        - Preserve ALL other content including user // comments
        - Add new standardized comment lines at the top
        
        Args:
            content (str): The content inside presessionrule tag
            
        Returns:
            str: Processed content
        """
        # Split content into lines
        lines = content.split('\n')
        
        # Remove ONLY old standardized header comments, preserve EVERYTHING else
        filtered_lines = []
        for line in lines:
            stripped = line.strip()
            # Remove only if it's an OLD standardized header comment
            if (stripped.startswith('//Created:') or
                stripped.startswith('//Reviewed By:') or
                stripped.startswith('//Base Name:')):
                continue  # Skip this line (remove old header)
            else:
                filtered_lines.append(line)  # Keep EVERYTHING else (including ALL // comments)
        
        # Get current date in format: DD-MM-YYYY
        current_date = datetime.now().strftime('%d-%m-%Y')
        
        # Create standardized comment lines
        created_comment = f"//Created:  {current_date}   By:  Migration Tool"
        reviewed_comment = f"//Reviewed By: {self.reviewer_name}"
        base_name_comment = f"//Base Name: {self.base_name}" if self.base_name else None
        
        # Find the first non-empty line to determine indentation
        indent = ""
        for line in filtered_lines:
            if line.strip():
                # Calculate indentation from first non-empty line
                indent = line[:len(line) - len(line.lstrip())]
                break
        
        # Build new content with comments at the top
        new_lines = []
        
        # Add comment lines with proper indentation
        new_lines.append(f"{indent}{created_comment}")
        new_lines.append(f"{indent}{reviewed_comment}")
        if base_name_comment:
            new_lines.append(f"{indent}{base_name_comment}")
        
        # Add the filtered content
        new_lines.extend(filtered_lines)
        
        return '\n'.join(new_lines)
    
    def update_xml_file(self, xml_file_path):
        """
        Update a single XML mapping file by processing presessionrule tag.
        Preserves all XML structure, namespaces, and formatting.
        
        Args:
            xml_file_path (str): Path to the XML file to update
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            print(f"\nProcessing file: {xml_file_path}")
            
            # Read the entire XML file as text to preserve formatting
            with open(xml_file_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
            
            # Store original content for comparison
            original_content = xml_content
            
            # Find presessionrule tag using regex
            # This pattern captures the opening tag (with optional namespace), content, and closing tag
            # Handles both <PreSessionRule> and <ns:PreSessionRule> formats
            pattern = r'(<(?:\w+:)?PreSessionRule>)(.*?)(</(?:\w+:)?PreSessionRule>)'
            
            # Search for the pattern (case-insensitive to handle PRESESSIONRULE or PreSessionRule)
            match = re.search(pattern, xml_content, re.DOTALL | re.IGNORECASE)
            
            if not match:
                print(f"Warning: No <PRESESSIONRULE> tag found in {xml_file_path}")
                return False
            
            # Extract the parts
            opening_tag = match.group(1)
            presession_content = match.group(2)
            closing_tag = match.group(3)
            
            # Process the content inside presessionrule tag
            updated_content = self.process_presessionrule_content(presession_content)
            
            # Replace the old presessionrule section with updated one
            updated_xml = xml_content[:match.start()] + opening_tag + updated_content + closing_tag + xml_content[match.end():]
            
            # Only write if content changed
            if updated_xml != original_content:
                # Write the updated content back to file
                with open(xml_file_path, 'w', encoding='utf-8') as f:
                    f.write(updated_xml)
                print(f"Successfully updated: {xml_file_path}")
                return True
            else:
                print(f"No changes needed for: {xml_file_path}")
                return True
                
        except Exception as e:
            print(f"Error processing {xml_file_path}: {e}")
            return False
    
    def process_single_file(self, xml_file_path):
        """
        Process a single XML file.
        
        Args:
            xml_file_path (str): Path to the XML file
            
        Returns:
            dict: Summary of processing results
        """
        results = {
            'total': 1,
            'success': 0,
            'failed': 0,
            'skipped': 0
        }
        
        print(f"\nProcessing single file: {xml_file_path}")
        
        if self.update_xml_file(xml_file_path):
            results['success'] += 1
        else:
            results['failed'] += 1
        
        return results
    
    def process_directory(self, directory_path):
        """
        Process all XML files in a directory.
        
        Args:
            directory_path (str): Path to directory containing XML files
            
        Returns:
            dict: Summary of processing results
        """
        results = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'skipped': 0
        }
        
        # Get all XML files in the directory
        xml_files = list(Path(directory_path).glob('*.xml'))
        
        if not xml_files:
            print(f"No XML files found in {directory_path}")
            return results
        
        print(f"\nFound {len(xml_files)} XML file(s) to process")
        
        for xml_file in xml_files:
            results['total'] += 1
            
            if self.update_xml_file(str(xml_file)):
                results['success'] += 1
            else:
                results['failed'] += 1
        
        return results
    
    def process_zip_file(self, zip_file_path):
        """
        Process XML files inside a ZIP archive.
        Extracts files to a temporary directory, processes them, and updates the ZIP.
        
        Args:
            zip_file_path (str): Path to the ZIP file
            
        Returns:
            dict: Summary of processing results
        """
        results = {
            'total': 0,
            'success': 0,
            'failed': 0,
            'skipped': 0
        }
        
        # Create a temporary directory
        temp_dir = tempfile.mkdtemp(prefix='presession_')
        
        try:
            print(f"\nExtracting ZIP file: {zip_file_path}")
            
            # Extract ZIP contents
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            # Find all XML and MXL files in the extracted directory (recursively)
            xml_files = list(Path(temp_dir).rglob('*.xml'))
            mxl_files = list(Path(temp_dir).rglob('*.mxl'))
            all_files = xml_files + mxl_files
            
            if not all_files:
                print(f"No XML/MXL files found in ZIP: {zip_file_path}")
                return results
            
            print(f"Found {len(all_files)} XML/MXL file(s) in ZIP")
            
            # Process each XML/MXL file
            for xml_file in all_files:
                results['total'] += 1
                
                if self.update_xml_file(str(xml_file)):
                    results['success'] += 1
                else:
                    results['failed'] += 1
            
            # Create a new ZIP file with updated contents
            if results['success'] > 0:
                print(f"\nUpdating ZIP file: {zip_file_path}")
                
                # Create backup of original ZIP
                backup_path = f"{zip_file_path}.backup"
                shutil.copy2(zip_file_path, backup_path)
                print(f"Backup created: {backup_path}")
                
                # Create new ZIP with updated files
                with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                    # Walk through temp directory and add all files
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            # Calculate relative path from temp_dir
                            arcname = os.path.relpath(file_path, temp_dir)
                            zip_ref.write(file_path, arcname)
                
                print(f"ZIP file updated successfully")
            
        except Exception as e:
            print(f"Error processing ZIP file: {e}")
            results['failed'] = results['total']
            results['success'] = 0
        
        finally:
            # Clean up temporary directory
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"Warning: Could not remove temporary directory: {e}")
        
        return results
    
    def process_mapping_files(self, path):
        """
        Process XML mapping files from various sources:
        - Single XML file
        - Directory containing XML files
        - ZIP file containing XML files
        
        Args:
            path (str): Path to XML file, directory, or ZIP file
            
        Returns:
            dict: Summary of processing results
        """
        path_obj = Path(path)
        
        # Check if path exists
        if not path_obj.exists():
            print(f"Error: Path does not exist: {path}")
            return {
                'total': 0,
                'success': 0,
                'failed': 0,
                'skipped': 0
            }
        
        # Determine the type of path and process accordingly
        if path_obj.is_file():
            if path_obj.suffix.lower() in ['.xml', '.mxl']:
                # Single XML file (including .mxl files)
                return self.process_single_file(str(path_obj))
            elif path_obj.suffix.lower() == '.zip':
                # ZIP file
                return self.process_zip_file(str(path_obj))
            else:
                print(f"Error: Unsupported file type: {path_obj.suffix}")
                return {
                    'total': 0,
                    'success': 0,
                    'failed': 0,
                    'skipped': 0
                }
        elif path_obj.is_dir():
            # Directory
            return self.process_directory(str(path_obj))
        else:
            print(f"Error: Invalid path type: {path}")
            return {
                'total': 0,
                'success': 0,
                'failed': 0,
                'skipped': 0
            }
    
    def run(self, path):
        """
        Main execution method.
        
        Args:
            path (str): Path to XML file, directory, or ZIP file
        """
        print("=" * 60)
        print("Pre-Session Rule Updater")
        print("=" * 60)
        
        # Step 1: Read Excel configuration
        if not self.read_excel_config():
            return
        
        # Step 2: Determine path type and process
        path_obj = Path(path)
        if path_obj.is_file():
            if path_obj.suffix.lower() == '.xml':
                print(f"\nProcessing single XML file: {path}")
            elif path_obj.suffix.lower() == '.zip':
                print(f"\nProcessing ZIP file: {path}")
        elif path_obj.is_dir():
            print(f"\nProcessing directory: {path}")
        
        results = self.process_mapping_files(path)
        
        # Step 3: Print summary
        print("\n" + "=" * 60)
        print("Processing Summary")
        print("=" * 60)
        print(f"Total files processed: {results['total']}")
        print(f"Successfully updated: {results['success']}")
        print(f"Failed: {results['failed']}")
        print(f"Skipped: {results['skipped']}")
        print("=" * 60)


def main():
    """
    Main entry point for the script.
    """
    # Check if command-line arguments are provided
    if len(sys.argv) < 3:
        print("Usage: python update_presession_rules.py <excel_file> <mapping_path> [mapping_results_file] [map_name]")
        print("\nArguments:")
        print("  excel_file           : Path to Excel file containing configuration")
        print("  mapping_path         : Path to XML file, directory, or ZIP file to process")
        print("  mapping_results_file : (Optional) Path to mapping_results.xlsx file")
        print("  map_name            : (Optional) Name of the map to look up Base Name")
        print("\nSupported mapping formats:")
        print("  - Single XML file (e.g., 'mapping.xml')")
        print("  - Directory containing XML files (e.g., 'mappings/')")
        print("  - ZIP file containing XML files (e.g., 'mappings.zip')")
        print("\nExample:")
        print("  python update_presession_rules.py Generic_checklist.xlsx 'IBM_Map_1 (1).zip'")
        print("  python update_presession_rules.py Generic_checklist.xlsx map.mxl mapping_results.xlsx MAP_NAME")
        return
    
    EXCEL_FILE = sys.argv[1]
    MAPPING_PATH = sys.argv[2]
    MAPPING_RESULTS_FILE = sys.argv[3] if len(sys.argv) > 3 else None
    MAP_NAME = sys.argv[4] if len(sys.argv) > 4 else None
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: Excel file not found: {EXCEL_FILE}")
        return
    
    # Check if mapping path exists
    if not os.path.exists(MAPPING_PATH):
        print(f"Error: Path not found: {MAPPING_PATH}")
        return
    
    # Check if mapping_results file exists (if provided)
    if MAPPING_RESULTS_FILE and not os.path.exists(MAPPING_RESULTS_FILE):
        print(f"Warning: mapping_results file not found: {MAPPING_RESULTS_FILE}")
        MAPPING_RESULTS_FILE = None
    
    # Create updater instance and run
    updater = PreSessionRuleUpdater(EXCEL_FILE, MAPPING_RESULTS_FILE, MAP_NAME)
    updater.run(MAPPING_PATH)


if __name__ == "__main__":
    main()

# Made with Bob
