#!/usr/bin/env python3
"""
Codelist Renamer for MXL Files
Renames codelists in MXL files based on mappings from codelist_report.xlsx
Supports both individual MXL files and ZIP archives containing MXL files

Usage:
    python3 codelist_renamer.py
    python3 codelist_renamer.py -i codelist_report.xlsx
    python3 codelist_renamer.py -i codelist_report.xlsx -o renamed_files/
    python3 codelist_renamer.py archive.zip
"""

import re
import openpyxl
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import argparse
import shutil
from datetime import datetime
import zipfile
import tempfile


class CodelistRenamer:
    def __init__(self, excel_file: str = 'codelist_report.xlsx'):
        self.excel_file = excel_file
        self.rename_mappings = {}  # {map_name: [(old_name, new_name), ...]}
        self.processed_files = []
        self.rename_count = 0
        
    def load_rename_mappings(self) -> bool:
        """Load rename mappings from Excel file"""
        try:
            wb = openpyxl.load_workbook(self.excel_file)
            ws = wb['Codelists']
            
            print(f"Loading rename mappings from {self.excel_file}...")
            
            # Skip header row, start from row 2
            for row in ws.iter_rows(min_row=2, values_only=True):
                if len(row) < 4:
                    continue
                    
                si_no, map_name, codelist_name, new_codelist_name = row[0], row[1], row[2], row[3]
                
                # Only process if new codelist name is provided and different from old name
                if new_codelist_name and isinstance(new_codelist_name, str) and new_codelist_name.strip() and new_codelist_name != codelist_name:
                    if map_name not in self.rename_mappings:
                        self.rename_mappings[map_name] = []
                    
                    new_name_clean = new_codelist_name.strip()
                    self.rename_mappings[map_name].append((codelist_name, new_name_clean))
                    print(f"  {map_name}: '{codelist_name}' -> '{new_name_clean}'")
            
            wb.close()
            
            if not self.rename_mappings:
                print("\n⚠ No rename mappings found in Excel file.")
                print("Please fill in the 'New Codelist Name' column in the Excel file.")
                return False
            
            print(f"\n✓ Loaded {sum(len(v) for v in self.rename_mappings.values())} rename mapping(s) for {len(self.rename_mappings)} file(s)")
            return True
            
        except FileNotFoundError:
            print(f"Error: Excel file '{self.excel_file}' not found")
            return False
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def rename_codelists_in_file(self, file_path: str, output_path: Optional[str] = None) -> Tuple[int, List[str]]:
        """
        Rename codelists in a single MXL file
        Returns: (number of replacements, list of renamed codelists)
        """
        file_name = Path(file_path).name
        
        # Check if this file has any rename mappings
        if file_name not in self.rename_mappings:
            print(f"  No rename mappings for {file_name}, skipping...")
            return 0, []
        
        try:
            # Read file content
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original_content = content
            replacements = 0
            renamed_codelists = []
            
            # Apply each rename mapping for this file
            for old_name, new_name in self.rename_mappings[file_name]:
                # Pattern 1: Extended Rule - from codelist where NAME = "CODELIST_NAME"
                pattern1 = rf'(from\s+codelist\s+where\s+NAME\s*=\s*")({re.escape(old_name)})(")'
                matches1 = len(re.findall(pattern1, content, re.IGNORECASE))
                content = re.sub(pattern1, rf'\1{new_name}\3', content, flags=re.IGNORECASE)
                
                # Pattern 2: In XML Value tags - <Value>CODELIST_NAME</Value>
                pattern2 = rf'(<Value>)({re.escape(old_name)})(</Value>)'
                matches2 = len(re.findall(pattern2, content))
                content = re.sub(pattern2, rf'\1{new_name}\3', content)
                
                # Pattern 3: In XML attributes - value="CODELIST_NAME"
                pattern3 = rf'(value\s*=\s*")({re.escape(old_name)})(")'
                matches3 = len(re.findall(pattern3, content, re.IGNORECASE))
                content = re.sub(pattern3, rf'\1{new_name}\3', content, flags=re.IGNORECASE)
                
                total_matches = matches1 + matches2 + matches3
                if total_matches > 0:
                    replacements += total_matches
                    renamed_codelists.append(f"{old_name} -> {new_name} ({total_matches} occurrence(s))")
                    print(f"    ✓ Renamed '{old_name}' to '{new_name}' ({total_matches} occurrence(s))")
            
            # Only write if changes were made
            if content != original_content:
                output_file = output_path if output_path else file_path
                
                # Create backup if overwriting original
                if output_file == file_path:
                    backup_path = f"{file_path}.backup"
                    shutil.copy2(file_path, backup_path)
                    print(f"    Created backup: {backup_path}")
                
                # Write modified content
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                print(f"    ✓ Saved changes to {output_file}")
            else:
                print(f"    No changes made to {file_name}")
            
            return replacements, renamed_codelists
            
        except Exception as e:
            print(f"  Error processing {file_path}: {e}")
            return 0, []
    
    def process_all_files(self, input_dir: str = 'mxl_files', output_dir: Optional[str] = None) -> None:
        """Process all MXL files that have rename mappings from specified input directory"""
        print("\n" + "="*70)
        print("PROCESSING MXL FILES")
        print("="*70)
        print(f"Input directory: {input_dir}")
        
        # Create output directory if specified
        if output_dir:
            Path(output_dir).mkdir(parents=True, exist_ok=True)
            print(f"Output directory: {output_dir}\n")
        else:
            print(f"Files will be renamed in-place (backups created)\n")
        
        # Process each file that has rename mappings
        for map_name in self.rename_mappings.keys():
            # Look for file in input directory
            file_path = Path(input_dir) / map_name
            
            if not file_path.exists():
                print(f"\n⚠ File not found: {file_path}")
                continue
            
            print(f"\nProcessing: {map_name}")
            
            # Determine output path
            output_path = None
            if output_dir:
                output_path = Path(output_dir) / map_name
            else:
                # Rename in-place (in the input directory)
                output_path = None
            
            # Rename codelists in file
            count, renamed = self.rename_codelists_in_file(str(file_path), str(output_path) if output_path else None)
            
            if count > 0:
                self.rename_count += count
                self.processed_files.append({
                    'file': map_name,
                    'count': count,
                    'renamed': renamed
                })
    def process_zip_file(self, zip_path: str, output_path: Optional[str] = None) -> None:
        """
        Extract MXL files from ZIP, rename codelists, and create new ZIP
        
        Args:
            zip_path: Path to input ZIP file
            output_path: Optional path for output ZIP file (default: adds '_renamed' suffix)
        """
        print(f"\n{'='*70}")
        print(f"PROCESSING ZIP FILE: {zip_path}")
        print('='*70)
        
        if not zipfile.is_zipfile(zip_path):
            print(f"Error: {zip_path} is not a valid ZIP file")
            return
        
        # Determine output ZIP path
        if output_path is None:
            zip_name = Path(zip_path).stem
            zip_ext = Path(zip_path).suffix
            output_path = f"{zip_name}_renamed{zip_ext}"
        
        # Create temporary directories
        temp_extract_dir = tempfile.mkdtemp(prefix='rename_extract_')
        temp_output_dir = tempfile.mkdtemp(prefix='rename_output_')
        
        try:
            # Extract ZIP file
            print(f"\nExtracting ZIP archive...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_dir)
            
            # Find all MXL files
            temp_path = Path(temp_extract_dir)
            mxl_files = list(temp_path.rglob('*.mxl'))
            
            if not mxl_files:
                print(f"⚠ No MXL files found in ZIP archive")
                return
            
            print(f"Found {len(mxl_files)} MXL file(s) in ZIP archive\n")
            
            # Process each MXL file
            files_processed = 0
            for mxl_file in mxl_files:
                # Get relative path within ZIP
                rel_path = mxl_file.relative_to(temp_path)
                file_name = mxl_file.name
                
                # Check if this file has rename mappings
                if file_name not in self.rename_mappings:
                    print(f"  Skipping {rel_path} (no rename mappings)")
                    # Copy unchanged file to output
                    output_file_path = Path(temp_output_dir) / rel_path
                    output_file_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(mxl_file, output_file_path)
                    continue
                
                print(f"\nProcessing: {rel_path}")
                
                # Create output path maintaining directory structure
                output_file_path = Path(temp_output_dir) / rel_path
                output_file_path.parent.mkdir(parents=True, exist_ok=True)
                
                # Rename codelists
                count, renamed = self.rename_codelists_in_file(
                    str(mxl_file), 
                    str(output_file_path)
                )
                
                if count > 0:
                    self.rename_count += count
                    self.processed_files.append({
                        'file': str(rel_path),
                        'count': count,
                        'renamed': renamed
                    })
                    files_processed += 1
                else:
                    # If no changes, still copy the file
                    if not output_file_path.exists():
                        shutil.copy2(mxl_file, output_file_path)
            
            # Copy all non-MXL files to output directory
            print(f"\nCopying non-MXL files...")
            for item in temp_path.rglob('*'):
                if item.is_file() and item.suffix.lower() != '.mxl':
                    rel_path = item.relative_to(temp_path)
                    output_file_path = Path(temp_output_dir) / rel_path
                    output_file_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(item, output_file_path)
            
            # Create new ZIP file with renamed content
            print(f"\nCreating output ZIP: {output_path}")
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for item in Path(temp_output_dir).rglob('*'):
                    if item.is_file():
                        arcname = item.relative_to(temp_output_dir)
                        zip_out.write(item, arcname)
            
            print(f"✓ Created renamed ZIP file: {output_path}")
            print(f"✓ Processed {files_processed} MXL file(s) with renames")
            
        finally:
            # Clean up temporary directories
            shutil.rmtree(temp_extract_dir, ignore_errors=True)
            shutil.rmtree(temp_output_dir, ignore_errors=True)
            print(f"\nCleaned up temporary files")
    
    
    def generate_summary_report(self, output_file: str = 'rename_summary.txt') -> None:
        """Generate a summary report of all renames"""
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("="*70 + "\n")
                f.write("CODELIST RENAME SUMMARY REPORT\n")
                f.write("="*70 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                f.write(f"Total Files Processed: {len(self.processed_files)}\n")
                f.write(f"Total Renames: {self.rename_count}\n\n")
                
                f.write("-"*70 + "\n")
                f.write("DETAILED CHANGES\n")
                f.write("-"*70 + "\n\n")
                
                for file_info in self.processed_files:
                    f.write(f"File: {file_info['file']}\n")
                    f.write(f"Total Renames: {file_info['count']}\n")
                    f.write("Changes:\n")
                    for rename in file_info['renamed']:
                        f.write(f"  • {rename}\n")
                    f.write("\n")
                
                f.write("="*70 + "\n")
            
            print(f"\n✓ Summary report saved: {output_file}")
            
        except Exception as e:
            print(f"Error generating summary report: {e}")
    
    def print_summary(self):
        """Print summary to console"""
        print("\n" + "="*70)
        print("RENAME SUMMARY")
        print("="*70)
        print(f"Total Files Processed: {len(self.processed_files)}")
        print(f"Total Renames: {self.rename_count}")
        
        if self.processed_files:
            print("\nFiles Modified:")
            for file_info in self.processed_files:
                print(f"  • {file_info['file']}: {file_info['count']} rename(s)")
        
        print("="*70)


def main():
    """Main execution function"""
    parser = argparse.ArgumentParser(
        description='Rename codelists in MXL files based on Excel mappings. Supports ZIP archives.',
        epilog='''
Examples:
  python3 codelist_renamer.py
  python3 codelist_renamer.py -i codelist_report.xlsx
  python3 codelist_renamer.py -i codelist_report.xlsx -o renamed_files/
  python3 codelist_renamer.py archive.zip
  python3 codelist_renamer.py archive.zip -o renamed_archive.zip
  python3 codelist_renamer.py --no-backup
        ''',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        'files',
        nargs='*',
        help='ZIP files to process. If not specified, processes MXL files in current directory'
    )
    
    parser.add_argument(
        '-i', '--input',
        default='codelist_report.xlsx',
        help='Input Excel file with rename mappings (default: codelist_report.xlsx)'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Output directory for renamed files or output ZIP file name'
    )
    
    parser.add_argument(
        '--no-backup',
        action='store_true',
        help='Do not create backup files when overwriting originals'
    )
    
    args = parser.parse_args()
    
    print("="*70)
    print("CODELIST RENAMER FOR MXL FILES")
    print("="*70)
    print(f"Excel file: {args.input}")
    
    # Create renamer instance
    renamer = CodelistRenamer(args.input)
    
    # Load rename mappings from Excel
    if not renamer.load_rename_mappings():
        return
    
    # Check if ZIP files were specified
    if args.files:
        # Process ZIP files
        for file_path in args.files:
            if not Path(file_path).exists():
                print(f"\n⚠ File not found: {file_path}")
                continue
            
            if file_path.lower().endswith('.zip'):
                # Determine output path for ZIP
                output_zip = args.output
                if output_zip is None and len(args.files) == 1:
                    # Single ZIP file, use default naming
                    output_zip = None  # Will add _renamed suffix
                elif output_zip is None:
                    # Multiple ZIP files, use default naming for each
                    zip_name = Path(file_path).stem
                    output_zip = f"{zip_name}_renamed.zip"
                
                renamer.process_zip_file(file_path, output_zip)
            else:
                print(f"\n⚠ Skipping non-ZIP file: {file_path}")
    else:
        # Process MXL files in mxl_files directory
        renamer.process_all_files(input_dir='mxl_files', output_dir=args.output)
    
    # Print summary
    renamer.print_summary()
    
    # Generate summary report
    renamer.generate_summary_report()
    
    print("\n✓ Processing complete!")


if __name__ == '__main__':
    main()

# Made with Bob
