#!/usr/bin/env python3
"""
Process All MXL Files
Processes all MXL files in the mxl_files directory using process_data_updater_v3.py
Generates inactive field error report after processing.
"""

import sys
import os
from pathlib import Path
import subprocess
from datetime import datetime
import pandas as pd
import re

def get_map_direction(mapping_file, mxl_file):
    """
    Get the direction (Inbound/Outbound) of a map from mapping_results.xlsx
    
    Args:
        mapping_file: Path to mapping_results.xlsx
        mxl_file: Path to MXL file
    
    Returns:
        str: 'Inbound', 'Outbound', or None if not found
    """
    try:
        # Extract map name from MXL file
        with open(mxl_file, 'r', encoding='UTF-8', errors='replace') as f:
            content = f.read()
        
        # Extract map name from Description tag
        desc_match = re.search(r'<Description>([^<]+)</Description>', content)
        if not desc_match:
            return None
        
        map_name = desc_match.group(1).strip()
        
        # Read mapping file
        df = pd.read_excel(mapping_file)
        
        # Find matching row
        matching_rows = df[df['Map Name'].str.strip() == map_name]
        
        if matching_rows.empty:
            return None
        
        # Get direction from first matching row
        direction = matching_rows.iloc[0]['Direction']
        return str(direction).strip() if pd.notna(direction) else None
        
    except Exception as e:
        print(f"Warning: Could not determine map direction: {e}")
        return None

def main():
    """Process all MXL files in the mxl_files directory."""
    
    # Check command line arguments
    if len(sys.argv) < 3:
        print("\nUsage: python process_all_mxl_files.py <mapping_results.xlsx> <Generic_checklist.xlsm> [--dry-run]")
        print("\nExample:")
        print("  python process_all_mxl_files.py mapping_results.xlsx Generic_checklistMain.xlsm")
        print("  python process_all_mxl_files.py mapping_results.xlsx Generic_checklistMain.xlsm --dry-run")
        sys.exit(1)
    
    mapping_file = sys.argv[1]
    checklist_file = sys.argv[2]
    dry_run = "--dry-run" in sys.argv
    
    # Validate files exist
    if not os.path.exists(mapping_file):
        print(f"Error: Mapping file not found: {mapping_file}")
        sys.exit(1)
    
    if not os.path.exists(checklist_file):
        print(f"Error: Checklist file not found: {checklist_file}")
        sys.exit(1)
    
    # Get all MXL files
    mxl_dir = Path("./mxl_files")
    if not mxl_dir.exists():
        print(f"Error: MXL directory not found: {mxl_dir}")
        sys.exit(1)
    
    mxl_files = list(mxl_dir.glob("*.mxl"))
    if not mxl_files:
        print(f"No MXL files found in {mxl_dir}")
        sys.exit(0)
    
    # Print header
    print("=" * 80)
    print("PROCESSING ALL MXL FILES")
    print("=" * 80)
    print(f"Mapping file: {mapping_file}")
    print(f"Checklist file: {checklist_file}")
    print(f"MXL directory: {mxl_dir}")
    print(f"Total files: {len(mxl_files)}")
    if dry_run:
        print("DRY RUN MODE - No changes will be written")
    print("=" * 80)
    print()
    
    # Process each file
    success_count = 0
    fail_count = 0
    
    for i, mxl_file in enumerate(mxl_files, 1):
        print(f"\n[{i}/{len(mxl_files)}] Processing: {mxl_file.name}")
        print("-" * 80)
        
        # Build command
        cmd = [
            sys.executable,
            "process_data_updater_v3.py",
            mapping_file,
            checklist_file,
            str(mxl_file)
        ]
        
        if dry_run:
            cmd.append("--dry-run")
        
        # Execute command
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False  # Don't raise exception, we'll check returncode
            )
            
            # Print output
            if result.stdout:
                print(result.stdout)
            
            # Check if process_data_updater_v3.py found matching 'Yes' rows
            # Exit code 2 means no matching rows were found
            if result.returncode == 2:
                print("\n" + "-" * 80)
                print("Skipping Inbound Maps Features processing (no matching 'Yes' rows found)")
                print("-" * 80)
                print(f"✓ Successfully processed {mxl_file.name}")
                success_count += 1
                continue
            elif result.returncode != 0:
                # Other non-zero exit codes indicate errors
                print(f"✗ Failed to process {mxl_file.name}")
                fail_count += 1
                continue
            
            # Check map direction and run appropriate features script
            map_direction = get_map_direction(mapping_file, mxl_file)
            
            if map_direction and map_direction.lower() == 'outbound':
                # Run outbound_mapsFeatures.py for Outbound maps
                print("\n" + "-" * 80)
                print(f"Skipping Inbound Maps Features processing (map direction: {map_direction})")
                print("Running Outbound Maps Features processing...")
                print("-" * 80)
                
                try:
                    outbound_cmd = [
                        sys.executable,
                        "outbound_mapsFeatures.py",
                        checklist_file,
                        str(mxl_file)
                    ]
                    
                    outbound_result = subprocess.run(
                        outbound_cmd,
                        capture_output=True,
                        text=True,
                        check=True
                    )
                
                    # Print output
                    if outbound_result.stdout:
                        print(outbound_result.stdout)
                    
                    print("✓ Outbound Maps Features processing completed")
                    
                except subprocess.CalledProcessError as e:
                    print(f"⚠ Outbound Maps Features processing failed: {e}")
                    if e.stdout:
                        print("Output:", e.stdout)
                    if e.stderr:
                        print("Error details:", e.stderr)
                    # Don't fail the entire process if outbound features fail
                except FileNotFoundError:
                    print("⚠ outbound_mapsFeatures.py not found, skipping")
                except Exception as e:
                    print(f"⚠ Unexpected error in Outbound Maps Features: {e}")
            else:
                # Run inbound_mapsFeatures.py for Inbound maps
                print("\n" + "-" * 80)
                print("Running Inbound Maps Features processing...")
                print("-" * 80)
                
                try:
                    inbound_cmd = [
                        sys.executable,
                        "inbound_mapsFeatures.py",
                        checklist_file,
                        str(mxl_file)
                    ]
                    
                    inbound_result = subprocess.run(
                        inbound_cmd,
                        capture_output=True,
                        text=True,
                        check=True
                    )
                
                    # Print output
                    if inbound_result.stdout:
                        print(inbound_result.stdout)
                    
                    print("✓ Inbound Maps Features processing completed")
                    
                except subprocess.CalledProcessError as e:
                    print(f"⚠ Inbound Maps Features processing failed: {e}")
                    if e.stdout:
                        print("Output:", e.stdout)
                    if e.stderr:
                        print("Error details:", e.stderr)
                    # Don't fail the entire process if inbound features fail
                except FileNotFoundError:
                    print("⚠ inbound_mapsFeatures.py not found, skipping")
                except Exception as e:
                    print(f"⚠ Unexpected error in Inbound Maps Features: {e}")
            
            success_count += 1
            print(f"\n✓ Successfully processed {mxl_file.name}")
            
        except subprocess.CalledProcessError as e:
            fail_count += 1
            print(f"✗ Failed to process {mxl_file.name}")
            print(f"Error: {e}")
            if e.stdout:
                print("Output:", e.stdout)
            if e.stderr:
                print("Error details:", e.stderr)
        except Exception as e:
            fail_count += 1
            print(f"✗ Unexpected error processing {mxl_file.name}: {e}")
    
    # Print summary
    print("\n" + "=" * 80)
    print("PROCESSING COMPLETE")
    print("=" * 80)
    print(f"Total files: {len(mxl_files)}")
    print(f"Successful: {success_count}")
    print(f"Failed: {fail_count}")
    print("=" * 80)
    
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

# Made with Bob
