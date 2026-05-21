#!/usr/bin/env python3
"""
Error Reporter Module
Tracks and reports errors for inactive fields and unmapped output fields.
Creates/updates an Excel error report with deduplication.
"""

import pandas as pd
import os
from typing import List, Dict, Set, Tuple


class ErrorReporter:
    """Handles error tracking and Excel report generation."""
    
    def __init__(self, report_file: str = "error_report.xlsx"):
        """
        Initialize the error reporter.
        
        Args:
            report_file: Path to the Excel error report file
        """
        self.report_file = report_file
        self.errors: List[Dict[str, str]] = []
        self.error_set: Set[Tuple[str, str, str]] = set()  # For deduplication: (map_name, field, error_type)
        
    def add_inactive_field_error(self, map_name: str, field_position: str, field_name: str = "", xpath: str = ""):
        """
        Add an error for an inactive field (no active segments found).
        
        Args:
            map_name: Name of the map
            field_position: Field position (e.g., "810.BIG-04" or "SPI_ASN.ASN-CUSTOMER_ID")
            field_name: Field name (e.g., "0396") - optional
            xpath: XPath column (e.g., "TranslationOutput/BusinessReference") - optional
        """
        error_key = (map_name, field_position, "inactive")
        
        if error_key not in self.error_set:
            # Build error message
            if xpath and field_name:
                # Include both field name and xpath
                error_msg = f"The field {field_position} is inactive for the field #{field_name} in {xpath}. Please check."
            elif xpath:
                # Include xpath only
                error_msg = f"The field {field_position} is inactive for the field in {xpath}. Please check."
            else:
                error_msg = f"The field {field_position} is inactive. Please check."
            
            self.errors.append({
                "Map Name": map_name,
                "Field": field_position,
                "Error Type": "Inactive Field",
                "Error Message": error_msg
            })
            self.error_set.add(error_key)
    
    def add_no_mapping_error(self, map_name: str, field_position: str, field_name: str = "", xpath: str = ""):
        """
        Add an error for an output field with no mapping.
        
        Args:
            map_name: Name of the map
            field_position: Field position (e.g., "810.BIG-04")
            field_name: Field name (e.g., "0396") - optional
            xpath: XPath column (e.g., "TranslationOutput/BusinessReference") - optional
        """
        error_key = (map_name, field_position, "no_mapping")
        
        if error_key not in self.error_set:
            # Build error message
            if xpath and field_name:
                # Include both field name and xpath
                error_msg = f"The field {field_position} has no mapping for the field #{field_name} in {xpath}. Please check."
            elif xpath:
                # Include xpath only
                error_msg = f"The field {field_position} has no mapping for the field in {xpath}. Please check."
            else:
                error_msg = f"The field {field_position} has no mapping. Please check."
            
            self.errors.append({
                "Map Name": map_name,
                "Field": field_position,
                "Error Type": "No Mapping",
                "Error Message": error_msg
            })
            self.error_set.add(error_key)
    
    def has_errors(self) -> bool:
        """Check if any errors have been recorded."""
        return len(self.errors) > 0
    
    def get_error_count(self) -> int:
        """Get the total number of errors."""
        return len(self.errors)
    
    def save_report(self) -> bool:
        """
        Save or append errors to the Excel report.
        Handles deduplication by checking existing errors.
        
        Returns:
            True if report was saved successfully, False otherwise
        """
        if not self.errors:
            return True  # No errors to save
        
        try:
            # Create DataFrame from current errors
            new_df = pd.DataFrame(self.errors)
            
            # Check if report file exists
            if os.path.exists(self.report_file):
                # Read existing report
                existing_df = pd.read_excel(self.report_file)
                
                # Create a set of existing errors for deduplication
                existing_errors = set()
                for _, row in existing_df.iterrows():
                    error_key = (
                        str(row.get("Map Name", "")),
                        str(row.get("Field", "")),
                        str(row.get("Error Type", ""))
                    )
                    existing_errors.add(error_key)
                
                # Filter out duplicates from new errors
                unique_new_errors = []
                for error in self.errors:
                    error_key = (
                        error["Map Name"],
                        error["Field"],
                        error["Error Type"]
                    )
                    if error_key not in existing_errors:
                        unique_new_errors.append(error)
                
                if unique_new_errors:
                    # Append unique new errors
                    new_df = pd.DataFrame(unique_new_errors)
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                    
                    # Save combined report
                    combined_df.to_excel(self.report_file, index=False)
                    print(f"✓ Appended {len(unique_new_errors)} new error(s) to {self.report_file}")
                else:
                    print(f"✓ No new errors to add (all errors already exist in {self.report_file})")
            else:
                # Create new report
                new_df.to_excel(self.report_file, index=False)
                print(f"✓ Created new error report: {self.report_file} with {len(self.errors)} error(s)")
            
            return True
            
        except Exception as e:
            print(f"✗ Error saving report: {e}")
            return False
    
    def clear_errors(self):
        """Clear all recorded errors."""
        self.errors.clear()
        self.error_set.clear()
    
    def print_summary(self):
        """Print a summary of errors."""
        if not self.errors:
            print("No errors recorded.")
            return
        
        print(f"\n{'='*80}")
        print(f"ERROR SUMMARY: {len(self.errors)} error(s) found")
        print(f"{'='*80}")
        
        # Group by error type
        inactive_count = sum(1 for e in self.errors if e["Error Type"] == "Inactive Field")
        no_mapping_count = sum(1 for e in self.errors if e["Error Type"] == "No Mapping")
        
        print(f"  - Inactive Fields: {inactive_count}")
        print(f"  - No Mapping: {no_mapping_count}")
        print(f"{'='*80}\n")

# Made with Bob
