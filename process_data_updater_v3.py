#!/usr/bin/env python3
"""
Process Data Updater V3 - Enhanced with Mapping and Checklist Integration
Matches MXL files with mapping_results.xlsx and Generic_checklist.xlsm
Updates ExplicitRule tags based on direction, document type, and field positions.
"""

from __future__ import annotations
import pandas as pd
import sys
import os
import shutil
import logging
import re
import subprocess
import xml.etree.ElementTree as ET
from typing import List, Tuple, Optional, Dict, Union
from datetime import datetime
from pathlib import Path
from modify_character_encoding import modify_character_encoding
from error_reporter import ErrorReporter
from inbound_mapsFeatures import process_mxl_file as process_inbound_features


class ProcessDataUpdaterV3:
    def __init__(self, mapping_file: str, checklist_file: str, mxl_file: str, dry_run: bool = False):
        """
        Initialize the updater with mapping, checklist, and MXL file paths.
        
        Args:
            mapping_file: Path to mapping_results.xlsx
            checklist_file: Path to Generic_checklist.xlsm
            mxl_file: Path to the MXL file to update
            dry_run: If True, only simulate changes without writing
        """
        self.mapping_file = mapping_file
        self.checklist_file = checklist_file
        self.mxl_file = mxl_file
        self.dry_run = dry_run
        self.map_name = None
        self.direction = None
        self.document_type = None
        
        # Setup logging
        self._setup_logging()
        
        # Initialize error reporter
        self.error_reporter = ErrorReporter()
        
    def _setup_logging(self):
        """Configure logging with appropriate levels and format."""
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        
        if self.dry_run:
            self.logger.info("=" * 80)
            self.logger.info("DRY RUN MODE - No changes will be written to files")
            self.logger.info("=" * 80)
    
    def modify_x_syntax_token(self) -> bool:
        """
        Modify the X syntax token in the MXL file by adding extended character range.
        
        Returns:
            True if modification was successful, False otherwise
        """
        try:
            # Read the MXL file
            with open(self.mxl_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check if the extended range already exists
            if '<Start>ÿ</Start>' in content and '<End>0x01</End>' in content:
                self.logger.info(f"Extended character range already exists in {self.mxl_file}")
                return True
            
            # Find the X token section
            # Pattern: <Token><Code>X</Code>...<Char>`</Char></Token>
            x_token_pattern = r'(<Token>\s*<Code>X</Code>.*?<Char>`</Char>)(\s*</Token>)'
            
            match = re.search(x_token_pattern, content, re.DOTALL)
            
            if not match:
                self.logger.error(f"Could not find X syntax token in {self.mxl_file}")
                return False
            
            # Prepare the extended range to add
            # Note: Start=0xFF (ÿ), End=0x01 - this is the correct order for character ranges
            extended_range = """
<Range>
<Start>ÿ</Start>
<End>0x01</End>
</Range>"""
            
            # Insert the extended range before </Token>
            modified_content = content[:match.end(1)] + extended_range + match.group(2) + content[match.end():]
            
            # Write back to file
            if not self.dry_run:
                with open(self.mxl_file, 'w', encoding='utf-8') as f:
                    f.write(modified_content)
            
            self.logger.info(f"✓ Successfully added extended character range to X syntax token")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to modify X syntax token: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    
    def check_and_trigger_generic_rules(self) -> bool:
        """
        Check Generic_rule sheet and trigger external scripts if conditions are met.
        
        Returns:
            bool: True if processing should continue, False if should abort
        """
        try:
            self.logger.info("Checking Generic_rule sheet for additional modifications...")
            
            # Read Generic_rule sheet
            df = pd.read_excel(self.checklist_file, sheet_name='Generic_rule', header=None)
            
            # Row 1 (index 0): Pre-session rule
            # Column A (index 0): "Pre-session rule"
            # Column B (index 1): "Yes" or "No"
            presession_flag = str(df.iloc[0, 1]).strip().lower()
            
            # Row 2 (index 1): Modify the X syntax token
            # Column A (index 0): "Modify the X syntax token"
            # Column B (index 1): "Yes" or "No"
            syntax_token_flag = str(df.iloc[1, 1]).strip().lower()
            
            # Row 8 (index 7): Freeformat
            # Column A (index 0): "Freeformat"
            # Column B (index 1): "BothSides" or other values
            freeformat_value = str(df.iloc[7, 1]).strip().lower()
            
            self.logger.info(f"Pre-session rule flag: {presession_flag}")
            self.logger.info(f"Modify X syntax token flag: {syntax_token_flag}")
            self.logger.info(f"Freeformat value: {freeformat_value}")
            
            # Check if pre-session rule should be triggered
            if presession_flag == 'yes':
                self.logger.info("=" * 80)
                self.logger.info("TRIGGERING PRE-SESSION RULE SCRIPT")
                self.logger.info("=" * 80)
                
                try:
                    # Get absolute paths for external script
                    abs_checklist = os.path.abspath(self.checklist_file)
                    abs_mxl = os.path.abspath(self.mxl_file)
                    abs_mapping = os.path.abspath(self.mapping_file)
                    
                    # Build command with optional map name
                    cmd = [sys.executable, 'update_presession_rules.py', abs_checklist, abs_mxl, abs_mapping]
                    if self.map_name:
                        cmd.append(self.map_name)
                    
                    # Call update_presession_rules.py with checklist file, mxl file, mapping_results file, and map name
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    self.logger.info("Pre-session rule script output:")
                    self.logger.info(result.stdout)
                    if result.stderr:
                        self.logger.warning(f"Pre-session rule script warnings: {result.stderr}")
                except subprocess.CalledProcessError as e:
                    self.logger.error(f"Pre-session rule script failed: {e}")
                    self.logger.error(f"Output: {e.stdout}")
                    self.logger.error(f"Error: {e.stderr}")
                    return False
                except FileNotFoundError:
                    self.logger.error("update_presession_rules.py not found in current directory")
                    return False
            
            # Check if syntax token modification should be triggered
            if syntax_token_flag == 'yes':
                self.logger.info("=" * 80)
                self.logger.info("MODIFYING X SYNTAX TOKEN")
                self.logger.info("=" * 80)
                
                # Call internal method to modify syntax token
                if not self.modify_x_syntax_token():
                    self.logger.error("Failed to modify X syntax token")
                    return False
            
            # Check if character encoding should be modified (rows 2, 3, 4)
            # Row 2 (index 2): Set input encoding
            # Row 3 (index 3): Set output encoding
            # Row 4 (index 4): If output is XML, set encoding to UTF8
            input_encoding_setting = str(df.iloc[2, 1]).strip() if pd.notna(df.iloc[2, 1]) else "Keep as is"
            output_encoding_setting = str(df.iloc[3, 1]).strip() if pd.notna(df.iloc[3, 1]) else "Keep as is"
            xml_utf8_setting = str(df.iloc[4, 1]).strip() if pd.notna(df.iloc[4, 1]) else "Keep as is"
            
            # Check if any encoding modification is needed
            if (input_encoding_setting != "Keep as is" or
                output_encoding_setting != "Keep as is" or
                xml_utf8_setting != "Keep as is"):
                self.logger.info("=" * 80)
                self.logger.info("MODIFYING CHARACTER ENCODING")
                self.logger.info("=" * 80)
                
                # Call external function to modify character encoding
                if not modify_character_encoding(self.mxl_file, self.map_name or "", self.checklist_file, self.mapping_file):
                    self.logger.error("Failed to modify character encoding")
                    return False
            
            # Check if freeformat script should be triggered
            if freeformat_value != 'no':
                self.logger.info("=" * 80)
                self.logger.info(f"TRIGGERING FREEFORMAT SCRIPT (Section: {freeformat_value.upper()})")
                self.logger.info("=" * 80)
                
                try:
                    # Get absolute path for external script
                    abs_mxl = os.path.abspath(self.mxl_file)
                    
                    # Call mxl_processor.py with mxl file and section parameter
                    result = subprocess.run(
                        [sys.executable, 'mxl_processor.py', abs_mxl, freeformat_value],
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    self.logger.info("Freeformat script output:")
                    self.logger.info(result.stdout)
                    if result.stderr:
                        self.logger.warning(f"Freeformat script warnings: {result.stderr}")
                except subprocess.CalledProcessError as e:
                    self.logger.error(f"Freeformat script failed: {e}")
                    self.logger.error(f"Output: {e.stdout}")
                    self.logger.error(f"Error: {e.stderr}")
                    return False
                except FileNotFoundError:
                    self.logger.error("mxl_processor.py not found in current directory")
                    return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error checking Generic_rule sheet: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def extract_map_name(self, content: str) -> str:
        """Extract map name from MXL content."""
        self.logger.debug(f"Searching for Description tag in content (length: {len(content)} chars)")
        match = re.search(r'<Description>([^<]+)</Description>', content)
        if match:
            self.map_name = match.group(1).strip()
            self.logger.info(f"Map name extracted: {self.map_name}")
            return self.map_name
        else:
            # Try to find any Description tag for debugging
            desc_search = re.search(r'<Description[^>]*>.*?</Description>', content, re.DOTALL)
            if desc_search:
                self.logger.error(f"Found Description tag but couldn't extract: {desc_search.group(0)[:200]}")
            else:
                self.logger.error("No Description tag found in MXL file")
        return ""
    
    def get_direction_and_doctype(self) -> Tuple[str, str]:
        """
        Get direction and document type from mapping_results.xlsx based on map name.
        Returns: (direction, document_type)
        """
        try:
            df = pd.read_excel(self.mapping_file, sheet_name='Mapping Details')
            
            # Find the row matching the map name
            matching_row = df[df['Map Name'] == self.map_name]
            
            if matching_row.empty:
                self.logger.warning(f"Map name '{self.map_name}' not found in mapping_results.xlsx")
                return ("", "")
            
            direction = str(matching_row.iloc[0]['Direction']).strip()
            
            # Extract document type from Input Format or Output Format
            input_format = str(matching_row.iloc[0]['Input Format'])
            output_format = str(matching_row.iloc[0]['Output Format'])
            
            # Extract document type (e.g., 850, 856, 204, DELINS, ORDERS, SALESORDER, etc.)
            # Strategy: First try Input Format, if not found then try Output Format
            # Special case: Skip "SAP" as it's just a marker, not a document type
            doc_type = ""
            
            # Format examples:
            # - "EDI (Edifact, DELINS, 003)" - 3 elements
            # - "EDI(X12, 850, 004010)" - 3 elements
            # - "EDI(X12, 850)" - 2 elements
            # - "Positional (SALESORDER)" - 1 element (use this)
            # - "XML (SAP)" - 1 element (skip this)
            
            # First, try Input Format - look for 2 or 3 element pattern
            # Pattern for 3 elements: (something, DOCTYPE, something)
            match = re.search(r'\([^,]+,\s*([A-Z0-9]+)', input_format)
            if match and match.group(1).upper() != 'SAP':
                doc_type = match.group(1)
                self.logger.info(f"Document type found in Input Format (3 elements): {doc_type}")
            else:
                # Pattern for 2 elements: (something, DOCTYPE)
                match = re.search(r'\(([^,]+),\s*([A-Z0-9]+)\)', input_format)
                if match and match.group(2).upper() != 'SAP':
                    doc_type = match.group(2)
                    self.logger.info(f"Document type found in Input Format (2 elements): {doc_type}")
                else:
                    # Pattern for 1 element: (DOCTYPE) - but skip if SAP
                    match = re.search(r'\(([A-Z0-9]+)\)', input_format)
                    if match and match.group(1).upper() != 'SAP':
                        doc_type = match.group(1)
                        self.logger.info(f"Document type found in Input Format (1 element): {doc_type}")
                    else:
                        # If not found in Input Format, try Output Format
                        # Pattern for 3 elements
                        match = re.search(r'\([^,]+,\s*([A-Z0-9]+)', output_format)
                        if match and match.group(1).upper() != 'SAP':
                            doc_type = match.group(1)
                            self.logger.info(f"Document type found in Output Format (3 elements): {doc_type}")
                        else:
                            # Pattern for 2 elements
                            match = re.search(r'\(([^,]+),\s*([A-Z0-9]+)\)', output_format)
                            if match and match.group(2).upper() != 'SAP':
                                doc_type = match.group(2)
                                self.logger.info(f"Document type found in Output Format (2 elements): {doc_type}")
                            else:
                                # Pattern for 1 element - but skip if SAP
                                match = re.search(r'\(([A-Z0-9]+)\)', output_format)
                                if match and match.group(1).upper() != 'SAP':
                                    doc_type = match.group(1)
                                    self.logger.info(f"Document type found in Output Format (1 element): {doc_type}")
                                else:
                                    # Fall back to finding any 3-digit code in Input Format first
                                    match = re.search(r'\b(\d{3})\b', input_format)
                                    if match:
                                        doc_type = match.group(1)
                                        self.logger.info(f"Document type (3-digit) found in Input Format: {doc_type}")
                                    else:
                                        # Finally try Output Format for 3-digit code
                                        match = re.search(r'\b(\d{3})\b', output_format)
                                        if match:
                                            doc_type = match.group(1)
                                            self.logger.info(f"Document type (3-digit) found in Output Format: {doc_type}")
            
            self.direction = direction
            self.document_type = doc_type
            
            self.logger.info(f"Direction: {direction}, Document Type: {doc_type}")
            return (direction, doc_type)
            
        except Exception as e:
            self.logger.error(f"Error reading mapping file: {e}")
            return ("", "")
    
    def get_process_data_updates(self) -> List[Tuple[str, str]]:
        """
        Get process data updates from Generic_checklist.xlsm based on direction and document type.
        Returns list of (xpath, field_position) tuples - each xpath is processed independently.
        
        Handles MapName column for custom changes:
        - If multiple rows match direction+doctype with Yes, checks MapName column
        - If current map name found in MapName column, uses that row
        - If current map name not found, uses row with empty MapName
        """
        try:
            df = pd.read_excel(self.checklist_file, sheet_name='Process_data_Updates')
            
            # Filter by direction and document type
            matching_rows = df[
                (df['Direction'].str.strip().str.lower() == self.direction.lower()) &
                (df['Document Type'].astype(str).str.strip() == self.document_type)
            ]
            
            if matching_rows.empty:
                self.logger.warning(f"No matching rows found for Direction='{self.direction}', Document Type='{self.document_type}'")
                return []
            
            # Filter for rows marked as 'Yes'
            yes_rows = matching_rows[matching_rows['Yes/No'].str.strip().str.upper() == 'YES']
            
            if yes_rows.empty:
                self.logger.warning(f"No rows marked 'Yes' for Direction='{self.direction}', Document Type='{self.document_type}'")
                return []
            
            # Check if there are multiple 'Yes' rows - need to use MapName logic
            if len(yes_rows) > 1:
                self.logger.info(f"Found {len(yes_rows)} rows marked 'Yes' for Direction='{self.direction}', Document Type='{self.document_type}'")
                self.logger.info("Checking MapName column for custom changes...")
                
                # Check if MapName column exists
                if 'MapName (If any custom changes)' in df.columns:
                    selected_row = None
                    
                    # First, try to find a row with matching map name
                    for index, row in yes_rows.iterrows():
                        map_name_cell = row.get('MapName (If any custom changes)', '')
                        if pd.notna(map_name_cell):
                            map_name_value = str(map_name_cell).strip()
                            if map_name_value and map_name_value.upper() in self.map_name.upper():
                                self.logger.info(f"Found matching MapName: '{map_name_value}' for current map '{self.map_name}'")
                                selected_row = row
                                break
                    
                    # If no matching map name found, use row with empty MapName
                    if selected_row is None:
                        for index, row in yes_rows.iterrows():
                            map_name_cell = row.get('MapName (If any custom changes)', '')
                            if pd.isna(map_name_cell) or str(map_name_cell).strip() == '':
                                self.logger.info(f"No matching MapName found for '{self.map_name}', using row with empty MapName")
                                selected_row = row
                                break
                    
                    # If still no row selected, log warning and use first row
                    if selected_row is None:
                        self.logger.warning(f"Could not determine appropriate row using MapName logic, using first 'Yes' row")
                        selected_row = yes_rows.iloc[0]
                    
                    # Convert selected row to DataFrame for consistent processing
                    matching_rows = pd.DataFrame([selected_row])
                else:
                    self.logger.warning("MapName column not found, using first 'Yes' row")
                    matching_rows = yes_rows.head(1)
            else:
                # Only one 'Yes' row, use it
                matching_rows = yes_rows
            
            # Collect all xpath-field pairs
            xpath_field_pairs = []
            
            for index, row in matching_rows.iterrows():
                yes_no = str(row.get('Yes/No', '')).strip().upper()
                
                if yes_no == 'YES':
                    # Check all TranslationOutput columns
                    xpath_columns = [
                        'TranslationOutput/BusinessReference',
                        'TranslationOutput/ExtraInfo1',
                        'TranslationOutput/ExtraInfo2',
                        'TranslationOutput/ExtraInfo3',
                        'TranslationOutput/SourceReferenceNumber'
                    ]
                    
                    # Collect all non-empty xpaths with their field positions
                    for xpath_col in xpath_columns:
                        if xpath_col in row and not pd.isna(row[xpath_col]):
                            field_pos = str(row[xpath_col]).strip()
                            if field_pos.upper() not in ['NAN', 'N/A', 'N/A - PASSTHRU', '']:
                                xpath_field_pairs.append((xpath_col, field_pos))
                                self.logger.info(f"  Found update: XPath='{xpath_col}', Field='{field_pos}'")
            
            self.logger.info(f"Total updates to process: {len(xpath_field_pairs)}")
            return xpath_field_pairs
            
        except Exception as e:
            self.logger.error(f"Error reading checklist file: {e}")
            return []
    
    def parse_field_position(self, field_pos: str) -> Optional[Tuple[str, Union[int, str]]]:
        """
        Parse field position string into segment name and field identifier.
        
        Supports multiple formats:
        - EDI: "BHT-03", "B4-02", "856.BSN-02" -> (segment_name, field_number)
        - NON-EDI: "IDOC.E1EDK01-BSART" -> (segment_name, field_name)
        - Hierarchical: "N104(N1*ST)" -> extracts "N1" as segment and "04" as field number
        - Loop prefix: "1000_HL:3.PRF-01" -> strips loop prefix
        
        Returns: (segment_name, field_identifier) where field_identifier is int for EDI, str for NON-EDI
        """
        original_field_pos = field_pos
        
        # Strip hierarchical loop prefix if present (e.g., "1000_HL:3.PRF-01" -> "PRF-01")
        if ':' in field_pos:
            # Extract the part after the colon
            field_pos = field_pos.split(':', 1)[1]
            self.logger.info(f"Stripped loop prefix: '{original_field_pos}' -> '{field_pos}'")
        
        # Handle hierarchical XML paths (e.g., "SPI_INVOICES.INVOICE.CUSTOMER_ID")
        # These represent nested XML elements and should use the last element as the field name
        if '.' in field_pos and '-' not in field_pos:
            parts = field_pos.split('.')
            if len(parts) >= 2:
                # Use the second-to-last part as segment/group name and last part as field name
                segment_name = parts[-2]
                field_name = parts[-1]
                self.logger.info(f"Parsed hierarchical XML path '{original_field_pos}' as segment='{segment_name}', field_name='{field_name}'")
                return (segment_name, field_name)
        
        # Strip document type prefix if present (e.g., "856.BSN-02" -> "BSN-02")
        if '.' in field_pos and '-' in field_pos:
            # Check if the part before the dot is numeric (document type)
            prefix = field_pos.split('.')[0]
            if prefix.isdigit():
                field_pos = field_pos.split('.', 1)[1]
                self.logger.info(f"Stripped document type prefix: '{original_field_pos}' -> '{field_pos}'")
        
        # Handle special case like "N104(N1*ST)" - extract segment from parentheses
        if '(' in field_pos and '*' in field_pos:
            # Extract the segment name from parentheses (e.g., "N1" from "N1*ST")
            paren_content = field_pos[field_pos.index('(')+1:field_pos.index(')')]
            if '*' in paren_content:
                segment_name = paren_content.split('*')[0].strip()
                # Extract field number from the part before parentheses (e.g., "04" from "N104")
                before_paren = field_pos.split('(')[0].strip()
                # Get the numeric part - take last 2 digits as field number
                field_num_str = ''.join([c for c in before_paren if c.isdigit()])
                if segment_name and field_num_str:
                    # Take last 2 digits as field number (e.g., "04" from "104")
                    if len(field_num_str) > 2:
                        field_num_str = field_num_str[-2:]
                    try:
                        field_number = int(field_num_str)
                        self.logger.info(f"Parsed '{original_field_pos}' as segment='{segment_name}', field={field_number}")
                        return (segment_name, field_number)
                    except ValueError:
                        pass
        
        # Remove any remaining parenthetical content for standard parsing
        if '(' in field_pos:
            field_pos = field_pos.split('(')[0].strip()
        
        if '-' in field_pos:
            parts = field_pos.split('-')
            segment_name = parts[0].strip()
            field_identifier = parts[1].strip()
            
            # Try to parse as integer (EDI format like "856.PRF-01" or "PRF-01")
            try:
                field_number = int(field_identifier)
                # Check if segment_name contains a dot (e.g., "856.PRF")
                if '.' in segment_name:
                    # Format: "856.PRF-01" -> name="856", tag="PRF", field=01
                    name_parts = segment_name.split('.')
                    return (segment_name, field_number)  # Keep full format for EDI
                else:
                    # Format: "PRF-01" -> segment="PRF", field=01
                    return (segment_name, field_number)
            except ValueError:
                # NON-EDI format - field_identifier is a field name
                # Check if segment_name contains dots (hierarchical path like "SPI_ASN.ASN")
                if '.' in segment_name:
                    # Format: "SPI_ASN.ASN-CUSTOMER_ID" -> hierarchical NON-EDI
                    self.logger.info(f"Parsed '{original_field_pos}' as hierarchical NON-EDI: path='{segment_name}', field_name='{field_identifier}'")
                else:
                    # Format: "CUSTOMER-ID" -> simple NON-EDI
                    self.logger.info(f"Parsed '{original_field_pos}' as NON-EDI: segment='{segment_name}', field_name='{field_identifier}'")
                return (segment_name, field_identifier)
        else:
            # Try to extract segment name and number without dash
            segment_name = ''.join([c for c in field_pos if c.isalpha()])
            field_num_str = ''.join([c for c in field_pos if c.isdigit()])
            
            if segment_name and field_num_str:
                try:
                    field_number = int(field_num_str)
                    return (segment_name, field_number)
                except ValueError:
                    pass
        
        self.logger.error(f"Could not parse field position: '{original_field_pos}'")
        return None
    
    def find_segment_and_field(self, content: str, segment_name: str, field_identifier: Union[int, str], xpath: str = "") -> Optional[Tuple[int, int, str, str]]:
        """
        Find a specific field in a segment or XMLElementGroup within the MXL content.
        
        Args:
            content: Full MXL file content
            segment_name: Segment name (EDI like "PRF" or "856.PRF") or XMLElementGroup path (NON-EDI like "SPI_ASN.ASN")
            field_identifier: Field number (int) for EDI, or field name (str) for NON-EDI
            xpath: XPath column name (e.g., "TranslationOutput/BusinessReference") - optional, for error reporting
            
        Returns: (field_start, field_end, field_name, field_content)
        """
        # Check if this is NON-EDI format (field_identifier is a string)
        if isinstance(field_identifier, str):
            # Try XML format first (XMLTag hierarchy)
            result = self.find_xml_element_field(content, segment_name, field_identifier, xpath)
            if result:
                return result
            
            # If XML format not found, try NON-XML format (Tag hierarchy)
            self.logger.info(f"XML format not found, trying NON-XML Tag hierarchy for '{segment_name}-{field_identifier}'")
            return self.find_tag_hierarchy_field(content, segment_name, field_identifier, xpath)
        
        # EDI format - field_identifier is an int
        field_number = field_identifier
        
        # Check if segment_name contains a dot (e.g., "856.PRF")
        actual_segment_name = segment_name
        name_value = None
        if '.' in segment_name:
            # Format: "856.PRF" -> name="856", tag="PRF"
            parts = segment_name.split('.')
            name_value = parts[0]
            actual_segment_name = parts[1]
            self.logger.info(f"Parsed EDI format: Name={name_value}, Tag={actual_segment_name}, Field={field_number}")
        
        # Search for ALL segments with the matching <Tag>
        # We need to find all occurrences and check each one for an active field at the target position
        tag_pattern = rf'<Tag>{re.escape(actual_segment_name)}</Tag>'
        tag_matches = list(re.finditer(tag_pattern, content))
        
        if not tag_matches:
            self.logger.warning(f"No segments with Tag='{actual_segment_name}' found")
            return None
        
        self.logger.info(f"Found {len(tag_matches)} segment(s) with Tag='{actual_segment_name}'")
        
        # Initialize field_content to track the last checked field (for error reporting)
        field_content = None
        
        # Try each segment occurrence until we find one with an active field at the target position
        for tag_idx, tag_match in enumerate(tag_matches):
            # Find the parent Segment/Loop containing this <Tag>
            search_start = max(0, tag_match.start() - 2000)
            before_tag = content[search_start:tag_match.start()]
            
            # Find the last opening tag before <Tag>
            parent_start = before_tag.rfind('<Segment>')
            if parent_start == -1:
                parent_start = before_tag.rfind('<Loop>')
            if parent_start == -1:
                continue
            
            parent_start = search_start + parent_start
            
            # Find the closing tag after <Tag>
            after_tag_start = tag_match.end()
            segment_end_match = re.search(r'</Segment>|</Loop>', content[after_tag_start:])
            if not segment_end_match:
                continue
            
            parent_end = after_tag_start + segment_end_match.end()
            segment_content = content[parent_start:parent_end]
            
            # If name_value is specified, check if this segment has the matching <Name>
            if name_value:
                name_pattern = rf'<Name>{re.escape(name_value)}(?::\d+)?</Name>'
                if not re.search(name_pattern, segment_content):
                    continue  # Skip this segment, doesn't match the Name requirement
            
            # Find all fields in this segment
            field_pattern = r'<Field>.*?</Field>'
            fields = list(re.finditer(field_pattern, segment_content, re.DOTALL))
            
            if field_number < 1 or field_number > len(fields):
                self.logger.debug(f"Segment occurrence #{tag_idx + 1}: Field #{field_number} not found (has {len(fields)} fields)")
                continue
            
            # Get the nth field (1-based index)
            target_field_index = field_number - 1
            field_match = fields[target_field_index]
            field_content = field_match.group(0)
            
            # Check if this field is active
            if self.is_field_active(field_content):
                self.logger.info(f"Found active field #{field_number} in segment occurrence #{tag_idx + 1} of Tag='{actual_segment_name}'")
                segment_start = parent_start
                segment_end = parent_end
                break
            else:
                self.logger.debug(f"Segment occurrence #{tag_idx + 1}: Field #{field_number} is inactive, trying next segment")
        else:
            # No active field found - but we can still extract the field name from the inactive field
            field_position = f"{segment_name}-{field_number:02d}" if isinstance(field_number, int) else f"{segment_name}-{field_number}"
            self.logger.warning(f"No active field #{field_number} found in any occurrence of segment '{actual_segment_name}'")
            
            # Try to extract field name from the inactive field for better error reporting
            inactive_field_name = ""
            if field_content:
                name_match = re.search(r'<Name>([^<]+)</Name>', field_content)
                if name_match:
                    inactive_field_name = name_match.group(1)
            
            if self.map_name:
                self.error_reporter.add_inactive_field_error(self.map_name, field_position, inactive_field_name, xpath)
            return None
        
        # Extract field name
        name_match = re.search(r'<Name>([^<]+)</Name>', field_content)
        field_name = name_match.group(1) if name_match else "Unknown"
        
        # Calculate absolute positions
        field_start = segment_start + field_match.start()
        field_end = segment_start + field_match.end()
        
        return (field_start, field_end, field_name, field_content)
    
    def find_xml_element_field(self, content: str, group_path: str, field_name: str, xpath: str = "") -> Optional[Tuple[int, int, str, str]]:
        """
        Find a field in a hierarchical XMLTag structure (XML format).
        
        Args:
            content: Full MXL file content
            group_path: Hierarchical path like "SPI_INVOICES.INVOICE" or simple name like "CUSTOMER"
            field_name: The field tag name to search for (e.g., "CUSTOMER_ID")
            
        Returns: (field_start, field_end, field_name, field_content)
        
        Example:
            For "SPI_INVOICES.INVOICE-CUSTOMER_ID":
            - Searches for <XMLTag>SPI_INVOICES</XMLTag>
            - Within that scope, searches for <XMLTag>INVOICE</XMLTag>
            - Within that scope, searches for <Tag>CUSTOMER_ID</Tag>
            - Updates the field
        """
        # Parse the hierarchical path
        if '.' in group_path:
            # Hierarchical format: "SPI_INVOICES.INVOICE" -> ["SPI_INVOICES", "INVOICE"]
            path_parts = group_path.split('.')
            self.logger.info(f"Parsing hierarchical XML path: {' > '.join(path_parts)} > {field_name}")
        else:
            # Simple format: just one level
            path_parts = [group_path]
            self.logger.info(f"Parsing simple XML: {group_path} > {field_name}")
        
        # Start searching from the beginning of content
        current_content = content
        current_offset = 0
        
        # Navigate through each level of XMLTag hierarchy
        for i, xml_tag_name in enumerate(path_parts):
            # Search for <XMLTag>xml_tag_name</XMLTag>
            xml_tag_pattern = rf'<XMLTag>{re.escape(xml_tag_name)}</XMLTag>'
            xml_tag_match = re.search(xml_tag_pattern, current_content)
            
            if not xml_tag_match:
                self.logger.warning(f"XMLTag '{xml_tag_name}' not found at level {i + 1}")
                return None
            
            # Find the scope of this XMLTag by looking for the next XMLTag at the same or higher level
            # The scope ends at the next sibling XMLTag or parent's closing tag
            # For simplicity, we'll search within a reasonable range after this XMLTag
            scope_start = xml_tag_match.end()
            
            # Find the next XMLTag at the same level (sibling) or closing tag
            # Look ahead up to 50000 characters for the scope
            scope_end = min(len(current_content), scope_start + 50000)
            
            # Update current_content to search within this XMLTag's scope
            absolute_scope_start = current_offset + scope_start
            current_content = content[absolute_scope_start:absolute_scope_start + (scope_end - scope_start)]
            current_offset = absolute_scope_start
            
            self.logger.info(f"Found XMLTag '{xml_tag_name}' at level {i + 1}, searching within scope (length: {len(current_content)})")
        
        # Now search for XMLRecord with <Tag>field_name</Tag> within the final XMLElementGroup
        # In NON-EDI XML format, the structure is:
        # <XMLElementGroup><XMLTag>CUSTOMER_ID</XMLTag>...<XMLRecord><Tag>CUSTOMER_ID</Tag><Field>...</Field></XMLRecord></XMLElementGroup>
        
        self.logger.info(f"Searching for <Tag>{field_name}</Tag> in final XMLElementGroup (content length: {len(current_content)})")
        
        # Look for all XMLRecords with matching <Tag>
        tag_pattern = rf'<Tag>{re.escape(field_name)}</Tag>'
        tag_matches = list(re.finditer(tag_pattern, current_content))
        
        self.logger.info(f"Found {len(tag_matches)} matches for <Tag>{field_name}</Tag>")
        
        if not tag_matches:
            self.logger.warning(f"No XMLRecord with Tag '{field_name}' found in the final XMLElementGroup")
            # Debug: Show a sample of the content
            sample = current_content[:500] if len(current_content) > 500 else current_content
            self.logger.debug(f"Content sample: {sample}")
            return None
        
        # Search for the first active field with this tag
        # Try each occurrence until we find an active one
        field_content = None  # Initialize to track last checked field for error reporting
        for tag_idx, tag_match in enumerate(tag_matches):
            # Find the parent <XMLRecord> element containing this <Tag>
            search_start = max(0, tag_match.start() - 3000)
            before_tag = current_content[search_start:tag_match.start()]
            
            record_start_pos = before_tag.rfind('<XMLRecord>')
            if record_start_pos == -1:
                continue
            
            record_start = search_start + record_start_pos
            
            # Find the closing </XMLRecord>
            after_tag_start = tag_match.end()
            record_end_match = re.search(r'</XMLRecord>', current_content[after_tag_start:])
            if not record_end_match:
                continue
            
            record_end = after_tag_start + record_end_match.end()
            record_content = current_content[record_start:record_end]
            
            # Now find the <Field> within this XMLRecord
            field_pattern = r'<Field>.*?</Field>'
            field_match = re.search(field_pattern, record_content, re.DOTALL)
            
            if not field_match:
                self.logger.warning(f"No Field found in XMLRecord with Tag '{field_name}'")
                continue
            
            field_content = field_match.group(0)
            
            # Check if this field is active
            if self.is_field_active(field_content):
                # Extract the field name
                name_match = re.search(r'<Name>([^<]+)</Name>', field_content)
                actual_field_name = name_match.group(1) if name_match else field_name
                
                # For SAP maps, verify the Name matches the expected pattern
                # Should match "SNDPRN" or "SNDPRN:2" but NOT "Temp_SNDPRN"
                # Pattern: Name should start with field_name followed by optional ":number"
                if actual_field_name:
                    # Check if name starts with field_name
                    if actual_field_name.startswith(field_name):
                        # Check what comes after field_name
                        suffix = actual_field_name[len(field_name):]
                        # Valid if: empty (exact match) or starts with ":" (e.g., ":2")
                        if suffix == "" or suffix.startswith(":"):
                            self.logger.info(f"Name '{actual_field_name}' matches pattern for field '{field_name}'")
                        else:
                            # Has other characters after field_name (e.g., "Temp_SNDPRN" where suffix would be "PRN")
                            self.logger.debug(f"Name '{actual_field_name}' does not match pattern (has suffix '{suffix}'), trying next occurrence")
                            continue
                    else:
                        # Name doesn't start with field_name at all
                        self.logger.debug(f"Name '{actual_field_name}' does not start with '{field_name}', trying next occurrence")
                        continue
                
                # Calculate absolute positions
                field_start_in_record = field_match.start()
                field_end_in_record = field_match.end()
                abs_field_start = current_offset + record_start + field_start_in_record
                abs_field_end = current_offset + record_start + field_end_in_record
                
                self.logger.info(f"Found active field with Tag '{field_name}' (occurrence #{tag_idx + 1}) and Name '{actual_field_name}'")
                return (abs_field_start, abs_field_end, actual_field_name, field_content)
            else:
                self.logger.debug(f"Tag '{field_name}' occurrence #{tag_idx + 1} is inactive, trying next occurrence")
        
        # If no active field found, log warning and report error
        field_position = f"{group_path}.{field_name}"
        self.logger.warning(f"No active field with Tag '{field_name}' found (checked {len(tag_matches)} occurrence(s))")
        
        # Try to extract field name from the inactive field for better error reporting
        inactive_field_name = ""
        if field_content:
            name_match = re.search(r'<Name>([^<]+)</Name>', field_content)
            if name_match:
                inactive_field_name = name_match.group(1)
        
        if self.map_name:
            self.error_reporter.add_inactive_field_error(self.map_name, field_position, inactive_field_name, xpath)
        return None
    def find_tag_hierarchy_field(self, content: str, group_path: str, field_name: str, xpath: str = "") -> Optional[Tuple[int, int, str, str]]:
        """
        Find a field in a hierarchical <Tag> structure (NON-XML, NON-EDI format).
        
        Args:
            content: Full MXL file content
            group_path: Hierarchical path like "SPI_INVOICES.INVOICE" or simple name like "CUSTOMER"
            field_name: The field name to search for (e.g., "CUSTOMER_ID")
            xpath: XPath column name (e.g., "TranslationOutput/SourceReferenceNumber") - optional
            
        Returns: (field_start, field_end, field_name, field_content)
        
        Example:
            For "SPI_INVOICES.INVOICE-CUSTOMER_ID":
            - Searches for <Tag>SPI_INVOICES</Tag>
            - Within that Segment/Loop, searches for <Tag>INVOICE</Tag>
            - Within that Segment/Loop, searches for first active field with <Name>CUSTOMER_ID</Name>
              (also matches CUSTOMER_ID:2, CUSTOMER_ID:02, CUSTOMER_ID_temp, etc.)
        """
        # Parse the hierarchical path
        if '.' in group_path:
            # Hierarchical format: "SPI_INVOICES.INVOICE" -> ["SPI_INVOICES", "INVOICE"]
            path_parts = group_path.split('.')
            self.logger.info(f"Parsing hierarchical NON-XML Tag path: {' > '.join(path_parts)} > {field_name}")
        else:
            # Simple format: just one level
            path_parts = [group_path]
            self.logger.info(f"Parsing simple NON-XML Tag: {group_path} > {field_name}")
        
        # Start searching from the beginning of content
        current_content = content
        current_offset = 0
        parent_tag = 'Segment'  # Initialize to default value
        
        # Navigate through each level of <Tag> hierarchy
        for i, tag_name in enumerate(path_parts):
            # Search for <Tag>tag_name</Tag>
            tag_pattern = rf'<Tag>{re.escape(tag_name)}</Tag>'
            tag_match = re.search(tag_pattern, current_content)
            
            if not tag_match:
                self.logger.warning(f"Tag '{tag_name}' not found at level {i + 1}")
                return None
            
            # Find the parent Segment or Loop containing this Tag
            # Search backwards to find the opening <Segment> or <Loop>
            search_start = max(0, tag_match.start() - 2000)
            before_tag = current_content[search_start:tag_match.start()]
            
            # Find the last <Segment> or <Loop> before this Tag
            segment_start_pos = before_tag.rfind('<Segment>')
            loop_start_pos = before_tag.rfind('<Loop>')
            
            # Use whichever is closer to the tag
            if segment_start_pos == -1 and loop_start_pos == -1:
                self.logger.warning(f"No Segment or Loop found before Tag '{tag_name}'")
                return None
            
            if segment_start_pos > loop_start_pos:
                parent_start_pos = segment_start_pos
                parent_tag = 'Segment'
            else:
                parent_start_pos = loop_start_pos
                parent_tag = 'Loop'
            
            parent_start = search_start + parent_start_pos
            
            # Find the matching closing tag after the Tag
            # Need to handle nested Segments/Loops properly
            after_tag_start = tag_match.end()
            search_pos = after_tag_start
            nesting_level = 1  # We're inside one Segment/Loop
            parent_end = -1
            
            while nesting_level > 0 and search_pos < len(current_content):
                # Find next opening or closing tags
                next_segment_open = current_content.find('<Segment>', search_pos)
                next_loop_open = current_content.find('<Loop>', search_pos)
                next_segment_close = current_content.find('</Segment>', search_pos)
                next_loop_close = current_content.find('</Loop>', search_pos)
                
                # Find the earliest tag
                candidates = []
                if next_segment_open != -1:
                    candidates.append(('open', next_segment_open, '<Segment>'))
                if next_loop_open != -1:
                    candidates.append(('open', next_loop_open, '<Loop>'))
                if next_segment_close != -1:
                    candidates.append(('close', next_segment_close, '</Segment>'))
                if next_loop_close != -1:
                    candidates.append(('close', next_loop_close, '</Loop>'))
                
                if not candidates:
                    self.logger.warning(f"No closing tag found for {parent_tag} '{tag_name}'")
                    return None
                
                # Sort by position to find the earliest tag
                candidates.sort(key=lambda x: x[1])
                tag_type, tag_pos, tag_str = candidates[0]
                
                if tag_type == 'open':
                    nesting_level += 1
                    search_pos = tag_pos + len(tag_str)
                else:
                    nesting_level -= 1
                    if nesting_level == 0:
                        parent_end = tag_pos + len(tag_str)
                        break
                    search_pos = tag_pos + len(tag_str)
            
            if nesting_level > 0 or parent_end == -1:
                self.logger.warning(f"Unmatched {parent_tag} for Tag '{tag_name}'")
                return None
            
            # Update current_content to search within this Segment/Loop
            # Calculate the absolute position of this parent in the original content
            absolute_parent_start = current_offset + parent_start
            current_content = content[absolute_parent_start:absolute_parent_start + (parent_end - parent_start)]
            current_offset = absolute_parent_start
            
            self.logger.info(f"Found Tag '{tag_name}' at level {i + 1} in {parent_tag} (content length: {len(current_content)})")
        
        # Now search for Field with <Name> matching field_name within the final Segment/Loop
        # Support flexible matching: CUSTOMER_ID, CUSTOMER_ID:2, CUSTOMER_ID:02, CUSTOMER_ID_temp, etc.
        
        self.logger.info(f"Searching for fields with <Name> matching '{field_name}' (with flexible suffixes)")
        
        # Find all Field elements
        field_pattern = r'<Field>.*?</Field>'
        field_matches = list(re.finditer(field_pattern, current_content, re.DOTALL))
        
        self.logger.info(f"Found {len(field_matches)} Field elements in final {parent_tag}")
        
        if not field_matches:
            self.logger.warning(f"No Field elements found in the final {parent_tag}")
            return None
        
        # Search for the first active field with matching Name
        field_content = None  # Initialize to track last checked field for error reporting
        for field_idx, field_match in enumerate(field_matches):
            field_content = field_match.group(0)
            
            # Extract the Name from this field
            name_match = re.search(r'<Name>([^<]+)</Name>', field_content)
            if not name_match:
                continue
            
            actual_field_name = name_match.group(1)
            
            # Check if name matches the pattern
            # Valid patterns: CUSTOMER_ID, CUSTOMER_ID:2, CUSTOMER_ID:02, CUSTOMER_ID_temp
            # Should start with field_name
            if not actual_field_name.startswith(field_name):
                continue
            
            # Check what comes after field_name
            suffix = actual_field_name[len(field_name):]
            # Valid suffixes: empty, :digits, _anything
            if suffix == "" or suffix.startswith(":") or suffix.startswith("_"):
                self.logger.info(f"Field #{field_idx + 1}: Name '{actual_field_name}' matches pattern for '{field_name}'")
                
                # Check if this field is active
                if self.is_field_active(field_content):
                    # Calculate absolute positions
                    abs_field_start = current_offset + field_match.start()
                    abs_field_end = current_offset + field_match.end()
                    
                    self.logger.info(f"Found active field with Name '{actual_field_name}' (field #{field_idx + 1})")
                    return (abs_field_start, abs_field_end, actual_field_name, field_content)
                else:
                    self.logger.debug(f"Field '{actual_field_name}' (field #{field_idx + 1}) is inactive, trying next field")
            else:
                self.logger.debug(f"Field #{field_idx + 1}: Name '{actual_field_name}' has invalid suffix '{suffix}', skipping")
        
        # If no active field found, log warning and report error
        field_position = f"{group_path}-{field_name}"
        self.logger.warning(f"No active field with Name matching '{field_name}' found (checked {len(field_matches)} field(s))")
        
        # Try to extract field name from the inactive field for better error reporting
        inactive_field_name = ""
        if field_content:
            name_match = re.search(r'<Name>([^<]+)</Name>', field_content)
            if name_match:
                inactive_field_name = name_match.group(1)
        
        if self.map_name:
            self.error_reporter.add_inactive_field_error(self.map_name, field_position, inactive_field_name, xpath)
        return None
    
    
    def is_field_active(self, field_content: str) -> bool:
        """
        Check if a field is active by looking for <Active>1</Active> tag.
        Returns True if active or if Active tag is not found (assume active by default).
        """
        active_match = re.search(r'<Active>(\d+)</Active>', field_content)
        if active_match:
            return active_match.group(1) == '1'
        return True  # If no Active tag found, assume active
    
    def is_field_in_output_section(self, content: str, field_start: int) -> bool:
        """
        Check if a field is in the OUTPUT section of the MXL file.
        
        Args:
            content: Full MXL file content
            field_start: Starting position of the field in content (absolute character position)
            
        Returns:
            True if field is in OUTPUT section, False if in INPUT section
        """
        # Find the position of </INPUT><OUTPUT> boundary
        output_match = re.search(r'</INPUT>\s*<OUTPUT>', content)
        if not output_match:
            # If no OUTPUT section found, assume it's in INPUT
            self.logger.debug(f"No OUTPUT section found in MXL file")
            return False
        
        # The OUTPUT section starts after </INPUT><OUTPUT>
        output_start = output_match.end()
        is_output = field_start >= output_start
        
        # Convert character positions to approximate line numbers for debugging
        lines_before_field = content[:field_start].count('\n') + 1
        lines_before_output = content[:output_start].count('\n') + 1
        
        self.logger.info(f"Field at char position {field_start} (approx line {lines_before_field}), OUTPUT starts at char position {output_start} (approx line {lines_before_output}), is_output={is_output}")
        
        return is_output
    
    def has_mapping(self, field_content: str) -> bool:
        """
        Check if an output field has mapping by looking for <Link>, <ImplicitRuleDef>, or <ExplicitRule> tags.
        
        Args:
            field_content: The field content to check
            
        Returns:
            True if field has any mapping, False otherwise
        """
        # Check for <Link> tag (even if empty like <Link></Link>)
        if re.search(r'<Link>.*?</Link>', field_content, re.DOTALL):
            return True
        
        # Check for <ImplicitRuleDef> tag
        if re.search(r'<ImplicitRuleDef>', field_content):
            return True
        
        # Check for <ExplicitRule> tag
        if re.search(r'<ExplicitRule>', field_content):
            return True
        
        return False
    
    def insert_or_update_explicit_rule_single(self, field_content: str, xpath: str, field_name: str,
                                             needs_review: bool, is_non_edi: bool) -> Optional[str]:
        """
        Insert or update ExplicitRule for a single xpath in field content.
        
        Args:
            field_content: The field content to update
            xpath: The xpath to add (e.g., "TranslationOutput/BusinessReference")
            field_name: The field name to use in SQL (e.g., "0396" or "BSART")
            needs_review: Whether to add "Please review the code" comment
            is_non_edi: Whether this is a NON-EDI format
        """
        # Create the SQL statement
        sql = f'update processdata set xpathresult = #{field_name} where xpath = "{xpath}";'
        if needs_review:
            sql += '\nPlease review the code.'
        
        # Check if ExplicitRule already exists
        explicit_rule_pattern = r'(<ExplicitRule>)(.*?)(</ExplicitRule>)'
        explicit_rule_match = re.search(explicit_rule_pattern, field_content, re.DOTALL)
        
        if explicit_rule_match:
            # ExplicitRule exists - check if this xpath already exists (commented or uncommented)
            existing_content = explicit_rule_match.group(2)
            # Match both commented and uncommented versions
            # Pattern matches: xpath = "..." OR // ... xpath = "..."
            xpath_pattern = f'(?://.*?)?xpath\\s*=\\s*"{re.escape(xpath)}"'
            
            if re.search(xpath_pattern, existing_content):
                self.logger.info(f"XPath '{xpath}' already exists in field '{field_name}' (may be commented), skipping")
                return None
            
            # Append new SQL - preserve existing content exactly as is (including // comments)
            if existing_content.strip():
                # Don't use rstrip() - preserve all whitespace and comments
                new_content = existing_content + "\n" + sql
            else:
                new_content = sql
            
            updated_field = field_content.replace(
                explicit_rule_match.group(0),
                f'{explicit_rule_match.group(1)}{new_content}{explicit_rule_match.group(3)}'
            )
            return updated_field
        else:
            # Need to add ExplicitRule - DO NOT remove <Link> or <ImplicitRule> tags
            # Priority: Insert after </Link> first, then </StoreLimit> if Link not found
            updated_field = field_content
            
            # Try to insert after </Link> first
            link_match = re.search(r'</Link>', updated_field)
            if link_match:
                insertion_point = link_match.end()
                updated_field = (
                    updated_field[:insertion_point] +
                    f'\n<ExplicitRule>{sql}</ExplicitRule>' +
                    updated_field[insertion_point:]
                )
                return updated_field
            
            # If no </Link>, try after </StoreLimit>
            store_limit_match = re.search(r'</StoreLimit>', updated_field)
            if store_limit_match:
                insertion_point = store_limit_match.end()
                updated_field = (
                    updated_field[:insertion_point] +
                    f'\n<ExplicitRule>{sql}</ExplicitRule>' +
                    updated_field[insertion_point:]
                )
                return updated_field
            
            self.logger.warning(f"Could not find insertion point for ExplicitRule in field '{field_name}'")
            return None
    
    def insert_or_update_explicit_rule(self, content: str, field_content: str, xpath_field_pairs: List[Tuple[str, str]]) -> Optional[str]:
        """
        Insert or update ExplicitRule in field content with multiple xpath-field pairs.
        Each xpath uses its own field position to find the field name.
        If ExplicitRule exists, append to it.
        If not, insert after <Link></Link> or </StoreLimit>.
        
        Args:
            content: Full MXL file content (to search for field names)
            field_content: The specific field content to update
            xpath_field_pairs: List of (xpath, field_position) tuples
        """
        # Generate SQL statements - each xpath gets its own field name
        new_sqls = []
        for xpath, field_pos in xpath_field_pairs:
            # Parse the field position and find the field name
            result = self.parse_field_position(field_pos)
            if not result:
                self.logger.warning(f"Could not parse field position '{field_pos}' for xpath '{xpath}'")
                continue
            
            segment_name, field_number = result
            field_info = self.find_segment_and_field(content, segment_name, field_number, xpath)
            
            if field_info:
                _, _, field_name, _ = field_info
                sql = f'update processdata set xpathresult = #{field_name} where xpath = "{xpath}";'
                new_sqls.append((xpath, sql))
            else:
                self.logger.warning(f"Could not find field for '{field_pos}' (xpath: '{xpath}')")
        
        if not new_sqls:
            return None
        
        # Check if ExplicitRule already exists
        explicit_rule_pattern = r'(<ExplicitRule>)(.*?)(</ExplicitRule>)'
        explicit_rule_match = re.search(explicit_rule_pattern, field_content, re.DOTALL)
        
        if explicit_rule_match:
            # ExplicitRule exists - check for duplicates and add missing ones
            existing_content = explicit_rule_match.group(2)
            
            # Filter out xpaths that already exist (commented or uncommented)
            sqls_to_add = []
            for xpath, sql in new_sqls:
                # Match both commented and uncommented versions
                # Pattern matches: xpath = "..." OR // ... xpath = "..."
                xpath_pattern = f'(?://.*?)?xpath\\s*=\\s*"{re.escape(xpath)}"'
                if sql not in existing_content and not re.search(xpath_pattern, existing_content):
                    sqls_to_add.append(sql)
                else:
                    self.logger.info(f"XPath '{xpath}' already exists (may be commented), skipping")
            
            if not sqls_to_add:
                return None  # All xpaths already exist
            
            # Append new SQLs - preserve existing content exactly as is (including // comments)
            if existing_content.strip():
                # Don't use rstrip() - preserve all whitespace and comments
                new_content = existing_content + "\n" + "\n".join(sqls_to_add)
            else:
                new_content = "\n".join(sqls_to_add)
            
            updated_field = field_content.replace(
                explicit_rule_match.group(0),
                f'{explicit_rule_match.group(1)}{new_content}{explicit_rule_match.group(3)}'
            )
            return updated_field
        else:
            # Need to add ExplicitRule with all SQLs
            all_sqls = "\n".join([sql for _, sql in new_sqls])
            
            # First, check for <Link></Link> or <Link>...</Link>
            link_pattern = r'<Link>.*?</Link>'
            link_match = re.search(link_pattern, field_content, re.DOTALL)
            
            if link_match:
                # Insert after </Link>
                insert_pos = link_match.end()
                new_explicit_rule = f'\n<ExplicitRule>{all_sqls}</ExplicitRule>'
                updated_field = field_content[:insert_pos] + new_explicit_rule + field_content[insert_pos:]
                return updated_field
            else:
                # Check for </StoreLimit>
                store_limit_pattern = r'</StoreLimit>'
                store_limit_match = re.search(store_limit_pattern, field_content)
                
                if store_limit_match:
                    # Insert after </StoreLimit>
                    insert_pos = store_limit_match.end()
                    new_explicit_rule = f'\n<ExplicitRule>{all_sqls}</ExplicitRule>'
                    updated_field = field_content[:insert_pos] + new_explicit_rule + field_content[insert_pos:]
                    return updated_field
        
        self.logger.warning(f"Could not find insertion point for ExplicitRule")
        return None
    
    def process_updates(self) -> bool:
        """
        Main processing function.
        
        Returns:
            bool: True if matching 'Yes' rows were found and processing occurred,
                  False if no matching 'Yes' rows were found
        """
        self.logger.info("=" * 80)
        self.logger.info("PROCESS DATA UPDATER V3 - ENHANCED WITH MAPPING INTEGRATION")
        self.logger.info("=" * 80)
        
        try:
            # Read MXL file with UTF-8 encoding
            with open(self.mxl_file, 'r', encoding='UTF-8', errors='replace') as f:
                content = f.read()
            
            original_content = content
            
            # Step 1: Extract map name
            map_name = self.extract_map_name(content)
            if not map_name:
                self.logger.error("Could not extract map name from MXL file. Aborting.")
                return False
            
            # Step 2: Get direction and document type from mapping file
            direction, doc_type = self.get_direction_and_doctype()
            if not direction or not doc_type:
                self.logger.error("Could not determine direction and document type. Aborting.")
                return False
            
            # Step 3: Check Generic_rule sheet and trigger external scripts
            # Note: Generic_rule processing runs regardless of Yes/No in process_data_updates
            if not self.check_and_trigger_generic_rules():
                self.logger.error("Generic rule processing failed. Aborting.")
                return False
            
            # Step 4: Get process data updates from checklist
            # Note: This respects Yes/No column - only processes rows marked as 'Yes'
            updates = self.get_process_data_updates()
            
            # Re-read MXL file as it may have been modified by external scripts
            with open(self.mxl_file, 'r', encoding='UTF-8', errors='replace') as f:
                content = f.read()
            
            # Step 5: Process each xpath-field pair independently (if any)
            if not updates:
                self.logger.info("No field mappings to process, but Generic_rule processing completed successfully.")
                self.logger.info("=" * 80)
                self.logger.info("PROCESSING COMPLETE")
                self.logger.info("=" * 80)
                return True
            
            # Process field mappings
            modifications_made = 0
            
            for xpath, field_pos in updates:
                result = self.parse_field_position(field_pos)
                if not result:
                    continue
                
                segment_name, field_identifier = result
                is_non_edi = isinstance(field_identifier, str)
                
                # Find the field in the current content
                field_info = self.find_segment_and_field(content, segment_name, field_identifier, xpath)
                if not field_info:
                    continue
                
                field_start, field_end, field_name, field_content = field_info
                
                # Check if this field is in the OUTPUT section of the MXL file
                # Only check for mapping if field is in OUTPUT section
                if self.is_field_in_output_section(content, field_start):
                    # Field is in OUTPUT section and active (we found it), but check if it has mapping
                    if not self.has_mapping(field_content):
                        # Active output field with no mapping - report error
                        if self.map_name:
                            self.error_reporter.add_no_mapping_error(self.map_name, field_pos, field_name, xpath)
                        self.logger.warning(f"Output field '{field_pos}' has no mapping (<Link>, <ImplicitRuleDef>, or <ExplicitRule>)")
                        continue  # Skip this field, don't try to add ExplicitRule
                
                # Check if field has qualifier (needs review)
                needs_review = '(' in field_pos and ')' in field_pos
                
                # Insert or update ExplicitRule for this single xpath
                updated_field = self.insert_or_update_explicit_rule_single(
                    field_content, xpath, field_name, needs_review, is_non_edi
                )
                
                if updated_field:
                    # Replace in content
                    content = content[:field_start] + updated_field + content[field_end:]
                    
                    modifications_made += 1
                    format_type = "NON-EDI" if is_non_edi else "EDI"
                    self.logger.info(f"✓ Updated {format_type} field '{field_name}' with xpath '{xpath}'")
            
            # Write back if changes were made
            if modifications_made > 0:
                if self.dry_run:
                    self.logger.info(f"DRY RUN: Would write {modifications_made} modifications")
                else:
                    with open(self.mxl_file, 'w', encoding='UTF-8') as f:
                        f.write(content)
                    self.logger.info(f"Successfully wrote {modifications_made} modifications to {self.mxl_file}")
            else:
                self.logger.info("No modifications were made")
            
            # Save error report if any errors were found
            if self.error_reporter.has_errors():
                self.logger.info("=" * 80)
                self.error_reporter.print_summary()
                if self.error_reporter.save_report():
                    self.logger.info(f"Error report saved to: {self.error_reporter.report_file}")
                else:
                    self.logger.error("Failed to save error report")
            
            self.logger.info("=" * 80)
            self.logger.info("PROCESSING COMPLETE")
            self.logger.info("=" * 80)
            return True
            
        except Exception as e:
            self.logger.error(f"Error during processing: {e}")
            import traceback
            traceback.print_exc()
            return True  # Return True because matching 'Yes' rows were found, even though processing failed
            raise


def main():
    if len(sys.argv) < 4:
        print("Usage: python process_data_updater_v3.py <mapping_results.xlsx> <Generic_checklist.xlsm> <mxl_file> [--dry-run]")
        print("\nExample:")
        print("  python process_data_updater_v3.py mapping_results.xlsx Generic_checklistMain.xlsm mxl_files/WTSUS_AD_I_850_4010.mxl")
        sys.exit(1)
    
    mapping_file = sys.argv[1]
    checklist_file = sys.argv[2]
    mxl_file = sys.argv[3]
    dry_run = '--dry-run' in sys.argv
    
    updater = ProcessDataUpdaterV3(mapping_file, checklist_file, mxl_file, dry_run)
    has_matching_rows = updater.process_updates()
    
    # Exit with code 2 if no matching 'Yes' rows were found
    # This signals to process_all_mxl_files.py to skip inbound_mapsFeatures
    if not has_matching_rows:
        sys.exit(2)
    
    # Process inbound maps features (syntax record, positional, XML, CSV) for Inbound maps
    if updater.direction and updater.direction.lower() == 'inbound':
        logging.info("=" * 80)
        logging.info("PROCESSING INBOUND MAPS FEATURES")
        logging.info("=" * 80)
        try:
            success = process_inbound_features(checklist_file, mxl_file)
            if success:
                logging.info("Inbound Maps Features processing completed successfully")
            else:
                logging.warning("Inbound Maps Features processing completed with warnings")
        except Exception as e:
            logging.error(f"Error processing Inbound Maps Features: {e}")
            import traceback
            logging.error(traceback.format_exc())


if __name__ == '__main__':
    main()

# Made with Bob
