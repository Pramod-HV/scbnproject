#!/usr/bin/env python3
"""
Character Encoding Modifier for MXL Files

This script modifies character encoding settings in Sterling B2Bi MXL map files
based on configuration in Generic_checklistMain.xlsm.

Usage:
    python modify_character_encoding.py

Configuration:
    - Generic_checklistMain.xlsm (Generic_rule sheet)
      - Row 2: Set input encoding
      - Row 3: Set output encoding  
      - Row 4: If output is XML, set encoding to UTF8
    
    - mapping_results.xlsx
      - Used to determine if output format is XML

Behavior:
    - "Keep as is" → No changes to file
    - "NONE" → Deletes entire encoding tag
    - Any other value → Sets encoding to that value
"""

import sys
import os
import re
import logging
from pathlib import Path
from datetime import datetime
import pandas as pd

# Configure logging - only create log file if running as main script
# When imported, use the calling module's logger
if __name__ == "__main__":
    log_filename = f"character_encoding_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
else:
    # When imported, just get a logger (will use parent's configuration)
    logger = logging.getLogger(__name__)


def modify_character_encoding(mxl_file: str, map_name: str, checklist_file: str, mapping_file: str) -> bool:
    """
    Modify character encoding in MXL files based on Generic_rule sheet settings.
    Uses regex-based text replacement to preserve original XML formatting.
    
    Args:
        mxl_file: Path to the MXL file
        map_name: Name of the map to check output type
        checklist_file: Path to Generic_checklistMain.xlsm
        mapping_file: Path to mapping_results.xlsx
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Read Generic_rule sheet
        generic_rule_df = pd.read_excel(checklist_file, sheet_name='Generic_rule', header=None)
        
        # Get encoding settings
        input_encoding_setting = str(generic_rule_df.iloc[2, 1]).strip() if pd.notna(generic_rule_df.iloc[2, 1]) else "Keep as is"
        input_encoding_custom = str(generic_rule_df.iloc[2, 2]).strip() if pd.notna(generic_rule_df.iloc[2, 2]) else ""
        
        output_encoding_setting = str(generic_rule_df.iloc[3, 1]).strip() if pd.notna(generic_rule_df.iloc[3, 1]) else "Keep as is"
        output_encoding_custom = str(generic_rule_df.iloc[3, 2]).strip() if pd.notna(generic_rule_df.iloc[3, 2]) else ""
        
        xml_utf8_setting = str(generic_rule_df.iloc[4, 1]).strip() if pd.notna(generic_rule_df.iloc[4, 1]) else "Keep as is"
        xml_utf8_custom = str(generic_rule_df.iloc[4, 2]).strip() if pd.notna(generic_rule_df.iloc[4, 2]) else ""
        
        # Check input and output formats from mapping_results
        is_xml_output = False
        is_xml_input = False
        try:
            mapping_df = pd.read_excel(mapping_file)
            map_row = mapping_df[mapping_df['Map Name'] == map_name]
            if not map_row.empty:
                output_format = str(map_row.iloc[0]['Output Format']).strip().upper()
                is_xml_output = 'XML' in output_format
                
                # Check input format
                input_format = str(map_row.iloc[0]['Input Format']).strip().upper()
                is_xml_input = 'XML' in input_format
        except Exception as e:
            logger.warning(f"Could not determine input/output format for {map_name}: {e}")
        
        # Read the file as text
        with open(mxl_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        modified = False
        
        # Determine what encoding values to use
        # None = Keep as is (no changes)
        # "EMPTY" = Empty the tag content (not delete the tag)
        # Any other value = Set the tag to that value
        input_encoding_value = None
        output_encoding_value = None
        
        # Process Input encoding - SKIP if input is XML
        if not is_xml_input:
            if input_encoding_setting != "Keep as is" and input_encoding_setting.lower() != "keep as is":
                if input_encoding_setting.upper() == "NONE":
                    input_encoding_value = "EMPTY"  # Special marker to empty the tag content
                elif input_encoding_setting == "Custom":
                    if input_encoding_custom and input_encoding_custom != "Enter the details":
                        input_encoding_value = input_encoding_custom
                else:
                    # Remove anything in brackets like "Cp 1252(#something)"
                    clean_value = input_encoding_setting.split('(')[0].strip()
                    # Double check it's not a variant of "Keep as is"
                    if clean_value.lower() != "keep as is":
                        input_encoding_value = clean_value
        else:
            logger.info(f"Skipping INPUT encoding modification for {map_name} (input is XML)")
        
        # Process Output encoding
        if is_xml_output:
            # For XML output, check row 4 first
            if xml_utf8_setting != "Keep as is" and xml_utf8_setting.lower() != "keep as is":
                if xml_utf8_setting.upper() == "UTF-8" or xml_utf8_setting.upper() == "UTF8":
                    output_encoding_value = "UTF-8"
                elif xml_utf8_setting.upper() == "NONE":
                    output_encoding_value = "EMPTY"  # Special marker to empty the tag content
                elif xml_utf8_setting == "Custom":
                    if xml_utf8_custom and xml_utf8_custom != "Enter the encoding":
                        output_encoding_value = xml_utf8_custom
                else:
                    # Use the value as-is (removing brackets if present)
                    clean_value = xml_utf8_setting.split('(')[0].strip()
                    # Double check it's not a variant of "Keep as is"
                    if clean_value.lower() != "keep as is":
                        output_encoding_value = clean_value
        else:
            # For non-XML output, process row 3
            if output_encoding_setting != "Keep as is" and output_encoding_setting.lower() != "keep as is":
                if output_encoding_setting.upper() == "NONE":
                    output_encoding_value = "EMPTY"  # Special marker to empty the tag content
                elif output_encoding_setting == "Custom":
                    if output_encoding_custom and output_encoding_custom != "Enter the details":
                        output_encoding_value = output_encoding_custom
                else:
                    # Remove anything in brackets
                    clean_value = output_encoding_setting.split('(')[0].strip()
                    # Double check it's not a variant of "Keep as is"
                    if clean_value.lower() != "keep as is":
                        output_encoding_value = clean_value
        
        # Apply changes using regex to preserve formatting
        # Handle both with and without namespace prefix (ns0:, ns1:, etc.)
        # Handle both <CharacterEncoding>value</CharacterEncoding> and <CharacterEncoding />
        if input_encoding_value is not None:
            if input_encoding_value == "EMPTY":
                # Empty the tag content (don't delete the tag itself)
                # Only process if tag exists
                # Try XMLSyntax (JavaEncoding) first
                xml_java_pattern = r'(<(?:\w+:)?JavaEncoding>)(.*?)(</(?:\w+:)?JavaEncoding>)'
                input_section_pattern = r'(<(?:\w+:)?INPUT>.*?)(</(?:\w+:)?INPUT>)'
                input_match = re.search(input_section_pattern, content, re.DOTALL)
                if input_match:
                    input_content = input_match.group(1)
                    java_match = re.search(xml_java_pattern, input_content, re.DOTALL)
                    if java_match:
                        # Empty the content between tags
                        new_input_content = input_content[:java_match.start(2)] + input_content[java_match.end(2):]
                        content = content[:input_match.start(1)] + new_input_content + content[input_match.start(2):]
                        modified = True
                        logger.info(f"Emptied Input JavaEncoding tag for {map_name}")
                    else:
                        # Try CharacterEncoding
                        char_enc_pattern = r'(<(?:\w+:)?CharacterEncoding>)(.*?)(</(?:\w+:)?CharacterEncoding>)'
                        char_match = re.search(char_enc_pattern, input_content, re.DOTALL)
                        if char_match:
                            # Empty the content between tags
                            new_input_content = input_content[:char_match.start(2)] + input_content[char_match.end(2):]
                            content = content[:input_match.start(1)] + new_input_content + content[input_match.start(2):]
                            modified = True
                            logger.info(f"Emptied Input CharacterEncoding tag for {map_name}")
            else:
                # Set the encoding value
                # Find INPUT section and replace CharacterEncoding/JavaEncoding within it
                # First check for XMLSyntax (uses JavaEncoding)
                xml_java_pattern = r'(<(?:\w+:)?INPUT>.*?<(?:\w+:)?XMLSyntax>.*?<(?:\w+:)?JavaEncoding>)(.*?)(</(?:\w+:)?JavaEncoding>.*?</(?:\w+:)?XMLSyntax>.*?</(?:\w+:)?INPUT>)'
                match = re.search(xml_java_pattern, content, re.DOTALL)
                if match:
                    content = content[:match.start(2)] + input_encoding_value + content[match.end(2):]
                    modified = True
                    logger.info(f"Set Input JavaEncoding to '{input_encoding_value}' for {map_name}")
                else:
                    # Try self-closing CharacterEncoding tag
                    input_pattern_self_closing = r'(<(?:\w+:)?INPUT>.*?<(?:\w+:)?CharacterEncoding\s*/>)(.*?</(?:\w+:)?INPUT>)'
                    match = re.search(input_pattern_self_closing, content, re.DOTALL)
                    if match:
                        # Replace self-closing tag with opening and closing tags
                        replacement = f'<CharacterEncoding>{input_encoding_value}</CharacterEncoding>'
                        content = content[:match.start(1)] + content[match.start(1):match.end(1)].replace('<CharacterEncoding />', replacement).replace('<CharacterEncoding/>', replacement) + content[match.end(1):]
                        modified = True
                        logger.info(f"Set Input CharacterEncoding to '{input_encoding_value}' for {map_name}")
                    else:
                        # Try regular opening/closing CharacterEncoding tags
                        input_pattern = r'(<(?:\w+:)?INPUT>.*?<(?:\w+:)?CharacterEncoding>)(.*?)(</(?:\w+:)?CharacterEncoding>.*?</(?:\w+:)?INPUT>)'
                        match = re.search(input_pattern, content, re.DOTALL)
                        if match:
                            content = content[:match.start(2)] + input_encoding_value + content[match.end(2):]
                            modified = True
                            logger.info(f"Set Input CharacterEncoding to '{input_encoding_value}' for {map_name}")
        
        if output_encoding_value is not None:
            if output_encoding_value == "EMPTY":
                # Empty the tag content (don't delete the tag itself)
                # Only process if tag exists
                # Try XMLSyntax (JavaEncoding) first
                xml_java_pattern = r'(<(?:\w+:)?JavaEncoding>)(.*?)(</(?:\w+:)?JavaEncoding>)'
                output_section_pattern = r'(<(?:\w+:)?OUTPUT>.*?)(</(?:\w+:)?OUTPUT>)'
                output_match = re.search(output_section_pattern, content, re.DOTALL)
                if output_match:
                    output_content = output_match.group(1)
                    java_match = re.search(xml_java_pattern, output_content, re.DOTALL)
                    if java_match:
                        # Empty the content between tags
                        new_output_content = output_content[:java_match.start(2)] + output_content[java_match.end(2):]
                        content = content[:output_match.start(1)] + new_output_content + content[output_match.start(2):]
                        modified = True
                        logger.info(f"Emptied Output JavaEncoding tag for {map_name}")
                    else:
                        # Try CharacterEncoding
                        char_enc_pattern = r'(<(?:\w+:)?CharacterEncoding>)(.*?)(</(?:\w+:)?CharacterEncoding>)'
                        char_match = re.search(char_enc_pattern, output_content, re.DOTALL)
                        if char_match:
                            # Empty the content between tags
                            new_output_content = output_content[:char_match.start(2)] + output_content[char_match.end(2):]
                            content = content[:output_match.start(1)] + new_output_content + content[output_match.start(2):]
                            modified = True
                            logger.info(f"Emptied Output CharacterEncoding tag for {map_name}")
            else:
                # Set the encoding value
                # Find OUTPUT section and replace CharacterEncoding/JavaEncoding within it
                # First check for XMLSyntax (uses JavaEncoding)
                xml_java_pattern = r'(<(?:\w+:)?OUTPUT>.*?<(?:\w+:)?XMLSyntax>.*?<(?:\w+:)?JavaEncoding>)(.*?)(</(?:\w+:)?JavaEncoding>.*?</(?:\w+:)?XMLSyntax>.*?</(?:\w+:)?OUTPUT>)'
                match = re.search(xml_java_pattern, content, re.DOTALL)
                if match:
                    content = content[:match.start(2)] + output_encoding_value + content[match.end(2):]
                    modified = True
                    logger.info(f"Set Output JavaEncoding to '{output_encoding_value}' for {map_name}")
                else:
                    # Try self-closing CharacterEncoding tag
                    output_pattern_self_closing = r'(<(?:\w+:)?OUTPUT>.*?<(?:\w+:)?CharacterEncoding\s*/>)(.*?</(?:\w+:)?OUTPUT>)'
                    match = re.search(output_pattern_self_closing, content, re.DOTALL)
                    if match:
                        # Replace self-closing tag with opening and closing tags
                        replacement = f'<CharacterEncoding>{output_encoding_value}</CharacterEncoding>'
                        content = content[:match.start(1)] + content[match.start(1):match.end(1)].replace('<CharacterEncoding />', replacement).replace('<CharacterEncoding/>', replacement) + content[match.end(1):]
                        modified = True
                        logger.info(f"Set Output CharacterEncoding to '{output_encoding_value}' for {map_name}")
                    else:
                        # Try regular opening/closing CharacterEncoding tags
                        output_pattern = r'(<(?:\w+:)?OUTPUT>.*?<(?:\w+:)?CharacterEncoding>)(.*?)(</(?:\w+:)?CharacterEncoding>.*?</(?:\w+:)?OUTPUT>)'
                        match = re.search(output_pattern, content, re.DOTALL)
                        if match:
                            content = content[:match.start(2)] + output_encoding_value + content[match.end(2):]
                            modified = True
                            logger.info(f"Set Output CharacterEncoding to '{output_encoding_value}' for {map_name}")
        
        # Save if modified
        if modified:
            # Remove ns0: prefix before saving to ensure clean format
            content = re.sub(r'<ns0:', '<', content)
            content = re.sub(r'</ns0:', '</', content)
            content = re.sub(r'\s+xmlns:ns0="[^"]*"', '', content)
            
            with open(mxl_file, 'w', encoding='utf-8') as f:
                f.write(content)
            logger.info(f"Successfully updated character encoding in {mxl_file}")
            return True
        else:
            logger.info(f"No encoding changes needed for {mxl_file}")
            return True
            
    except Exception as e:
        logger.error(f"Error modifying character encoding in {mxl_file}: {e}")
        return False


def main():
    """Main entry point."""
    logger.info("=" * 80)
    logger.info("CHARACTER ENCODING MODIFIER")
    logger.info("=" * 80)
    
    # Configuration
    mxl_dir = Path("./mxl_files")
    checklist_file = "Generic_checklistMain.xlsm"
    mapping_file = "mapping_results.xlsx"
    
    # Validate files exist
    if not Path(checklist_file).exists():
        logger.error(f"Checklist file not found: {checklist_file}")
        return 1
    
    if not Path(mapping_file).exists():
        logger.error(f"Mapping file not found: {mapping_file}")
        return 1
    
    if not mxl_dir.exists():
        logger.error(f"MXL directory not found: {mxl_dir}")
        return 1
    
    # Process all MXL files
    mxl_files = list(mxl_dir.glob("*.mxl"))
    if not mxl_files:
        logger.warning(f"No MXL files found in {mxl_dir}")
        return 0
    
    logger.info(f"Found {len(mxl_files)} MXL files")
    
    success_count = 0
    fail_count = 0
    
    for mxl_file in mxl_files:
        map_name = mxl_file.stem
        logger.info(f"\nProcessing: {map_name}")
        
        if modify_character_encoding(str(mxl_file), map_name, checklist_file, mapping_file):
            success_count += 1
        else:
            fail_count += 1
    
    # Summary
    logger.info("")
    logger.info("=" * 80)
    logger.info("PROCESSING COMPLETE")
    logger.info("=" * 80)
    logger.info(f"Total files: {len(mxl_files)}")
    logger.info(f"Successful: {success_count}")
    logger.info(f"Failed: {fail_count}")
    logger.info(f"Log file: {log_filename}")
    logger.info("=" * 80)
    
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())

# Made with Bob
