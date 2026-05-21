#!/usr/bin/env python3
"""
Remove namespace prefixes (ns0:, ns1:, etc.) from MXL files.
This script fixes files that have been corrupted with namespace prefixes.
"""

import re
import sys
from pathlib import Path


def remove_namespace_prefixes(file_path: Path) -> bool:
    """
    Remove namespace prefixes from an MXL file.
    
    Args:
        file_path: Path to the MXL file
        
    Returns:
        True if file was modified, False otherwise
    """
    try:
        # Read file content
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        
        # Remove namespace prefixes from opening tags (e.g., <ns0:Tag> -> <Tag>)
        content = re.sub(r'<ns\d+:', '<', content)
        
        # Remove namespace prefixes from closing tags (e.g., </ns0:Tag> -> </Tag>)
        content = re.sub(r'</ns\d+:', '</', content)
        
        # Remove xmlns:ns declarations (e.g., xmlns:ns0="...")
        content = re.sub(r'\s+xmlns:ns\d+="[^"]*"', '', content)
        
        # Check if any changes were made
        if content == original_content:
            print(f"✓ No namespace prefixes found in: {file_path.name}")
            return False
        
        # Write back the cleaned content
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"✓ Removed namespace prefixes from: {file_path.name}")
        return True
        
    except Exception as e:
        print(f"✗ Error processing {file_path.name}: {e}")
        return False


def main():
    """Main function to process MXL files."""
    if len(sys.argv) < 2:
        print("Usage: python remove_namespace_prefixes.py <mxl_file_or_directory>")
        print("\nExamples:")
        print("  python remove_namespace_prefixes.py mxl_files/")
        print("  python remove_namespace_prefixes.py mxl_files/CBI_XML_AMAZON_O_856_4010.mxl")
        sys.exit(1)
    
    input_path = Path(sys.argv[1])
    
    if not input_path.exists():
        print(f"Error: Path does not exist: {input_path}")
        sys.exit(1)
    
    # Collect MXL files
    if input_path.is_file():
        if input_path.suffix.lower() == '.mxl':
            mxl_files = [input_path]
        else:
            print(f"Error: File is not an MXL file: {input_path}")
            sys.exit(1)
    else:
        mxl_files = list(input_path.glob('*.mxl'))
        if not mxl_files:
            print(f"No MXL files found in: {input_path}")
            sys.exit(1)
    
    print(f"\n{'='*80}")
    print(f"REMOVING NAMESPACE PREFIXES FROM {len(mxl_files)} FILE(S)")
    print(f"{'='*80}\n")
    
    modified_count = 0
    for mxl_file in mxl_files:
        if remove_namespace_prefixes(mxl_file):
            modified_count += 1
    
    print(f"\n{'='*80}")
    print(f"COMPLETE: Modified {modified_count} of {len(mxl_files)} file(s)")
    print(f"{'='*80}\n")


if __name__ == '__main__':
    main()

# Made with Bob
