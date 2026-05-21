"""
MXL File Processor - Task 2: Update Format Tags

This module provides functionality to update <Format> tags in .mxl files
while preserving XML structure and namespaces.

Author: Sreenandha Ramesh
"""

import xml.etree.ElementTree as ET
from typing import List, Union, Dict, Any


def read_mxl(file_path: str) -> List[str]:
    """
    Read and process an .mxl file to extract field names with format 'X'.

    This function parses an XML-based .mxl file, locates all <Field> elements,
    and returns the names of fields where the <Format> tag (nested within
    <StoreLimit>) has the value 'X'.

    The function handles XML namespaces correctly and processes the file safely,
    returning an empty list if no matching fields are found or if errors occur.

    Args:
        file_path (str): Path to the .mxl file to be processed.

    Returns:
        List[str]: A list of field names (from <Name> tags) where the
                   corresponding <Format> tag equals 'X'. Returns an empty
                   list if no matches are found or if the file cannot be read.

    Examples:
        >>> field_names = read_mxl('data.mxl')
        >>> print(field_names)
        ['OrderType', 'ShipmentDate', 'CustomerID']

        >>> # File not found or invalid XML
        >>> field_names = read_mxl('nonexistent.mxl')
        >>> print(field_names)
        []

    Notes:
        - The function is read-only and does not modify the input file.
        - XML namespaces are handled automatically using namespace-aware parsing.
        - Invalid XML or missing files are handled gracefully without raising
          exceptions.
        - The <Format> tag is expected to be nested within <StoreLimit> inside
          each <Field> element.

    XML Structure Expected:
        ```xml
        <Field>
            <Name>FieldName</Name>
            <StoreLimit>
                <Format>X</Format>
            </StoreLimit>
        </Field>
        ```
    """
    field_names = []

    try:
        # Parse the XML file
        tree = ET.parse(file_path)
        root = tree.getroot()

        # Extract namespace from root element if present
        # Format: {namespace}tag or just tag if no namespace
        namespace = ''
        if root.tag.startswith('{'):
            namespace = root.tag[1:root.tag.index('}')]

        # Create namespace dictionary for XPath queries
        ns = {'ns': namespace} if namespace else {}

        # Find all Field elements using namespace-aware search
        if namespace:
            # Use namespace prefix in XPath
            fields = root.findall('.//ns:Field', ns)
        else:
            # No namespace, use simple XPath
            fields = root.findall('.//Field')

        # Process each Field element
        for field in fields:
            # Find the Format tag within StoreLimit
            if namespace:
                format_elem = field.find('.//ns:StoreLimit/ns:Format', ns)
                name_elem = field.find('ns:Name', ns)
            else:
                format_elem = field.find('.//StoreLimit/Format')
                name_elem = field.find('Name')

            # Check if Format exists and has value 'X'
            if format_elem is not None and format_elem.text == 'X':
                # Extract the Name value if it exists
                if name_elem is not None and name_elem.text:
                    field_names.append(name_elem.text)

    except FileNotFoundError:
        # File does not exist - return empty list
        pass
    except ET.ParseError:
        # Invalid XML - return empty list
        pass
    except Exception:
        # Any other error - return empty list
        pass

    return field_names


def read_mxl_by_section(file_path: str, section: str = 'BothSides') -> List[str]:
    """
    Read and process an .mxl file to extract field names with format 'X' from specific sections.
    
    Args:
        file_path (str): Path to the .mxl file to be processed.
        section (str): Section to process - 'InputOnly', 'OutputOnly', or 'BothSides' (default)
        
    Returns:
        List[str]: A list of field names where Format='X' in the specified section(s)
    """
    field_names = []
    
    try:
        # Parse the XML file
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Extract namespace
        namespace = ''
        if root.tag.startswith('{'):
            namespace = root.tag[1:root.tag.index('}')]
        
        ns = {'ns': namespace} if namespace else {}
        
        # Determine which sections to search
        sections_to_search = []
        if section.lower() == 'inputonly':
            sections_to_search = ['INPUT']
        elif section.lower() == 'outputonly':
            sections_to_search = ['OUTPUT']
        else:  # BothSides or any other value
            sections_to_search = ['INPUT', 'OUTPUT']
        
        # Search in specified sections
        for section_name in sections_to_search:
            if namespace:
                section_elem = root.find(f'.//ns:{section_name}', ns)
            else:
                section_elem = root.find(f'.//{section_name}')
            
            if section_elem is not None:
                # Find all Field elements within this section
                if namespace:
                    fields = section_elem.findall('.//ns:Field', ns)
                else:
                    fields = section_elem.findall('.//Field')
                
                # Process each Field element
                for field in fields:
                    # Find the Format tag within StoreLimit
                    if namespace:
                        format_elem = field.find('.//ns:StoreLimit/ns:Format', ns)
                        name_elem = field.find('ns:Name', ns)
                    else:
                        format_elem = field.find('.//StoreLimit/Format')
                        name_elem = field.find('Name')
                    
                    # Check if Format exists and has value 'X'
                    if format_elem is not None and format_elem.text == 'X':
                        # Extract the Name value if it exists
                        if name_elem is not None and name_elem.text:
                            field_names.append(name_elem.text)
    
    except FileNotFoundError:
        pass
    except ET.ParseError:
        pass
    except Exception:
        pass
    
    return field_names


def write_mxl(file_path: str, change: Union[List[str], None] = None, section: str = 'BothSides') -> None:
    """
    Update <Format> tags from <Format>X</Format> to <Format></Format>
    for fields in an .mxl (XML) file in specified sections.

    Args:
        file_path (str): Path to the .mxl file to update
        change (Union[List[str], None], optional):
            - None (default): Uses read_mxl_by_section() to find fields with Format='X'
            - List[str]: Updates only the specified field names
        section (str): Section to process - 'InputOnly', 'OutputOnly', or 'BothSides' (default)

    Returns:
        None
    """
    
    # Determine which fields to update
    if change is None:
        fields_to_update = read_mxl_by_section(file_path, section)
    elif isinstance(change, list):
        fields_to_update = change
    else:
        raise ValueError("change parameter must be a list of field names or None")

    if not fields_to_update:
        print(f"No fields with Format='X' found to update in section '{section}'")
        return
    
    from_value = 'X'
    to_value = ''

    # Parse the XML
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Detect namespace dynamically (if any)
    namespace = ''
    if root.tag.startswith('{'):
        namespace = root.tag[1:root.tag.index('}')]
    
    # Register the namespace to prevent ns0: prefixes in output
    if namespace:
        ET.register_namespace('', namespace)

    ns = {'ns': namespace} if namespace else {}

    updated_count = 0

    # Find Field elements safely (namespace-aware)
    if namespace:
        fields = root.findall('.//ns:Field', ns)
    else:
        fields = root.findall('.//Field')

    for field in fields:
        # Locate Name and Format safely
        if namespace:
            name_elem = field.find('ns:Name', ns)
            format_elem = field.find('.//ns:StoreLimit/ns:Format', ns)
        else:
            name_elem = field.find('Name')
            format_elem = field.find('.//StoreLimit/Format')

        # Validate field name and update logic
        if name_elem is not None and name_elem.text:
            # Check if this field is in the list to update
            if name_elem.text in fields_to_update:
                # Update only if Format matches from_value
                if format_elem is not None and format_elem.text == from_value:
                    format_elem.text = to_value
                    updated_count += 1
                    print(f"Updated field: {name_elem.text}")

    # Write back to the same file, preserving namespaces
    tree.write(
        file_path,
        encoding='UTF-8',
        xml_declaration=True,
        method='xml'
    )

    print(f"Successfully updated {updated_count} field(s) in {file_path} (section: {section})")

if __name__ == '__main__':
    # Example usage
    import sys

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        section = sys.argv[2] if len(sys.argv) > 2 else 'BothSides'
        
        results = read_mxl_by_section(file_path, section)

        if results:
            print(f"Found {len(results)} field(s) with Format='X' in section '{section}':")
            for i, field_name in enumerate(results, 1):
                print(f"  {i}. {field_name}")
            write_mxl(file_path, results, section)
        else:
            print(f"No fields with Format='X' found in section '{section}' or file could not be read.")
    else:
        print("Usage: python mxl_processor.py <path_to_mxl_file> [section]")
        print("  section: InputOnly, OutputOnly, or BothSides (default)")

# Made with Bob
