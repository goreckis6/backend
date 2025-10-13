#!/usr/bin/env python3
"""
CSV to XML converter.
Converts CSV files to XML (Extensible Markup Language) format.
"""

import argparse
import os
import sys
import pandas as pd
import traceback
import xml.etree.ElementTree as ET
from xml.dom import minidom

def sanitize_tag_name(name):
    """Sanitize tag name for XML compatibility."""
    # XML tags must start with letter or underscore
    # Can contain letters, digits, hyphens, underscores, and periods
    sanitized = ''.join(c if c.isalnum() or c in '-_.' else '_' for c in str(name))
    
    # Remove leading/trailing special chars
    sanitized = sanitized.strip('-_.')
    
    # Ensure it starts with letter or underscore
    if sanitized and not (sanitized[0].isalpha() or sanitized[0] == '_'):
        sanitized = 'tag_' + sanitized
    
    return sanitized or 'tag'

def escape_xml_text(text):
    """Escape special characters for XML text content."""
    if pd.isna(text):
        return ''
    
    s = str(text)
    # XML requires escaping of: < > & " '
    s = s.replace('&', '&amp;')
    s = s.replace('<', '&lt;')
    s = s.replace('>', '&gt;')
    s = s.replace('"', '&quot;')
    s = s.replace("'", '&apos;')
    
    return s

def convert_csv_to_xml(csv_file, xml_file, root_element='data', row_element='row', 
                       pretty_print=True, include_header=True):
    """
    Convert CSV to XML format.
    
    Args:
        csv_file (str): Path to input CSV file
        xml_file (str): Path to output XML file
        root_element (str): Name of the root XML element
        row_element (str): Name of each row element
        pretty_print (bool): Format with indentation
        include_header (bool): Include CSV headers as attributes/elements
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting CSV to XML...", flush=True)
    print(f"Input: {csv_file}", flush=True)
    print(f"Output: {xml_file}", flush=True)
    print(f"Root element: {root_element}", flush=True)
    print(f"Row element: {row_element}", flush=True)
    
    try:
        # Read CSV file
        print("Reading CSV file...", flush=True)
        df = pd.read_csv(csv_file, encoding='utf-8', low_memory=False)
        
        rows, cols = df.shape
        print(f"CSV loaded: {rows} rows × {cols} columns", flush=True)
        
        if rows == 0:
            raise ValueError("CSV file is empty (no data rows)")
        
        # Sanitize column names for XML tags
        original_columns = df.columns.tolist()
        sanitized_columns = [sanitize_tag_name(col) for col in original_columns]
        df.columns = sanitized_columns
        
        print(f"Columns: {', '.join(sanitized_columns[:10])}", flush=True)
        if len(sanitized_columns) > 10:
            print(f"... and {len(sanitized_columns) - 10} more columns", flush=True)
        
        # Create root element
        root = ET.Element(sanitize_tag_name(root_element))
        
        # Add metadata as comment
        comment_text = f" Generated from CSV file: {os.path.basename(csv_file)} | Rows: {rows}, Columns: {cols} "
        root.append(ET.Comment(comment_text))
        
        # Process each row
        for index, row_data in df.iterrows():
            # Create row element
            row_elem = ET.SubElement(root, sanitize_tag_name(row_element))
            
            # Add columns as child elements
            for col in sanitized_columns:
                value = row_data[col]
                col_elem = ET.SubElement(row_elem, col)
                col_elem.text = escape_xml_text(value)
            
            # Progress logging for large files
            if (index + 1) % 1000 == 0:
                print(f"Processed {index + 1} of {rows} rows...", flush=True)
        
        # Convert to string
        xml_string = ET.tostring(root, encoding='utf-8', method='xml')
        
        # Pretty print if requested
        if pretty_print:
            print("Formatting XML with indentation...", flush=True)
            dom = minidom.parseString(xml_string)
            xml_string = dom.toprettyxml(indent="  ", encoding='utf-8')
        else:
            # Add XML declaration
            xml_string = b'<?xml version="1.0" encoding="UTF-8"?>\n' + xml_string
        
        # Write to XML file
        print(f"Writing XML file...", flush=True)
        with open(xml_file, 'wb') as f:
            f.write(xml_string)
        
        # Verify output file
        if not os.path.exists(xml_file):
            raise FileNotFoundError(f"XML file was not created: {xml_file}")
        
        file_size = os.path.getsize(xml_file)
        csv_size = os.path.getsize(csv_file)
        
        print(f"XML file created successfully!", flush=True)
        print(f"CSV size: {csv_size:,} bytes ({csv_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"XML size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Total rows: {rows}", flush=True)
        
        return True
        
    except FileNotFoundError as e:
        print(f"ERROR: File not found: {e}", flush=True)
        return False
    except pd.errors.EmptyDataError:
        print("ERROR: CSV file is empty", flush=True)
        return False
    except pd.errors.ParserError as e:
        print(f"ERROR: CSV parsing error: {e}", flush=True)
        return False
    except Exception as e:
        print(f"ERROR: Conversion failed: {e}", flush=True)
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to XML')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('xml_file', help='Output XML file path')
    parser.add_argument('--root-element', 
                       default='data',
                       help='Root XML element name (default: data)')
    parser.add_argument('--row-element', 
                       default='row',
                       help='Row XML element name (default: row)')
    parser.add_argument('--no-pretty-print',
                       action='store_true',
                       help='Disable pretty printing (compact output)')
    
    args = parser.parse_args()
    
    print("=== CSV to XML Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    print(f"Pandas version: {pd.__version__}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}", flush=True)
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.xml_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert
    success = convert_csv_to_xml(
        args.csv_file, 
        args.xml_file, 
        args.root_element,
        args.row_element,
        not args.no_pretty_print
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()

