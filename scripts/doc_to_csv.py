#!/usr/bin/env python3
"""
DOC to CSV converter.
Converts Microsoft Word DOC files to CSV format by extracting tables.
"""

import argparse
import os
import sys
import subprocess
import traceback
import tempfile
import shutil
from pathlib import Path

def convert_doc_to_docx(doc_file, docx_file):
    """
    Convert DOC to DOCX using LibreOffice.
    
    Args:
        doc_file (str): Path to input DOC file
        docx_file (str): Path to output DOCX file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting DOC to DOCX with LibreOffice...", flush=True)
    
    try:
        # Create output directory
        output_dir = os.path.dirname(docx_file)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # LibreOffice command for headless conversion
        cmd = [
            'libreoffice',
            '--headless',
            '--invisible',
            '--nocrashreport',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
            '--convert-to', 'docx',
            '--outdir', output_dir,
            doc_file
        ]
        
        # Set LibreOffice environment
        env = os.environ.copy()
        env['SAL_USE_VCLPLUGIN'] = 'svp'
        env['HOME'] = '/tmp'
        
        print(f"Executing: {' '.join(cmd)}", flush=True)
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            env=env
        )
        
        if result.stdout:
            print(f"LibreOffice stdout: {result.stdout}", flush=True)
        if result.stderr:
            print(f"LibreOffice stderr: {result.stderr}", flush=True)
        
        # LibreOffice creates filename.docx, we need to rename it
        base_name = os.path.splitext(os.path.basename(doc_file))[0]
        actual_docx = os.path.join(output_dir, f"{base_name}.docx")
        
        if os.path.exists(actual_docx):
            if actual_docx != docx_file:
                shutil.move(actual_docx, docx_file)
            print(f"DOCX file created: {docx_file}", flush=True)
            return True
        else:
            print(f"ERROR: DOCX file not created: {actual_docx}", flush=True)
            # List directory contents for debugging
            print(f"Directory contents: {os.listdir(output_dir)}", flush=True)
            return False
            
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out", flush=True)
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion failed: {e}", flush=True)
        traceback.print_exc()
        return False

def extract_tables_from_docx(docx_file):
    """
    Extract tables from DOCX file using python-docx.
    
    Args:
        docx_file (str): Path to DOCX file
    
    Returns:
        list: List of tables (each table is a list of rows)
    """
    print(f"Extracting tables from DOCX...", flush=True)
    
    try:
        from docx import Document
        
        doc = Document(docx_file)
        tables = []
        
        for table_idx, table in enumerate(doc.tables):
            print(f"Processing table {table_idx + 1}...", flush=True)
            table_data = []
            
            for row_idx, row in enumerate(table.rows):
                row_data = []
                for cell in row.cells:
                    # Get text from cell, strip whitespace
                    cell_text = cell.text.strip()
                    row_data.append(cell_text)
                
                table_data.append(row_data)
            
            if table_data:
                tables.append(table_data)
                print(f"Table {table_idx + 1}: {len(table_data)} rows × {len(table_data[0]) if table_data else 0} columns", flush=True)
        
        print(f"Found {len(tables)} table(s)", flush=True)
        return tables
        
    except ImportError:
        print("ERROR: python-docx not installed. Install with: pip install python-docx", flush=True)
        return []
    except Exception as e:
        print(f"ERROR: Table extraction failed: {e}", flush=True)
        traceback.print_exc()
        return []

def escape_csv_field(field):
    """Escape CSV field according to RFC 4180."""
    field = str(field) if field is not None else ''
    
    # If field contains comma, quote, or newline, wrap in quotes
    if ',' in field or '"' in field or '\n' in field or '\r' in field:
        # Escape quotes by doubling them
        field = field.replace('"', '""')
        return f'"{field}"'
    
    return field

def convert_doc_to_csv(doc_file, csv_file, delimiter=',', include_headers=True, encoding='utf-8'):
    """
    Convert DOC to CSV format.
    
    Args:
        doc_file (str): Path to input DOC file
        csv_file (str): Path to output CSV file
        delimiter (str): CSV delimiter
        include_headers (bool): Include headers in output
        encoding (str): Output encoding
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting DOC to CSV...", flush=True)
    print(f"Input: {doc_file}", flush=True)
    print(f"Output: {csv_file}", flush=True)
    
    try:
        # Check if input file exists
        if not os.path.exists(doc_file):
            raise FileNotFoundError(f"DOC file not found: {doc_file}")
        
        file_size = os.path.getsize(doc_file)
        print(f"DOC file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        
        # Create temporary directory for intermediate files
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_file = os.path.join(tmpdir, 'document.docx')
            
            # Convert DOC to DOCX
            success = convert_doc_to_docx(doc_file, docx_file)
            if not success:
                raise Exception("Failed to convert DOC to DOCX with LibreOffice")
            
            # Extract tables from DOCX
            tables = extract_tables_from_docx(docx_file)
            
            if not tables:
                raise ValueError("No tables found in DOC file. The document may not contain tabular data.")
            
            # Use the first (or largest) table
            if len(tables) > 1:
                print(f"Multiple tables found. Using the largest table.", flush=True)
                table = max(tables, key=len)
            else:
                table = tables[0]
            
            print(f"Selected table has {len(table)} rows", flush=True)
            
            if len(table) == 0:
                raise ValueError("Table is empty")
            
            # Write CSV file
            print(f"Writing CSV file...", flush=True)
            with open(csv_file, 'w', encoding=encoding, newline='') as f:
                for row_idx, row in enumerate(table):
                    # Skip headers if not wanted
                    if not include_headers and row_idx == 0:
                        continue
                    
                    # Escape and join fields
                    escaped_fields = [escape_csv_field(cell) for cell in row]
                    csv_line = delimiter.join(escaped_fields)
                    f.write(csv_line + '\n')
                    
                    # Progress logging for large tables
                    if (row_idx + 1) % 1000 == 0:
                        print(f"Processed {row_idx + 1} of {len(table)} rows...", flush=True)
        
        # Verify output file
        if not os.path.exists(csv_file):
            raise FileNotFoundError(f"CSV file was not created: {csv_file}")
        
        output_size = os.path.getsize(csv_file)
        print(f"CSV file created successfully!", flush=True)
        print(f"CSV size: {output_size:,} bytes ({output_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Total rows exported: {len(table) if include_headers else len(table) - 1}", flush=True)
        
        return True
        
    except FileNotFoundError as e:
        print(f"ERROR: File not found: {e}", flush=True)
        return False
    except ValueError as e:
        print(f"ERROR: {e}", flush=True)
        return False
    except Exception as e:
        print(f"ERROR: Conversion failed: {e}", flush=True)
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert DOC to CSV')
    parser.add_argument('doc_file', help='Input DOC file path')
    parser.add_argument('csv_file', help='Output CSV file path')
    parser.add_argument('--delimiter', 
                       default=',',
                       help='CSV delimiter (default: ,)')
    parser.add_argument('--no-headers',
                       action='store_true',
                       help='Exclude headers from output')
    parser.add_argument('--encoding',
                       default='utf-8',
                       choices=['utf-8', 'ascii', 'utf-16'],
                       help='Output encoding (default: utf-8)')
    
    args = parser.parse_args()
    
    print("=== DOC to CSV Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.doc_file):
        print(f"ERROR: Input DOC file not found: {args.doc_file}", flush=True)
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.csv_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert delimiter escape sequences
    delimiter = args.delimiter
    if delimiter == '\\t':
        delimiter = '\t'
    
    # Convert
    success = convert_doc_to_csv(
        args.doc_file, 
        args.csv_file,
        delimiter,
        not args.no_headers,
        args.encoding
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()

