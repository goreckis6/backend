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

def convert_doc_to_html(doc_file, html_file):
    """
    Convert DOC to HTML using LibreOffice.
    
    Args:
        doc_file (str): Path to input DOC file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting DOC to HTML with LibreOffice...", flush=True)
    
    try:
        # Create output directory
        output_dir = os.path.dirname(html_file)
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
            '--convert-to', 'html',
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
        
        # LibreOffice creates filename.html, we need to rename it
        base_name = os.path.splitext(os.path.basename(doc_file))[0]
        actual_html = os.path.join(output_dir, f"{base_name}.html")
        
        if os.path.exists(actual_html):
            if actual_html != html_file:
                shutil.move(actual_html, html_file)
            print(f"HTML file created: {html_file}", flush=True)
            return True
        else:
            print(f"ERROR: HTML file not created: {actual_html}", flush=True)
            return False
            
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out", flush=True)
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion failed: {e}", flush=True)
        traceback.print_exc()
        return False

def extract_tables_from_html(html_file):
    """
    Extract tables from HTML file.
    
    Args:
        html_file (str): Path to HTML file
    
    Returns:
        list: List of tables (each table is a list of rows)
    """
    print(f"Extracting tables from HTML...", flush=True)
    
    try:
        from html.parser import HTMLParser
        
        class TableExtractor(HTMLParser):
            def __init__(self):
                super().__init__()
                self.tables = []
                self.current_table = []
                self.current_row = []
                self.current_cell = []
                self.in_table = False
                self.in_row = False
                self.in_cell = False
            
            def handle_starttag(self, tag, attrs):
                if tag == 'table':
                    self.in_table = True
                    self.current_table = []
                elif tag == 'tr' and self.in_table:
                    self.in_row = True
                    self.current_row = []
                elif tag in ['td', 'th'] and self.in_row:
                    self.in_cell = True
                    self.current_cell = []
            
            def handle_endtag(self, tag):
                if tag == 'table' and self.in_table:
                    if self.current_table:
                        self.tables.append(self.current_table)
                    self.in_table = False
                elif tag == 'tr' and self.in_row:
                    if self.current_row:
                        self.current_table.append(self.current_row)
                    self.in_row = False
                elif tag in ['td', 'th'] and self.in_cell:
                    cell_text = ' '.join(self.current_cell).strip()
                    self.current_row.append(cell_text)
                    self.in_cell = False
            
            def handle_data(self, data):
                if self.in_cell:
                    self.current_cell.append(data.strip())
        
        with open(html_file, 'r', encoding='utf-8', errors='ignore') as f:
            html_content = f.read()
        
        parser = TableExtractor()
        parser.feed(html_content)
        
        print(f"Found {len(parser.tables)} table(s)", flush=True)
        return parser.tables
        
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
            html_file = os.path.join(tmpdir, 'document.html')
            
            # Convert DOC to HTML
            success = convert_doc_to_html(doc_file, html_file)
            if not success:
                raise Exception("Failed to convert DOC to HTML")
            
            # Extract tables from HTML
            tables = extract_tables_from_html(html_file)
            
            if not tables:
                raise ValueError("No tables found in DOC file. The document may not contain tabular data.")
            
            # Use the first (or largest) table
            if len(tables) > 1:
                print(f"Multiple tables found. Using the largest table.", flush=True)
                table = max(tables, key=len)
            else:
                table = tables[0]
            
            print(f"Table has {len(table)} rows", flush=True)
            
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
        print(f"Total rows: {len(table)}", flush=True)
        
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

