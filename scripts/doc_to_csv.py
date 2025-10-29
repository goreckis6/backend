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
        
        # LibreOffice command for headless conversion with UTF-8 encoding
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
        
        # Set LibreOffice environment with UTF-8 locale
        env = os.environ.copy()
        env['SAL_USE_VCLPLUGIN'] = 'svp'
        env['HOME'] = '/tmp'
        env['LANG'] = 'en_US.UTF-8'
        env['LC_ALL'] = 'en_US.UTF-8'
        
        print(f"Executing: {' '.join(cmd)}", flush=True)
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            env=env,
            encoding='utf-8',
            errors='replace'
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

def extract_tables_from_html(html_file):
    """
    Extract tables from HTML file using pandas.
    
    Args:
        html_file (str): Path to HTML file
    
    Returns:
        list: List of DataFrames (tables)
    """
    print(f"Extracting tables from HTML with pandas...", flush=True)
    
    try:
        import pandas as pd
        from io import StringIO
        
        # Try different parsers in order of preference
        parsers = ['lxml', 'html5lib', 'html.parser']
        tables = None
        
        for parser in parsers:
            try:
                print(f"Trying parser: {parser}", flush=True)
                # Read HTML file content
                with open(html_file, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                # Wrap HTML content in StringIO to avoid deprecation warning
                html_io = StringIO(html_content)
                
                # Read all tables from HTML
                tables = pd.read_html(html_io, flavor=parser)
                print(f"Successfully parsed with {parser}", flush=True)
                break
            except (ImportError, ValueError) as e:
                print(f"Parser {parser} failed: {e}", flush=True)
                continue
        
        if tables is None or len(tables) == 0:
            raise ValueError("No tables found in HTML")
        
        print(f"Found {len(tables)} table(s)", flush=True)
        
        for idx, table in enumerate(tables):
            print(f"Table {idx + 1}: {len(table)} rows × {len(table.columns)} columns", flush=True)
        
        return tables
        
    except ImportError as e:
        print(f"ERROR: Required library not installed: {e}", flush=True)
        print("Make sure pandas and lxml are installed: pip install pandas lxml", flush=True)
        return []
    except ValueError as e:
        print(f"ERROR: No tables found in HTML: {e}", flush=True)
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
    Strategy: DOC → HTML (LibreOffice) → Extract tables (pandas) → CSV
    
    Args:
        doc_file (str): Path to input DOC file
        csv_file (str): Path to output CSV file
        delimiter (str): CSV delimiter
        include_headers (bool): Include headers in output
        encoding (str): Output encoding
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"=== Converting DOC to CSV ===", flush=True)
    print(f"Input: {doc_file}", flush=True)
    print(f"Output: {csv_file}", flush=True)
    
    try:
        import pandas as pd
        import csv
        
        # Check if input file exists
        if not os.path.exists(doc_file):
            raise FileNotFoundError(f"DOC file not found: {doc_file}")
        
        file_size = os.path.getsize(doc_file)
        print(f"DOC file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        
        # Create temporary directory for intermediate files
        with tempfile.TemporaryDirectory() as tmpdir:
            html_file = os.path.join(tmpdir, 'document.html')
            
            # Step 1: Convert DOC to HTML using LibreOffice
            print("Step 1/3: Converting DOC to HTML...", flush=True)
            success = convert_doc_to_html(doc_file, html_file)
            if not success:
                raise Exception("Failed to convert DOC to HTML with LibreOffice")
            
            # Step 2: Extract tables from HTML using pandas
            print("Step 2/3: Extracting tables from HTML...", flush=True)
            tables = extract_tables_from_html(html_file)
            
            if not tables or len(tables) == 0:
                # If no tables found, read the entire HTML as plain text and convert to CSV
                print("WARNING: No tables found. Converting document text to CSV format...", flush=True)
                import pandas as pd
                from bs4 import BeautifulSoup
                
                with open(html_file, 'r', encoding='utf-8', errors='replace') as f:
                    soup = BeautifulSoup(f, 'html.parser', from_encoding='utf-8')
                    # Get all text, split by lines
                    text = soup.get_text()
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                
                if not lines:
                    raise ValueError("Document appears to be empty or contains no text.")
                
                # Create a DataFrame with single column
                table_df = pd.DataFrame({'Content': lines})
                print(f"Created single-column CSV with {len(table_df)} lines", flush=True)
            else:
                # Use the first (or largest) table
                if len(tables) > 1:
                    print(f"Multiple tables found. Using the largest table.", flush=True)
                    table_df = max(tables, key=len)
                else:
                    table_df = tables[0]
                
                print(f"Selected table: {len(table_df)} rows × {len(table_df.columns)} columns", flush=True)
                
                if len(table_df) == 0:
                    raise ValueError("Table is empty")
            
            # Step 3: Write CSV file using pandas
            print("Step 3/3: Writing CSV file...", flush=True)
            
            # Determine quoting style for CSV
            quoting = csv.QUOTE_MINIMAL
            
            # Use UTF-8 with BOM for better compatibility with Excel and other tools
            if encoding.lower() == 'utf-8':
                actual_encoding = 'utf-8-sig'  # UTF-8 with BOM
            else:
                actual_encoding = encoding
            
            # Write to CSV
            table_df.to_csv(
                csv_file,
                sep=delimiter,
                encoding=actual_encoding,
                index=False,
                header=include_headers,
                quoting=quoting
            )
        
        # Verify output file
        if not os.path.exists(csv_file):
            raise FileNotFoundError(f"CSV file was not created: {csv_file}")
        
        output_size = os.path.getsize(csv_file)
        print(f"=== CONVERSION SUCCESSFUL ===", flush=True)
        print(f"CSV file created: {csv_file}", flush=True)
        print(f"CSV size: {output_size:,} bytes ({output_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Total rows exported: {len(table_df)}", flush=True)
        
        return True
        
    except ImportError as e:
        print(f"ERROR: Required library not available: {e}", flush=True)
        print("Make sure pandas is installed: pip install pandas", flush=True)
        return False
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

