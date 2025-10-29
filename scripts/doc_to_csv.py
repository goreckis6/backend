#!/usr/bin/env python3
"""
DOC to CSV Converter
Converts Microsoft Word DOC files to CSV format
Uses LibreOffice for best table-based DOC conversion
"""

import os
import sys
import argparse
import traceback
import subprocess
import tempfile
import shutil
import csv
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


def find_libreoffice():
    """Find LibreOffice binary"""
    libreoffice_paths = [
        'libreoffice',
        '/usr/bin/libreoffice',
        '/usr/local/bin/libreoffice',
        '/opt/libreoffice/program/soffice',
        'soffice'
    ]
    
    for path in libreoffice_paths:
        try:
            result = subprocess.run(
                [path, '--version'],
                capture_output=True,
                check=True,
                timeout=5
            )
            print(f"Found LibreOffice at: {path}")
            return path
        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
            continue
    
    return None


def convert_doc_to_csv_libreoffice(doc_file, output_file, delimiter=',', extract_tables=True, include_paragraphs=True):
    """
    Convert DOC file to CSV format using LibreOffice
    Strategy: DOC -> CSV via LibreOffice (best for table-based DOC files)
    
    Args:
        doc_file (str): Path to input DOC file
        output_file (str): Path to output CSV file
        delimiter (str): CSV delimiter (comma, semicolon, tab, pipe)
        extract_tables (bool): Extract tables from DOC
        include_paragraphs (bool): Include paragraphs as data rows
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOC to CSV conversion using LibreOffice...")
    print(f"Input: {doc_file}")
    print(f"Output: {output_file}")
    print(f"Delimiter: {delimiter}")
    print(f"Extract tables: {extract_tables}")
    print(f"Include paragraphs: {include_paragraphs}")
    
    try:
        # Check if DOC file exists
        if not os.path.exists(doc_file):
            print(f"ERROR: DOC file does not exist: {doc_file}")
            return False
        
        file_size = os.path.getsize(doc_file)
        print(f"DOC file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: Input file is empty")
            return False
        
        # Find LibreOffice
        libreoffice = find_libreoffice()
        
        if not libreoffice:
            print("ERROR: LibreOffice not found. Please ensure LibreOffice is installed.")
            return False
        
        # Create temporary directory for LibreOffice conversion
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Create output directory if needed
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # LibreOffice can convert DOC to CSV
            # Strategy: Convert DOC to CSV using LibreOffice's filter
            
            # Map delimiter to LibreOffice filter option
            filter_map = {
                ',': 'Text - txt - csv (StarCalc)',
                ';': 'Text - txt - csv (StarCalc)',
                '\t': 'Text - txt - csv (StarCalc)',
                '|': 'Text - txt - csv (StarCalc)'
            }
            
            # Use LibreOffice to convert DOC to CSV
            # LibreOffice --convert-to csv:Text - txt - csv (StarCalc) --outdir temp_dir doc_file
            cmd = [
                libreoffice,
                '--headless',
                '--invisible',
                '--nocrashreport',
                '--nodefault',
                '--nofirststartwizard',
                '--nolockcheck',
                '--nologo',
                '--norestore',
                '--convert-to', 'csv',
                '--outdir', temp_dir,
                doc_file
            ]
            
            # Set LibreOffice environment
            env = os.environ.copy()
            env['SAL_USE_VCLPLUGIN'] = 'svp'
            env['HOME'] = '/tmp'
            env['LANG'] = 'en_US.UTF-8'
            env['LC_ALL'] = 'en_US.UTF-8'
            
            print(f"Running LibreOffice: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300,  # 5 minute timeout
                env=env,
                encoding='utf-8',
                errors='replace'
            )
            
            if result.stdout:
                print(f"LibreOffice stdout: {result.stdout}")
            if result.stderr:
                print(f"LibreOffice stderr: {result.stderr}")
            
            # LibreOffice creates filename.csv in temp_dir
            base_name = os.path.splitext(os.path.basename(doc_file))[0]
            libreoffice_csv = os.path.join(temp_dir, f"{base_name}.csv")
            
            if os.path.exists(libreoffice_csv):
                print(f"LibreOffice created CSV: {libreoffice_csv}")
                
                # Read the CSV created by LibreOffice and adjust delimiter if needed
                if HAS_PANDAS:
                    # Use pandas to handle delimiter conversion
                    try:
                        # Try to read with different delimiters to find the right one
                        df = None
                        for delim in [',', ';', '\t']:
                            try:
                                df = pd.read_csv(libreoffice_csv, delimiter=delim, encoding='utf-8', header=None)
                                print(f"Successfully read CSV with delimiter '{delim}'")
                                break
                            except Exception:
                                continue
                        
                        if df is None:
                            # Try reading with comma as default
                            df = pd.read_csv(libreoffice_csv, encoding='utf-8', header=None, on_bad_lines='skip')
                        
                        # Adjust delimiter if needed
                        if delimiter != ',':
                            # Convert to desired delimiter
                            if delimiter == ';':
                                df.to_csv(output_file, sep=';', index=False, header=False, encoding='utf-8')
                            elif delimiter == '\t':
                                df.to_csv(output_file, sep='\t', index=False, header=False, encoding='utf-8')
                            elif delimiter == '|':
                                df.to_csv(output_file, sep='|', index=False, header=False, encoding='utf-8')
                            else:
                                df.to_csv(output_file, sep=',', index=False, header=False, encoding='utf-8')
                        else:
                            # Just copy/rename if delimiter matches
                            shutil.copy2(libreoffice_csv, output_file)
                    except Exception as e:
                        print(f"Warning: Could not process CSV with pandas: {e}")
                        # Fallback: just copy the file
                        shutil.copy2(libreoffice_csv, output_file)
                else:
                    # No pandas available, just copy the file
                    shutil.copy2(libreoffice_csv, output_file)
                
                # Verify output file
                if os.path.exists(output_file):
                    output_size = os.path.getsize(output_file)
                    print(f"CSV file created successfully: {output_size} bytes")
                    return True
                else:
                    print(f"ERROR: CSV file was not created at {output_file}")
                    return False
            else:
                print(f"ERROR: LibreOffice did not create CSV file: {libreoffice_csv}")
                print(f"Contents of temp directory {temp_dir}: {os.listdir(temp_dir)}")
                return False
                
        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"Warning: Could not clean up temp directory: {e}")
            
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out after 5 minutes")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DOC to CSV: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert DOC file to CSV format using LibreOffice')
    parser.add_argument('doc_file', help='Path to input DOC file')
    parser.add_argument('output_file', help='Path to output CSV file')
    parser.add_argument('--delimiter', default=',', choices=[',', ';', '\t', '|'],
                        help='CSV delimiter (default: comma)')
    parser.add_argument('--no-tables', action='store_true',
                        help='Do not extract tables (LibreOffice extracts tables by default)')
    parser.add_argument('--no-paragraphs', action='store_true',
                        help='Do not include paragraphs (not applicable for LibreOffice CSV export)')
    
    args = parser.parse_args()
    
    print("=== DOC to CSV Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_doc_to_csv_libreoffice(
        args.doc_file,
        args.output_file,
        delimiter=args.delimiter,
        extract_tables=not args.no_tables,
        include_paragraphs=not args.no_paragraphs
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()
