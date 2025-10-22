#!/usr/bin/env python3
"""
CSV to EPUB Converter using Calibre
Converts CSV files to EPUB using Calibre for reliable EPUB generation.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback
import subprocess
import tempfile
import shutil

def create_epub_from_csv_calibre(csv_file, output_file, title="CSV Data", author="Unknown", chunk_size=1000):
    """
    Convert CSV file to EPUB using Calibre for reliable EPUB generation.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output EPUB file
        title (str): Book title
        author (str): Book author
        chunk_size (int): Number of rows to process at once
    """
    print("=" * 60)
    print("CSV TO EPUB CONVERSION (CALIBRE)")
    print("=" * 60)
    print(f"Starting CSV to EPUB conversion with Calibre...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print(f"Chunk size: {chunk_size}")
    print("=" * 60)
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file, dtype=str, na_filter=False)
        print(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        print(f"DataFrame shape: {df.shape}")
        print(f"DataFrame columns: {list(df.columns)}")
        
        # Create temporary directory for HTML files
        temp_dir = tempfile.mkdtemp(prefix='csv_epub_')
        print(f"Using temporary directory: {temp_dir}")
        
        try:
            # Create HTML file
            html_file = os.path.join(temp_dir, 'data.html')
            print(f"Creating HTML file: {html_file}")
            
            # Limit data for performance
            max_rows = min(5000, len(df))
            print(f"Processing {max_rows} rows for HTML generation...")
            
            # Create HTML content
            html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            line-height: 1.6;
            color: #333;
        }}
        h1 {{
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            margin-bottom: 30px;
        }}
        .info {{
            background: #ecf0f1;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
            border-left: 4px solid #3498db;
        }}
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 8px;
            text-align: left;
            font-weight: bold;
            font-size: 14px;
        }}
        td {{
            padding: 10px 8px;
            border-bottom: 1px solid #eee;
            font-size: 13px;
        }}
        tr:nth-child(even) {{
            background: #f8f9fa;
        }}
        tr:hover {{
            background: #e3f2fd;
        }}
        .cell-content {{
            max-width: 200px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }}
        .cell-content:hover {{
            white-space: normal;
            max-width: none;
            background: #fff;
            padding: 5px;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .note {{
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
            border-left: 4px solid #ffc107;
        }}
    </style>
</head>
<body>
    <h1>{title}</h1>
    
    <div class="info">
        <h3>Dataset Information</h3>
        <p><strong>Source File:</strong> {os.path.basename(csv_file)}</p>
        <p><strong>Total Records:</strong> {len(df):,}</p>
        <p><strong>Total Columns:</strong> {len(df.columns)}</p>
        <p><strong>Records Shown:</strong> {max_rows:,}</p>
        <p><strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>
    
    {f'<div class="note"><strong>Note:</strong> This table contains the first {max_rows} rows of data. For the complete dataset, please refer to the original CSV file.</div>' if len(df) > max_rows else ''}
    
    <table>
        <thead>
            <tr>
"""
            
            # Add column headers
            for col in df.columns:
                html_content += f"                <th>{col}</th>\n"
            html_content += "            </tr>\n        </thead>\n        <tbody>\n"
            
            # Add data rows
            for idx in range(max_rows):
                if idx % 500 == 0:
                    print(f"Processing row {idx + 1}/{max_rows}")
                
                row = df.iloc[idx]
                html_content += "            <tr>\n"
                for value in row:
                    # Escape HTML and limit length
                    cell_value = str(value) if value else ""
                    # Basic HTML escaping
                    cell_value = cell_value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                    html_content += f"                <td><div class=\"cell-content\" title=\"{cell_value}\">{cell_value}</div></td>\n"
                html_content += "            </tr>\n"
            
            html_content += """        </tbody>
    </table>
</body>
</html>"""
            
            # Write HTML file
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            print(f"HTML file created: {os.path.getsize(html_file)} bytes")
            
            # Use Calibre to convert HTML to EPUB
            print("Converting HTML to EPUB using Calibre...")
            
            # Calibre command
            calibre_cmd = [
                'ebook-convert',
                html_file,
                output_file,
                '--title', title,
                '--authors', author,
                '--language', 'en',
                '--publisher', 'MorphyIMG CSV Converter',
                '--pubdate', datetime.now().strftime('%Y-%m-%d'),
                '--epub-version', '3',
                '--pretty-print',
                '--enable-heuristics'
            ]
            
            print(f"Calibre command: {' '.join(calibre_cmd)}")
            
            # Run Calibre conversion
            try:
                result = subprocess.run(
                    calibre_cmd,
                    capture_output=True,
                    text=True,
                    timeout=300  # 5 minute timeout
                )
                
                print(f"Calibre return code: {result.returncode}")
                if result.stdout:
                    print(f"Calibre stdout: {result.stdout}")
                if result.stderr:
                    print(f"Calibre stderr: {result.stderr}")
                
                if result.returncode != 0:
                    raise Exception(f"Calibre conversion failed with return code {result.returncode}: {result.stderr}")
                
            except subprocess.TimeoutExpired:
                raise Exception("Calibre conversion timed out after 5 minutes")
            except FileNotFoundError:
                raise Exception("Calibre (ebook-convert) not found. Please install Calibre.")
            
            # Verify output file
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"EPUB file created successfully: {file_size / (1024*1024):.2f} MB ({file_size} bytes)")
                
                if file_size < 1024:  # Less than 1KB
                    print(f"WARNING: File is very small ({file_size} bytes)")
                    return False
                
                # Basic validation - check if it's a valid ZIP file
                try:
                    import zipfile
                    with zipfile.ZipFile(output_file, 'r') as zip_file:
                        file_list = zip_file.namelist()
                        print(f"EPUB contains {len(file_list)} files")
                        print(f"Key files: {[f for f in file_list if f.endswith(('.opf', '.xhtml', '.html', '.ncx'))]}")
                except Exception as e:
                    print(f"WARNING: Could not validate EPUB as ZIP: {e}")
                
                return True
            else:
                print("ERROR: EPUB file was not created by Calibre")
                return False
                
        finally:
            # Clean up temporary directory
            print(f"Cleaning up temporary directory: {temp_dir}")
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except Exception as e:
        print(f"ERROR: Failed to create EPUB from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to EPUB using Calibre')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output EPUB file path')
    parser.add_argument('--title', default='CSV Data', help='Book title')
    parser.add_argument('--author', default='Unknown', help='Book author')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Chunk size (ignored in Calibre version)')
    parser.add_argument('--no-toc', action='store_true', help='Do not include table of contents (ignored in Calibre version)')
    
    args = parser.parse_args()
    
    print("=== CSV to EPUB Converter (Calibre) ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)
    
    # Check if Calibre is available
    try:
        result = subprocess.run(['ebook-convert', '--version'], capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print(f"Calibre version: {result.stdout.strip()}")
        else:
            print("WARNING: Could not get Calibre version")
    except Exception as e:
        print(f"WARNING: Could not check Calibre availability: {e}")
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to EPUB
    success = create_epub_from_csv_calibre(
        args.csv_file,
        args.output_file,
        args.title,
        args.author,
        args.chunk_size
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()
