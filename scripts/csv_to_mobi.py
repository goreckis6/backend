#!/usr/bin/env python3
"""
CSV to MOBI Converter
Converts CSV files to MOBI format using pandas and ebooklib
"""

import sys
import os
import pandas as pd
import argparse
from pathlib import Path
import tempfile
import subprocess
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def convert_csv_to_mobi(csv_path, output_path, book_title=None, author=None, include_headers=True, chunk_size=1000):
    """
    Convert CSV file to MOBI format
    
    Args:
        csv_path (str): Path to input CSV file
        output_path (str): Path to output MOBI file
        book_title (str): Title for the MOBI book
        author (str): Author for the MOBI book
        include_headers (bool): Whether to include CSV headers
        chunk_size (int): Number of rows to process at a time
    """
    try:
        logger.info(f"Starting CSV to MOBI conversion: {csv_path} -> {output_path}")
        
        # Read CSV file
        logger.info("Reading CSV file...")
        df = pd.read_csv(csv_path)
        
        if df.empty:
            raise ValueError("CSV file is empty")
        
        logger.info(f"CSV file contains {len(df)} rows and {len(df.columns)} columns")
        
        # Set default values
        if not book_title:
            book_title = Path(csv_path).stem.replace('_', ' ').title()
        if not author:
            author = "CSV Converter"
        
        # Create temporary HTML file for intermediate conversion
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as html_file:
            html_path = html_file.name
            
            # Write HTML content
            html_file.write(f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{book_title}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            margin: 20px;
            color: #333;
        }}
        h1 {{
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #34495e;
            margin-top: 30px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            font-size: 14px;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        th {{
            background-color: #f2f2f2;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
        }}
        .author {{
            color: #7f8c8d;
            font-style: italic;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>{book_title}</h1>
        <p class="author">by {author}</p>
    </div>
    
    <h2>Data Table</h2>
    <table>
""")
            
            # Add table headers if requested
            if include_headers and not df.columns.empty:
                html_file.write("        <thead>\n            <tr>\n")
                for col in df.columns:
                    html_file.write(f"                <th>{col}</th>\n")
                html_file.write("            </tr>\n        </thead>\n")
            
            # Add table body
            html_file.write("        <tbody>\n")
            
            # Process data in chunks for large files
            total_rows = len(df)
            processed_rows = 0
            
            for start_idx in range(0, total_rows, chunk_size):
                end_idx = min(start_idx + chunk_size, total_rows)
                chunk_df = df.iloc[start_idx:end_idx]
                
                for _, row in chunk_df.iterrows():
                    html_file.write("            <tr>\n")
                    for value in row:
                        # Handle NaN values and escape HTML
                        if pd.isna(value):
                            cell_value = ""
                        else:
                            cell_value = str(value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        html_file.write(f"                <td>{cell_value}</td>\n")
                    html_file.write("            </tr>\n")
                
                processed_rows += len(chunk_df)
                logger.info(f"Processed {processed_rows}/{total_rows} rows")
            
            html_file.write("""        </tbody>
    </table>
</body>
</html>""")
        
        logger.info(f"Created HTML file: {html_path}")
        
        # Convert HTML to MOBI using ebook-convert (Calibre)
        logger.info("Converting HTML to MOBI using ebook-convert...")
        
        # Try to find ebook-convert
        ebook_convert_paths = [
            'ebook-convert',
            '/usr/bin/ebook-convert',
            '/usr/local/bin/ebook-convert',
            '/opt/calibre/bin/ebook-convert'
        ]
        
        ebook_convert = None
        for path in ebook_convert_paths:
            try:
                subprocess.run([path, '--version'], capture_output=True, check=True)
                ebook_convert = path
                break
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
        
        if not ebook_convert:
            # Fallback: try to install calibre or use alternative method
            logger.warning("ebook-convert not found. Attempting to install calibre...")
            try:
                subprocess.run(['apt-get', 'update'], check=True)
                subprocess.run(['apt-get', 'install', '-y', 'calibre'], check=True)
                ebook_convert = 'ebook-convert'
            except subprocess.CalledProcessError:
                logger.error("Failed to install calibre. Cannot convert to MOBI format.")
                raise RuntimeError("ebook-convert (Calibre) is required for MOBI conversion but not available")
        
        # Convert HTML to MOBI
        convert_cmd = [
            ebook_convert,
            html_path,
            output_path,
            '--title', book_title,
            '--authors', author,
            '--language', 'en',
            '--mobi-file-type', 'old',
            '--disable-font-rescaling'
        ]
        
        logger.info(f"Running command: {' '.join(convert_cmd)}")
        result = subprocess.run(convert_cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode != 0:
            logger.error(f"ebook-convert failed with return code {result.returncode}")
            logger.error(f"stderr: {result.stderr}")
            raise RuntimeError(f"ebook-convert failed: {result.stderr}")
        
        logger.info("MOBI conversion completed successfully")
        
        # Clean up temporary HTML file
        try:
            os.unlink(html_path)
            logger.info("Cleaned up temporary HTML file")
        except OSError:
            logger.warning("Failed to clean up temporary HTML file")
        
        # Verify output file was created
        if not os.path.exists(output_path):
            raise RuntimeError("MOBI file was not created")
        
        file_size = os.path.getsize(output_path)
        logger.info(f"MOBI file created successfully: {output_path} ({file_size} bytes)")
        
        return True
        
    except Exception as e:
        logger.error(f"Error converting CSV to MOBI: {str(e)}")
        raise

def main():
    parser = argparse.ArgumentParser(description='Convert CSV file to MOBI format')
    parser.add_argument('input_csv', help='Input CSV file path')
    parser.add_argument('output_mobi', help='Output MOBI file path')
    parser.add_argument('--title', help='Book title (default: filename)')
    parser.add_argument('--author', help='Book author (default: "CSV Converter")')
    parser.add_argument('--no-headers', action='store_true', help='Do not include CSV headers')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Number of rows to process at a time')
    
    args = parser.parse_args()
    
    try:
        success = convert_csv_to_mobi(
            csv_path=args.input_csv,
            output_path=args.output_mobi,
            book_title=args.title,
            author=args.author,
            include_headers=not args.no_headers,
            chunk_size=args.chunk_size
        )
        
        if success:
            print(f"Successfully converted {args.input_csv} to {args.output_mobi}")
            sys.exit(0)
        else:
            print("Conversion failed")
            sys.exit(1)
            
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
