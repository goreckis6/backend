#!/usr/bin/env python3
"""
Simple CSV to EPUB Converter
Creates a minimal but valid EPUB file from CSV data.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback
import uuid

try:
    from ebooklib import epub
except ImportError as e:
    print(f"ERROR: Required EPUB library not available: {e}")
    print("Please install ebooklib: pip install ebooklib")
    sys.exit(1)

def create_simple_epub_from_csv(csv_file, output_file, title="CSV Data", author="Unknown"):
    """
    Create a simple but valid EPUB from CSV data.
    """
    print("=" * 50)
    print("SIMPLE CSV TO EPUB CONVERSION")
    print("=" * 50)
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print("=" * 50)
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file, dtype=str, na_filter=False)
        print(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        
        # Create EPUB book
        print("Creating EPUB book...")
        book = epub.EpubBook()
        
        # Set basic metadata
        book.set_identifier(str(uuid.uuid4()))
        book.set_title(title)
        book.set_language('en')
        book.add_author(author)
        
        # Create a simple HTML content
        print("Creating HTML content...")
        
        # Limit data to prevent huge files
        max_rows = min(1000, len(df))
        print(f"Processing {max_rows} rows...")
        
        # Create simple HTML table
        html_content = f"""
        <html>
        <head>
            <title>{title}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h1 {{ color: #333; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
            </style>
        </head>
        <body>
            <h1>{title}</h1>
            <p>Generated from CSV file: {os.path.basename(csv_file)}</p>
            <p>Total records: {len(df):,}</p>
            <p>Records shown: {max_rows:,}</p>
            <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            
            <table>
                <thead>
                    <tr>
        """
        
        # Add column headers
        for col in df.columns:
            html_content += f"<th>{col}</th>"
        html_content += "</tr></thead><tbody>"
        
        # Add data rows
        for idx in range(max_rows):
            if idx % 100 == 0:
                print(f"Processing row {idx + 1}/{max_rows}")
            
            row = df.iloc[idx]
            html_content += "<tr>"
            for value in row:
                # Escape HTML and limit length
                cell_value = str(value) if value else ""
                if len(cell_value) > 100:
                    cell_value = cell_value[:97] + "..."
                # Basic HTML escaping
                cell_value = cell_value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                html_content += f"<td>{cell_value}</td>"
            html_content += "</tr>"
        
        html_content += """
                </tbody>
            </table>
        </body>
        </html>
        """
        
        print(f"HTML content size: {len(html_content)} characters")
        
        # Create chapter
        chapter = epub.EpubHtml(title='Data', file_name='data.xhtml', lang='en')
        chapter.content = html_content
        book.add_item(chapter)
        
        # Create spine
        book.spine = ['data']
        
        # Add navigation
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        
        # Save EPUB
        print(f"Saving EPUB file to: {output_file}")
        try:
            epub.write_epub(output_file, book, {})
            print("EPUB file saved successfully")
        except Exception as e:
            print(f"ERROR saving EPUB: {e}")
            return False
        
        # Verify file
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"EPUB file created: {file_size / (1024*1024):.2f} MB ({file_size} bytes)")
            
            if file_size < 1024:  # Less than 1KB
                print(f"WARNING: File is very small ({file_size} bytes)")
                return False
            
            # Basic validation - check if it's a valid ZIP file
            try:
                import zipfile
                with zipfile.ZipFile(output_file, 'r') as zip_file:
                    file_list = zip_file.namelist()
                    print(f"EPUB contains {len(file_list)} files")
                    print(f"Files: {file_list}")
            except Exception as e:
                print(f"WARNING: Could not validate EPUB as ZIP: {e}")
            
            return True
        else:
            print("ERROR: EPUB file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to EPUB (simple version)')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output EPUB file path')
    parser.add_argument('--title', default='CSV Data', help='Book title')
    parser.add_argument('--author', default='Unknown', help='Book author')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Chunk size (ignored in simple version)')
    parser.add_argument('--no-toc', action='store_true', help='Do not include table of contents (ignored in simple version)')
    
    args = parser.parse_args()
    
    print("=== Simple CSV to EPUB Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to EPUB
    success = create_simple_epub_from_csv(
        args.csv_file,
        args.output_file,
        args.title,
        args.author
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()
