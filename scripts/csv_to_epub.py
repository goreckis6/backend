#!/usr/bin/env python3
"""
CSV to EPUB Converter
Converts CSV files to EPUB ebook format using ebooklib library.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback
from io import StringIO
import zipfile
import tempfile
import shutil

try:
    from ebooklib import epub
    from ebooklib.epub import EpubBook, EpubHtml, EpubNcx, EpubNav
except ImportError as e:
    print(f"ERROR: Required EPUB library not available: {e}")
    print("Please install ebooklib: pip install ebooklib")
    sys.exit(1)

def create_epub_from_csv(csv_file, output_file, title="CSV Data", author="Unknown", include_toc=True, chunk_size=1000):
    """
    Convert CSV file to EPUB format with performance optimizations.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output EPUB file
        title (str): Book title
        author (str): Book author
        include_toc (bool): Whether to include table of contents
        chunk_size (int): Number of rows to process at once
    """
    print("=" * 50)
    print("CSV TO EPUB CONVERSION STARTED")
    print("=" * 50)
    print(f"Starting CSV to EPUB conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print(f"Chunk size: {chunk_size}")
    print("=" * 50)
    
    try:
        # Read CSV file with optimizations
        print("Reading CSV file with optimizations...")
        print(f"File exists: {os.path.exists(csv_file)}")
        print(f"File readable: {os.access(csv_file, os.R_OK)}")
        print(f"File size: {os.path.getsize(csv_file)} bytes")
        
        # Get file size for progress tracking
        file_size = os.path.getsize(csv_file)
        print(f"File size: {file_size / (1024*1024):.2f} MB")
        
        # Read CSV with optimized settings
        df = pd.read_csv(
            csv_file,
            dtype=str,  # Read all as strings to avoid type inference overhead
            na_filter=False,  # Disable NaN filtering for speed
            low_memory=False  # Use more memory for speed
        )
        
        print(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        print("=" * 50)
        print("CSV READING COMPLETED - STARTING PROCESSING")
        print("=" * 50)
        
        # Create EPUB book
        print("Creating EPUB book...")
        book = epub.EpubBook()
        
        # Set metadata
        book.set_identifier(f"csv-{datetime.now().strftime('%Y%m%d%H%M%S')}")
        book.set_title(title)
        book.set_language('en')
        book.add_author(author)
        book.add_metadata('DC', 'description', f'Data converted from CSV file: {os.path.basename(csv_file)}')
        book.add_metadata('DC', 'publisher', 'MorphyIMG CSV to EPUB Converter')
        book.add_metadata('DC', 'date', datetime.now().strftime('%Y-%m-%d'))
        
        # Create cover page
        print("Creating cover page...")
        cover_html = f"""
        <html>
        <head>
            <title>Cover</title>
            <style>
                body {{ 
                    font-family: Arial, sans-serif; 
                    text-align: center; 
                    padding: 50px;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    margin: 0;
                    height: 100vh;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                }}
                h1 {{ font-size: 3em; margin-bottom: 20px; }}
                h2 {{ font-size: 1.5em; margin-bottom: 30px; opacity: 0.9; }}
                .info {{ font-size: 1.2em; margin: 10px 0; }}
                .stats {{ 
                    background: rgba(255,255,255,0.1); 
                    padding: 20px; 
                    border-radius: 10px; 
                    margin-top: 30px;
                }}
            </style>
        </head>
        <body>
            <h1>{title}</h1>
            <h2>by {author}</h2>
            <div class="info">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
            <div class="stats">
                <div class="info">Rows: {len(df):,}</div>
                <div class="info">Columns: {len(df.columns)}</div>
                <div class="info">Source: {os.path.basename(csv_file)}</div>
            </div>
        </body>
        </html>
        """
        
        cover = epub.EpubHtml(title='Cover', file_name='cover.xhtml', lang='en')
        cover.content = cover_html
        book.add_item(cover)
        
        # Create table of contents
        if include_toc:
            print("Creating table of contents...")
            toc_html = f"""
            <html>
            <head>
                <title>Table of Contents</title>
                <style>
                    body {{ font-family: Arial, sans-serif; padding: 20px; }}
                    h1 {{ color: #333; border-bottom: 2px solid #667eea; padding-bottom: 10px; }}
                    .toc-item {{ margin: 10px 0; padding: 10px; background: #f5f5f5; border-radius: 5px; }}
                    .toc-item a {{ text-decoration: none; color: #667eea; font-weight: bold; }}
                    .toc-item a:hover {{ color: #764ba2; }}
                    .stats {{ background: #e8f4f8; padding: 15px; border-radius: 5px; margin: 20px 0; }}
                </style>
            </head>
            <body>
                <h1>Table of Contents</h1>
                <div class="stats">
                    <h3>Data Summary</h3>
                    <p><strong>Total Rows:</strong> {len(df):,}</p>
                    <p><strong>Total Columns:</strong> {len(df.columns)}</p>
                    <p><strong>Columns:</strong> {', '.join(df.columns[:10])}{'...' if len(df.columns) > 10 else ''}</p>
                </div>
                <div class="toc-item">
                    <a href="cover.xhtml">Cover Page</a>
                </div>
                <div class="toc-item">
                    <a href="data.xhtml">Data Table</a>
                </div>
            </body>
            </html>
            """
            
            toc = epub.EpubHtml(title='Table of Contents', file_name='toc.xhtml', lang='en')
            toc.content = toc_html
            book.add_item(toc)
        
        # Create data table page
        print("Creating data table page...")
        data_html = f"""
        <html>
        <head>
            <title>Data Table</title>
            <style>
                body {{ font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }}
                h1 {{ color: #333; border-bottom: 2px solid #667eea; padding-bottom: 10px; }}
                table {{ 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 20px 0;
                    font-size: 0.9em;
                }}
                th {{ 
                    background: #667eea; 
                    color: white; 
                    padding: 12px 8px; 
                    text-align: left; 
                    font-weight: bold;
                    position: sticky;
                    top: 0;
                }}
                td {{ 
                    padding: 8px; 
                    border-bottom: 1px solid #ddd; 
                    vertical-align: top;
                }}
                tr:nth-child(even) {{ background: #f9f9f9; }}
                tr:hover {{ background: #f0f8ff; }}
                .summary {{ 
                    background: #e8f4f8; 
                    padding: 15px; 
                    border-radius: 5px; 
                    margin: 20px 0;
                }}
                .note {{ 
                    background: #fff3cd; 
                    border: 1px solid #ffeaa7; 
                    padding: 10px; 
                    border-radius: 5px; 
                    margin: 10px 0;
                }}
            </style>
        </head>
        <body>
            <h1>Data Table</h1>
            <div class="summary">
                <h3>Dataset Information</h3>
                <p><strong>Source File:</strong> {os.path.basename(csv_file)}</p>
                <p><strong>Total Records:</strong> {len(df):,}</p>
                <p><strong>Total Columns:</strong> {len(df.columns)}</p>
                <p><strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </div>
            <div class="note">
                <strong>Note:</strong> This table contains the first {min(1000, len(df))} rows of data. 
                {'For the complete dataset, please refer to the original CSV file.' if len(df) > 1000 else ''}
            </div>
        """
        
        # Add table header
        data_html += "<table><thead><tr>"
        for col in df.columns:
            data_html += f"<th>{col}</th>"
        data_html += "</tr></thead><tbody>"
        
        # Add table data (include more rows for better file size)
        # For files with < 5000 rows, include all data; otherwise limit to 5000 for performance
        max_rows = min(5000, len(df))
        print(f"Adding table data (first {max_rows} rows out of {len(df)} total)...")
        print(f"DataFrame shape: {df.shape}")
        print(f"DataFrame columns: {list(df.columns)}")
        
        rows_added = 0
        for idx in range(max_rows):
            if idx % 100 == 0:
                print(f"Processing row {idx + 1}/{max_rows}")
            
            row = df.iloc[idx]
            data_html += "<tr>"
            for value in row:
                # Escape HTML and handle long text
                cell_value = str(value) if value else ""
                if len(cell_value) > 100:
                    cell_value = cell_value[:97] + "..."
                data_html += f"<td>{cell_value}</td>"
            data_html += "</tr>"
            rows_added += 1
        
        print(f"Successfully added {rows_added} rows to HTML table")
        
        data_html += """
            </tbody></table>
            </body>
            </html>
        """
        
        data_page = epub.EpubHtml(title='Data Table', file_name='data.xhtml', lang='en')
        data_page.content = data_html
        book.add_item(data_page)
        
        print(f"Data HTML content size: {len(data_html)} characters")
        print(f"Data HTML preview (first 200 chars): {data_html[:200]}...")
        
        # Create spine (reading order)
        print("Creating book spine...")
        book.spine = ['cover']
        if include_toc:
            book.spine.append('toc')
        book.spine.append('data')
        
        # Add navigation (NCX) for better compatibility
        print("Adding navigation (NCX)...")
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        
        # Add navigation
        if include_toc:
            print("Adding table of contents...")
            # Create TOC page
            toc_html = f"""
            <html>
            <head>
                <title>Table of Contents</title>
                <style>
                    body {{ font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }}
                    h1 {{ color: #333; border-bottom: 2px solid #667eea; padding-bottom: 10px; }}
                    ul {{ list-style-type: none; padding: 0; }}
                    li {{ margin: 10px 0; }}
                    a {{ text-decoration: none; color: #667eea; font-size: 1.1em; }}
                    a:hover {{ color: #764ba2; }}
                </style>
            </head>
            <body>
                <h1>Table of Contents</h1>
                <ul>
                    <li><a href="cover.xhtml">Cover</a></li>
                    <li><a href="data.xhtml">Data Table</a></li>
                </ul>
            </body>
            </html>
            """
            toc_page = epub.EpubHtml(title='Table of Contents', file_name='toc.xhtml', lang='en')
            toc_page.content = toc_html
            book.add_item(toc_page)
            
            book.toc = [
                epub.Link('cover.xhtml', 'Cover', 'cover'),
                epub.Link('toc.xhtml', 'Table of Contents', 'toc'),
                epub.Link('data.xhtml', 'Data Table', 'data')
            ]
        
        # Add CSS style
        print("Adding CSS styles...")
        style = epub.EpubItem(
            uid="style_default",
            file_name="style/default.css",
            media_type="text/css",
            content="""
            body { font-family: Arial, sans-serif; }
            h1, h2, h3 { color: #333; }
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f2f2f2; }
            """
        )
        book.add_item(style)
        
        # Memory cleanup for large files
        print("Cleaning up memory...")
        del df  # Free up memory from the large DataFrame
        
        # Save EPUB file with proper options
        print(f"Saving EPUB file to {output_file}...")
        try:
            # Create EPUB with proper options for better compatibility
            epub.write_epub(output_file, book, {
                'epub2_guide': True,
                'epub3_landmark': True,
                'epub3_nav': True
            })
            print("EPUB file saved successfully with full options")
        except Exception as e:
            print(f"Error saving EPUB with full options: {e}")
            try:
                # Try with basic options
                epub.write_epub(output_file, book, {
                    'epub2_guide': True
                })
                print("EPUB file saved with basic options")
            except Exception as e2:
                print(f"Error saving EPUB with basic options: {e2}")
                # Try with minimal options
                epub.write_epub(output_file, book, {})
                print("EPUB file saved with minimal options")
        
        # Verify file was created and is valid
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"EPUB file created successfully: {file_size / (1024*1024):.2f} MB")
            
            # Check if file is too small (less than 10KB is suspicious)
            if file_size < 10240:  # 10KB
                print(f"WARNING: EPUB file is very small ({file_size} bytes). This might indicate an issue.")
            
            # Basic validation - check if it's a valid ZIP file (EPUB is a ZIP)
            try:
                with zipfile.ZipFile(output_file, 'r') as zip_file:
                    file_list = zip_file.namelist()
                    print(f"EPUB contains {len(file_list)} files:")
                    for file in file_list:
                        print(f"  - {file}")
                    
                    # Check for required EPUB files
                    required_files = ['META-INF/container.xml', 'OEBPS/content.opf']
                    missing_files = [f for f in required_files if f not in file_list]
                    if missing_files:
                        print(f"WARNING: Missing required EPUB files: {missing_files}")
                    else:
                        print("EPUB structure validation passed")
                        
            except zipfile.BadZipFile:
                print("ERROR: Generated file is not a valid ZIP/EPUB file")
                return False
            except Exception as e:
                print(f"WARNING: Could not validate EPUB structure: {e}")
            
            return True
        else:
            print("ERROR: EPUB file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to create EPUB from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to EPUB format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output EPUB file path')
    parser.add_argument('--title', default='CSV Data', help='Book title')
    parser.add_argument('--author', default='Unknown', help='Book author')
    parser.add_argument('--no-toc', action='store_true', help='Do not include table of contents')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Chunk size for processing large files')
    
    args = parser.parse_args()
    
    print("=== CSV to EPUB Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)
    
    # Check pandas availability
    try:
        print(f"Pandas version: {pd.__version__}")
    except Exception as e:
        print(f"ERROR: Pandas not available: {e}")
        sys.exit(1)
    
    # Check ebooklib availability
    try:
        import ebooklib
        print(f"ebooklib library is available")
    except Exception as e:
        print(f"ERROR: ebooklib not available: {e}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to EPUB
    success = create_epub_from_csv(
        args.csv_file,
        args.output_file,
        args.title,
        args.author,
        not args.no_toc,
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
