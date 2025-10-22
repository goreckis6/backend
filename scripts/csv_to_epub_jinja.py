#!/usr/bin/env python3
"""
CSV to EPUB Converter using Jinja2 Templates
Converts CSV files to EPUB ebook format using jinja2 for templating and ebooklib for EPUB generation.
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
import uuid

try:
    from jinja2 import Template
    from ebooklib import epub
    from ebooklib.epub import EpubBook, EpubHtml, EpubNcx, EpubNav
except ImportError as e:
    print(f"ERROR: Required libraries not available: {e}")
    print("Please install: pip install jinja2 ebooklib")
    sys.exit(1)

def create_epub_from_csv_jinja(csv_file, output_file, title="CSV Data", author="Unknown", include_toc=True, chunk_size=1000):
    """
    Convert CSV file to EPUB format using Jinja2 templates for better HTML generation.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output EPUB file
        title (str): Book title
        author (str): Book author
        include_toc (bool): Whether to include table of contents
        chunk_size (int): Number of rows to process at once
    """
    print("=" * 60)
    print("CSV TO EPUB CONVERSION (JINJA2 TEMPLATES)")
    print("=" * 60)
    print(f"Starting CSV to EPUB conversion with Jinja2...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print(f"Chunk size: {chunk_size}")
    print("=" * 60)
    
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
        print(f"DataFrame shape: {df.shape}")
        print(f"DataFrame columns: {list(df.columns)}")
        print("=" * 60)
        print("CSV READING COMPLETED - STARTING EPUB GENERATION")
        print("=" * 60)
        
        # Create EPUB book
        print("Creating EPUB book structure...")
        book = epub.EpubBook()
        
        # Set metadata
        book.set_identifier(str(uuid.uuid4()))
        book.set_title(title)
        book.set_language('en')
        book.add_author(author)
        book.add_metadata('DC', 'description', f'Data converted from CSV file: {os.path.basename(csv_file)}')
        book.add_metadata('DC', 'publisher', 'MorphyIMG CSV Converter')
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
                h2 {{ font-size: 1.5em; margin-bottom: 30px; }}
                .info {{ font-size: 1.2em; line-height: 1.6; }}
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
            <h2>Data Report</h2>
            <div class="info">
                <p>Generated from CSV file</p>
                <p>Author: {author}</p>
                <p>Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            </div>
            <div class="stats">
                <p><strong>Total Records:</strong> {len(df):,}</p>
                <p><strong>Total Columns:</strong> {len(df.columns)}</p>
                <p><strong>Source File:</strong> {os.path.basename(csv_file)}</p>
            </div>
        </body>
        </html>
        """
        
        cover_page = epub.EpubHtml(title='Cover', file_name='cover.xhtml', lang='en')
        cover_page.content = cover_html
        book.add_item(cover_page)
        
        # Create data page using Jinja2 template
        print("Creating data page with Jinja2 template...")
        
        # Jinja2 template for data table
        data_template = Template("""
        <html>
        <head>
            <title>Data Table</title>
            <style>
                body { 
                    font-family: Arial, sans-serif; 
                    margin: 20px; 
                    line-height: 1.6;
                    background: #f5f5f5;
                }
                .container {
                    background: white;
                    padding: 30px;
                    border-radius: 10px;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }
                h1 { 
                    color: #333; 
                    border-bottom: 3px solid #667eea; 
                    padding-bottom: 10px;
                    margin-bottom: 30px;
                }
                .summary { 
                    background: #e8f4f8; 
                    padding: 20px; 
                    border-radius: 8px; 
                    margin: 20px 0;
                    border-left: 4px solid #667eea;
                }
                .note { 
                    background: #fff3cd; 
                    border: 1px solid #ffeaa7; 
                    padding: 15px; 
                    border-radius: 8px; 
                    margin: 20px 0;
                    border-left: 4px solid #ffc107;
                }
                table { 
                    border-collapse: collapse; 
                    width: 100%; 
                    margin: 20px 0;
                    background: white;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                }
                th { 
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white; 
                    padding: 15px 10px; 
                    text-align: left; 
                    font-weight: bold;
                    font-size: 14px;
                }
                td { 
                    padding: 12px 10px; 
                    border-bottom: 1px solid #eee;
                    font-size: 13px;
                }
                tr:nth-child(even) { background: #f9f9f9; }
                tr:hover { background: #f0f8ff; }
                .cell-content {
                    max-width: 200px;
                    overflow: hidden;
                    text-overflow: ellipsis;
                    white-space: nowrap;
                }
                .cell-content:hover {
                    white-space: normal;
                    max-width: none;
                    background: #fff;
                    padding: 5px;
                    border-radius: 4px;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Data Table</h1>
                
                <div class="summary">
                    <h3>Dataset Information</h3>
                    <p><strong>Source File:</strong> {{ csv_filename }}</p>
                    <p><strong>Total Records:</strong> {{ total_records:, }}</p>
                    <p><strong>Total Columns:</strong> {{ total_columns }}</p>
                    <p><strong>Records Shown:</strong> {{ records_shown:, }}</p>
                    <p><strong>Generated:</strong> {{ generated_time }}</p>
                </div>
                
                {% if total_records > records_shown %}
                <div class="note">
                    <strong>Note:</strong> This table contains the first {{ records_shown }} rows of data. 
                    For the complete dataset, please refer to the original CSV file.
                </div>
                {% endif %}
                
                <table>
                    <thead>
                        <tr>
                            {% for column in columns %}
                            <th>{{ column }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in data_rows %}
                        <tr>
                            {% for value in row %}
                            <td>
                                <div class="cell-content" title="{{ value }}">
                                    {{ value }}
                                </div>
                            </td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </body>
        </html>
        """)
        
        # Prepare data for template
        max_rows = min(5000, len(df))  # Process up to 5000 rows
        print(f"Processing {max_rows} rows for EPUB generation...")
        
        # Convert DataFrame to list of lists for Jinja2
        data_rows = []
        for idx in range(max_rows):
            if idx % 500 == 0:
                print(f"Preparing row {idx + 1}/{max_rows}")
            row = df.iloc[idx].tolist()
            # Truncate very long values for display
            row = [str(val)[:200] + "..." if len(str(val)) > 200 else str(val) for val in row]
            data_rows.append(row)
        
        # Render template with data
        print("Rendering HTML with Jinja2 template...")
        data_html = data_template.render(
            csv_filename=os.path.basename(csv_file),
            total_records=len(df),
            total_columns=len(df.columns),
            records_shown=max_rows,
            generated_time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            columns=df.columns.tolist(),
            data_rows=data_rows
        )
        
        print(f"Generated HTML content: {len(data_html)} characters")
        
        # Create data page
        data_page = epub.EpubHtml(title='Data Table', file_name='data.xhtml', lang='en')
        data_page.content = data_html
        book.add_item(data_page)
        
        # Create table of contents if requested
        if include_toc:
            print("Creating table of contents...")
            toc_template = Template("""
            <html>
            <head>
                <title>Table of Contents</title>
                <style>
                    body { 
                        font-family: Arial, sans-serif; 
                        padding: 30px; 
                        line-height: 1.8;
                        background: #f5f5f5;
                    }
                    .container {
                        background: white;
                        padding: 40px;
                        border-radius: 10px;
                        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                        max-width: 600px;
                        margin: 0 auto;
                    }
                    h1 { 
                        color: #333; 
                        border-bottom: 3px solid #667eea; 
                        padding-bottom: 15px;
                        margin-bottom: 30px;
                    }
                    ul { 
                        list-style: none; 
                        padding: 0; 
                    }
                    li { 
                        margin: 15px 0; 
                        padding: 10px;
                        border-radius: 5px;
                        transition: background 0.3s;
                    }
                    li:hover {
                        background: #f0f8ff;
                    }
                    a { 
                        text-decoration: none; 
                        color: #667eea; 
                        font-size: 1.2em;
                        font-weight: 500;
                    }
                    a:hover { 
                        color: #764ba2; 
                    }
                    .description {
                        font-size: 0.9em;
                        color: #666;
                        margin-top: 5px;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>Table of Contents</h1>
                    <ul>
                        <li><a href="cover.xhtml">📖 Cover</a>
                            <div class="description">Book cover and summary information</div>
                        </li>
                        <li><a href="data.xhtml">📊 Data Table</a>
                            <div class="description">Complete dataset with {{ total_records:, }} records</div>
                        </li>
                    </ul>
                </div>
            </body>
            </html>
            """)
            
            toc_html = toc_template.render(
                total_records=len(df)
            )
            
            toc_page = epub.EpubHtml(title='Table of Contents', file_name='toc.xhtml', lang='en')
            toc_page.content = toc_html
            book.add_item(toc_page)
        
        # Create spine (reading order)
        print("Creating book spine...")
        book.spine = ['cover']
        if include_toc:
            book.spine.append('toc')
        book.spine.append('data')
        
        # Add navigation
        print("Adding navigation...")
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        
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
        
        # Memory cleanup
        print("Cleaning up memory...")
        del df
        del data_rows
        
        # Save EPUB file
        print(f"Saving EPUB file to {output_file}...")
        try:
            # Try with full options first
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
            
            # Check if file is too small
            if file_size < 10240:  # 10KB
                print(f"WARNING: EPUB file is very small ({file_size} bytes). This might indicate an issue.")
            
            # Validate EPUB structure
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
    parser = argparse.ArgumentParser(description='Convert CSV to EPUB format using Jinja2 templates')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output EPUB file path')
    parser.add_argument('--title', default='CSV Data', help='Book title')
    parser.add_argument('--author', default='Unknown', help='Book author')
    parser.add_argument('--no-toc', action='store_true', help='Do not include table of contents')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Chunk size for processing large files')
    
    args = parser.parse_args()
    
    print("=== CSV to EPUB Converter (Jinja2) ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        print(f"Jinja2 version: {jinja2.__version__}")
        print(f"Ebooklib version: {epub.__version__}")
    except Exception as e:
        print(f"ERROR: Library version check failed: {e}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to EPUB
    success = create_epub_from_csv_jinja(
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
