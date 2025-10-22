#!/usr/bin/env python3
"""
Manual CSV to EPUB Converter
Creates a minimal but valid EPUB by manually constructing the ZIP structure.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback
import zipfile
import tempfile
import shutil
import uuid
import xml.etree.ElementTree as ET

def create_manual_epub_from_csv(csv_file, output_file, title="CSV Data", author="Unknown"):
    """
    Create a minimal but valid EPUB by manually constructing the structure.
    """
    print("=" * 60)
    print("MANUAL CSV TO EPUB CONVERSION")
    print("=" * 60)
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print("=" * 60)
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file, dtype=str, na_filter=False)
        print(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        
        # Limit data for performance
        max_rows = min(1000, len(df))
        print(f"Processing {max_rows} rows...")
        
        # Create temporary directory
        temp_dir = tempfile.mkdtemp(prefix='epub_manual_')
        print(f"Using temporary directory: {temp_dir}")
        
        try:
            # Create META-INF directory
            meta_inf_dir = os.path.join(temp_dir, 'META-INF')
            os.makedirs(meta_inf_dir, exist_ok=True)
            
            # Create OEBPS directory
            oebps_dir = os.path.join(temp_dir, 'OEBPS')
            os.makedirs(oebps_dir, exist_ok=True)
            
            # Create mimetype file
            mimetype_file = os.path.join(temp_dir, 'mimetype')
            with open(mimetype_file, 'w') as f:
                f.write('application/epub+zip')
            
            # Create container.xml
            container_xml = os.path.join(meta_inf_dir, 'container.xml')
            container_content = '''<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
    <rootfiles>
        <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
    </rootfiles>
</container>'''
            with open(container_xml, 'w', encoding='utf-8') as f:
                f.write(container_content)
            
            # Create content.opf
            content_opf = os.path.join(oebps_dir, 'content.opf')
            book_id = str(uuid.uuid4())
            current_time = datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
            
            opf_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" unique-identifier="book-id" version="2.0">
    <metadata xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf">
        <dc:title>{title}</dc:title>
        <dc:creator opf:role="aut">{author}</dc:creator>
        <dc:language>en</dc:language>
        <dc:identifier id="book-id" opf:scheme="UUID">{book_id}</dc:identifier>
        <dc:date opf:event="publication">{current_time}</dc:date>
        <dc:publisher>MorphyIMG CSV Converter</dc:publisher>
    </metadata>
    <manifest>
        <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
        <item id="content" href="content.xhtml" media-type="application/xhtml+xml"/>
        <item id="css" href="style.css" media-type="text/css"/>
    </manifest>
    <spine toc="ncx">
        <itemref idref="content"/>
    </spine>
</package>'''
            
            with open(content_opf, 'w', encoding='utf-8') as f:
                f.write(opf_content)
            
            # Create CSS file
            css_file = os.path.join(oebps_dir, 'style.css')
            css_content = '''body {
    font-family: Arial, sans-serif;
    margin: 20px;
    line-height: 1.6;
    color: #333;
}
h1 {
    color: #2c3e50;
    border-bottom: 2px solid #3498db;
    padding-bottom: 10px;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin: 20px 0;
}
th, td {
    border: 1px solid #ddd;
    padding: 8px;
    text-align: left;
}
th {
    background-color: #f2f2f2;
    font-weight: bold;
}
tr:nth-child(even) {
    background-color: #f9f9f9;
}
.info {
    background: #ecf0f1;
    padding: 15px;
    border-radius: 5px;
    margin: 20px 0;
    border-left: 4px solid #3498db;
}'''
            
            with open(css_file, 'w', encoding='utf-8') as f:
                f.write(css_content)
            
            # Create content.xhtml
            content_xhtml = os.path.join(oebps_dir, 'content.xhtml')
            
            # Generate HTML table
            html_table = '<table>\n<thead>\n<tr>\n'
            for col in df.columns:
                html_table += f'<th>{col}</th>\n'
            html_table += '</tr>\n</thead>\n<tbody>\n'
            
            for idx in range(max_rows):
                if idx % 100 == 0:
                    print(f"Processing row {idx + 1}/{max_rows}")
                
                row = df.iloc[idx]
                html_table += '<tr>\n'
                for value in row:
                    # Escape HTML
                    cell_value = str(value) if value else ""
                    cell_value = cell_value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
                    html_table += f'<td>{cell_value}</td>\n'
                html_table += '</tr>\n'
            
            html_table += '</tbody>\n</table>'
            
            xhtml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>{title}</title>
    <link rel="stylesheet" type="text/css" href="style.css"/>
</head>
<body>
    <h1>{title}</h1>
    
    <div class="info">
        <h3>Dataset Information</h3>
        <p><strong>Source File:</strong> {os.path.basename(csv_file)}</p>
        <p><strong>Total Records:</strong> {len(df):,}</p>
        <p><strong>Records Shown:</strong> {max_rows:,}</p>
        <p><strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>
    
    {html_table}
</body>
</html>'''
            
            with open(content_xhtml, 'w', encoding='utf-8') as f:
                f.write(xhtml_content)
            
            # Create toc.ncx
            toc_ncx = os.path.join(oebps_dir, 'toc.ncx')
            toc_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE ncx PUBLIC "-//NISO//DTD ncx 2005-1//EN" "http://www.daisy.org/z3986/2005/ncx-2005-1.dtd">
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
    <head>
        <meta name="dtb:uid" content="{book_id}"/>
        <meta name="dtb:depth" content="1"/>
        <meta name="dtb:totalPageCount" content="0"/>
        <meta name="dtb:maxPageNumber" content="0"/>
    </head>
    <docTitle>
        <text>{title}</text>
    </docTitle>
    <navMap>
        <navPoint id="navpoint-1" playOrder="1">
            <navLabel>
                <text>{title}</text>
            </navLabel>
            <content src="content.xhtml"/>
        </navPoint>
    </navMap>
</ncx>'''
            
            with open(toc_ncx, 'w', encoding='utf-8') as f:
                f.write(toc_content)
            
            # Create EPUB ZIP file
            print("Creating EPUB ZIP file...")
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as epub_zip:
                # Add mimetype first (uncompressed)
                epub_zip.write(mimetype_file, 'mimetype', compress_type=zipfile.ZIP_STORED)
                
                # Add all other files
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file == 'mimetype':
                            continue  # Already added
                        file_path = os.path.join(root, file)
                        arc_path = os.path.relpath(file_path, temp_dir)
                        epub_zip.write(file_path, arc_path)
            
            print("EPUB ZIP file created successfully")
            
            # Verify the file
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"EPUB file created: {file_size / (1024*1024):.2f} MB ({file_size} bytes)")
                
                if file_size < 1024:
                    print(f"WARNING: File is very small ({file_size} bytes)")
                    return False
                
                # Validate ZIP structure
                try:
                    with zipfile.ZipFile(output_file, 'r') as zip_file:
                        file_list = zip_file.namelist()
                        print(f"EPUB contains {len(file_list)} files:")
                        for file in file_list:
                            print(f"  - {file}")
                        
                        # Check for required files
                        required_files = ['mimetype', 'META-INF/container.xml', 'OEBPS/content.opf', 'OEBPS/content.xhtml']
                        missing_files = [f for f in required_files if f not in file_list]
                        if missing_files:
                            print(f"WARNING: Missing required files: {missing_files}")
                        else:
                            print("EPUB structure validation passed")
                            
                except Exception as e:
                    print(f"WARNING: Could not validate EPUB structure: {e}")
                
                return True
            else:
                print("ERROR: EPUB file was not created")
                return False
                
        finally:
            # Clean up
            print(f"Cleaning up temporary directory: {temp_dir}")
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except Exception as e:
        print(f"ERROR: Failed to create EPUB from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to EPUB (manual construction)')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output EPUB file path')
    parser.add_argument('--title', default='CSV Data', help='Book title')
    parser.add_argument('--author', default='Unknown', help='Book author')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Chunk size (ignored)')
    parser.add_argument('--no-toc', action='store_true', help='Do not include table of contents (ignored)')
    
    args = parser.parse_args()
    
    print("=== Manual CSV to EPUB Converter ===")
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
    success = create_manual_epub_from_csv(
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
