#!/usr/bin/env python3
"""
CSV to ODT Converter
Converts CSV files to OpenDocument Text (ODT) format using pandas and odfpy.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback

try:
    from odf.opendocument import OpenDocumentText
    from odf.style import Style, TextProperties, ParagraphProperties, TableColumnProperties
    from odf.text import P, H, S
    from odf.table import Table, TableColumn, TableRow, TableCell
    from odf import style
except ImportError as e:
    print(f"ERROR: Required ODF library not available: {e}")
    print("Please install odfpy: pip install odfpy")
    sys.exit(1)

def create_odt_from_csv(csv_file, output_file, title="CSV Data", author="Unknown"):
    """
    Convert CSV file to ODT format.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output ODT file
        title (str): Document title
        author (str): Document author
    """
    print(f"Starting CSV to ODT conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print("Processing all data (no row limits)")
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file)
        print(f"CSV loaded successfully: {len(df)} rows, {len(df.columns)} columns")
        
        # Process all rows - no limits
        print(f"Processing all {len(df)} rows (including any repeated data)")
        
        # Create ODT document
        print("Creating ODT document...")
        doc = OpenDocumentText()
        
        # Skip metadata for now - focus on document content
        print("Skipping metadata setup, focusing on document content...")
        
        # Create styles
        print("Creating document styles...")
        
        # Title style
        title_style = Style(name="Title", family="paragraph")
        title_style.addElement(TextProperties(fontsize="18pt", fontweight="bold"))
        title_style.addElement(ParagraphProperties(textalign="center", margintop="0.5in", marginbottom="0.3in"))
        doc.styles.addElement(title_style)
        
        # Header style
        header_style = Style(name="Header", family="paragraph")
        header_style.addElement(TextProperties(fontweight="bold", fontsize="12pt"))
        header_style.addElement(ParagraphProperties(margintop="0.2in", marginbottom="0.1in"))
        doc.styles.addElement(header_style)
        
        # Table header style
        table_header_style = Style(name="TableHeader", family="table-cell")
        table_header_style.addElement(TextProperties(fontweight="bold", fontsize="10pt"))
        table_header_style.addElement(ParagraphProperties(textalign="center"))
        doc.styles.addElement(table_header_style)
        
        # Table cell style
        table_cell_style = Style(name="TableCell", family="table-cell")
        table_cell_style.addElement(TextProperties(fontsize="10pt"))
        table_cell_style.addElement(ParagraphProperties(textalign="left"))
        doc.styles.addElement(table_cell_style)
        
        # Add title
        print("Adding document title...")
        title_para = P(stylename="Title")
        title_para.addText(title)
        doc.text.addElement(title_para)
        
        # Add author info
        author_para = P(stylename="Header")
        author_para.addText(f"Author: {author}")
        doc.text.addElement(author_para)
        
        # Add creation date
        date_para = P(stylename="Header")
        date_para.addText(f"Created: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        doc.text.addElement(date_para)
        
        # Add empty line
        doc.text.addElement(P())
        
        # Create table
        print("Creating data table...")
        table = Table()
        
        # Add table columns (without style attributes for now)
        for _ in range(len(df.columns)):
            table.addElement(TableColumn())
        
        # Add header row
        print("Adding table headers...")
        header_row = TableRow()
        for col in df.columns:
            cell = TableCell(stylename="TableHeader")
            p = P()
            p.addText(str(col))
            cell.addElement(p)
            header_row.addElement(cell)
        table.addElement(header_row)
        
        # Add data rows
        print(f"Adding {len(df)} data rows...")
        for idx, row in df.iterrows():
            if idx % 100 == 0:
                print(f"Processing row {idx + 1}/{len(df)}")
            
            data_row = TableRow()
            for value in row:
                cell = TableCell(stylename="TableCell")
                # Handle NaN values and convert to string
                cell_value = str(value) if pd.notna(value) else ""
                p = P()
                p.addText(cell_value)
                cell.addElement(p)
                data_row.addElement(cell)
            table.addElement(data_row)
        
        # Add table to document
        doc.text.addElement(table)
        
        # Add summary
        print("Adding document summary...")
        doc.text.addElement(P())
        summary_para = P(stylename="Header")
        summary_para.addText(f"Summary: {len(df)} rows, {len(df.columns)} columns")
        doc.text.addElement(summary_para)
        
        # Save document
        print(f"Saving ODT document to {output_file}...")
        doc.save(output_file)
        
        # Verify file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"ODT file created successfully: {file_size} bytes")
            return True
        else:
            print("ERROR: ODT file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to create ODT from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to ODT format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output ODT file path')
    parser.add_argument('--title', default='CSV Data', help='Document title')
    parser.add_argument('--author', default='Unknown', help='Document author')
    
    args = parser.parse_args()
    
    print("=== CSV to ODT Converter ===")
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
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to ODT
    success = create_odt_from_csv(
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
