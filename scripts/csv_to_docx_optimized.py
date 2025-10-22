#!/usr/bin/env python3
"""
Optimized CSV to DOCX Converter
High-performance CSV to DOCX conversion with optimizations for large files.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback
from io import StringIO

try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.shared import OxmlElement, qn
    from docx.oxml.ns import qn as qn_ns
except ImportError as e:
    print(f"ERROR: Required DOCX library not available: {e}")
    print("Please install python-docx: pip install python-docx")
    sys.exit(1)

def create_docx_from_csv_optimized(csv_file, output_file, title="CSV Data", author="Unknown", chunk_size=1000):
    """
    Optimized CSV to DOCX conversion with performance improvements.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output DOCX file
        title (str): Document title
        author (str): Document author
        chunk_size (int): Number of rows to process at once
    """
    print("=" * 50)
    print("CSV TO DOCX CONVERSION STARTED")
    print("=" * 50)
    print(f"Starting optimized CSV to DOCX conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
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
        
        # Read CSV with optimized settings for large files
        df = pd.read_csv(
            csv_file,
            dtype=str,  # Read all as strings to avoid type inference overhead
            na_filter=False,  # Disable NaN filtering for speed
            low_memory=False,  # Use more memory for speed
            engine='c',  # Use C engine for better performance
            chunksize=None  # Read entire file at once for processing
        )
        
        print(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        print("=" * 50)
        print("CSV READING COMPLETED - STARTING PROCESSING")
        print("=" * 50)
        
        # Estimate processing time
        estimated_time = len(df) / 1000  # Rough estimate: 1000 rows per second
        print(f"Estimated processing time: {estimated_time:.1f} seconds for {len(df)} rows")
        
        # Create DOCX document
        print("Creating optimized DOCX document...")
        doc = Document()
        
        # Set document properties
        doc.core_properties.title = title
        doc.core_properties.author = author
        doc.core_properties.created = datetime.now()
        
        # Add minimal title section
        print("Adding document header...")
        title_para = doc.add_heading(title, 0)
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add creation info
        info_para = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Rows: {len(df)} | Columns: {len(df.columns)}")
        info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()  # Empty line
        
        # Create table with optimizations
        print("Creating optimized data table...")
        table = doc.add_table(rows=1, cols=len(df.columns))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'
        
        # Add header row with minimal styling
        print("Adding table headers...")
        header_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            header_cells[i].text = str(col)
            # Minimal header styling
            for paragraph in header_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                    run.font.size = Pt(9)  # Smaller font for speed
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Process data in chunks for better performance
        print(f"Processing {len(df)} rows in chunks of {chunk_size}...")
        
        total_rows = len(df)
        
        # Pre-allocate table rows for much better performance
        print(f"Pre-allocating {total_rows} table rows...")
        for _ in range(total_rows):
            table.add_row()
        
        print("Pre-allocation complete, now filling data...")
        
        # Process data in chunks (single-threaded for reliability)
        for chunk_start in range(0, total_rows, chunk_size):
            chunk_end = min(chunk_start + chunk_size, total_rows)
            chunk_df = df.iloc[chunk_start:chunk_end]
            
            print(f"Processing rows {chunk_start + 1}-{chunk_end} of {total_rows} ({(chunk_end/total_rows*100):.1f}%)")
            
            # Process chunk data more efficiently
            for idx, (_, row) in enumerate(chunk_df.iterrows()):
                row_index = chunk_start + idx
                row_cells = table.rows[row_index + 1].cells  # +1 because header is row 0
                
                # Process cells with minimal styling - no individual cell styling for speed
                for i, value in enumerate(row):
                    # Convert to string efficiently
                    cell_value = str(value) if value else ""
                    row_cells[i].text = cell_value
        
        # Add minimal summary
        print("Adding document summary...")
        doc.add_paragraph()
        summary_para = doc.add_paragraph(f"Total: {len(df)} rows × {len(df.columns)} columns")
        summary_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Memory cleanup for large files
        print("Cleaning up memory...")
        del df  # Free up memory from the large DataFrame
        
        # Save document with optimizations
        print(f"Saving DOCX document...")
        doc.save(output_file)
        
        # Verify file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"DOCX file created successfully: {file_size / (1024*1024):.2f} MB")
            return True
        else:
            print("ERROR: DOCX file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to create DOCX from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to DOCX format (optimized)')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output DOCX file path')
    parser.add_argument('--title', default='CSV Data', help='Document title')
    parser.add_argument('--author', default='Unknown', help='Document author')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Chunk size for processing large files')
    
    args = parser.parse_args()
    
    print("=== Optimized CSV to DOCX Converter ===")
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
    
    # Check python-docx availability
    try:
        from docx import Document
        print("python-docx library is available")
    except Exception as e:
        print(f"ERROR: python-docx not available: {e}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to DOCX with optimizations
    success = create_docx_from_csv_optimized(
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
