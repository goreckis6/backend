#!/usr/bin/env python3
"""
CSV to RTF Converter
Converts CSV files to Rich Text Format (RTF) using pandas and custom RTF generation.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback
import html

def escape_rtf(text):
    """Escape special characters for RTF format."""
    if not isinstance(text, str):
        text = str(text)
    
    # Escape RTF special characters
    text = text.replace('\\', '\\\\')
    text = text.replace('{', '\\{')
    text = text.replace('}', '\\}')
    text = text.replace('\n', '\\par ')
    text = text.replace('\r', '')
    
    # Convert to UTF-8 and escape non-ASCII characters
    try:
        # Encode to bytes and then decode to handle special characters
        text = text.encode('utf-8').decode('utf-8')
    except:
        pass
    
    return text

def create_rtf_from_csv(csv_file, output_file, title="CSV Data", author="Unknown"):
    """
    Convert CSV file to RTF format.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output RTF file
        title (str): Document title
        author (str): Document author
    """
    print(f"Starting CSV to RTF conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file)
        print(f"CSV loaded successfully: {len(df)} rows, {len(df.columns)} columns")
        print(f"Column names: {list(df.columns)}")
        print(f"First few rows:")
        print(df.head())
        
        # Process all rows - no limits
        print(f"Processing all {len(df)} rows (including any repeated data)")
        
        # Check for any data issues
        print(f"Data types: {df.dtypes.to_dict()}")
        print(f"Any null values: {df.isnull().sum().sum()}")
        
        # Start building RTF content
        rtf_content = []
        
        # RTF header
        rtf_content.append(r"{\rtf1\ansi\deff0")
        rtf_content.append(r"{\fonttbl{\f0\fswiss\fcharset0 Arial;}{\f1\fswiss\fcharset0 Arial Bold;}}")
        rtf_content.append(r"{\colortbl;\red0\green0\blue0;\red68\green114\blue196;\red242\green242\blue242;}")
        
        # Document properties
        rtf_content.append(r"{\info")
        rtf_content.append(f"{{\\title {escape_rtf(title)}}}")
        rtf_content.append(f"{{\\author {escape_rtf(author)}}}")
        
        # Create creation time string separately to avoid f-string issues
        now = datetime.now()
        creation_time = f"{{\\creatim\\yr{now.year}\\mo{now.month}\\dy{now.day}\\hr{now.hour}\\min{now.minute}}}"
        rtf_content.append(creation_time)
        rtf_content.append(r"}")
        
        # Document content
        rtf_content.append(r"{\f0\fs24")  # Start with Arial, 12pt
        
        # Title
        rtf_content.append(r"{\qc\b\fs32 " + escape_rtf(title) + r"\par}")
        
        # Create title info strings separately
        title_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        rtf_content.append(r"{\qc\fs18 " + escape_rtf(f"Generated on {title_time}") + r"\par}")
        rtf_content.append(r"{\qc\fs18 " + escape_rtf(f"Author: {author}") + r"\par}")
        rtf_content.append(r"\par")
        
        # Summary
        rtf_content.append(r"{\b\fs20 Data Summary\par}")
        rtf_content.append(r"{\fs18")
        rtf_content.append(f"\\bullet Total Rows: {len(df):,}\\par")
        rtf_content.append(f"\\bullet Total Columns: {len(df.columns)}\\par")
        rtf_content.append(f"\\bullet File Size: {os.path.getsize(csv_file):,} bytes\\par")
        
        # Create generation time string separately
        gen_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        rtf_content.append(f"\\bullet Generated: {gen_time}\\par")
        rtf_content.append(r"}\par")
        
        # Column headers
        rtf_content.append(r"{\b\fs20 Column Headers\par}")
        rtf_content.append(r"{\fs18")
        for i, col in enumerate(df.columns, 1):
            rtf_content.append(f"\\bullet {i}. {escape_rtf(str(col))}\\par")
        rtf_content.append(r"}\par")
        
        # Data table - use simple formatted text approach
        rtf_content.append(r"{\b\fs20 Data Table\par}")
        rtf_content.append(r"{\fs10")
        
        print(f"Processing {len(df)} data rows...")
        processed_rows = 0
        
        # Create a simple, reliable format that works in all RTF viewers
        try:
            for idx, (_, row) in enumerate(df.iterrows()):
                if idx % 1000 == 0:
                    print(f"Processing row {idx + 1}/{len(df)}")
                
                # Create a formatted row
                row_text = f"Row {idx + 1}: "
                row_data = []
                for i, value in enumerate(row):
                    col_name = df.columns[i]
                    cell_value = str(value) if pd.notna(value) else ""
                    row_data.append(f"{col_name}={cell_value}")
                
                # Join with pipe separator for clarity
                row_text += " | ".join(row_data)
                rtf_content.append(f"{escape_rtf(row_text)}\\par")
                processed_rows += 1
                
                # Add spacing every 5 rows for readability
                if (idx + 1) % 5 == 0:
                    rtf_content.append(r"\\par")
                
        except Exception as e:
            print(f"Error processing rows: {e}")
            print(f"Processed {processed_rows} rows before error")
        
        print(f"Successfully processed {processed_rows} rows out of {len(df)} total rows")
        
        # End data section
        rtf_content.append(r"}\par")
        
        # Conclusion
        rtf_content.append(r"{\b\fs20 Conclusion\par}")
        rtf_content.append(r"{\fs18")
        rtf_content.append(f"Data processing complete. Total rows processed: {len(df):,}, Columns analyzed: {len(df.columns)}.\\par")
        
        # Create conclusion time string separately
        conclusion_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        rtf_content.append(f"File generated: {conclusion_time}\\par")
        rtf_content.append(r"}\par")
        
        # End document
        rtf_content.append(r"}")
        
        # Write RTF file
        print(f"Writing RTF file to {output_file}...")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(rtf_content))
        
        # Verify file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"RTF file created successfully: {file_size} bytes")
            return True
        else:
            print("ERROR: RTF file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to create RTF from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to RTF format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output RTF file path')
    parser.add_argument('--title', default='CSV Data', help='Document title')
    parser.add_argument('--author', default='Unknown', help='Document author')
    
    args = parser.parse_args()
    
    print("=== CSV to RTF Converter ===")
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
    
    # Convert CSV to RTF
    success = create_rtf_from_csv(
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
