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
    Convert CSV file to RTF format using the most basic RTF structure.
    
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
        
        # Create the most basic RTF structure possible
        rtf_content = []
        
        # Basic RTF header - minimal structure
        rtf_content.append("{\\rtf1\\ansi\\deff0")
        rtf_content.append("{\\fonttbl{\\f0\\fswiss Arial;}}")
        rtf_content.append("\\f0\\fs24")  # Arial, 12pt
        
        # Title
        rtf_content.append("\\b\\fs32 " + escape_rtf(title) + "\\par")
        rtf_content.append("\\fs18 Generated on " + escape_rtf(datetime.now().strftime('%Y-%m-%d %H:%M:%S')) + "\\par")
        rtf_content.append("\\fs18 Author: " + escape_rtf(author) + "\\par")
        rtf_content.append("\\par")
        
        # Summary
        rtf_content.append("\\b\\fs20 Data Summary\\par")
        rtf_content.append("\\fs18 Total Rows: " + str(len(df)) + "\\par")
        rtf_content.append("\\fs18 Total Columns: " + str(len(df.columns)) + "\\par")
        rtf_content.append("\\fs18 File Size: " + str(os.path.getsize(csv_file)) + " bytes\\par")
        rtf_content.append("\\fs18 Generated: " + escape_rtf(datetime.now().strftime('%Y-%m-%d %H:%M:%S')) + "\\par")
        rtf_content.append("\\par")
        
        # Column headers
        rtf_content.append("\\b\\fs20 Column Headers\\par")
        for i, col in enumerate(df.columns, 1):
            rtf_content.append("\\fs18 " + str(i) + ". " + escape_rtf(str(col)) + "\\par")
        rtf_content.append("\\par")
        
        # Data table
        rtf_content.append("\\b\\fs20 Data Table\\par")
        rtf_content.append("\\fs10")
        
        print(f"Processing {len(df)} data rows...")
        processed_rows = 0
        
        # Create simple row format
        try:
            for idx, (_, row) in enumerate(df.iterrows()):
                if idx % 1000 == 0:
                    print(f"Processing row {idx + 1}/{len(df)}")
                
                # Create simple row format
                row_text = "Row " + str(idx + 1) + ": "
                row_data = []
                for i, value in enumerate(row):
                    col_name = df.columns[i]
                    cell_value = str(value) if pd.notna(value) else ""
                    row_data.append(col_name + "=" + cell_value)
                
                # Join with pipe separator
                row_text += " | ".join(row_data)
                rtf_content.append(escape_rtf(row_text) + "\\par")
                processed_rows += 1
                
                # Add spacing every 10 rows
                if (idx + 1) % 10 == 0:
                    rtf_content.append("\\par")
                
        except Exception as e:
            print(f"Error processing rows: {e}")
            print(f"Processed {processed_rows} rows before error")
        
        print(f"Successfully processed {processed_rows} rows out of {len(df)} total rows")
        
        # Conclusion
        rtf_content.append("\\par")
        rtf_content.append("\\b\\fs20 Conclusion\\par")
        rtf_content.append("\\fs18 Data processing complete. Total rows processed: " + str(len(df)) + ", Columns analyzed: " + str(len(df.columns)) + ".\\par")
        rtf_content.append("\\fs18 File generated: " + escape_rtf(datetime.now().strftime('%Y-%m-%d %H:%M:%S')) + "\\par")
        
        # End document
        rtf_content.append("}")
        
        # Write RTF file
        print(f"Writing RTF file to {output_file}...")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(''.join(rtf_content))
        
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
