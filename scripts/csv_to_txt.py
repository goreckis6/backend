#!/usr/bin/env python3
"""
CSV to TXT Converter
Converts CSV files to plain text format using pandas and custom text generation.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback

def create_txt_from_csv(csv_file, output_file, title="CSV Data", author="Unknown"):
    """
    Convert CSV file to TXT format using a clean, readable text structure.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output TXT file
        title (str): Document title
        author (str): Document author
    """
    print(f"Starting CSV to TXT conversion...")
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
        
        # Create text content
        txt_content = []
        
        # Header
        txt_content.append("=" * 80)
        txt_content.append(f" {title}")
        txt_content.append("=" * 80)
        txt_content.append("")
        txt_content.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        txt_content.append(f"Author: {author}")
        txt_content.append("")
        
        # Summary
        txt_content.append("DATA SUMMARY")
        txt_content.append("-" * 40)
        txt_content.append(f"Total Rows: {len(df):,}")
        txt_content.append(f"Total Columns: {len(df.columns)}")
        txt_content.append(f"File Size: {os.path.getsize(csv_file):,} bytes")
        txt_content.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        txt_content.append("")
        
        # Column headers
        txt_content.append("COLUMN HEADERS")
        txt_content.append("-" * 40)
        for i, col in enumerate(df.columns, 1):
            txt_content.append(f"{i:2d}. {col}")
        txt_content.append("")
        
        # Data table
        txt_content.append("DATA TABLE")
        txt_content.append("-" * 40)
        
        print(f"Processing {len(df)} data rows...")
        processed_rows = 0
        
        # Create clean text format
        try:
            for idx, (_, row) in enumerate(df.iterrows()):
                if idx % 1000 == 0:
                    print(f"Processing row {idx + 1}/{len(df)}")
                
                # Create clean row format
                row_text = f"Row {idx + 1:6d}: "
                row_data = []
                for i, value in enumerate(row):
                    col_name = df.columns[i]
                    cell_value = str(value) if pd.notna(value) else ""
                    row_data.append(f"{col_name}={cell_value}")
                
                # Join with pipe separator
                row_text += " | ".join(row_data)
                txt_content.append(row_text)
                processed_rows += 1
                
                # Add spacing every 20 rows for readability
                if (idx + 1) % 20 == 0:
                    txt_content.append("")
                
        except Exception as e:
            print(f"Error processing rows: {e}")
            print(f"Processed {processed_rows} rows before error")
        
        print(f"Successfully processed {processed_rows} rows out of {len(df)} total rows")
        
        # Conclusion
        txt_content.append("")
        txt_content.append("CONCLUSION")
        txt_content.append("-" * 40)
        txt_content.append(f"Data processing complete. Total rows processed: {len(df):,}, Columns analyzed: {len(df.columns)}.")
        txt_content.append(f"File generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        txt_content.append("")
        txt_content.append("=" * 80)
        
        # Write TXT file
        print(f"Writing TXT file to {output_file}...")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(txt_content))
        
        # Verify file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"TXT file created successfully: {file_size} bytes")
            return True
        else:
            print("ERROR: TXT file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to create TXT from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to TXT format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output TXT file path')
    parser.add_argument('--title', default='CSV Data', help='Document title')
    parser.add_argument('--author', default='Unknown', help='Document author')
    
    args = parser.parse_args()
    
    print("=== CSV to TXT Converter ===")
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
    
    # Convert CSV to TXT
    success = create_txt_from_csv(
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

