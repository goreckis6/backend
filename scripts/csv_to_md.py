#!/usr/bin/env python3
"""
CSV to Markdown Converter
Converts CSV files to Markdown table format with various options.
"""

import pandas as pd
import argparse
import sys
import os
from pathlib import Path

def convert_csv_to_markdown(
    input_file: str,
    output_file: str,
    include_headers: bool = True,
    table_alignment: str = 'left',
    chunk_size: int = 1000
):
    """
    Convert CSV file to Markdown table format.
    
    Args:
        input_file: Path to input CSV file
        output_file: Path to output Markdown file
        include_headers: Whether to include column headers
        table_alignment: Table alignment ('left', 'center', 'right')
        chunk_size: Number of rows to process at a time for large files
    """
    try:
        print(f"Converting CSV to Markdown: {input_file} -> {output_file}")
        
        # Read CSV file
        try:
            df = pd.read_csv(input_file)
        except pd.errors.ParserError as e:
            raise ValueError(f"CSV file appears to be corrupted or malformed. Please check the file format and try again. Error: {str(e)}")
        except Exception as e:
            raise ValueError(f"Could not read CSV file. The file may be corrupted or in an unsupported format. Error: {str(e)}")
        
        if df.empty:
            raise ValueError("CSV file is empty")
        
        print(f"CSV loaded: {len(df)} rows, {len(df.columns)} columns")
        
        # Handle NaN values
        df = df.fillna('')
        
        # Convert to Markdown table
        markdown_content = []
        
        if include_headers:
            # Add header row
            headers = df.columns.tolist()
            markdown_content.append('| ' + ' | '.join(headers) + ' |')
            
            # Add separator row based on alignment
            if table_alignment == 'left':
                separator = '| ' + ' | '.join(['---'] * len(headers)) + ' |'
            elif table_alignment == 'center':
                separator = '| ' + ' | '.join([':---:'] * len(headers)) + ' |'
            else:  # right
                separator = '| ' + ' | '.join(['---:'] * len(headers)) + ' |'
            
            markdown_content.append(separator)
        
        # Add data rows in chunks for large files
        total_rows = len(df)
        for start_idx in range(0, total_rows, chunk_size):
            end_idx = min(start_idx + chunk_size, total_rows)
            chunk_df = df.iloc[start_idx:end_idx]
            
            for _, row in chunk_df.iterrows():
                # Escape pipe characters in data
                row_data = [str(cell).replace('|', '\\|') for cell in row]
                markdown_content.append('| ' + ' | '.join(row_data) + ' |')
            
            print(f"Processed rows {start_idx + 1}-{end_idx} of {total_rows}")
        
        # Write to output file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_content))
        
        print(f"Markdown conversion completed: {len(markdown_content)} lines written")
        return True
        
    except Exception as e:
        print(f"Error converting CSV to Markdown: {str(e)}", file=sys.stderr)
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to Markdown table format')
    parser.add_argument('input_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output Markdown file path')
    parser.add_argument('--include-headers', action='store_true', default=True,
                       help='Include column headers in the table')
    parser.add_argument('--no-headers', dest='include_headers', action='store_false',
                       help='Exclude column headers from the table')
    parser.add_argument('--table-alignment', choices=['left', 'center', 'right'], 
                       default='left', help='Table alignment style')
    parser.add_argument('--chunk-size', type=int, default=1000,
                       help='Number of rows to process at a time for large files')
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.input_file):
        print(f"Error: Input file does not exist: {args.input_file}", file=sys.stderr)
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert CSV to Markdown
    success = convert_csv_to_markdown(
        input_file=args.input_file,
        output_file=args.output_file,
        include_headers=args.include_headers,
        table_alignment=args.table_alignment,
        chunk_size=args.chunk_size
    )
    
    if success:
        print("Conversion completed successfully")
        sys.exit(0)
    else:
        print("Conversion failed", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()
