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
        
        # Read CSV file with robust error handling
        try:
            # Try reading with default settings first
            df = pd.read_csv(input_file)
        except pd.errors.ParserError as e:
            print(f"Default CSV parsing failed: {e}")
            print("Trying alternative parsing methods...")
            
            # Try to identify the problematic line
            try:
                with open(input_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    print(f"Total lines in file: {len(lines)}")
                    for i, line in enumerate(lines[:10], 1):  # Check first 10 lines
                        print(f"Line {i}: {repr(line[:100])}")  # Show first 100 chars
            except Exception as debug_error:
                print(f"Could not read file for debugging: {debug_error}")
            
            # Try different delimiters
            for delimiter in [',', ';', '\t', '|']:
                try:
                    print(f"Trying delimiter: '{delimiter}'")
                    df = pd.read_csv(input_file, delimiter=delimiter)
                    print(f"Success with delimiter: '{delimiter}'")
                    break
                except pd.errors.ParserError:
                    continue
            
            # If still failing, try with error handling
            try:
                df = pd.read_csv(input_file, on_bad_lines='skip', engine='python')
                print("Success with error handling (skipping bad lines)")
            except Exception as final_error:
                raise ValueError(f"Could not parse CSV file with any method. Last error: {final_error}")
        
        # Additional fallback for encoding issues
        if df.empty:
            try:
                print("Trying with different encoding...")
                df = pd.read_csv(input_file, encoding='latin-1')
            except:
                try:
                    df = pd.read_csv(input_file, encoding='utf-8-sig')
                except:
                    raise ValueError("Could not read CSV file with any encoding")
        
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
