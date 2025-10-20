#!/usr/bin/env python3
"""
CSV to JSON converter using Python + pandas
Converts CSV files to JSON format with various options
"""

import argparse
import os
import sys
import json
import traceback
import pandas as pd

def convert_csv_to_json(csv_file, output_file, orient='records', indent=2, date_format='iso'):
    """
    Convert CSV file to JSON format using pandas
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output JSON file
        orient (str): JSON orientation ('records', 'split', 'index', 'columns', 'values')
        indent (int): JSON indentation (None for compact, 2 for pretty)
        date_format (str): Date format ('iso', 'epoch')
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CSV to JSON conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Orient: {orient}")
    print(f"Indent: {indent}")
    print(f"Date format: {date_format}")
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        
        print(f"CSV loaded successfully")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"First few rows:\n{df.head()}")
        
        # Handle date columns if specified
        if date_format == 'epoch':
            # Convert datetime columns to epoch timestamps
            for col in df.select_dtypes(include=['datetime64']).columns:
                df[col] = df[col].astype('int64') // 10**9
        
        # Convert to JSON
        print(f"Converting to JSON with orient='{orient}'...")
        
        if orient == 'records':
            # Array of objects (most common)
            json_data = df.to_json(orient='records', date_format=date_format, indent=indent)
        elif orient == 'split':
            # Split into columns, index, and data
            json_data = df.to_json(orient='split', date_format=date_format, indent=indent)
        elif orient == 'index':
            # Dict of index -> {column -> value}
            json_data = df.to_json(orient='index', date_format=date_format, indent=indent)
        elif orient == 'columns':
            # Dict of column -> {index -> value}
            json_data = df.to_json(orient='columns', date_format=date_format, indent=indent)
        elif orient == 'values':
            # Just the values array
            json_data = df.to_json(orient='values', date_format=date_format, indent=indent)
        else:
            print(f"ERROR: Unsupported orient: {orient}")
            return False
        
        # Write to output file
        print("Writing JSON file...")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(json_data)
        
        # Verify the output file
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"JSON file created successfully: {file_size} bytes")
            
            # Verify it's valid JSON
            try:
                with open(output_file, 'r', encoding='utf-8') as f:
                    json.load(f)
                print("Verified JSON file is valid")
                return True
            except Exception as verify_error:
                print(f"ERROR: Output file is not valid JSON: {verify_error}")
                return False
        else:
            print("ERROR: JSON file was not created")
            return False
            
    except pd.errors.EmptyDataError:
        print("ERROR: CSV file is empty")
        return False
    except pd.errors.ParserError as e:
        print(f"ERROR: CSV parsing error: {e}")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert CSV to JSON: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to JSON format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output JSON file path')
    parser.add_argument('--orient', choices=['records', 'split', 'index', 'columns', 'values'], 
                        default='records', help='JSON orientation (default: records)')
    parser.add_argument('--indent', type=int, default=2,
                        help='JSON indentation (default: 2, use 0 for compact)')
    parser.add_argument('--date-format', choices=['iso', 'epoch'], default='iso',
                        help='Date format (default: iso)')
    
    args = parser.parse_args()
    
    print("=== CSV to JSON Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        import pandas
        print(f"pandas version: {pandas.__version__}")
    except ImportError as e:
        print(f"ERROR: pandas not available: {e}")
        print("Please install pandas: pip install pandas")
        sys.exit(1)
    
    # Validate indent parameter
    if args.indent < 0:
        print(f"ERROR: Indent must be non-negative, got: {args.indent}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to JSON
    success = convert_csv_to_json(
        args.csv_file,
        args.output_file,
        orient=args.orient,
        indent=args.indent if args.indent > 0 else None,
        date_format=args.date_format
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()


