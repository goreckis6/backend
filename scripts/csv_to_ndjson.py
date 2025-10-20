#!/usr/bin/env python3
"""
CSV to NDJSON converter using Python + pandas
Converts CSV files to NDJSON (Newline Delimited JSON) format for streaming data processing.
"""

import argparse
import os
import sys
import pandas as pd
import json
import traceback

def convert_csv_to_ndjson(csv_file, output_file, include_headers=True):
    """
    Convert CSV file to NDJSON format using pandas.

    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output NDJSON file
        include_headers (bool): Whether to include headers in conversion

    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CSV to NDJSON conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Include headers: {include_headers}")

    try:
        # Read CSV file
        print("Reading CSV file with pandas...")
        df = pd.read_csv(csv_file)
        print(f"DataFrame shape: {df.shape}")
        print(f"DataFrame columns: {df.columns.tolist()}")

        # Convert DataFrame to NDJSON (one JSON object per line)
        print(f"Converting to NDJSON...")
        
        # Write NDJSON to output file
        with open(output_file, 'w', encoding='utf-8') as f:
            for _, row in df.iterrows():
                # Convert row to dictionary and then to JSON string
                json_line = json.dumps(row.to_dict(), ensure_ascii=False)
                f.write(json_line + '\n')

        # Verify the output file
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"NDJSON file created successfully: {file_size} bytes")
            
            # Count lines
            with open(output_file, 'r', encoding='utf-8') as f:
                line_count = sum(1 for _ in f)
            print(f"Total lines in NDJSON: {line_count}")
            
            return True
        else:
            print("ERROR: NDJSON file was not created")
            return False

    except FileNotFoundError:
        print(f"ERROR: Input CSV file not found: {csv_file}")
        return False
    except pd.errors.EmptyDataError:
        print(f"ERROR: Input CSV file is empty: {csv_file}")
        return False
    except pd.errors.ParserError as e:
        print(f"ERROR: CSV parsing error: {e}")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert CSV to NDJSON: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to NDJSON format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output NDJSON file path')
    parser.add_argument('--include-headers', type=str, default='true',
                        choices=['true', 'false'],
                        help='Include headers in conversion (default: true)')

    args = parser.parse_args()

    print("=== CSV to NDJSON Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")

    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)

    # Check required libraries
    try:
        import pandas as pd
        print(f"pandas version: {pd.__version__}")
    except ImportError as e:
        print(f"ERROR: pandas not available: {e}")
        print("Please install pandas: pip install pandas")
        sys.exit(1)

    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Convert CSV to NDJSON
    include_headers = args.include_headers.lower() == 'true'
    success = convert_csv_to_ndjson(
        args.csv_file,
        args.output_file,
        include_headers=include_headers
    )

    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()


