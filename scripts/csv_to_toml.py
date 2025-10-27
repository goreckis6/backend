#!/usr/bin/env python3
"""
CSV to TOML converter using Python + pandas + tomli-w
Converts CSV files to TOML format with various options
"""

import argparse
import os
import sys
import traceback
import pandas as pd
import tomli_w

def convert_csv_to_toml(csv_file, output_file, indent=2):
    """
    Convert CSV file to TOML format using pandas
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output TOML file
        indent (int): TOML indentation (2 or 4 spaces)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CSV to TOML conversion...", flush=True)
    print(f"Input: {csv_file}", flush=True)
    print(f"Output: {output_file}", flush=True)
    print(f"Indent: {indent}", flush=True)
    
    try:
        # Read CSV file
        print("Reading CSV file...", flush=True)
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        
        rows, cols = df.shape
        print(f"CSV loaded: {rows} rows × {cols} columns", flush=True)
        print(f"Columns: {', '.join(df.columns.tolist())}", flush=True)
        
        # Convert DataFrame to TOML structure
        # TOML structure: array of tables
        # [[data]]
        #   col1 = "value1"
        #   col2 = "value2"
        
        print("Converting to TOML structure...", flush=True)
        
        toml_data = {
            "data": []
        }
        
        # Convert each row to a dictionary
        for idx, row in df.iterrows():
            row_dict = {}
            for col in df.columns:
                value = row[col]
                
                # Handle NaN values
                if pd.isna(value):
                    row_dict[col] = None
                # Handle different data types
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    row_dict[col] = value
                elif isinstance(value, bool):
                    row_dict[col] = value
                else:
                    # Convert to string
                    row_dict[col] = str(value)
            
            toml_data["data"].append(row_dict)
        
        # Write to TOML file
        print(f"Writing TOML file...", flush=True)
        
        with open(output_file, 'wb') as f:
            tomli_w.dump(toml_data, f)
        
        # Verify output file
        if not os.path.exists(output_file):
            raise FileNotFoundError(f"TOML file was not created: {output_file}")
        
        file_size = os.path.getsize(output_file)
        print(f"TOML file created successfully: {file_size:,} bytes ({file_size / 1024:.2f} KB)", flush=True)
        
        return True
        
    except FileNotFoundError as e:
        print(f"ERROR: File not found: {e}", flush=True)
        return False
    except pd.errors.EmptyDataError:
        print("ERROR: CSV file is empty", flush=True)
        return False
    except pd.errors.ParserError as e:
        print(f"ERROR: CSV parsing error: {e}", flush=True)
        return False
    except Exception as e:
        print(f"ERROR: Conversion failed: {e}", flush=True)
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to TOML')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('toml_file', help='Output TOML file path')
    parser.add_argument('--indent', 
                        type=int,
                        choices=[2, 4],
                        default=2,
                        help='TOML indentation (2 or 4 spaces)')
    
    args = parser.parse_args()
    
    print("=== CSV to TOML Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    print(f"Pandas version: {pd.__version__}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}", flush=True)
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.toml_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert
    success = convert_csv_to_toml(args.csv_file, args.toml_file, args.indent)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()


