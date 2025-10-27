#!/usr/bin/env python3
"""
CSV to YAML converter.
Converts CSV files to YAML (YAML Ain't Markup Language) format.
"""

import argparse
import os
import sys
import pandas as pd
import traceback
import yaml

def convert_csv_to_yaml(csv_file, yaml_file, structure='list', root_key='data', flow_style=False, indent=2, allow_unicode=True):
    """
    Convert CSV to YAML format.
    
    Args:
        csv_file (str): Path to input CSV file
        yaml_file (str): Path to output YAML file
        structure (str): Output structure ('list' or 'dict')
        root_key (str): Root key name for the data
        flow_style (bool): Use flow style (compact) instead of block style
        indent (int): YAML indentation (2, 4, or 6 spaces)
        allow_unicode (bool): Allow Unicode characters in output
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting CSV to YAML...", flush=True)
    print(f"Input: {csv_file}", flush=True)
    print(f"Output: {yaml_file}", flush=True)
    print(f"Structure: {structure}", flush=True)
    
    try:
        # Read CSV file
        print("Reading CSV file...", flush=True)
        df = pd.read_csv(csv_file, encoding='utf-8', low_memory=False)
        
        rows, cols = df.shape
        print(f"CSV loaded: {rows} rows × {cols} columns", flush=True)
        
        if rows == 0:
            raise ValueError("CSV file is empty (no data rows)")
        
        # Keep original column names (PyYAML handles sanitization)
        original_columns = df.columns.tolist()
        
        print(f"Columns: {', '.join(original_columns[:10])}", flush=True)
        if len(original_columns) > 10:
            print(f"... and {len(original_columns) - 10} more columns", flush=True)
        
        # Build data structure
        yaml_data = {
            'metadata': {
                'source_file': os.path.basename(csv_file),
                'generated': pd.Timestamp.now().isoformat(),
                'rows': int(rows),
                'columns': int(cols)
            }
        }
        
        if structure == 'list':
            # List of dictionaries (most common for CSV data)
            data_list = []
            for index, row in df.iterrows():
                row_dict = {}
                for col in original_columns:
                    value = row[col]
                    if pd.isna(value):
                        row_dict[col] = None
                    elif isinstance(value, (int, float)) and not isinstance(value, bool):
                        row_dict[col] = value
                    elif isinstance(value, bool):
                        row_dict[col] = value
                    else:
                        row_dict[col] = str(value)
                data_list.append(row_dict)
                
                # Progress logging for large files
                if (index + 1) % 1000 == 0:
                    print(f"Processed {index + 1} of {rows} rows...", flush=True)
            
            yaml_data[root_key] = data_list
        else:  # dict structure
            # Dictionary with indexed entries
            data_dict = {}
            for index, row in df.iterrows():
                entry_key = f"entry_{index}"
                row_dict = {}
                for col in original_columns:
                    value = row[col]
                    if pd.isna(value):
                        row_dict[col] = None
                    elif isinstance(value, (int, float)) and not isinstance(value, bool):
                        row_dict[col] = value
                    elif isinstance(value, bool):
                        row_dict[col] = value
                    else:
                        row_dict[col] = str(value)
                data_dict[entry_key] = row_dict
                
                # Progress logging
                if (index + 1) % 1000 == 0:
                    print(f"Processed {index + 1} of {rows} rows...", flush=True)
            
            yaml_data[root_key] = data_dict
        
        # Write to YAML file using PyYAML
        print(f"Writing YAML file...", flush=True)
        with open(yaml_file, 'w', encoding='utf-8' if allow_unicode else 'ascii', errors='ignore' if not allow_unicode else 'strict') as f:
            yaml.dump(
                yaml_data, 
                f, 
                default_flow_style=flow_style,
                indent=int(indent),
                allow_unicode=allow_unicode,
                sort_keys=False
            )
        
        # Verify output file
        if not os.path.exists(yaml_file):
            raise FileNotFoundError(f"YAML file was not created: {yaml_file}")
        
        file_size = os.path.getsize(yaml_file)
        csv_size = os.path.getsize(csv_file)
        
        print(f"YAML file created successfully!", flush=True)
        print(f"CSV size: {csv_size:,} bytes ({csv_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"YAML size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Total entries: {rows}", flush=True)
        
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
    parser = argparse.ArgumentParser(description='Convert CSV to YAML')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('yaml_file', help='Output YAML file path')
    parser.add_argument('--structure', 
                       choices=['list', 'dict'],
                       default='list',
                       help='YAML structure type (default: list)')
    parser.add_argument('--root-key', 
                       default='data',
                       help='Root key name (default: data)')
    parser.add_argument('--flow-style',
                       action='store_true',
                       help='Use flow style (compact) instead of block style')
    parser.add_argument('--indent',
                       type=int,
                       default=2,
                       choices=[2, 4, 6],
                       help='YAML indentation (2, 4, or 6 spaces)')
    parser.add_argument('--default-flow-style',
                       type=str,
                       default='false',
                       help='Use flow style (true/false)')
    parser.add_argument('--allow-unicode',
                       type=str,
                       default='true',
                       help='Allow Unicode characters (true/false)')
    
    args = parser.parse_args()
    
    # Convert string booleans to actual booleans
    flow_style = args.default_flow_style.lower() == 'true'
    allow_unicode = args.allow_unicode.lower() == 'true'
    
    print("=== CSV to YAML Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    print(f"Pandas version: {pd.__version__}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}", flush=True)
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.yaml_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert
    success = convert_csv_to_yaml(
        args.csv_file, 
        args.yaml_file, 
        args.structure,
        args.root_key,
        flow_style,
        args.indent,
        allow_unicode
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()


