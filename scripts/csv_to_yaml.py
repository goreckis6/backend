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

def sanitize_key(key):
    """Sanitize key for YAML compatibility."""
    # YAML keys can contain most characters, but we sanitize for safety
    sanitized = str(key).strip()
    # Replace problematic characters
    sanitized = sanitized.replace(':', '_')
    sanitized = sanitized.replace('#', '_')
    return sanitized or 'key'

def escape_yaml_string(value):
    """Escape special characters in YAML strings."""
    if pd.isna(value):
        return 'null'
    
    s = str(value)
    
    # Check if string needs quoting
    needs_quoting = False
    
    # Check for special YAML values that need quoting
    if s.lower() in ['yes', 'no', 'true', 'false', 'on', 'off', 'null', '~']:
        needs_quoting = True
    
    # Check for special characters
    if any(c in s for c in [':', '{', '}', '[', ']', ',', '&', '*', '#', '?', '|', '-', '<', '>', '=', '!', '%', '@', '`']):
        needs_quoting = True
    
    # Check if starts with special characters
    if s and s[0] in [' ', '\t', '\n', '\r', '"', "'", '>', '|', '-', '?', ':']:
        needs_quoting = True
    
    # Check if it looks like a number but should be a string
    try:
        float(s)
        needs_quoting = True
    except ValueError:
        pass
    
    if needs_quoting or '\n' in s or '"' in s or "'" in s:
        # Use double quotes and escape
        s = s.replace('\\', '\\\\')
        s = s.replace('"', '\\"')
        s = s.replace('\n', '\\n')
        s = s.replace('\r', '\\r')
        s = s.replace('\t', '\\t')
        return f'"{s}"'
    
    return s

def convert_value_to_yaml(value):
    """Convert Python value to YAML representation."""
    if pd.isna(value):
        return 'null'
    
    if isinstance(value, bool):
        return 'true' if value else 'false'
    
    if isinstance(value, (int, float)):
        return str(value)
    
    # String value
    return escape_yaml_string(value)

def convert_csv_to_yaml(csv_file, yaml_file, structure='list', root_key='data', flow_style=False):
    """
    Convert CSV to YAML format.
    
    Args:
        csv_file (str): Path to input CSV file
        yaml_file (str): Path to output YAML file
        structure (str): Output structure ('list' or 'dict')
        root_key (str): Root key name for the data
        flow_style (bool): Use flow style (compact) instead of block style
    
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
        
        # Sanitize column names
        original_columns = df.columns.tolist()
        sanitized_columns = [sanitize_key(col) for col in original_columns]
        df.columns = sanitized_columns
        
        print(f"Columns: {', '.join(sanitized_columns[:10])}", flush=True)
        if len(sanitized_columns) > 10:
            print(f"... and {len(sanitized_columns) - 10} more columns", flush=True)
        
        # Start building YAML
        yaml_lines = []
        
        # Add header comment
        yaml_lines.append(f"# YAML generated from CSV file: {os.path.basename(csv_file)}")
        yaml_lines.append(f"# Generated: {pd.Timestamp.now()}")
        yaml_lines.append(f"# Rows: {rows}, Columns: {cols}")
        yaml_lines.append("")
        
        if structure == 'list':
            # List of dictionaries (most common for CSV data)
            yaml_lines.append(f"{root_key}:")
            
            for index, row in df.iterrows():
                yaml_lines.append(f"  - ")
                
                for col in sanitized_columns:
                    value = row[col]
                    yaml_value = convert_value_to_yaml(value)
                    yaml_lines.append(f"    {col}: {yaml_value}")
                
                # Progress logging for large files
                if (index + 1) % 1000 == 0:
                    print(f"Processed {index + 1} of {rows} rows...", flush=True)
        
        else:  # dict structure
            # Dictionary with indexed entries
            yaml_lines.append(f"{root_key}:")
            
            for index, row in df.iterrows():
                entry_key = f"entry_{index}"
                yaml_lines.append(f"  {entry_key}:")
                
                for col in sanitized_columns:
                    value = row[col]
                    yaml_value = convert_value_to_yaml(value)
                    yaml_lines.append(f"    {col}: {yaml_value}")
                
                # Progress logging
                if (index + 1) % 1000 == 0:
                    print(f"Processed {index + 1} of {rows} rows...", flush=True)
        
        # Write to YAML file
        print(f"Writing YAML file...", flush=True)
        with open(yaml_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(yaml_lines))
            f.write('\n')  # Add final newline
        
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
    
    args = parser.parse_args()
    
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
        args.flow_style
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()


