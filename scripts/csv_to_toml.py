#!/usr/bin/env python3
"""
CSV to TOML converter.
Converts CSV files to TOML (Tom's Obvious Minimal Language) format.
"""

import argparse
import os
import sys
import pandas as pd
import traceback

def escape_toml_string(value):
    """Escape special characters in TOML strings."""
    if pd.isna(value):
        return '""'
    
    s = str(value)
    # Escape backslashes and quotes
    s = s.replace('\\', '\\\\')
    s = s.replace('"', '\\"')
    s = s.replace('\n', '\\n')
    s = s.replace('\r', '\\r')
    s = s.replace('\t', '\\t')
    
    return f'"{s}"'

def sanitize_key(key):
    """Sanitize key for TOML compatibility."""
    # Replace spaces and special chars with underscores
    sanitized = ''.join(c if c.isalnum() or c == '_' else '_' for c in str(key))
    # Remove leading/trailing underscores
    sanitized = sanitized.strip('_')
    # Ensure it doesn't start with a number
    if sanitized and sanitized[0].isdigit():
        sanitized = 'key_' + sanitized
    return sanitized.lower() or 'key'

def convert_value_to_toml(value):
    """Convert Python value to TOML representation."""
    if pd.isna(value):
        return '""'  # TOML doesn't have null, use empty string
    
    if isinstance(value, bool):
        return 'true' if value else 'false'
    
    if isinstance(value, (int, float)):
        return str(value)
    
    # String value
    return escape_toml_string(value)

def convert_csv_to_toml(csv_file, toml_file, structure='array', section_name='data'):
    """
    Convert CSV to TOML format.
    
    Args:
        csv_file (str): Path to input CSV file
        toml_file (str): Path to output TOML file
        structure (str): Output structure ('array' or 'table')
        section_name (str): Name of the TOML section
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting CSV to TOML...", flush=True)
    print(f"Input: {csv_file}", flush=True)
    print(f"Output: {toml_file}", flush=True)
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
        
        # Start building TOML
        toml_lines = []
        
        # Add header comment
        toml_lines.append(f"# TOML generated from CSV file: {os.path.basename(csv_file)}")
        toml_lines.append(f"# Generated: {pd.Timestamp.now()}")
        toml_lines.append(f"# Rows: {rows}, Columns: {cols}")
        toml_lines.append("")
        
        if structure == 'array':
            # Array of tables structure (most common for CSV data)
            for index, row in df.iterrows():
                toml_lines.append(f"[[{section_name}]]")
                
                for col in sanitized_columns:
                    value = row[col]
                    toml_value = convert_value_to_toml(value)
                    toml_lines.append(f"{col} = {toml_value}")
                
                toml_lines.append("")  # Blank line between records
                
                # Progress logging for large files
                if (index + 1) % 1000 == 0:
                    print(f"Processed {index + 1} of {rows} rows...", flush=True)
        
        else:  # table structure
            # Single table with named entries
            for index, row in df.iterrows():
                entry_name = f"{section_name}.entry_{index}"
                toml_lines.append(f"[{entry_name}]")
                
                for col in sanitized_columns:
                    value = row[col]
                    toml_value = convert_value_to_toml(value)
                    toml_lines.append(f"{col} = {toml_value}")
                
                toml_lines.append("")
                
                # Progress logging
                if (index + 1) % 1000 == 0:
                    print(f"Processed {index + 1} of {rows} rows...", flush=True)
        
        # Write to TOML file
        print(f"Writing TOML file...", flush=True)
        with open(toml_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(toml_lines))
        
        # Verify output file
        if not os.path.exists(toml_file):
            raise FileNotFoundError(f"TOML file was not created: {toml_file}")
        
        file_size = os.path.getsize(toml_file)
        csv_size = os.path.getsize(csv_file)
        
        print(f"TOML file created successfully!", flush=True)
        print(f"CSV size: {csv_size:,} bytes ({csv_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"TOML size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
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
    parser = argparse.ArgumentParser(description='Convert CSV to TOML')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('toml_file', help='Output TOML file path')
    parser.add_argument('--structure', 
                       choices=['array', 'table'],
                       default='array',
                       help='TOML structure type (default: array)')
    parser.add_argument('--section-name', 
                       default='data',
                       help='TOML section name (default: data)')
    
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
    success = convert_csv_to_toml(
        args.csv_file, 
        args.toml_file, 
        args.structure,
        args.section_name
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()


