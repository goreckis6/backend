#!/usr/bin/env python3
"""
AVRO to CSV converter using Python + fastavro + pandas
Converts AVRO files to CSV format with various options
"""

import argparse
import os
import sys
import traceback
import pandas as pd
import fastavro

def convert_avro_to_csv(avro_file, output_file, delimiter=',', encoding='utf-8', include_header=True):
    """
    Convert AVRO file to CSV format using fastavro and pandas
    
    Args:
        avro_file (str): Path to input AVRO file
        output_file (str): Path to output CSV file
        delimiter (str): CSV delimiter (default: comma)
        encoding (str): Output encoding (default: utf-8)
        include_header (bool): Whether to include header row (default: True)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting AVRO to CSV conversion...")
    print(f"Input: {avro_file}")
    print(f"Output: {output_file}")
    print(f"Delimiter: '{delimiter}'")
    print(f"Encoding: {encoding}")
    print(f"Include header: {include_header}")
    
    try:
        # Check if input file exists and is readable
        if not os.path.exists(avro_file):
            print(f"ERROR: Input AVRO file not found: {avro_file}")
            return False
            
        file_size = os.path.getsize(avro_file)
        print(f"AVRO file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: AVRO file is empty")
            return False
        
        # Read AVRO file using pandas
        print("Reading AVRO file with pandas...")
        
        try:
            # Use pandas to read AVRO file directly
            df = pd.read_parquet(avro_file, engine='pyarrow')
            print(f"AVRO loaded successfully using pandas/pyarrow")
        except Exception as pandas_error:
            print(f"Pandas/pyarrow failed: {pandas_error}")
            print("Trying with fastavro...")
            
            # Fallback to fastavro if pandas fails
            try:
                records = []
                with open(avro_file, 'rb') as avro_file_handle:
                    avro_reader = fastavro.reader(avro_file_handle)
                    schema = avro_reader.schema
                    print(f"AVRO schema: {schema}")
                    
                    for record in avro_reader:
                        records.append(record)
                
                if len(records) == 0:
                    print("WARNING: AVRO file contains no records")
                    # Create empty CSV file
                    with open(output_file, 'w', encoding=encoding) as f:
                        if include_header:
                            f.write("")
                    print("Empty CSV file created")
                    return True
                
                # Convert to DataFrame
                df = pd.DataFrame(records)
                print(f"AVRO loaded successfully using fastavro")
                
            except Exception as avro_error:
                print(f"ERROR: Failed to read AVRO file: {avro_error}")
                print(f"ERROR: This might not be a valid AVRO file or the file is corrupted")
                return False
        
        print(f"DataFrame shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"Records count: {len(df)}")
        
        if len(df) == 0:
            print("WARNING: AVRO file contains no records")
            with open(output_file, 'w', encoding=encoding) as f:
                if include_header:
                    f.write("")
            print("Empty CSV file created")
            return True
        
        print(f"First few rows:\n{df.head()}")
        
        # Handle nested objects/arrays by converting to JSON strings
        for col in df.columns:
            if df[col].dtype == 'object':
                sample_value = df[col].dropna().iloc[0] if not df[col].dropna().empty else None
                if sample_value is not None and isinstance(sample_value, (dict, list)):
                    print(f"Converting complex objects in column '{col}' to JSON strings")
                    df[col] = df[col].apply(lambda x: pd.io.json.dumps(x) if pd.notna(x) and isinstance(x, (dict, list)) else x)
        
        # Write to CSV
        print("Writing CSV file...")
        df.to_csv(
            output_file, 
            index=False, 
            sep=delimiter, 
            encoding=encoding,
            header=include_header
        )
        
        # Verify the output file
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"CSV file created successfully: {file_size} bytes")
            
            # Verify it can be read back
            try:
                test_df = pd.read_csv(output_file, sep=delimiter, encoding=encoding)
                print(f"Verified CSV file can be read back: {test_df.shape}")
                return True
            except Exception as verify_error:
                print(f"ERROR: Output file verification failed: {verify_error}")
                return False
        else:
            print("ERROR: CSV file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to convert AVRO to CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert AVRO to CSV format')
    parser.add_argument('avro_file', help='Input AVRO file path')
    parser.add_argument('output_file', help='Output CSV file path')
    parser.add_argument('--delimiter', default=',', help='CSV delimiter (default: comma)')
    parser.add_argument('--encoding', default='utf-8', help='Output encoding (default: utf-8)')
    parser.add_argument('--no-header', action='store_true', help='Exclude header row')
    
    args = parser.parse_args()
    
    print("=== AVRO to CSV Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.avro_file):
        print(f"ERROR: Input AVRO file not found: {args.avro_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        import pandas
        import fastavro
        print(f"pandas version: {pandas.__version__}")
        print(f"fastavro version: {fastavro.__version__}")
    except ImportError as e:
        print(f"ERROR: Required library not available: {e}")
        print("Please install required libraries: pip install pandas fastavro")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert AVRO to CSV
    success = convert_avro_to_csv(
        args.avro_file,
        args.output_file,
        delimiter=args.delimiter,
        encoding=args.encoding,
        include_header=not args.no_header
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()
