#!/usr/bin/env python3
"""
CSV to Parquet converter.
Converts CSV files to Apache Parquet columnar storage format.
"""

import argparse
import os
import sys
import pandas as pd
import traceback

def convert_csv_to_parquet(csv_file, parquet_file, compression='snappy'):
    """
    Convert CSV to Parquet using pandas and pyarrow.
    
    Args:
        csv_file (str): Path to input CSV file
        parquet_file (str): Path to output Parquet file
        compression (str): Compression codec (snappy, gzip, brotli, none)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting CSV to Parquet...", flush=True)
    print(f"Input: {csv_file}", flush=True)
    print(f"Output: {parquet_file}", flush=True)
    print(f"Compression: {compression}", flush=True)
    
    try:
        # Read CSV file
        print("Reading CSV file...", flush=True)
        df = pd.read_csv(csv_file, encoding='utf-8', low_memory=False)
        
        rows, cols = df.shape
        print(f"CSV loaded: {rows} rows × {cols} columns", flush=True)
        
        # Display column info
        print(f"Columns: {', '.join(df.columns.tolist()[:10])}", flush=True)
        if len(df.columns) > 10:
            print(f"... and {len(df.columns) - 10} more columns", flush=True)
        
        # Convert to Parquet
        print(f"Writing Parquet file with {compression} compression...", flush=True)
        df.to_parquet(
            parquet_file,
            engine='pyarrow',
            compression=compression,
            index=False
        )
        
        # Verify output file
        if not os.path.exists(parquet_file):
            raise FileNotFoundError(f"Parquet file was not created: {parquet_file}")
        
        file_size = os.path.getsize(parquet_file)
        csv_size = os.path.getsize(csv_file)
        compression_ratio = (1 - file_size / csv_size) * 100 if csv_size > 0 else 0
        
        print(f"Parquet file created successfully!", flush=True)
        print(f"CSV size: {csv_size:,} bytes ({csv_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Parquet size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Compression ratio: {compression_ratio:.1f}%", flush=True)
        
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
    parser = argparse.ArgumentParser(description='Convert CSV to Parquet')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('parquet_file', help='Output Parquet file path')
    parser.add_argument('--compression', 
                       choices=['snappy', 'gzip', 'brotli', 'none'],
                       default='snappy',
                       help='Compression codec (default: snappy)')
    
    args = parser.parse_args()
    
    print("=== CSV to Parquet Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    print(f"Pandas version: {pd.__version__}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}", flush=True)
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.parquet_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert
    success = convert_csv_to_parquet(args.csv_file, args.parquet_file, args.compression)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()

