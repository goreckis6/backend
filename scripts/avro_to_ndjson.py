#!/usr/bin/env python3
"""
AVRO to NDJSON converter using Python + fastavro
Converts AVRO files to Newline Delimited JSON format
"""

import argparse
import os
import sys
import json
import traceback
import fastavro

def convert_avro_to_ndjson(avro_file, output_file, encoding='utf-8', date_format='iso'):
    """
    Convert AVRO file to NDJSON format using fastavro
    
    Args:
        avro_file (str): Path to input AVRO file
        output_file (str): Path to output NDJSON file
        encoding (str): Output encoding (default: utf-8)
        date_format (str): Date format ('iso', 'epoch')
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting AVRO to NDJSON conversion...")
    print(f"Input: {avro_file}")
    print(f"Output: {output_file}")
    print(f"Encoding: {encoding}")
    print(f"Date format: {date_format}")
    
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
        
        # Check file magic bytes to verify it's an AVRO file
        with open(avro_file, 'rb') as f:
            magic_bytes = f.read(4)
            print(f"File magic bytes (first 4 bytes): {magic_bytes.hex()}")
            print(f"Expected AVRO magic bytes: 4f626a01 (Obj\\x01)")
            
            # AVRO files should start with 'Obj\x01'
            if magic_bytes != b'Obj\x01':
                print(f"ERROR: File does not have AVRO magic bytes")
                print(f"This file is not a valid AVRO binary file")
                
                # Try to read first 100 bytes as text to see what it is
                f.seek(0)
                first_bytes = f.read(min(100, file_size))
                try:
                    first_text = first_bytes.decode('utf-8', errors='ignore')
                    print(f"First 100 bytes as text: {first_text[:100]}")
                except:
                    print(f"Could not decode first bytes as text")
                
                return False
        
        # Read AVRO file and write NDJSON
        print("Reading AVRO file and converting to NDJSON...")
        record_count = 0
        
        try:
            with open(avro_file, 'rb') as avro_file_handle:
                avro_reader = fastavro.reader(avro_file_handle)
                schema = avro_reader.schema
                print(f"AVRO schema: {schema}")
                
                with open(output_file, 'w', encoding=encoding) as ndjson_file:
                    # Reset file position to beginning
                    avro_file_handle.seek(0)
                    avro_reader = fastavro.reader(avro_file_handle)
                    
                    # Process each record and write as NDJSON line
                    for record in avro_reader:
                        # Handle date formatting if needed
                        if date_format == 'epoch':
                            record = convert_dates_to_epoch(record)
                        
                        # Write as JSON line
                        json_line = json.dumps(record, ensure_ascii=False)
                        ndjson_file.write(json_line + '\n')
                        record_count += 1
                        
        except Exception as avro_error:
            print(f"ERROR: Failed to read AVRO file: {avro_error}")
            print(f"ERROR: This might not be a valid AVRO file or the file is corrupted")
            return False
        
        print(f"AVRO to NDJSON conversion completed")
        print(f"Records processed: {record_count}")
        
        if record_count == 0:
            print("WARNING: AVRO file contained no records")
            # Create empty NDJSON file
            with open(output_file, 'w', encoding=encoding) as f:
                pass  # Empty file
            print("Empty NDJSON file created")
            return True
        
        # Verify the output file
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"NDJSON file created successfully: {file_size} bytes")
            
            # Verify it can be read back (check first few lines)
            try:
                with open(output_file, 'r', encoding=encoding) as f:
                    lines_read = 0
                    for line in f:
                        if line.strip():  # Skip empty lines
                            json.loads(line.strip())
                            lines_read += 1
                            if lines_read >= 5:  # Check first 5 non-empty lines
                                break
                print(f"Verified NDJSON file format (checked {lines_read} lines)")
                return True
            except Exception as verify_error:
                print(f"ERROR: Output file verification failed: {verify_error}")
                return False
        else:
            print("ERROR: NDJSON file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to convert AVRO to NDJSON: {e}")
        traceback.print_exc()
        return False

def convert_dates_to_epoch(data):
    """
    Recursively convert datetime objects to epoch timestamps
    """
    import datetime
    
    if isinstance(data, datetime.datetime):
        return int(data.timestamp())
    elif isinstance(data, datetime.date):
        return int(datetime.datetime.combine(data, datetime.time()).timestamp())
    elif isinstance(data, list):
        return [convert_dates_to_epoch(item) for item in data]
    elif isinstance(data, dict):
        return {key: convert_dates_to_epoch(value) for key, value in data.items()}
    else:
        return data

def main():
    parser = argparse.ArgumentParser(description='Convert AVRO to NDJSON format')
    parser.add_argument('avro_file', help='Input AVRO file path')
    parser.add_argument('output_file', help='Output NDJSON file path')
    parser.add_argument('--encoding', default='utf-8', help='Output encoding (default: utf-8)')
    parser.add_argument('--date-format', choices=['iso', 'epoch'], default='iso',
                        help='Date format (default: iso)')
    
    args = parser.parse_args()
    
    print("=== AVRO to NDJSON Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.avro_file):
        print(f"ERROR: Input AVRO file not found: {args.avro_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        import fastavro
        print(f"fastavro version: {fastavro.__version__}")
    except ImportError as e:
        print(f"ERROR: fastavro not available: {e}")
        print("Please install fastavro: pip install fastavro")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert AVRO to NDJSON
    success = convert_avro_to_ndjson(
        args.avro_file,
        args.output_file,
        encoding=args.encoding,
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
