#!/usr/bin/env python3
"""
AVRO to JSON converter using Python + fastavro
Converts AVRO files to JSON format with various options
"""

import argparse
import os
import sys
import json
import traceback
import fastavro
import pandas as pd

def convert_avro_to_json(avro_file, output_file, orient='records', indent=2, date_format='iso'):
    """
    Convert AVRO file to JSON format using fastavro
    
    Args:
        avro_file (str): Path to input AVRO file
        output_file (str): Path to output JSON file
        orient (str): JSON orientation ('records', 'split', 'index', 'columns', 'values')
        indent (int): JSON indentation (None for compact, 2 for pretty)
        date_format (str): Date format ('iso', 'epoch')
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting AVRO to JSON conversion...")
    print(f"Input: {avro_file}")
    print(f"Output: {output_file}")
    print(f"Orient: {orient}")
    print(f"Indent: {indent}")
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
        
        # Read AVRO file using fastavro
        print("Reading AVRO file with fastavro...")
        
        try:
            records = []
            with open(avro_file, 'rb') as avro_file_handle:
                avro_reader = fastavro.reader(avro_file_handle)
                schema = avro_reader.schema
                print(f"AVRO schema: {schema}")
                
                for record in avro_reader:
                    records.append(record)
            
            print(f"AVRO loaded successfully using fastavro")
            
        except Exception as avro_error:
            print(f"ERROR: Failed to read AVRO file: {avro_error}")
            print(f"ERROR: This file is not a valid AVRO binary file")
            print(f"ERROR: Please upload a file in Apache AVRO binary format")
            print(f"ERROR: AVRO files should start with 'Obj\\x01' magic bytes")
            return False
        
        print(f"Records count: {len(records)}")
        
        if len(records) == 0:
            print("WARNING: AVRO file contains no records")
            # Create empty JSON file based on orient
            if orient == 'records':
                json_data = []
            elif orient == 'split':
                json_data = {"columns": [], "index": [], "data": []}
            else:
                json_data = {}
            
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=indent if indent > 0 else None, ensure_ascii=False)
            print("Empty JSON file created")
            return True
        
        # Convert to JSON based on orientation
        print(f"Converting to JSON with orient='{orient}'...")
        
        if orient == 'records':
            # Array of objects (most common)
            json_data = records
        elif orient == 'split':
            # Split into columns, index, and data
            if records:
                columns = list(records[0].keys())
                data = []
                for record in records:
                    data.append([record.get(col) for col in columns])
                json_data = {
                    "columns": columns,
                    "index": list(range(len(records))),
                    "data": data
                }
            else:
                json_data = {"columns": [], "index": [], "data": []}
        elif orient == 'index':
            # Dict of index -> {column -> value}
            json_data = {}
            for i, record in enumerate(records):
                json_data[i] = record
        elif orient == 'columns':
            # Dict of column -> {index -> value}
            if records:
                columns = list(records[0].keys())
                json_data = {}
                for col in columns:
                    json_data[col] = {}
                    for i, record in enumerate(records):
                        json_data[col][i] = record.get(col)
            else:
                json_data = {}
        elif orient == 'values':
            # Just the values array
            if records:
                columns = list(records[0].keys())
                json_data = []
                for record in records:
                    json_data.append([record.get(col) for col in columns])
            else:
                json_data = []
        else:
            print(f"ERROR: Unsupported orient: {orient}")
            return False
        
        # Handle date formatting if needed
        if date_format == 'epoch':
            json_data = convert_dates_to_epoch(json_data)
        
        # Write to JSON file
        print("Writing JSON file...")
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=indent if indent > 0 else None, ensure_ascii=False)
        
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
            
    except Exception as e:
        print(f"ERROR: Failed to convert AVRO to JSON: {e}")
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
    parser = argparse.ArgumentParser(description='Convert AVRO to JSON format')
    parser.add_argument('avro_file', help='Input AVRO file path')
    parser.add_argument('output_file', help='Output JSON file path')
    parser.add_argument('--orient', choices=['records', 'split', 'index', 'columns', 'values'], 
                        default='records', help='JSON orientation (default: records)')
    parser.add_argument('--indent', type=int, default=2,
                        help='JSON indentation (default: 2, use 0 for compact)')
    parser.add_argument('--date-format', choices=['iso', 'epoch'], default='iso',
                        help='Date format (default: iso)')
    
    args = parser.parse_args()
    
    print("=== AVRO to JSON Converter ===")
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
    
    # Validate indent parameter
    if args.indent < 0:
        print(f"ERROR: Indent must be non-negative, got: {args.indent}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert AVRO to JSON
    success = convert_avro_to_json(
        args.avro_file,
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
