#!/usr/bin/env python3
"""
CSV to AVRO converter using Python + pandas + fastavro
Converts CSV files to AVRO format with automatic schema detection
"""

import argparse
import os
import sys
import traceback
import pandas as pd
import fastavro
import json
from typing import Any, Dict, List, Union

def infer_avro_schema(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Infer AVRO schema from pandas DataFrame
    
    Args:
        df: pandas DataFrame
    
    Returns:
        Dict: AVRO schema
    """
    schema = {
        "type": "record",
        "name": "Record",
        "namespace": "com.example",
        "fields": []
    }
    
    for column in df.columns:
        field = {"name": column}
        
        # Get the first non-null value to determine type
        first_value = df[column].dropna().iloc[0] if not df[column].dropna().empty else None
        
        if df[column].dtype == 'object':
            if first_value is None:
                field["type"] = ["null", "string"]
            elif isinstance(first_value, bool):
                field["type"] = ["null", "boolean"]
            elif isinstance(first_value, (int, float)):
                # Check if it's actually a number stored as string
                try:
                    float(first_value)
                    field["type"] = ["null", "double"]
                except (ValueError, TypeError):
                    field["type"] = ["null", "string"]
            elif isinstance(first_value, str):
                # Check if it's a JSON string
                try:
                    json.loads(first_value)
                    field["type"] = ["null", "string"]  # Store as string, can be parsed later
                except (json.JSONDecodeError, TypeError):
                    field["type"] = ["null", "string"]
            else:
                field["type"] = ["null", "string"]
        elif df[column].dtype in ['int8', 'int16', 'int32']:
            field["type"] = ["null", "int"]
        elif df[column].dtype in ['int64']:
            field["type"] = ["null", "long"]
        elif df[column].dtype in ['float32', 'float64']:
            field["type"] = ["null", "double"]
        elif df[column].dtype == 'bool':
            field["type"] = ["null", "boolean"]
        elif 'datetime' in str(df[column].dtype):
            field["type"] = ["null", "string"]  # Convert datetime to ISO string
        else:
            field["type"] = ["null", "string"]  # Default to string
        
        schema["fields"].append(field)
    
    return schema

def convert_csv_to_avro(csv_file, output_file, encoding='utf-8-sig', delimiter=','):
    """
    Convert CSV file to AVRO format using pandas and fastavro
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output AVRO file
        encoding (str): Input CSV encoding (default: utf-8-sig)
        delimiter (str): CSV delimiter (default: comma)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CSV to AVRO conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Encoding: {encoding}")
    print(f"Delimiter: '{delimiter}'")
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file, encoding=encoding, sep=delimiter)
        
        print(f"CSV loaded successfully")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"First few rows:\n{df.head()}")
        
        if len(df) == 0:
            print("WARNING: CSV file contains no records")
            # Create empty AVRO file with minimal schema
            schema = {
                "type": "record",
                "name": "Record",
                "namespace": "com.example",
                "fields": []
            }
            with open(output_file, 'wb') as avro_file:
                fastavro.writer(avro_file, schema, [])
            print("Empty AVRO file created")
            return True
        
        # Convert datetime columns to strings for AVRO compatibility
        for col in df.columns:
            if 'datetime' in str(df[col].dtype):
                print(f"Converting datetime column '{col}' to ISO string format")
                df[col] = df[col].astype(str)
        
        # Infer AVRO schema
        print("Inferring AVRO schema...")
        schema = infer_avro_schema(df)
        print(f"Inferred schema: {json.dumps(schema, indent=2)}")
        
        # Convert DataFrame to records
        print("Converting DataFrame to records...")
        records = df.to_dict('records')
        
        # Clean up records for AVRO compatibility
        cleaned_records = []
        for record in records:
            cleaned_record = {}
            for key, value in record.items():
                if pd.isna(value):
                    cleaned_record[key] = None
                elif isinstance(value, (int, float)) and pd.isna(value):
                    cleaned_record[key] = None
                else:
                    cleaned_record[key] = value
            cleaned_records.append(cleaned_record)
        
        print(f"Converted {len(cleaned_records)} records")
        
        # Write AVRO file
        print("Writing AVRO file...")
        with open(output_file, 'wb') as avro_file:
            fastavro.writer(avro_file, schema, cleaned_records)
        
        # Verify the output file
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"AVRO file created successfully: {file_size} bytes")
            
            # Verify it can be read back
            try:
                with open(output_file, 'rb') as avro_file:
                    reader = fastavro.reader(avro_file)
                    read_schema = reader.schema
                    record_count = sum(1 for _ in reader)
                
                print(f"Verified AVRO file can be read back: {record_count} records")
                print(f"Schema matches: {read_schema == schema}")
                return True
            except Exception as verify_error:
                print(f"ERROR: Output file verification failed: {verify_error}")
                return False
        else:
            print("ERROR: AVRO file was not created")
            return False
            
    except pd.errors.EmptyDataError:
        print("ERROR: CSV file is empty")
        return False
    except pd.errors.ParserError as e:
        print(f"ERROR: CSV parsing error: {e}")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert CSV to AVRO: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to AVRO format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output AVRO file path')
    parser.add_argument('--encoding', default='utf-8-sig', help='CSV encoding (default: utf-8-sig)')
    parser.add_argument('--delimiter', default=',', help='CSV delimiter (default: comma)')
    
    args = parser.parse_args()
    
    print("=== CSV to AVRO Converter ===")
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
    
    # Convert CSV to AVRO
    success = convert_csv_to_avro(
        args.csv_file,
        args.output_file,
        encoding=args.encoding,
        delimiter=args.delimiter
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

