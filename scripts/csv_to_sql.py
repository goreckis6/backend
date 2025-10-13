#!/usr/bin/env python3
"""
CSV to SQL converter.
Converts CSV files to SQL INSERT statements for database import.
"""

import argparse
import os
import sys
import pandas as pd
import traceback

def sanitize_column_name(name):
    """Sanitize column name for SQL compatibility."""
    # Replace spaces and special chars with underscores
    sanitized = ''.join(c if c.isalnum() else '_' for c in str(name))
    # Remove leading/trailing underscores
    sanitized = sanitized.strip('_')
    # Ensure it doesn't start with a number
    if sanitized and sanitized[0].isdigit():
        sanitized = 'col_' + sanitized
    return sanitized.lower() or 'column'

def escape_sql_string(value):
    """Escape single quotes in SQL strings."""
    if pd.isna(value):
        return 'NULL'
    if isinstance(value, (int, float)):
        return str(value)
    # Escape single quotes by doubling them
    return f"'{str(value).replace(chr(39), chr(39) + chr(39))}'"

def convert_csv_to_sql(csv_file, sql_file, table_name='data_table', dialect='mysql', include_create=True):
    """
    Convert CSV to SQL INSERT statements.
    
    Args:
        csv_file (str): Path to input CSV file
        sql_file (str): Path to output SQL file
        table_name (str): Name of the SQL table
        dialect (str): SQL dialect (mysql, postgresql, sqlite)
        include_create (bool): Include CREATE TABLE statement
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting CSV to SQL...", flush=True)
    print(f"Input: {csv_file}", flush=True)
    print(f"Output: {sql_file}", flush=True)
    print(f"Table: {table_name}", flush=True)
    print(f"Dialect: {dialect}", flush=True)
    
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
        sanitized_columns = [sanitize_column_name(col) for col in original_columns]
        df.columns = sanitized_columns
        
        print(f"Columns: {', '.join(sanitized_columns[:10])}", flush=True)
        if len(sanitized_columns) > 10:
            print(f"... and {len(sanitized_columns) - 10} more columns", flush=True)
        
        # Start building SQL
        sql_lines = []
        
        # Add header comment
        sql_lines.append(f"-- SQL generated from CSV file: {os.path.basename(csv_file)}")
        sql_lines.append(f"-- Generated: {pd.Timestamp.now()}")
        sql_lines.append(f"-- Rows: {rows}, Columns: {cols}")
        sql_lines.append("")
        
        # Create CREATE TABLE statement if requested
        if include_create:
            sql_lines.append(f"-- Create table statement")
            
            if dialect == 'postgresql':
                sql_lines.append(f"DROP TABLE IF EXISTS {table_name} CASCADE;")
            elif dialect == 'mysql':
                sql_lines.append(f"DROP TABLE IF EXISTS `{table_name}`;")
            else:  # sqlite
                sql_lines.append(f"DROP TABLE IF EXISTS {table_name};")
            
            sql_lines.append("")
            
            # Infer column types
            column_defs = []
            for col in sanitized_columns:
                # Simple type inference
                dtype = df[col].dtype
                if dtype == 'int64':
                    sql_type = 'INTEGER' if dialect == 'sqlite' else 'BIGINT'
                elif dtype == 'float64':
                    sql_type = 'REAL' if dialect == 'sqlite' else 'DOUBLE PRECISION'
                elif dtype == 'bool':
                    sql_type = 'BOOLEAN'
                else:
                    sql_type = 'TEXT' if dialect == 'sqlite' else 'VARCHAR(255)'
                
                if dialect == 'mysql':
                    column_defs.append(f"  `{col}` {sql_type}")
                else:
                    column_defs.append(f"  {col} {sql_type}")
            
            if dialect == 'mysql':
                create_sql = f"CREATE TABLE `{table_name}` (\n"
            else:
                create_sql = f"CREATE TABLE {table_name} (\n"
            
            create_sql += ",\n".join(column_defs)
            create_sql += "\n);"
            
            sql_lines.append(create_sql)
            sql_lines.append("")
        
        # Add INSERT statements
        sql_lines.append(f"-- Insert statements")
        sql_lines.append("")
        
        # Process in chunks for large files
        chunk_size = 1000
        total_chunks = (rows + chunk_size - 1) // chunk_size
        
        for chunk_idx in range(total_chunks):
            start_idx = chunk_idx * chunk_size
            end_idx = min((chunk_idx + 1) * chunk_size, rows)
            chunk_df = df.iloc[start_idx:end_idx]
            
            if chunk_idx > 0 and chunk_idx % 10 == 0:
                print(f"Processing rows {start_idx}-{end_idx} of {rows}...", flush=True)
            
            for _, row in chunk_df.iterrows():
                values = [escape_sql_string(val) for val in row]
                
                if dialect == 'mysql':
                    insert_sql = f"INSERT INTO `{table_name}` (`{'`, `'.join(sanitized_columns)}`) VALUES ({', '.join(values)});"
                else:
                    insert_sql = f"INSERT INTO {table_name} ({', '.join(sanitized_columns)}) VALUES ({', '.join(values)});"
                
                sql_lines.append(insert_sql)
        
        # Write to SQL file
        print(f"Writing SQL file...", flush=True)
        with open(sql_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(sql_lines))
        
        # Verify output file
        if not os.path.exists(sql_file):
            raise FileNotFoundError(f"SQL file was not created: {sql_file}")
        
        file_size = os.path.getsize(sql_file)
        csv_size = os.path.getsize(csv_file)
        
        print(f"SQL file created successfully!", flush=True)
        print(f"CSV size: {csv_size:,} bytes ({csv_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"SQL size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        print(f"Total INSERT statements: {rows}", flush=True)
        
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
    parser = argparse.ArgumentParser(description='Convert CSV to SQL')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('sql_file', help='Output SQL file path')
    parser.add_argument('--table-name', 
                       default='data_table',
                       help='SQL table name (default: data_table)')
    parser.add_argument('--dialect', 
                       choices=['mysql', 'postgresql', 'sqlite'],
                       default='mysql',
                       help='SQL dialect (default: mysql)')
    parser.add_argument('--no-create-table',
                       action='store_true',
                       help='Skip CREATE TABLE statement')
    
    args = parser.parse_args()
    
    print("=== CSV to SQL Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    print(f"Pandas version: {pd.__version__}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}", flush=True)
        sys.exit(1)
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.sql_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Convert
    success = convert_csv_to_sql(
        args.csv_file, 
        args.sql_file, 
        args.table_name,
        args.dialect,
        not args.no_create_table
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()

