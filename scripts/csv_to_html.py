#!/usr/bin/env python3
"""
CSV to HTML Converter
Converts CSV files to HTML format with customizable styling options.
"""

import pandas as pd
import argparse
import sys
import os
from pathlib import Path


def create_html_table(df, table_class='simple', include_headers=True):
    """
    Convert a pandas DataFrame to HTML table with specified styling.
    
    Args:
        df: pandas DataFrame
        table_class: CSS class for table styling ('simple', 'striped', 'bordered')
        include_headers: Whether to include column headers
    
    Returns:
        str: HTML table string
    """
    # Define CSS styles based on table class
    css_styles = {
        'simple': """
        <style>
        .simple-table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
        }
        .simple-table th, .simple-table td {
            padding: 8px 12px;
            text-align: left;
            border: none;
        }
        .simple-table th {
            background-color: #f8f9fa;
            font-weight: bold;
        }
        .simple-table tr:nth-child(even) {
            background-color: #f8f9fa;
        }
        </style>
        """,
        'striped': """
        <style>
        .striped-table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
        }
        .striped-table th, .striped-table td {
            padding: 10px 15px;
            text-align: left;
            border: none;
        }
        .striped-table th {
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
        }
        .striped-table tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .striped-table tr:hover {
            background-color: #e8f5e8;
        }
        </style>
        """,
        'bordered': """
        <style>
        .bordered-table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
        }
        .bordered-table th, .bordered-table td {
            padding: 10px 15px;
            text-align: left;
            border: 1px solid #ddd;
        }
        .bordered-table th {
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
        }
        .bordered-table tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .bordered-table tr:hover {
            background-color: #e8f5e8;
        }
        </style>
        """
    }
    
    # Generate HTML table
    if include_headers:
        html_table = df.to_html(
            classes=[f'{table_class}-table'],
            table_id='csv-table',
            escape=False,
            index=False
        )
    else:
        # Create table without headers
        html_table = f'<table class="{table_class}-table" id="csv-table">\n'
        html_table += '<tbody>\n'
        for _, row in df.iterrows():
            html_table += '  <tr>\n'
            for value in row:
                html_table += f'    <td>{value}</td>\n'
            html_table += '  </tr>\n'
        html_table += '</tbody>\n</table>'
    
    # Wrap in complete HTML document
    html_document = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CSV Data Table</title>
    {css_styles.get(table_class, css_styles['simple'])}
</head>
<body>
    <div style="max-width: 100%; overflow-x: auto; margin: 20px;">
        {html_table}
    </div>
</body>
</html>"""
    
    return html_document


def convert_csv_to_html(input_file, output_file, table_class='simple', include_headers=True, chunk_size=1000):
    """
    Convert CSV file to HTML format.
    
    Args:
        input_file: Path to input CSV file
        output_file: Path to output HTML file
        table_class: CSS class for table styling
        include_headers: Whether to include column headers
        chunk_size: Number of rows to process at a time for large files
    """
    try:
        # Read CSV file
        print(f"Reading CSV file: {input_file}")
        
        # Check file size to determine if we need chunking
        file_size = os.path.getsize(input_file)
        file_size_mb = file_size / (1024 * 1024)
        
        if file_size_mb > 10:  # For files larger than 10MB, use chunking
            print(f"Large file detected ({file_size_mb:.2f} MB), using chunked processing...")
            
            # Read in chunks and combine
            chunk_list = []
            for chunk in pd.read_csv(input_file, chunksize=chunk_size):
                chunk_list.append(chunk)
            
            if not chunk_list:
                raise ValueError("CSV file is empty")
            
            df = pd.concat(chunk_list, ignore_index=True)
            print(f"Processed {len(df)} rows in chunks")
        else:
            df = pd.read_csv(input_file)
            print(f"Processed {len(df)} rows")
        
        # Clean the data
        df = df.fillna('')  # Replace NaN values with empty strings
        
        # Generate HTML
        print(f"Generating HTML with {table_class} styling...")
        html_content = create_html_table(df, table_class, include_headers)
        
        # Write HTML file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"HTML file created successfully: {output_file}")
        print(f"Table contains {len(df)} rows and {len(df.columns)} columns")
        
        return True
        
    except Exception as e:
        print(f"Error converting CSV to HTML: {str(e)}", file=sys.stderr)
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert CSV file to HTML format')
    parser.add_argument('input_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output HTML file path')
    parser.add_argument('--table-class', choices=['simple', 'striped', 'bordered'], 
                       default='simple', help='CSS class for table styling')
    parser.add_argument('--no-headers', action='store_true', 
                       help='Exclude column headers from the table')
    parser.add_argument('--chunk-size', type=int, default=1000,
                       help='Number of rows to process at a time for large files')
    
    args = parser.parse_args()
    
    # Validate input file
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' does not exist", file=sys.stderr)
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to HTML
    success = convert_csv_to_html(
        args.input_file,
        args.output_file,
        args.table_class,
        not args.no_headers,
        args.chunk_size
    )
    
    if success:
        print("Conversion completed successfully!")
        sys.exit(0)
    else:
        print("Conversion failed!", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
