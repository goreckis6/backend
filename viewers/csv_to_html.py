#!/usr/bin/env python3
"""
CSV to HTML converter for web preview using Python + pandas.
Converts CSV files to HTML format for browser viewing with table styling.
"""

import argparse
import os
import sys
import traceback
import pandas as pd
import csv

def detect_delimiter(csv_file, sample_size=5):
    """
    Detect the delimiter used in the CSV file.
    
    Args:
        csv_file (str): Path to CSV file
        sample_size (int): Number of lines to sample
    
    Returns:
        str: Detected delimiter
    """
    with open(csv_file, 'r', encoding='utf-8', errors='replace') as f:
        sample = ''.join([f.readline() for _ in range(sample_size)])
    
    sniffer = csv.Sniffer()
    try:
        delimiter = sniffer.sniff(sample).delimiter
        print(f"Detected delimiter: {repr(delimiter)}")
        return delimiter
    except:
        print("Could not detect delimiter, using comma as default")
        return ','

def convert_csv_to_html_pandas(csv_file, html_file, max_rows=2000):
    """
    Convert CSV to HTML using pandas with table styling.
    
    Args:
        csv_file (str): Path to input CSV file
        html_file (str): Path to output HTML file
        max_rows (int): Maximum rows to display (default: 2000)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting CSV to HTML conversion with pandas...")
    print(f"Input: {csv_file}")
    print(f"Output: {html_file}")
    print(f"Max Rows: {max_rows}")
    
    try:
        # Detect delimiter
        delimiter = detect_delimiter(csv_file)
        
        # Read CSV file
        print("Reading CSV file with pandas...")
        df = pd.read_csv(csv_file, sep=delimiter, encoding='utf-8', on_bad_lines='skip')
        
        # Check if CSV is empty
        if df.empty:
            print("WARNING: CSV file is empty")
        
        # Get stats
        rows, cols = df.shape
        print(f"CSV has {rows:,} rows and {cols} columns")
        
        # Truncate if too many rows
        truncated = False
        if rows > max_rows:
            truncated = True
            df = df.head(max_rows)
        
        html_parts = []
        html_parts.append('''<style>
        .csv-stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .csv-stat-box {
            background: #f0fdf4;
            padding: 12px 18px;
            border-radius: 8px;
            border-left: 4px solid #10b981;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .csv-stat-label {
            font-size: 12px;
            color: #6b7280;
            margin-bottom: 4px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .csv-stat-value {
            font-size: 24px;
            font-weight: 700;
            color: #059669;
        }
        .csv-warning-banner {
            background: #fef3c7;
            border-left: 4px solid #f59e0b;
            padding: 14px 18px;
            border-radius: 8px;
            margin-bottom: 20px;
            color: #92400e;
            font-size: 14px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .csv-table-wrapper {
            overflow-x: auto;
            margin: 20px 0;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        table.data-table {
            border-collapse: collapse;
            width: 100%;
            background: white;
            margin: 0;
        }
        table.data-table th {
            background: linear-gradient(to bottom, #10b981, #059669);
            color: white;
            padding: 12px 14px;
            text-align: left;
            font-weight: 600;
            font-size: 14px;
            border: 1px solid #047857;
            white-space: nowrap;
        }
        table.data-table td {
            padding: 10px 14px;
            border: 1px solid #e5e7eb;
            color: #374151;
            font-size: 14px;
        }
        table.data-table tr:nth-child(even) {
            background: #f9fafb;
        }
        table.data-table tr:hover {
            background: #f0fdf4;
        }
        .csv-empty {
            text-align: center;
            padding: 60px 20px;
            color: #9ca3af;
            font-size: 16px;
            background: #f9fafb;
            border-radius: 8px;
        }
    </style>''')
        
        # Add stats with better styling
        html_parts.append('        <div class="csv-stats">\n')
        html_parts.append(f'            <div class="csv-stat-box"><div class="csv-stat-label">Rows</div><div class="csv-stat-value">{rows:,}</div></div>\n')
        html_parts.append(f'            <div class="csv-stat-box"><div class="csv-stat-label">Columns</div><div class="csv-stat-value">{cols}</div></div>\n')
        html_parts.append('        </div>\n')
        
        # Show warning if truncated
        if truncated:
            html_parts.append(f'        <div class="csv-warning-banner">⚠️ This CSV file has {rows:,} rows. Showing first {max_rows:,} rows only. Download the file for full content.</div>\n')
        
        # Check if empty
        if df.empty:
            html_parts.append('        <div class="csv-empty">This CSV file is empty</div>\n')
        else:
            # Convert DataFrame to HTML table
            table_html = df.to_html(
                index=False,
                na_rep='',
                border=0,
                classes='data-table',
                escape=True,
                table_id='csv-data-table'
            )
            
            html_parts.append('        <div class="csv-table-wrapper">\n')
            html_parts.append('        ' + table_html + '\n')
            html_parts.append('        </div>\n')
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(''.join(html_parts))
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: Pandas conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to HTML for web preview')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-rows', type=int, default=2000,
                        help='Maximum rows to display (default: 2000)')
    
    args = parser.parse_args()
    
    print("=== CSV to HTML Preview Converter ===")
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
        print(f"Pandas version: {pandas.__version__}")
    except ImportError as e:
        print(f"ERROR: Pandas not available: {e}")
        print("Please install: pip install pandas")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to HTML
    success = convert_csv_to_html_pandas(
        args.csv_file,
        args.html_file,
        max_rows=args.max_rows
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

