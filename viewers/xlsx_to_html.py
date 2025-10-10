#!/usr/bin/env python3
"""
Excel to HTML converter for web preview using Python + openpyxl/pandas.
Converts Excel files to HTML format for browser viewing with table styling.
"""

import argparse
import os
import sys
import traceback
import pandas as pd

def convert_xlsx_to_html_pandas(xlsx_file, html_file, max_rows=2000):
    """
    Convert Excel to HTML using pandas with table styling.
    
    Args:
        xlsx_file (str): Path to input Excel file
        html_file (str): Path to output HTML file
        max_rows (int): Maximum rows to display (default: 2000)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting Excel to HTML conversion with pandas...")
    print(f"Input: {xlsx_file}")
    print(f"Output: {html_file}")
    print(f"Max Rows: {max_rows}")
    
    try:
        # Read Excel file
        print("Reading Excel file with pandas...")
        
        # Determine file type and engine
        file_extension = os.path.splitext(xlsx_file)[1].lower()
        print(f"File extension: {file_extension}")
        
        # Read all sheets with appropriate engine
        if file_extension == '.xls':
            # Old Excel format - requires xlrd
            print("Using xlrd engine for XLS format...")
            try:
                xl_file = pd.ExcelFile(xlsx_file, engine='xlrd')
            except ImportError:
                print("WARNING: xlrd not available, trying openpyxl (may fail for XLS)...")
                xl_file = pd.ExcelFile(xlsx_file)
            except Exception as e:
                print(f"ERROR with xlrd: {e}, trying default engine...")
                xl_file = pd.ExcelFile(xlsx_file)
        else:
            # Modern Excel format (XLSX) - use openpyxl
            print("Using openpyxl engine for XLSX format...")
            xl_file = pd.ExcelFile(xlsx_file, engine='openpyxl')
        
        sheet_names = xl_file.sheet_names
        print(f"Found {len(sheet_names)} sheet(s): {sheet_names}")
        
        html_parts = []
        html_parts.append('''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Excel Preview</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            margin: 0;
            padding: 0;
            background: #f5f5f5;
        }
        .header-bar {
            background: linear-gradient(to right, #10b981, #059669);
            color: white;
            padding: 15px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .header-title {
            font-size: 20px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .header-actions {
            display: flex;
            gap: 10px;
        }
        .btn {
            padding: 8px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s;
        }
        .btn-print {
            background: white;
            color: #059669;
        }
        .btn-print:hover {
            background: #f0fdf4;
            transform: scale(1.05);
        }
        .btn-close {
            background: rgba(255,255,255,0.2);
            color: white;
            border: 1px solid rgba(255,255,255,0.3);
        }
        .btn-close:hover {
            background: rgba(255,255,255,0.3);
            transform: scale(1.05);
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
        }
        h1 {
            color: #333;
            border-bottom: 3px solid #10b981;
            padding-bottom: 10px;
            margin-bottom: 20px;
            margin-top: 0;
        }
        @media print {
            .header-bar {
                display: none;
            }
            body {
                background: white;
            }
            .container {
                padding: 0;
                box-shadow: none;
            }
        }
        .sheet-nav {
            background: #f9fafb;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 30px;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }
        .sheet-tab {
            background: white;
            padding: 8px 16px;
            border-radius: 6px;
            border: 2px solid #e5e7eb;
            cursor: pointer;
            font-weight: 500;
            color: #6b7280;
            transition: all 0.2s;
        }
        .sheet-tab:hover {
            border-color: #10b981;
            color: #10b981;
        }
        .sheet-tab.active {
            background: #10b981;
            color: white;
            border-color: #10b981;
        }
        .sheet-section {
            display: none;
            margin-bottom: 40px;
        }
        .sheet-section.active {
            display: block;
        }
        .sheet-title {
            color: #10b981;
            font-size: 24px;
            margin-bottom: 15px;
            font-weight: 600;
        }
        .info-banner {
            background: #dbeafe;
            border-left: 4px solid #3b82f6;
            padding: 12px 16px;
            border-radius: 6px;
            margin-bottom: 20px;
            color: #1e40af;
        }
        .warning-banner {
            background: #fef3c7;
            border-left: 4px solid #f59e0b;
            padding: 12px 16px;
            border-radius: 6px;
            margin-bottom: 20px;
            color: #92400e;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        th {
            background: linear-gradient(to bottom, #10b981, #059669);
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            border: 1px solid #059669;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        td {
            padding: 10px 12px;
            border: 1px solid #e5e7eb;
            color: #374151;
        }
        tr:nth-child(even) {
            background: #f9fafb;
        }
        tr:hover {
            background: #f0fdf4;
        }
        .stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
        }
        .stat-box {
            background: #f0fdf4;
            padding: 10px 16px;
            border-radius: 6px;
            border-left: 3px solid #10b981;
        }
        .stat-label {
            font-size: 12px;
            color: #6b7280;
            margin-bottom: 4px;
        }
        .stat-value {
            font-size: 20px;
            font-weight: 600;
            color: #059669;
        }
        .empty-sheet {
            text-align: center;
            padding: 60px;
            color: #9ca3af;
        }
    </style>
</head>
<body>
    <div class="header-bar">
        <div class="header-title">
            <span>📊</span>
            <span>Excel Spreadsheet Preview</span>
        </div>
        <div class="header-actions">
            <button onclick="window.print()" class="btn btn-print">
                🖨️ Print
            </button>
            <button onclick="window.close()" class="btn btn-close">
                ✖️ Close
            </button>
        </div>
    </div>
    <div class="container">
''')
        
        # Add sheet navigation if multiple sheets
        if len(sheet_names) > 1:
            html_parts.append('        <div class="sheet-nav">\n')
            for idx, sheet_name in enumerate(sheet_names):
                active_class = 'active' if idx == 0 else ''
                html_parts.append(f'            <div class="sheet-tab {active_class}" onclick="showSheet(\'{sheet_name}\')">{sheet_name}</div>\n')
            html_parts.append('        </div>\n')
        
        # Process each sheet
        for sheet_idx, sheet_name in enumerate(sheet_names):
            active_class = 'active' if sheet_idx == 0 else ''
            html_parts.append(f'        <div class="sheet-section {active_class}" id="sheet-{sheet_name}">\n')
            html_parts.append(f'            <h2 class="sheet-title">{sheet_name}</h2>\n')
            
            try:
                df = pd.read_excel(xlsx_file, sheet_name=sheet_name)
                
                # Check if sheet is empty
                if df.empty:
                    html_parts.append('            <div class="empty-sheet">This sheet is empty</div>\n')
                    html_parts.append('        </div>\n')
                    continue
                
                # Get stats
                rows, cols = df.shape
                html_parts.append('            <div class="stats">\n')
                html_parts.append(f'                <div class="stat-box"><div class="stat-label">Rows</div><div class="stat-value">{rows:,}</div></div>\n')
                html_parts.append(f'                <div class="stat-box"><div class="stat-label">Columns</div><div class="stat-value">{cols}</div></div>\n')
                html_parts.append('            </div>\n')
                
                # Truncate if too many rows
                truncated = False
                if rows > max_rows:
                    truncated = True
                    df = df.head(max_rows)
                
                # Show warning if truncated
                if truncated:
                    html_parts.append(f'            <div class="warning-banner">⚠️ This sheet has {rows:,} rows. Showing first {max_rows:,} rows only. Download the file for full content.</div>\n')
                
                # Convert DataFrame to HTML table
                table_html = df.to_html(
                    index=False,
                    na_rep='',
                    border=0,
                    classes='data-table',
                    escape=True
                )
                
                html_parts.append('            ' + table_html + '\n')
                
            except Exception as e:
                print(f"Error processing sheet '{sheet_name}': {e}")
                html_parts.append(f'            <div class="warning-banner">⚠️ Error loading sheet: {str(e)}</div>\n')
            
            html_parts.append('        </div>\n')
        
        # Add JavaScript for sheet switching
        if len(sheet_names) > 1:
            html_parts.append('''
    <script>
        function showSheet(sheetName) {
            // Hide all sheets
            document.querySelectorAll('.sheet-section').forEach(section => {
                section.classList.remove('active');
            });
            // Show selected sheet
            document.getElementById('sheet-' + sheetName).classList.add('active');
            
            // Update tab styles
            document.querySelectorAll('.sheet-tab').forEach(tab => {
                tab.classList.remove('active');
            });
            event.target.classList.add('active');
        }
    </script>
''')
        
        html_parts.append('''    </div>
</body>
</html>''')
        
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
    parser = argparse.ArgumentParser(description='Convert Excel to HTML for web preview')
    parser.add_argument('xlsx_file', help='Input Excel file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-rows', type=int, default=2000,
                        help='Maximum rows to display per sheet (default: 2000)')
    
    args = parser.parse_args()
    
    print("=== Excel to HTML Preview Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.xlsx_file):
        print(f"ERROR: Input Excel file not found: {args.xlsx_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        import pandas
        print(f"Pandas version: {pandas.__version__}")
    except ImportError as e:
        print(f"ERROR: Pandas not available: {e}")
        print("Please install: pip install pandas openpyxl xlrd")
        sys.exit(1)
    
    # Check file type and required engines
    file_extension = os.path.splitext(args.xlsx_file)[1].lower()
    if file_extension == '.xls':
        try:
            import xlrd
            print(f"xlrd available: {xlrd.__VERSION__}")
        except ImportError:
            print(f"WARNING: xlrd not available for XLS files")
            print("Please install: pip install xlrd")
            print("Will attempt conversion but may fail...")
    else:
        try:
            import openpyxl
            print(f"openpyxl available: {openpyxl.__version__}")
        except ImportError:
            print(f"WARNING: openpyxl not available for XLSX files")
            print("Please install: pip install openpyxl")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert Excel to HTML
    success = convert_xlsx_to_html_pandas(
        args.xlsx_file,
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

