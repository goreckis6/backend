#!/usr/bin/env python3
"""
ODS to HTML converter for web preview using Python + pandas/odfpy.
Converts OpenDocument Spreadsheet files to HTML format for browser viewing.
"""

import argparse
import os
import sys
import traceback
import pandas as pd

def convert_ods_to_html_pandas(ods_file, html_file, max_rows=2000):
    """
    Convert ODS to HTML using pandas with table styling.
    
    Args:
        ods_file (str): Path to input ODS file
        html_file (str): Path to output HTML file
        max_rows (int): Maximum rows to display per sheet (default: 2000)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting ODS to HTML conversion with pandas...")
    print(f"Input: {ods_file}")
    print(f"Output: {html_file}")
    print(f"Max Rows: {max_rows}")
    
    try:
        # Read ODS file - pandas can read ODS with odfpy engine
        print("Reading ODS file with pandas...")
        
        # Read all sheets
        xl_file = pd.ExcelFile(ods_file, engine='odf')
        sheet_names = xl_file.sheet_names
        print(f"Found {len(sheet_names)} sheet(s): {sheet_names}")
        
        html_parts = []
        html_parts.append('''<style>
        .ods-container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
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
        .ods-stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .ods-stat-box {
            background: #f0fdf4;
            padding: 12px 18px;
            border-radius: 8px;
            border-left: 4px solid #10b981;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .ods-stat-label {
            font-size: 12px;
            color: #6b7280;
            margin-bottom: 4px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .ods-stat-value {
            font-size: 24px;
            font-weight: 700;
            color: #059669;
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
        .empty-sheet {
            text-align: center;
            padding: 60px;
            color: #9ca3af;
        }
        @media print {
            .ods-container {
                padding: 0;
                box-shadow: none;
            }
            th {
                position: static;
            }
            .sheet-nav {
                display: none;
            }
        }
    </style>
    <div class="ods-container">
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
                df = pd.read_excel(ods_file, sheet_name=sheet_name, engine='odf')
                
                # Check if sheet is empty
                if df.empty:
                    html_parts.append('            <div class="empty-sheet">This sheet is empty</div>\n')
                    html_parts.append('        </div>\n')
                    continue
                
                # Get stats
                rows, cols = df.shape
                html_parts.append('            <div class="ods-stats">\n')
                html_parts.append(f'                <div class="ods-stat-box"><div class="ods-stat-label">Rows</div><div class="ods-stat-value">{rows:,}</div></div>\n')
                html_parts.append(f'                <div class="ods-stat-box"><div class="ods-stat-label">Columns</div><div class="ods-stat-value">{cols}</div></div>\n')
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
        
        html_parts.append('    </div>\n')
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(''.join(html_parts))
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: ODS conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert ODS to HTML for web preview')
    parser.add_argument('ods_file', help='Input ODS file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-rows', type=int, default=2000,
                        help='Maximum rows to display per sheet (default: 2000)')
    
    args = parser.parse_args()
    
    print("=== ODS to HTML Preview Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.ods_file):
        print(f"ERROR: Input ODS file not found: {args.ods_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        import pandas
        print(f"Pandas version: {pandas.__version__}")
    except ImportError as e:
        print(f"ERROR: Pandas not available: {e}")
        print("Please install: pip install pandas odfpy")
        sys.exit(1)
    
    try:
        import odf
        print(f"odfpy available")
    except ImportError as e:
        print(f"ERROR: odfpy not available: {e}")
        print("Please install: pip install odfpy")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert ODS to HTML
    success = convert_ods_to_html_pandas(
        args.ods_file,
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

