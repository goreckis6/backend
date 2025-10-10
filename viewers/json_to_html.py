#!/usr/bin/env python3
"""
JSON to HTML converter for web preview using Python.
Converts JSON files to formatted HTML with syntax highlighting.
"""

import argparse
import os
import sys
import json
import traceback
import html

def format_json_value(value, indent=0):
    """
    Format JSON value with syntax highlighting.
    
    Args:
        value: JSON value to format
        indent (int): Current indentation level
    
    Returns:
        str: HTML-formatted value
    """
    spaces = '  ' * indent
    
    if value is None:
        return '<span class="json-null">null</span>'
    elif isinstance(value, bool):
        return f'<span class="json-boolean">{str(value).lower()}</span>'
    elif isinstance(value, (int, float)):
        return f'<span class="json-number">{value}</span>'
    elif isinstance(value, str):
        escaped = html.escape(value)
        return f'<span class="json-string">"{escaped}"</span>'
    elif isinstance(value, list):
        if not value:
            return '<span class="json-bracket">[</span><span class="json-bracket">]</span>'
        items = []
        for item in value:
            items.append(f'{spaces}  {format_json_value(item, indent + 1)}')
        return f'<span class="json-bracket">[</span>\n' + ',\n'.join(items) + f'\n{spaces}<span class="json-bracket">]</span>'
    elif isinstance(value, dict):
        if not value:
            return '<span class="json-brace">{{</span><span class="json-brace">}}</span>'
        items = []
        for key, val in value.items():
            escaped_key = html.escape(str(key))
            items.append(f'{spaces}  <span class="json-key">"{escaped_key}"</span>: {format_json_value(val, indent + 1)}')
        return f'<span class="json-brace">{{</span>\n' + ',\n'.join(items) + f'\n{spaces}<span class="json-brace">}}</span>'
    
    return html.escape(str(value))

def convert_json_to_html(json_file, html_file, max_size_mb=10):
    """
    Convert JSON to HTML with syntax highlighting.
    
    Args:
        json_file (str): Path to input JSON file
        html_file (str): Path to output HTML file
        max_size_mb (int): Maximum file size to display in MB (default: 10)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting JSON to HTML conversion...")
    print(f"Input: {json_file}")
    print(f"Output: {html_file}")
    print(f"Max Size: {max_size_mb} MB")
    
    try:
        # Check file size
        file_size = os.path.getsize(json_file)
        print(f"JSON file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read JSON file
        print("Reading JSON file...")
        with open(json_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse JSON
        print("Parsing JSON...")
        try:
            json_data = json.loads(content)
            is_valid = True
            error_msg = None
        except json.JSONDecodeError as e:
            print(f"JSON parsing error: {e}")
            is_valid = False
            error_msg = str(e)
            json_data = None
        
        # Check if file is too large for formatted display
        truncated = False
        if file_size > max_size_mb * 1024 * 1024:
            truncated = True
            print(f"WARNING: File is too large ({file_size / 1024 / 1024:.2f} MB), showing raw preview")
        
        html_parts = []
        html_parts.append('''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>JSON Preview</title>
    <style>
        body {
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            margin: 0;
            padding: 0;
            background: #1e293b;
            color: #e2e8f0;
        }
        .header-bar {
            background: linear-gradient(to right, #3b82f6, #2563eb);
            color: white;
            padding: 15px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.3);
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
            color: #2563eb;
        }
        .btn-print:hover {
            background: #dbeafe;
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
            padding: 30px;
        }
        .stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
        }
        .stat-box {
            background: #334155;
            padding: 10px 16px;
            border-radius: 6px;
            border-left: 3px solid #3b82f6;
        }
        .stat-label {
            font-size: 12px;
            color: #94a3b8;
            margin-bottom: 4px;
        }
        .stat-value {
            font-size: 20px;
            font-weight: 600;
            color: #60a5fa;
        }
        .warning-banner {
            background: #78350f;
            border-left: 4px solid #f59e0b;
            padding: 12px 16px;
            border-radius: 6px;
            margin-bottom: 20px;
            color: #fef3c7;
        }
        .error-banner {
            background: #7f1d1d;
            border-left: 4px solid #ef4444;
            padding: 12px 16px;
            border-radius: 6px;
            margin-bottom: 20px;
            color: #fecaca;
        }
        .json-container {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 24px;
            overflow-x: auto;
            box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);
        }
        pre {
            margin: 0;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: 14px;
            line-height: 1.6;
        }
        .json-key {
            color: #60a5fa;
            font-weight: 500;
        }
        .json-string {
            color: #34d399;
        }
        .json-number {
            color: #f472b6;
        }
        .json-boolean {
            color: #a78bfa;
            font-weight: 600;
        }
        .json-null {
            color: #9ca3af;
            font-style: italic;
        }
        .json-bracket, .json-brace {
            color: #fbbf24;
            font-weight: bold;
        }
        @media print {
            .header-bar {
                display: none;
            }
            body {
                background: white;
                color: black;
            }
            .container {
                padding: 0;
            }
            .json-container {
                background: white;
                border: 1px solid #ccc;
                box-shadow: none;
            }
            .json-key { color: #0000ff; }
            .json-string { color: #008000; }
            .json-number { color: #ff0000; }
            .json-boolean { color: #800080; }
            .json-null { color: #666666; }
            .json-bracket, .json-brace { color: #000000; }
        }
    </style>
</head>
<body>
    <div class="header-bar">
        <div class="header-title">
            <span>{ }</span>
            <span>JSON File Preview</span>
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
        
        # Add stats
        html_parts.append('        <div class="stats">\n')
        html_parts.append(f'            <div class="stat-box"><div class="stat-label">File Size</div><div class="stat-value">{file_size / 1024:.1f} KB</div></div>\n')
        html_parts.append(f'            <div class="stat-box"><div class="stat-label">Status</div><div class="stat-value">{"Valid ✓" if is_valid else "Invalid ✗"}</div></div>\n')
        if is_valid and json_data is not None:
            if isinstance(json_data, dict):
                html_parts.append(f'            <div class="stat-box"><div class="stat-label">Type</div><div class="stat-value">Object</div></div>\n')
                html_parts.append(f'            <div class="stat-box"><div class="stat-label">Keys</div><div class="stat-value">{len(json_data)}</div></div>\n')
            elif isinstance(json_data, list):
                html_parts.append(f'            <div class="stat-box"><div class="stat-label">Type</div><div class="stat-value">Array</div></div>\n')
                html_parts.append(f'            <div class="stat-box"><div class="stat-label">Items</div><div class="stat-value">{len(json_data)}</div></div>\n')
        html_parts.append('        </div>\n')
        
        # Show error if invalid
        if not is_valid:
            html_parts.append(f'        <div class="error-banner">❌ JSON Parsing Error: {html.escape(error_msg or "Invalid JSON format")}</div>\n')
        
        # Show warning if truncated
        if truncated:
            html_parts.append(f'        <div class="warning-banner">⚠️ This JSON file is large ({file_size / 1024 / 1024:.2f} MB). Showing raw content. Download for better viewing.</div>\n')
        
        # Display JSON content
        html_parts.append('        <div class="json-container">\n')
        html_parts.append('            <pre>')
        
        if is_valid and json_data is not None and not truncated:
            # Format with syntax highlighting
            formatted = format_json_value(json_data)
            html_parts.append(formatted)
        else:
            # Show raw content with escaping
            escaped_content = html.escape(content[:100000])  # Limit to first 100KB
            if len(content) > 100000:
                escaped_content += '\n\n... (content truncated, download file to view full content)'
            html_parts.append(escaped_content)
        
        html_parts.append('</pre>\n')
        html_parts.append('        </div>\n')
        
        html_parts.append('''    </div>
</body>
</html>''')
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(''.join(html_parts))
        
        output_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {output_size:,} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: JSON conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert JSON to HTML for web preview')
    parser.add_argument('json_file', help='Input JSON file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-size-mb', type=int, default=10,
                        help='Maximum file size for formatted display in MB (default: 10)')
    
    args = parser.parse_args()
    
    print("=== JSON to HTML Preview Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.json_file):
        print(f"ERROR: Input JSON file not found: {args.json_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert JSON to HTML
    success = convert_json_to_html(
        args.json_file,
        args.html_file,
        max_size_mb=args.max_size_mb
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

