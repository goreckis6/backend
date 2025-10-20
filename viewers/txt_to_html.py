#!/usr/bin/env python3
"""
TXT to HTML converter for web preview with pagination support.
Converts text files to HTML format for browser viewing.
"""

import argparse
import os
import sys
import traceback

def convert_txt_to_html(txt_file, html_file, max_lines=10000):
    """
    Convert TXT to HTML with line numbers and pagination.
    
    Args:
        txt_file (str): Path to input TXT file
        html_file (str): Path to output HTML file
        max_lines (int): Maximum lines to display (default 10000)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting TXT to HTML...")
    print(f"Input: {txt_file}")
    print(f"Output: {html_file}")
    print(f"Max lines: {max_lines}")
    
    try:
        # Read file with different encodings
        encodings = ['utf-8', 'utf-16', 'latin-1', 'cp1252']
        content = None
        used_encoding = None
        
        for encoding in encodings:
            try:
                with open(txt_file, 'r', encoding=encoding, errors='replace') as f:
                    content = f.read()
                used_encoding = encoding
                print(f"Successfully read file with {encoding} encoding")
                break
            except Exception as e:
                print(f"Failed to read with {encoding}: {e}")
                continue
        
        if content is None:
            print("ERROR: Could not read file with any encoding")
            return False
        
        # Split into lines
        lines = content.split('\n')
        total_lines = len(lines)
        print(f"Total lines: {total_lines}")
        
        # Limit lines for large files
        if total_lines > max_lines:
            print(f"WARNING: File has {total_lines} lines, limiting to {max_lines}")
            lines = lines[:max_lines]
            truncated = True
        else:
            truncated = False
        
        # Escape HTML special characters
        def escape_html(text):
            return (text
                    .replace('&', '&amp;')
                    .replace('<', '&lt;')
                    .replace('>', '&gt;')
                    .replace('"', '&quot;')
                    .replace("'", '&#39;'))
        
        # Generate HTML with line numbers
        html_lines = []
        for i, line in enumerate(lines):
            line_num = str(i + 1).rjust(6)
            escaped_line = escape_html(line)
            html_lines.append(f'<div class="line"><span class="line-number">{line_num}</span><span class="line-content">{escaped_line}</span></div>')
        
        html_content = '\n'.join(html_lines)
        
        # Create full HTML
        full_html = f'''<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Text File Preview</title>
</head>
<body>
  <div class="lines-container">
    {html_content}
  </div>
  {f'<div class="truncated-warning">File truncated: showing first {max_lines} of {total_lines} lines</div>' if truncated else ''}
</body>
</html>'''
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        print(f"Encoding used: {used_encoding}")
        print(f"Lines processed: {len(lines)}")
        
        return True
        
    except Exception as e:
        print(f"ERROR: Failed to convert TXT to HTML: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert TXT to HTML for web preview')
    parser.add_argument('txt_file', help='Input TXT file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-lines', type=int, default=10000,
                        help='Maximum lines to display (default: 10000)')
    
    args = parser.parse_args()
    
    print("=== TXT to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.txt_file):
        print(f"ERROR: Input TXT file not found: {args.txt_file}")
        sys.exit(1)
    
    # Convert TXT to HTML
    success = convert_txt_to_html(args.txt_file, args.html_file, args.max_lines)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()


