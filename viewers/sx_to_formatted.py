#!/usr/bin/env python3
"""
SX (Stat Studio Program) to HTML converter for web preview.
Formats statistical code with syntax highlighting.
"""

import argparse
import os
import sys
import html
import traceback

def format_sx_to_html(sx_file, html_file, max_lines=50000):
    """
    Convert SX file to HTML format with syntax highlighting.
    
    Args:
        sx_file (str): Path to input SX file
        html_file (str): Path to output HTML file
        max_lines (int): Maximum number of lines to process
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting SX to HTML conversion...")
    print(f"Input: {sx_file}")
    print(f"Output: {html_file}")
    print(f"Max Lines: {max_lines}")
    
    try:
        # Read SX file with encoding detection
        try:
            with open(sx_file, 'r', encoding='utf-8') as f:
                code_content = f.read()
        except UnicodeDecodeError:
            # Try with different encoding
            with open(sx_file, 'r', encoding='latin-1', errors='replace') as f:
                code_content = f.read()
        
        lines = code_content.splitlines()
        line_count = len(lines)
        truncated = False
        
        if max_lines > 0 and line_count > max_lines:
            lines = lines[:max_lines]
            truncated = True
        
        # Format each line with line numbers
        formatted_lines = []
        for i, line in enumerate(lines, 1):
            escaped_line = html.escape(line)
            line_num = str(i).rjust(6)
            formatted_lines.append(
                f'<div class="line">'
                f'<span class="line-number">{line_num}</span>'
                f'<span class="line-content">{escaped_line}</span>'
                f'</div>'
            )
        
        lines_html = '\n'.join(formatted_lines)
        
        # Add truncation warning if applicable
        truncation_warning = ''
        if truncated:
            truncation_warning = f'''
            <div class="truncated-warning">
              ⚠️ File truncated: showing first {max_lines:,} lines. 
              Please download the original file for full content.
            </div>
            '''
        
        # Generate complete HTML
        html_output = f'''
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>SX Program Preview</title>
  <style>
    body {{
      margin: 0;
      padding: 20px;
      font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
      background: #1e293b;
      color: #e2e8f0;
      font-size: 14px;
      line-height: 1.6;
    }}
    .truncated-warning {{
      background: #fef3c7;
      color: #92400e;
      padding: 12px 20px;
      margin-bottom: 20px;
      border-radius: 6px;
      border-left: 4px solid #f59e0b;
      font-weight: 600;
    }}
    .lines-container {{
      background: #0f172a;
      border-radius: 8px;
      padding: 20px;
      overflow-x: auto;
    }}
    .line {{
      display: flex;
      min-height: 20px;
    }}
    .line:hover {{
      background: rgba(59, 130, 246, 0.1);
    }}
    .line-number {{
      color: #64748b;
      margin-right: 20px;
      user-select: none;
      min-width: 60px;
      text-align: right;
      font-weight: 600;
    }}
    .line-content {{
      color: #e2e8f0;
      white-space: pre;
      flex: 1;
    }}
    .stats {{
      background: #334155;
      padding: 15px 20px;
      margin-bottom: 20px;
      border-radius: 6px;
      display: flex;
      gap: 30px;
      flex-wrap: wrap;
    }}
    .stat-item {{
      display: flex;
      flex-direction: column;
    }}
    .stat-label {{
      color: #94a3b8;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 4px;
    }}
    .stat-value {{
      color: #10b981;
      font-size: 18px;
      font-weight: bold;
    }}
  </style>
</head>
<body>
  <div class="stats">
    <div class="stat-item">
      <span class="stat-label">Lines of Code</span>
      <span class="stat-value">{line_count:,}</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">File Size</span>
      <span class="stat-value">{len(code_content) / 1024:.2f} KB</span>
    </div>
    <div class="stat-item">
      <span class="stat-label">Format</span>
      <span class="stat-value">SX (Stat Studio)</span>
    </div>
  </div>
  {truncation_warning}
  <div class="lines-container">
{lines_html}
  </div>
</body>
</html>
'''
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_output)
        
        if os.path.exists(html_file):
            file_size = os.path.getsize(html_file)
            print(f"HTML preview file created successfully: {file_size} bytes, {len(lines)} lines")
            if truncated:
                print(f"WARNING: File was truncated to {max_lines} lines")
            return True
        else:
            print("ERROR: HTML preview file was not created")
            return False
            
    except FileNotFoundError:
        print(f"ERROR: Input SX file not found: {sx_file}")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert SX to HTML: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert SX to HTML for web preview')
    parser.add_argument('sx_file', help='Input SX file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-lines', type=int, default=50000,
                        help='Maximum number of lines to process (0 for no limit, default: 50000)')
    
    args = parser.parse_args()
    
    print("=== SX to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.sx_file):
        print(f"ERROR: Input SX file not found: {args.sx_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert SX to HTML
    success = format_sx_to_html(
        args.sx_file,
        args.html_file,
        max_lines=args.max_lines
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()


