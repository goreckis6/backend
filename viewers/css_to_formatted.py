#!/usr/bin/env python3
"""
CSS formatter for web preview using Python.
Formats CSS files with syntax highlighting and rule analysis.
"""

import argparse
import os
import sys
import traceback
import html
import re

def analyze_css(css_content):
    """
    Analyze CSS content for statistics.
    
    Args:
        css_content (str): CSS content
    
    Returns:
        dict: Statistics about the CSS
    """
    stats = {
        'rules': 0,
        'selectors': 0,
        'properties': 0,
        'media_queries': 0,
        'comments': 0,
        'imports': 0
    }
    
    # Count rules (selector { ... })
    stats['rules'] = len(re.findall(r'[^{}]+\{[^{}]*\}', css_content))
    
    # Count selectors (rough estimate)
    stats['selectors'] = len(re.findall(r'[^{}]+(?=\{)', css_content))
    
    # Count properties (property: value;)
    stats['properties'] = len(re.findall(r'[\w-]+\s*:\s*[^;]+;', css_content))
    
    # Count media queries
    stats['media_queries'] = len(re.findall(r'@media[^{]+\{', css_content))
    
    # Count comments
    stats['comments'] = len(re.findall(r'/\*.*?\*/', css_content, re.DOTALL))
    
    # Count imports
    stats['imports'] = len(re.findall(r'@import\s+', css_content))
    
    return stats

def escape_and_highlight_css(css_content):
    """
    Escape HTML and add syntax highlighting to CSS.
    
    Args:
        css_content (str): CSS content
    
    Returns:
        str: HTML with syntax highlighting
    """
    lines = css_content.split('\n')
    highlighted_lines = []
    in_comment = False
    
    for line in lines:
        if not line.strip():
            highlighted_lines.append('')
            continue
        
        # Escape HTML first
        escaped_line = html.escape(line)
        
        # Handle multi-line comments
        if '/*' in escaped_line and '*/' not in escaped_line:
            in_comment = True
        
        if in_comment:
            escaped_line = f'<span class="css-comment">{escaped_line}</span>'
            if '*/' in escaped_line:
                in_comment = False
            highlighted_lines.append(escaped_line)
            continue
        
        # Highlight single-line comments
        escaped_line = re.sub(
            r'/\*.*?\*/',
            lambda m: f'<span class="css-comment">{m.group(0)}</span>',
            escaped_line
        )
        
        # Highlight @-rules (@import, @media, @keyframes, etc.)
        escaped_line = re.sub(
            r'(@[\w-]+)',
            r'<span class="css-at-rule">\1</span>',
            escaped_line
        )
        
        # Highlight selectors (before {)
        escaped_line = re.sub(
            r'^(\s*)([^{:]+)(\{)',
            r'\1<span class="css-selector">\2</span><span class="css-brace">\3</span>',
            escaped_line
        )
        
        # Highlight closing braces
        escaped_line = re.sub(
            r'(\})',
            r'<span class="css-brace">\1</span>',
            escaped_line
        )
        
        # Highlight properties (property: value;)
        escaped_line = re.sub(
            r'([\w-]+)(\s*:\s*)([^;]+)(;)',
            r'<span class="css-property">\1</span>\2<span class="css-value">\3</span><span class="css-semicolon">\4</span>',
            escaped_line
        )
        
        # Highlight color values (#hex, rgb, rgba, hsl)
        escaped_line = re.sub(
            r'(#[0-9a-fA-F]{3,8})',
            r'<span class="css-color">\1</span>',
            escaped_line
        )
        
        # Highlight units (px, em, rem, %, etc.)
        escaped_line = re.sub(
            r'(\d+)(px|em|rem|%|vh|vw|pt|cm|mm|in)',
            r'<span class="css-number">\1</span><span class="css-unit">\2</span>',
            escaped_line
        )
        
        highlighted_lines.append(escaped_line)
    
    return '\n'.join(highlighted_lines)

def convert_css_to_formatted(css_file, output_file, max_size_mb=10):
    """
    Format CSS with syntax highlighting.
    
    Args:
        css_file (str): Path to input CSS file
        output_file (str): Path to output HTML file
        max_size_mb (int): Maximum file size to display in MB (default: 10)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting CSS formatting...")
    print(f"Input: {css_file}")
    print(f"Output: {output_file}")
    print(f"Max Size: {max_size_mb} MB")
    
    try:
        # Check file size
        file_size = os.path.getsize(css_file)
        print(f"CSS file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read CSS file
        print("Reading CSS file...")
        with open(css_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # Analyze CSS
        print("Analyzing CSS...")
        stats = analyze_css(content)
        print(f"CSS analysis: {stats}")
        
        # Check if file is too large for formatted display
        truncated = False
        if file_size > max_size_mb * 1024 * 1024:
            truncated = True
            print(f"WARNING: File is too large ({file_size / 1024 / 1024:.2f} MB), showing limited preview")
            content = content[:100000]  # First 100KB only
        
        # Highlight CSS
        highlighted_css = escape_and_highlight_css(content)
        
        output_parts = []
        output_parts.append('''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>CSS Preview</title>
    <style>
        body {
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            margin: 0;
            padding: 0;
            background: #1e293b;
            color: #e2e8f0;
        }
        .header-bar {
            background: linear-gradient(to right, #6366f1, #8b5cf6);
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
            color: #6366f1;
        }
        .btn-print:hover {
            background: #e0e7ff;
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
            flex-wrap: wrap;
        }
        .stat-box {
            background: #334155;
            padding: 10px 16px;
            border-radius: 6px;
            border-left: 3px solid #6366f1;
        }
        .stat-label {
            font-size: 12px;
            color: #94a3b8;
            margin-bottom: 4px;
        }
        .stat-value {
            font-size: 20px;
            font-weight: 600;
            color: #818cf8;
        }
        .warning-banner {
            background: #78350f;
            border-left: 4px solid #f59e0b;
            padding: 12px 16px;
            border-radius: 6px;
            margin-bottom: 20px;
            color: #fef3c7;
        }
        .css-container {
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
        .css-selector {
            color: #fbbf24;
            font-weight: 500;
        }
        .css-property {
            color: #60a5fa;
        }
        .css-value {
            color: #34d399;
        }
        .css-brace {
            color: #f472b6;
            font-weight: bold;
        }
        .css-semicolon {
            color: #9ca3af;
        }
        .css-comment {
            color: #9ca3af;
            font-style: italic;
        }
        .css-at-rule {
            color: #a78bfa;
            font-weight: 600;
        }
        .css-color {
            color: #f472b6;
            font-weight: 600;
            background: rgba(244, 114, 182, 0.1);
            padding: 0 4px;
            border-radius: 3px;
        }
        .css-number {
            color: #fb923c;
        }
        .css-unit {
            color: #fbbf24;
            font-weight: 500;
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
            .css-container {
                background: white;
                border: 1px solid #ccc;
                box-shadow: none;
            }
            .css-selector { color: #000080; }
            .css-property { color: #0000ff; }
            .css-value { color: #008000; }
            .css-brace { color: #000000; }
            .css-comment { color: #666666; }
        }
    </style>
</head>
<body>
    <div class="header-bar">
        <div class="header-title">
            <span>{ }</span>
            <span>CSS Stylesheet Preview</span>
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
        output_parts.append('        <div class="stats">\n')
        output_parts.append(f'            <div class="stat-box"><div class="stat-label">File Size</div><div class="stat-value">{file_size / 1024:.1f} KB</div></div>\n')
        output_parts.append(f'            <div class="stat-box"><div class="stat-label">Rules</div><div class="stat-value">{stats["rules"]}</div></div>\n')
        output_parts.append(f'            <div class="stat-box"><div class="stat-label">Selectors</div><div class="stat-value">{stats["selectors"]}</div></div>\n')
        output_parts.append(f'            <div class="stat-box"><div class="stat-label">Properties</div><div class="stat-value">{stats["properties"]}</div></div>\n')
        
        if stats['media_queries'] > 0:
            output_parts.append(f'            <div class="stat-box"><div class="stat-label">Media Queries</div><div class="stat-value">{stats["media_queries"]}</div></div>\n')
        
        if stats['imports'] > 0:
            output_parts.append(f'            <div class="stat-box"><div class="stat-label">@imports</div><div class="stat-value">{stats["imports"]}</div></div>\n')
        
        output_parts.append('        </div>\n')
        
        # Show warning if truncated
        if truncated:
            output_parts.append(f'        <div class="warning-banner">⚠️ This CSS file is large ({file_size / 1024 / 1024:.2f} MB). Showing first 100KB only. Download for full content.</div>\n')
        
        # Display CSS content
        output_parts.append('        <div class="css-container">\n')
        output_parts.append('            <pre>')
        output_parts.append(highlighted_css)
        output_parts.append('</pre>\n')
        output_parts.append('        </div>\n')
        
        output_parts.append('''    </div>
</body>
</html>''')
        
        # Write output file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(''.join(output_parts))
        
        output_size = os.path.getsize(output_file)
        print(f"Formatted CSS file created successfully: {output_size:,} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: CSS formatting error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Format CSS for web preview')
    parser.add_argument('css_file', help='Input CSS file path')
    parser.add_argument('output_file', help='Output formatted HTML file path')
    parser.add_argument('--max-size-mb', type=int, default=10,
                        help='Maximum file size for formatted display in MB (default: 10)')
    
    args = parser.parse_args()
    
    print("=== CSS Formatter for Preview ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.css_file):
        print(f"ERROR: Input CSS file not found: {args.css_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Format CSS
    success = convert_css_to_formatted(
        args.css_file,
        args.output_file,
        max_size_mb=args.max_size_mb
    )
    
    if success:
        print("=== FORMATTING SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== FORMATTING FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

