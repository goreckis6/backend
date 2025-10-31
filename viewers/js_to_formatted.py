#!/usr/bin/env python3
"""
JavaScript formatter for web preview using Python.
Formats JS files with syntax highlighting and code analysis.
"""

import argparse
import os
import sys
import traceback
import html
import re

def analyze_js(js_content):
    """
    Analyze JavaScript content for statistics.
    
    Args:
        js_content (str): JavaScript content
    
    Returns:
        dict: Statistics about the JavaScript
    """
    stats = {
        'lines': 0,
        'functions': 0,
        'classes': 0,
        'imports': 0,
        'exports': 0,
        'comments': 0,
        'variables': 0
    }
    
    lines = js_content.split('\n')
    stats['lines'] = len([line for line in lines if line.strip()])
    
    # Count functions (function, arrow functions, methods)
    stats['functions'] = len(re.findall(r'\bfunction\s+\w+', js_content))
    stats['functions'] += len(re.findall(r'const\s+\w+\s*=\s*\(.*?\)\s*=>', js_content))
    stats['functions'] += len(re.findall(r'let\s+\w+\s*=\s*\(.*?\)\s*=>', js_content))
    
    # Count classes
    stats['classes'] = len(re.findall(r'\bclass\s+\w+', js_content))
    
    # Count imports
    stats['imports'] = len(re.findall(r'\bimport\s+', js_content))
    stats['imports'] += len(re.findall(r'\brequire\s*\(', js_content))
    
    # Count exports
    stats['exports'] = len(re.findall(r'\bexport\s+', js_content))
    stats['exports'] += len(re.findall(r'module\.exports', js_content))
    
    # Count comments
    stats['comments'] = len(re.findall(r'//[^\n]*', js_content))
    stats['comments'] += len(re.findall(r'/\*.*?\*/', js_content, re.DOTALL))
    
    # Count variable declarations
    stats['variables'] = len(re.findall(r'\b(const|let|var)\s+\w+', js_content))
    
    return stats

def escape_and_highlight_js(js_content):
    """
    Escape HTML and add syntax highlighting to JavaScript.
    
    Args:
        js_content (str): JavaScript content
    
    Returns:
        str: HTML with syntax highlighting
    """
    lines = js_content.split('\n')
    highlighted_lines = []
    in_multiline_comment = False
    
    for line in lines:
        if not line.strip():
            highlighted_lines.append('')
            continue
        
        # Escape HTML first
        escaped_line = html.escape(line)
        
        # Handle multi-line comments
        if '/*' in escaped_line and '*/' not in escaped_line:
            in_multiline_comment = True
        
        if in_multiline_comment:
            escaped_line = f'<span class="js-comment">{escaped_line}</span>'
            if '*/' in escaped_line:
                in_multiline_comment = False
            highlighted_lines.append(escaped_line)
            continue
        
        # Highlight single-line comments
        escaped_line = re.sub(
            r'(//[^\n]*)',
            r'<span class="js-comment">\1</span>',
            escaped_line
        )
        
        # Highlight multi-line comments (single line)
        escaped_line = re.sub(
            r'/\*.*?\*/',
            lambda m: f'<span class="js-comment">{m.group(0)}</span>',
            escaped_line
        )
        
        # Highlight strings (double and single quotes)
        escaped_line = re.sub(
            r'(&quot;[^&quot;]*&quot;)',
            r'<span class="js-string">\1</span>',
            escaped_line
        )
        escaped_line = re.sub(
            r"('[^']*')",
            r'<span class="js-string">\1</span>',
            escaped_line
        )
        escaped_line = re.sub(
            r'(`[^`]*`)',
            r'<span class="js-template">\1</span>',
            escaped_line
        )
        
        # Highlight keywords
        js_keywords = [
            'function', 'const', 'let', 'var', 'if', 'else', 'for', 'while', 'do',
            'switch', 'case', 'break', 'continue', 'return', 'class', 'extends',
            'import', 'export', 'from', 'default', 'async', 'await', 'try', 'catch',
            'throw', 'new', 'this', 'super', 'static', 'typeof', 'instanceof',
            'true', 'false', 'null', 'undefined', 'void', 'delete'
        ]
        
        for keyword in js_keywords:
            escaped_line = re.sub(
                rf'\b({keyword})\b',
                r'<span class="js-keyword">\1</span>',
                escaped_line
            )
        
        # Highlight numbers
        escaped_line = re.sub(
            r'\b(\d+\.?\d*)\b',
            r'<span class="js-number">\1</span>',
            escaped_line
        )
        
        # Highlight function calls
        escaped_line = re.sub(
            r'\b(\w+)(\s*\()',
            r'<span class="js-function">\1</span>\2',
            escaped_line
        )
        
        highlighted_lines.append(escaped_line)
    
    return '\n'.join(highlighted_lines)

def convert_js_to_formatted(js_file, output_file, max_size_mb=10):
    """
    Format JavaScript with syntax highlighting.
    
    Args:
        js_file (str): Path to input JS file
        output_file (str): Path to output HTML file
        max_size_mb (int): Maximum file size to display in MB (default: 10)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting JavaScript formatting...")
    print(f"Input: {js_file}")
    print(f"Output: {output_file}")
    print(f"Max Size: {max_size_mb} MB")
    
    try:
        # Check file size
        file_size = os.path.getsize(js_file)
        print(f"JS file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read JS file
        print("Reading JavaScript file...")
        with open(js_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # Analyze JavaScript
        print("Analyzing JavaScript...")
        stats = analyze_js(content)
        print(f"JS analysis: {stats}")
        
        # Check if file is too large for formatted display
        truncated = False
        if file_size > max_size_mb * 1024 * 1024:
            truncated = True
            print(f"WARNING: File is too large ({file_size / 1024 / 1024:.2f} MB), showing limited preview")
            content = content[:100000]  # First 100KB only
        
        # Highlight JavaScript
        highlighted_js = escape_and_highlight_js(content)
        
        output_parts = []
        output_parts.append('''<style>
        .js-stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .js-stat-box {
            background: #fefce8;
            padding: 12px 18px;
            border-radius: 8px;
            border-left: 4px solid #eab308;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .js-stat-label {
            font-size: 12px;
            color: #64748b;
            margin-bottom: 4px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .js-stat-value {
            font-size: 24px;
            font-weight: 700;
            color: #ca8a04;
        }
        .js-warning-banner {
            background: #fef3c7;
            border-left: 4px solid #f59e0b;
            padding: 14px 18px;
            border-radius: 8px;
            margin-bottom: 20px;
            color: #92400e;
            font-size: 14px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .js-container {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 24px;
            overflow-x: auto;
            box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);
            margin: 20px 0;
        }
        .js-container pre {
            margin: 0;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: 14px;
            line-height: 1.6;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            color: #e2e8f0;
        }
        .js-keyword {
            color: #c084fc;
            font-weight: 600;
        }
        .js-function {
            color: #fbbf24;
            font-weight: 500;
        }
        .js-string {
            color: #34d399;
        }
        .js-template {
            color: #5eead4;
        }
        .js-number {
            color: #f472b6;
        }
        .js-comment {
            color: #9ca3af;
            font-style: italic;
        }
        @media print {
            .js-container {
                background: white;
                border: 1px solid #ccc;
                box-shadow: none;
            }
            .js-container pre {
                color: black;
            }
            .js-keyword { color: #0000ff; font-weight: bold; }
            .js-function { color: #000080; }
            .js-string { color: #008000; }
            .js-number { color: #ff0000; }
            .js-comment { color: #666666; }
        }
    </style>''')
        
        # Add stats with better styling
        output_parts.append('        <div class="js-stats">\n')
        output_parts.append(f'            <div class="js-stat-box"><div class="js-stat-label">File Size</div><div class="js-stat-value">{file_size / 1024:.1f} KB</div></div>\n')
        output_parts.append(f'            <div class="js-stat-box"><div class="js-stat-label">Lines</div><div class="js-stat-value">{stats["lines"]}</div></div>\n')
        output_parts.append(f'            <div class="js-stat-box"><div class="js-stat-label">Functions</div><div class="js-stat-value">{stats["functions"]}</div></div>\n')
        
        if stats['classes'] > 0:
            output_parts.append(f'            <div class="js-stat-box"><div class="js-stat-label">Classes</div><div class="js-stat-value">{stats["classes"]}</div></div>\n')
        
        if stats['imports'] > 0:
            output_parts.append(f'            <div class="js-stat-box"><div class="js-stat-label">Imports</div><div class="js-stat-value">{stats["imports"]}</div></div>\n')
        
        if stats['exports'] > 0:
            output_parts.append(f'            <div class="js-stat-box"><div class="js-stat-label">Exports</div><div class="js-stat-value">{stats["exports"]}</div></div>\n')
        
        output_parts.append('        </div>\n')
        
        # Show warning if truncated
        if truncated:
            output_parts.append(f'        <div class="js-warning-banner">⚠️ This JavaScript file is large ({file_size / 1024 / 1024:.2f} MB). Showing first 100KB only. Download for full content.</div>\n')
        
        # Display JavaScript content
        output_parts.append('        <div class="js-container">\n')
        output_parts.append('            <pre>')
        output_parts.append(highlighted_js)
        output_parts.append('</pre>\n')
        output_parts.append('        </div>\n')
        
        # Write output file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(''.join(output_parts))
        
        output_size = os.path.getsize(output_file)
        print(f"Formatted JavaScript file created successfully: {output_size:,} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: JavaScript formatting error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Format JavaScript for web preview')
    parser.add_argument('js_file', help='Input JavaScript file path')
    parser.add_argument('output_file', help='Output formatted HTML file path')
    parser.add_argument('--max-size-mb', type=int, default=10,
                        help='Maximum file size for formatted display in MB (default: 10)')
    
    args = parser.parse_args()
    
    print("=== JavaScript Formatter for Preview ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.js_file):
        print(f"ERROR: Input JavaScript file not found: {args.js_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Format JavaScript
    success = convert_js_to_formatted(
        args.js_file,
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


