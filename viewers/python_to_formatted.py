#!/usr/bin/env python3
"""
Python code formatter for web preview.
Formats Python files with syntax highlighting and code analysis.
"""

import argparse
import os
import sys
import traceback
import html
import re
import ast

def analyze_python(py_content):
    """
    Analyze Python content for statistics.
    
    Args:
        py_content (str): Python content
    
    Returns:
        dict: Statistics about the Python code
    """
    stats = {
        'lines': 0,
        'functions': 0,
        'classes': 0,
        'imports': 0,
        'decorators': 0,
        'comments': 0,
        'docstrings': 0
    }
    
    lines = py_content.split('\n')
    stats['lines'] = len([line for line in lines if line.strip() and not line.strip().startswith('#')])
    
    # Try to parse AST for accurate analysis
    try:
        tree = ast.parse(py_content)
        
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) or isinstance(node, ast.AsyncFunctionDef):
                stats['functions'] += 1
                # Check for docstring
                if (node.body and isinstance(node.body[0], ast.Expr) and 
                    isinstance(node.body[0].value, (ast.Str, ast.Constant))):
                    stats['docstrings'] += 1
            elif isinstance(node, ast.ClassDef):
                stats['classes'] += 1
                # Check for docstring
                if (node.body and isinstance(node.body[0], ast.Expr) and 
                    isinstance(node.body[0].value, (ast.Str, ast.Constant))):
                    stats['docstrings'] += 1
            elif isinstance(node, ast.Import) or isinstance(node, ast.ImportFrom):
                stats['imports'] += 1
    except SyntaxError:
        # If AST parsing fails, use regex fallback
        stats['functions'] = len(re.findall(r'^\s*def\s+\w+', py_content, re.MULTILINE))
        stats['functions'] += len(re.findall(r'^\s*async\s+def\s+\w+', py_content, re.MULTILINE))
        stats['classes'] = len(re.findall(r'^\s*class\s+\w+', py_content, re.MULTILINE))
        stats['imports'] = len(re.findall(r'^\s*(import|from)\s+', py_content, re.MULTILINE))
    
    # Count decorators
    stats['decorators'] = len(re.findall(r'^\s*@\w+', py_content, re.MULTILINE))
    
    # Count comments
    stats['comments'] = len(re.findall(r'#[^\n]*', py_content))
    
    return stats

def escape_and_highlight_python(py_content):
    """
    Escape HTML and add syntax highlighting to Python code.
    
    Args:
        py_content (str): Python content
    
    Returns:
        str: HTML with syntax highlighting
    """
    lines = py_content.split('\n')
    highlighted_lines = []
    in_docstring = False
    docstring_delimiter = None
    
    for line in lines:
        if not line.strip():
            highlighted_lines.append('')
            continue
        
        # Escape HTML first
        escaped_line = html.escape(line)
        
        # Handle docstrings (triple quotes)
        if '"""' in escaped_line or "'''" in escaped_line:
            if not in_docstring:
                in_docstring = True
                docstring_delimiter = '"""' if '"""' in escaped_line else "'''"
            elif docstring_delimiter in escaped_line:
                in_docstring = False
        
        if in_docstring or '"""' in escaped_line or "'''" in escaped_line:
            escaped_line = f'<span class="py-docstring">{escaped_line}</span>'
            highlighted_lines.append(escaped_line)
            continue
        
        # Highlight comments
        escaped_line = re.sub(
            r'(#[^\n]*)',
            r'<span class="py-comment">\1</span>',
            escaped_line
        )
        
        # Highlight strings (single, double, f-strings)
        escaped_line = re.sub(
            r'(f&quot;[^&quot;]*&quot;)',
            r'<span class="py-fstring">\1</span>',
            escaped_line
        )
        escaped_line = re.sub(
            r'(&quot;[^&quot;]*&quot;)',
            r'<span class="py-string">\1</span>',
            escaped_line
        )
        escaped_line = re.sub(
            r"('[^']*')",
            r'<span class="py-string">\1</span>',
            escaped_line
        )
        
        # Highlight decorators
        escaped_line = re.sub(
            r'(@\w+)',
            r'<span class="py-decorator">\1</span>',
            escaped_line
        )
        
        # Highlight keywords
        py_keywords = [
            'def', 'class', 'if', 'elif', 'else', 'for', 'while', 'try', 'except',
            'finally', 'with', 'as', 'return', 'yield', 'import', 'from', 'pass',
            'break', 'continue', 'lambda', 'async', 'await', 'raise', 'assert',
            'del', 'global', 'nonlocal', 'True', 'False', 'None', 'and', 'or',
            'not', 'is', 'in', 'self', '__init__', '__name__', '__main__'
        ]
        
        for keyword in py_keywords:
            escaped_line = re.sub(
                rf'\b({re.escape(keyword)})\b',
                r'<span class="py-keyword">\1</span>',
                escaped_line
            )
        
        # Highlight built-in functions
        builtins = ['print', 'len', 'range', 'str', 'int', 'float', 'list', 'dict', 
                    'set', 'tuple', 'open', 'type', 'isinstance', 'super']
        
        for builtin in builtins:
            escaped_line = re.sub(
                rf'\b({builtin})(\s*\()',
                r'<span class="py-builtin">\1</span>\2',
                escaped_line
            )
        
        # Highlight numbers
        escaped_line = re.sub(
            r'\b(\d+\.?\d*)\b',
            r'<span class="py-number">\1</span>',
            escaped_line
        )
        
        # Highlight function/class definitions
        escaped_line = re.sub(
            r'\b(def|class)\s+(\w+)',
            r'<span class="py-keyword">\1</span> <span class="py-function">\2</span>',
            escaped_line
        )
        
        highlighted_lines.append(escaped_line)
    
    return '\n'.join(highlighted_lines)

def convert_python_to_formatted(py_file, output_file, max_size_mb=10):
    """
    Format Python with syntax highlighting.
    
    Args:
        py_file (str): Path to input Python file
        output_file (str): Path to output HTML file
        max_size_mb (int): Maximum file size to display in MB (default: 10)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting Python formatting...")
    print(f"Input: {py_file}")
    print(f"Output: {output_file}")
    print(f"Max Size: {max_size_mb} MB")
    
    try:
        # Check file size
        file_size = os.path.getsize(py_file)
        print(f"Python file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read Python file
        print("Reading Python file...")
        with open(py_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # Analyze Python
        print("Analyzing Python...")
        stats = analyze_python(content)
        print(f"Python analysis: {stats}")
        
        # Check if file is too large for formatted display
        truncated = False
        if file_size > max_size_mb * 1024 * 1024:
            truncated = True
            print(f"WARNING: File is too large ({file_size / 1024 / 1024:.2f} MB), showing limited preview")
            content = content[:100000]  # First 100KB only
        
        # Highlight Python
        highlighted_py = escape_and_highlight_python(content)
        
        output_parts = []
        output_parts.append('''<style>
        .py-stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .py-stat-box {
            background: #f0f9ff;
            padding: 12px 18px;
            border-radius: 8px;
            border-left: 4px solid #3b82f6;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .py-stat-label {
            font-size: 12px;
            color: #64748b;
            margin-bottom: 4px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .py-stat-value {
            font-size: 24px;
            font-weight: 700;
            color: #2563eb;
        }
        .py-warning-banner {
            background: #fef3c7;
            border-left: 4px solid #f59e0b;
            padding: 14px 18px;
            border-radius: 8px;
            margin-bottom: 20px;
            color: #92400e;
            font-size: 14px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .py-container {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 24px;
            overflow-x: auto;
            box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);
            margin: 20px 0;
        }
        .py-container pre {
            margin: 0;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: 14px;
            line-height: 1.6;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            color: #e2e8f0;
        }
        .py-keyword {
            color: #c084fc;
            font-weight: 600;
        }
        .py-function {
            color: #fbbf24;
            font-weight: 500;
        }
        .py-builtin {
            color: #60a5fa;
            font-weight: 500;
        }
        .py-string {
            color: #34d399;
        }
        .py-fstring {
            color: #5eead4;
        }
        .py-number {
            color: #f472b6;
        }
        .py-comment {
            color: #9ca3af;
            font-style: italic;
        }
        .py-docstring {
            color: #86efac;
            font-style: italic;
        }
        .py-decorator {
            color: #fbbf24;
            font-weight: 600;
        }
        @media print {
            .py-container {
                background: white;
                border: 1px solid #ccc;
                box-shadow: none;
            }
            .py-container pre {
                color: black;
            }
            .py-keyword { color: #0000ff; font-weight: bold; }
            .py-function { color: #000080; }
            .py-builtin { color: #0000ff; }
            .py-string { color: #008000; }
            .py-number { color: #ff0000; }
            .py-comment { color: #666666; }
            .py-decorator { color: #ff6600; }
        }
    </style>''')
        
        # Add stats with better styling
        output_parts.append('        <div class="py-stats">\n')
        output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">File Size</div><div class="py-stat-value">{file_size / 1024:.1f} KB</div></div>\n')
        output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">Lines</div><div class="py-stat-value">{stats["lines"]}</div></div>\n')
        output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">Functions</div><div class="py-stat-value">{stats["functions"]}</div></div>\n')
        
        if stats['classes'] > 0:
            output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">Classes</div><div class="py-stat-value">{stats["classes"]}</div></div>\n')
        
        if stats['imports'] > 0:
            output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">Imports</div><div class="py-stat-value">{stats["imports"]}</div></div>\n')
        
        if stats['decorators'] > 0:
            output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">Decorators</div><div class="py-stat-value">{stats["decorators"]}</div></div>\n')
        
        if stats['docstrings'] > 0:
            output_parts.append(f'            <div class="py-stat-box"><div class="py-stat-label">Docstrings</div><div class="py-stat-value">{stats["docstrings"]}</div></div>\n')
        
        output_parts.append('        </div>\n')
        
        # Show warning if truncated
        if truncated:
            output_parts.append(f'        <div class="py-warning-banner">⚠️ This Python file is large ({file_size / 1024 / 1024:.2f} MB). Showing first 100KB only. Download for full content.</div>\n')
        
        # Display Python content
        output_parts.append('        <div class="py-container">\n')
        output_parts.append('            <pre>')
        output_parts.append(highlighted_py)
        output_parts.append('</pre>\n')
        output_parts.append('        </div>\n')
        
        # Write output file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(''.join(output_parts))
        
        output_size = os.path.getsize(output_file)
        print(f"Formatted Python file created successfully: {output_size:,} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: Python formatting error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Format Python for web preview')
    parser.add_argument('py_file', help='Input Python file path')
    parser.add_argument('output_file', help='Output formatted HTML file path')
    parser.add_argument('--max-size-mb', type=int, default=10,
                        help='Maximum file size for formatted display in MB (default: 10)')
    
    args = parser.parse_args()
    
    print("=== Python Formatter for Preview ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.py_file):
        print(f"ERROR: Input Python file not found: {args.py_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Format Python
    success = convert_python_to_formatted(
        args.py_file,
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


