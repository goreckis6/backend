#!/usr/bin/env python3
"""
HTML formatter for web preview using Python.
Formats HTML files with syntax highlighting and validation.
"""

import argparse
import os
import sys
import traceback
import html as html_module
from html.parser import HTMLParser

class HTMLElementCounter(HTMLParser):
    """Count HTML elements and extract statistics."""
    
    def __init__(self):
        super().__init__()
        self.element_count = 0
        self.tag_counts = {}
        self.has_doctype = False
        
    def handle_starttag(self, tag, attrs):
        self.element_count += 1
        self.tag_counts[tag] = self.tag_counts.get(tag, 0) + 1
    
    def handle_startendtag(self, tag, attrs):
        self.element_count += 1
        self.tag_counts[tag] = self.tag_counts.get(tag, 0) + 1
        
    def handle_decl(self, decl):
        if 'doctype' in decl.lower():
            self.has_doctype = True

def escape_and_highlight_html(html_content):
    """
    Escape HTML and add syntax highlighting.
    
    Args:
        html_content (str): HTML content
    
    Returns:
        str: HTML with syntax highlighting
    """
    lines = html_content.split('\n')
    highlighted_lines = []
    
    for line in lines:
        if not line.strip():
            highlighted_lines.append('')
            continue
        
        # Escape HTML first
        escaped_line = html_module.escape(line)
        
        # Highlight DOCTYPE
        if '&lt;!DOCTYPE' in escaped_line or '&lt;!doctype' in escaped_line:
            escaped_line = escaped_line.replace('&lt;!DOCTYPE', '<span class="html-doctype">&lt;!DOCTYPE')
            escaped_line = escaped_line.replace('&lt;!doctype', '<span class="html-doctype">&lt;!doctype')
            escaped_line = escaped_line.replace('&gt;', '&gt;</span>')
        
        # Highlight comments
        if '&lt;!--' in escaped_line:
            escaped_line = escaped_line.replace('&lt;!--', '<span class="html-comment">&lt;!--')
            escaped_line = escaped_line.replace('--&gt;', '--&gt;</span>')
        
        # Highlight tags
        import re
        # Opening tags: <tagname
        escaped_line = re.sub(
            r'&lt;(/?)(\w+)',
            r'<span class="html-bracket">&lt;\1</span><span class="html-tag">\2</span>',
            escaped_line
        )
        # Closing brackets: >
        escaped_line = re.sub(
            r'(/?)&gt;',
            r'<span class="html-bracket">\1&gt;</span>',
            escaped_line
        )
        
        # Highlight attribute values
        escaped_line = re.sub(
            r'([\w\-]+)=&quot;([^&quot;]*)&quot;',
            r'<span class="html-attr-name">\1</span>=<span class="html-attr-value">&quot;\2&quot;</span>',
            escaped_line
        )
        escaped_line = re.sub(
            r"([\w\-]+)='([^']*)'",
            r'<span class="html-attr-name">\1</span>=<span class="html-attr-value">\'\2\'</span>',
            escaped_line
        )
        
        highlighted_lines.append(escaped_line)
    
    return '\n'.join(highlighted_lines)

def convert_html_to_formatted(html_file, output_file, max_size_mb=10):
    """
    Format HTML with syntax highlighting.
    
    Args:
        html_file (str): Path to input HTML file
        output_file (str): Path to output HTML file
        max_size_mb (int): Maximum file size to display in MB (default: 10)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting HTML formatting...")
    print(f"Input: {html_file}")
    print(f"Output: {output_file}")
    print(f"Max Size: {max_size_mb} MB")
    
    try:
        # Check file size
        file_size = os.path.getsize(html_file)
        print(f"HTML file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read HTML file
        print("Reading HTML file...")
        with open(html_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # Parse HTML to get statistics
        print("Analyzing HTML...")
        parser = HTMLElementCounter()
        is_valid = True
        error_msg = None
        
        try:
            parser.feed(content)
        except Exception as e:
            print(f"HTML parsing warning: {e}")
            is_valid = False
            error_msg = str(e)
        
        # Check if file is too large for formatted display
        truncated = False
        if file_size > max_size_mb * 1024 * 1024:
            truncated = True
            print(f"WARNING: File is too large ({file_size / 1024 / 1024:.2f} MB), showing limited preview")
            content = content[:100000]  # First 100KB only
        
        # Highlight HTML
        highlighted_html = escape_and_highlight_html(content)
        
        output_parts = []
        output_parts.append('''<style>
        .html-stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }
        .html-stat-box {
            background: #fff7ed;
            padding: 12px 18px;
            border-radius: 8px;
            border-left: 4px solid #f97316;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .html-stat-label {
            font-size: 12px;
            color: #64748b;
            margin-bottom: 4px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .html-stat-value {
            font-size: 24px;
            font-weight: 700;
            color: #ea580c;
        }
        .html-warning-banner {
            background: #fef3c7;
            border-left: 4px solid #f59e0b;
            padding: 14px 18px;
            border-radius: 8px;
            margin-bottom: 20px;
            color: #92400e;
            font-size: 14px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .html-error-banner {
            background: #fee2e2;
            border-left: 4px solid #ef4444;
            padding: 14px 18px;
            border-radius: 8px;
            margin-bottom: 20px;
            color: #991b1b;
            font-size: 14px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .html-container {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 24px;
            overflow-x: auto;
            box-shadow: inset 0 2px 10px rgba(0,0,0,0.3);
            margin: 20px 0;
        }
        .html-container pre {
            margin: 0;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: 14px;
            line-height: 1.6;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            color: #e2e8f0;
        }
        .html-doctype {
            color: #a78bfa;
            font-weight: 600;
        }
        .html-tag {
            color: #60a5fa;
            font-weight: 500;
        }
        .html-bracket {
            color: #fbbf24;
            font-weight: bold;
        }
        .html-attr-name {
            color: #34d399;
        }
        .html-attr-value {
            color: #f472b6;
        }
        .html-comment {
            color: #9ca3af;
            font-style: italic;
        }
        @media print {
            .html-container {
                background: white;
                border: 1px solid #ccc;
                box-shadow: none;
            }
            .html-container pre {
                color: black;
            }
            .html-tag { color: #0000ff; }
            .html-attr-name { color: #ff0000; }
            .html-attr-value { color: #008000; }
            .html-bracket { color: #000000; }
            .html-comment { color: #666666; }
        }
    </style>''')
        
        # Add stats with better styling
        output_parts.append('        <div class="html-stats">\n')
        output_parts.append(f'            <div class="html-stat-box"><div class="html-stat-label">File Size</div><div class="html-stat-value">{file_size / 1024:.1f} KB</div></div>\n')
        output_parts.append(f'            <div class="html-stat-box"><div class="html-stat-label">Status</div><div class="html-stat-value" style="font-size: 18px;">{"✓ Valid" if is_valid else "⚠ Warning"}</div></div>\n')
        output_parts.append(f'            <div class="html-stat-box"><div class="html-stat-label">Elements</div><div class="html-stat-value">{parser.element_count}</div></div>\n')
        output_parts.append(f'            <div class="html-stat-box"><div class="html-stat-label">DOCTYPE</div><div class="html-stat-value" style="font-size: 18px;">{"✓ Yes" if parser.has_doctype else "No"}</div></div>\n')
        
        # Show top tags
        if parser.tag_counts:
            top_tags = sorted(parser.tag_counts.items(), key=lambda x: x[1], reverse=True)[:3]
            top_tags_str = ', '.join([f'{tag} ({count})' for tag, count in top_tags])
            output_parts.append(f'            <div class="html-stat-box"><div class="html-stat-label">Top Tags</div><div class="html-stat-value" style="font-size: 14px;">{html_module.escape(top_tags_str)}</div></div>\n')
        
        output_parts.append('        </div>\n')
        
        # Show error if invalid
        if not is_valid and error_msg:
            output_parts.append(f'        <div class="html-error-banner">⚠️ HTML Parsing Warning: {html_module.escape(error_msg)}</div>\n')
        
        # Show warning if truncated
        if truncated:
            output_parts.append(f'        <div class="html-warning-banner">⚠️ This HTML file is large ({file_size / 1024 / 1024:.2f} MB). Showing first 100KB only. Download for full content.</div>\n')
        
        # Display HTML content
        output_parts.append('        <div class="html-container">\n')
        output_parts.append('            <pre>')
        output_parts.append(highlighted_html)
        output_parts.append('</pre>\n')
        output_parts.append('        </div>\n')
        
        # Write output file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(''.join(output_parts))
        
        output_size = os.path.getsize(output_file)
        print(f"Formatted HTML file created successfully: {output_size:,} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: HTML formatting error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Format HTML for web preview')
    parser.add_argument('html_file', help='Input HTML file path')
    parser.add_argument('output_file', help='Output formatted HTML file path')
    parser.add_argument('--max-size-mb', type=int, default=10,
                        help='Maximum file size for formatted display in MB (default: 10)')
    
    args = parser.parse_args()
    
    print("=== HTML Formatter for Preview ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.html_file):
        print(f"ERROR: Input HTML file not found: {args.html_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Format HTML
    success = convert_html_to_formatted(
        args.html_file,
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


