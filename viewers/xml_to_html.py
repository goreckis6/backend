#!/usr/bin/env python3
"""
XML to HTML converter for web preview using Python.
Converts XML files to formatted HTML with syntax highlighting.
"""

import argparse
import os
import sys
import traceback
import html
import xml.etree.ElementTree as ET
from xml.dom import minidom

def prettify_xml(xml_string):
    """
    Prettify XML string with proper indentation.
    
    Args:
        xml_string (str): XML content
    
    Returns:
        str: Prettified XML
    """
    try:
        dom = minidom.parseString(xml_string)
        return dom.toprettyxml(indent='  ')
    except:
        # If parsing fails, return original
        return xml_string

def escape_and_highlight_xml(xml_content):
    """
    Escape HTML and add syntax highlighting to XML.
    
    Args:
        xml_content (str): XML content
    
    Returns:
        str: HTML with syntax highlighting
    """
    lines = xml_content.split('\n')
    highlighted_lines = []
    
    for line in lines:
        if not line.strip():
            highlighted_lines.append('')
            continue
        
        # Escape HTML first
        escaped_line = html.escape(line)
        
        # Highlight XML declaration
        if '<?xml' in escaped_line:
            escaped_line = escaped_line.replace('&lt;?xml', '<span class="xml-declaration">&lt;?xml')
            escaped_line = escaped_line.replace('?&gt;', '?&gt;</span>')
        
        # Highlight comments
        if '&lt;!--' in escaped_line:
            escaped_line = escaped_line.replace('&lt;!--', '<span class="xml-comment">&lt;!--')
            escaped_line = escaped_line.replace('--&gt;', '--&gt;</span>')
        
        # Highlight CDATA
        if '&lt;![CDATA[' in escaped_line:
            escaped_line = escaped_line.replace('&lt;![CDATA[', '<span class="xml-cdata">&lt;![CDATA[')
            escaped_line = escaped_line.replace(']]&gt;', ']]&gt;</span>')
        
        # Highlight tags and attributes (simple approach)
        # Opening tags: <tagname
        escaped_line = escaped_line.replace('&lt;', '<span class="xml-bracket">&lt;</span><span class="xml-tag">')
        # Closing tags: >
        escaped_line = escaped_line.replace('&gt;', '</span><span class="xml-bracket">&gt;</span>')
        # Self-closing: />
        escaped_line = escaped_line.replace('/</span><span class="xml-bracket">&gt;</span>', '<span class="xml-bracket">/&gt;</span>')
        
        # Highlight attribute values (anything in quotes)
        import re
        # Match attribute="value" or attribute='value'
        escaped_line = re.sub(
            r'([\w\-:]+)=&quot;([^&quot;]*)&quot;',
            r'<span class="xml-attr-name">\1</span>=<span class="xml-attr-value">&quot;\2&quot;</span>',
            escaped_line
        )
        escaped_line = re.sub(
            r"([\w\-:]+)='([^']*)'",
            r'<span class="xml-attr-name">\1</span>=<span class="xml-attr-value">\'\2\'</span>',
            escaped_line
        )
        
        highlighted_lines.append(escaped_line)
    
    return '\n'.join(highlighted_lines)

def convert_xml_to_html(xml_file, html_file, max_size_mb=10):
    """
    Convert XML to HTML with syntax highlighting.
    
    Args:
        xml_file (str): Path to input XML file
        html_file (str): Path to output HTML file
        max_size_mb (int): Maximum file size to display in MB (default: 10)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting XML to HTML conversion...")
    print(f"Input: {xml_file}")
    print(f"Output: {html_file}")
    print(f"Max Size: {max_size_mb} MB")
    
    try:
        # Check file size
        file_size = os.path.getsize(xml_file)
        print(f"XML file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read XML file
        print("Reading XML file...")
        with open(xml_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        # Validate XML
        print("Validating XML...")
        is_valid = True
        error_msg = None
        element_count = 0
        root_tag = "Unknown"
        
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            root_tag = root.tag
            # Count all elements
            element_count = len(list(root.iter()))
            print(f"XML is well-formed. Root: {root_tag}, Elements: {element_count}")
            
            # Prettify XML
            content = prettify_xml(content)
        except ET.ParseError as e:
            print(f"XML parsing error: {e}")
            is_valid = False
            error_msg = str(e)
        except Exception as e:
            print(f"XML processing error: {e}")
            is_valid = False
            error_msg = str(e)
        
        # Check if file is too large for formatted display
        truncated = False
        if file_size > max_size_mb * 1024 * 1024:
            truncated = True
            print(f"WARNING: File is too large ({file_size / 1024 / 1024:.2f} MB), showing limited preview")
            content = content[:100000]  # First 100KB only
        
        # Highlight XML
        highlighted_xml = escape_and_highlight_xml(content)
        
        html_parts = []
        html_parts.append('''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>XML Preview</title>
    <style>
        body {
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            margin: 0;
            padding: 0;
            background: #1e293b;
            color: #e2e8f0;
        }
        .header-bar {
            background: linear-gradient(to right, #f97316, #ea580c);
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
            color: #ea580c;
        }
        .btn-print:hover {
            background: #ffedd5;
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
            border-left: 3px solid #f97316;
        }
        .stat-label {
            font-size: 12px;
            color: #94a3b8;
            margin-bottom: 4px;
        }
        .stat-value {
            font-size: 20px;
            font-weight: 600;
            color: #fb923c;
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
        .xml-container {
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
        .xml-declaration {
            color: #a78bfa;
            font-weight: 600;
        }
        .xml-tag {
            color: #60a5fa;
            font-weight: 500;
        }
        .xml-bracket {
            color: #fbbf24;
            font-weight: bold;
        }
        .xml-attr-name {
            color: #34d399;
        }
        .xml-attr-value {
            color: #f472b6;
        }
        .xml-comment {
            color: #9ca3af;
            font-style: italic;
        }
        .xml-cdata {
            color: #fbbf24;
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
            .xml-container {
                background: white;
                border: 1px solid #ccc;
                box-shadow: none;
            }
            .xml-tag { color: #0000ff; }
            .xml-attr-name { color: #ff0000; }
            .xml-attr-value { color: #008000; }
            .xml-bracket { color: #000000; }
            .xml-comment { color: #666666; }
        }
    </style>
</head>
<body>
    <div class="header-bar">
        <div class="header-title">
            <span>&lt;/&gt;</span>
            <span>XML File Preview</span>
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
        if is_valid:
            html_parts.append(f'            <div class="stat-box"><div class="stat-label">Root Tag</div><div class="stat-value">{html.escape(root_tag)}</div></div>\n')
            html_parts.append(f'            <div class="stat-box"><div class="stat-label">Elements</div><div class="stat-value">{element_count}</div></div>\n')
        html_parts.append('        </div>\n')
        
        # Show error if invalid
        if not is_valid:
            html_parts.append(f'        <div class="error-banner">❌ XML Parsing Error: {html.escape(error_msg or "Invalid XML format")}</div>\n')
        
        # Show warning if truncated
        if truncated:
            html_parts.append(f'        <div class="warning-banner">⚠️ This XML file is large ({file_size / 1024 / 1024:.2f} MB). Showing first 100KB only. Download for full content.</div>\n')
        
        # Display XML content
        html_parts.append('        <div class="xml-container">\n')
        html_parts.append('            <pre>')
        html_parts.append(highlighted_xml)
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
        print(f"ERROR: XML conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert XML to HTML for web preview')
    parser.add_argument('xml_file', help='Input XML file path')
    parser.add_argument('html_file', help='Output HTML file path')
    parser.add_argument('--max-size-mb', type=int, default=10,
                        help='Maximum file size for formatted display in MB (default: 10)')
    
    args = parser.parse_args()
    
    print("=== XML to HTML Preview Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.xml_file):
        print(f"ERROR: Input XML file not found: {args.xml_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert XML to HTML
    success = convert_xml_to_html(
        args.xml_file,
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

