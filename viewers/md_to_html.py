#!/usr/bin/env python3
"""
Markdown to HTML converter for web preview.
Converts Markdown files to HTML format for browser viewing.
"""

import argparse
import os
import sys
import subprocess
import traceback

def convert_md_to_html_pandoc(md_file, html_file):
    """
    Convert Markdown to HTML using Pandoc.
    
    Args:
        md_file (str): Path to input Markdown file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting Markdown to HTML conversion with Pandoc...")
    print(f"Input: {md_file}")
    print(f"Output: {html_file}")
    
    try:
        # Check if pandoc is available
        try:
            subprocess.run(['pandoc', '--version'], 
                          capture_output=True, 
                          check=True)
            print("Pandoc found")
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("WARNING: Pandoc not found")
            return False
        
        # Convert Markdown to HTML using pandoc with GitHub-flavored markdown
        cmd = [
            'pandoc',
            md_file,
            '-f', 'gfm',  # GitHub Flavored Markdown
            '-t', 'html',
            '--standalone',
            '--highlight-style', 'github',
            '-o', html_file
        ]
        
        print(f"Running command: {' '.join(cmd)}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=60
        )
        
        if result.stdout:
            print(f"Pandoc stdout: {result.stdout}")
        if result.stderr:
            print(f"Pandoc stderr: {result.stderr}")
        
        # Verify output file exists
        if os.path.exists(html_file):
            file_size = os.path.getsize(html_file)
            print(f"HTML file created successfully: {file_size} bytes")
            return True
        else:
            print("ERROR: HTML file was not created by Pandoc")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"Pandoc conversion failed: {e}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        return False
    except subprocess.TimeoutExpired:
        print("ERROR: Pandoc conversion timed out")
        return False
    except Exception as e:
        print(f"ERROR: Pandoc conversion error: {e}")
        traceback.print_exc()
        return False

def convert_md_to_html_python(md_file, html_file):
    """
    Convert Markdown to HTML using Python markdown library.
    Fallback method if Pandoc is not available.
    
    Args:
        md_file (str): Path to input Markdown file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting Markdown to HTML conversion with Python markdown...")
    
    try:
        import markdown
        
        # Read Markdown file with encoding detection
        try:
            with open(md_file, 'r', encoding='utf-8') as f:
                md_content = f.read()
        except UnicodeDecodeError:
            # Try with different encoding
            with open(md_file, 'r', encoding='utf-8', errors='replace') as f:
                md_content = f.read()
        
        # Convert to HTML with extensions
        html_content = markdown.markdown(
            md_content,
            extensions=[
                'tables',
                'fenced_code',
                'codehilite',
                'nl2br',
                'sane_lists'
            ]
        )
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        return True
        
    except ImportError as e:
        print(f"WARNING: Python markdown library not found: {e}")
        print("Install with: pip install markdown")
        return False
    except Exception as e:
        print(f"ERROR: Python markdown conversion error: {e}")
        traceback.print_exc()
        return False

def convert_md_to_html_simple(md_file, html_file):
    """
    Simple Markdown to HTML conversion (basic text with line breaks).
    Last resort fallback if Pandoc and Python markdown are not available.
    
    Args:
        md_file (str): Path to input Markdown file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Using simple Markdown text extraction...")
    
    try:
        # Read Markdown file
        with open(md_file, 'r', encoding='utf-8', errors='replace') as f:
            md_content = f.read()
        
        # Very basic conversion - preserve line breaks and escape HTML
        import html as html_lib
        escaped_content = html_lib.escape(md_content)
        # Convert line breaks to <br> tags
        html_content = escaped_content.replace('\n', '<br>\n')
        
        # Wrap in basic HTML
        full_html = f'''
<div style="font-family: monospace; white-space: pre-wrap; word-wrap: break-word;">
<p><em>Note: Basic text preview. For full Markdown rendering, Pandoc or Python markdown library is required.</em></p>
<hr>
{html_content}
</div>
'''
        
        # Write HTML file
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        file_size = os.path.getsize(html_file)
        print(f"Simple HTML file created: {file_size} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: Simple conversion failed: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert Markdown to HTML for web preview')
    parser.add_argument('md_file', help='Input Markdown file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== Markdown to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.md_file):
        print(f"ERROR: Input Markdown file not found: {args.md_file}")
        sys.exit(1)
    
    # Try conversion methods in order of preference
    success = False
    
    # 1. Try Pandoc first (best quality, supports GFM)
    success = convert_md_to_html_pandoc(args.md_file, args.html_file)
    
    # 2. Try Python markdown if Pandoc failed
    if not success:
        print("\nTrying Python markdown library...")
        success = convert_md_to_html_python(args.md_file, args.html_file)
    
    # 3. Fall back to simple text extraction
    if not success:
        print("\nFalling back to simple text extraction...")
        success = convert_md_to_html_simple(args.md_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        print("ERROR: All conversion methods failed")
        sys.exit(1)

if __name__ == "__main__":
    main()

