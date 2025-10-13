#!/usr/bin/env python3
"""
RTF to HTML converter for web preview using Python + LibreOffice/Pandoc.
Converts RTF files to HTML format for browser viewing.
"""

import argparse
import os
import sys
import subprocess
import traceback

def convert_rtf_to_html_pandoc(rtf_file, html_file):
    """
    Convert RTF to HTML using Pandoc.
    
    Args:
        rtf_file (str): Path to input RTF file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting RTF to HTML conversion with Pandoc...")
    print(f"Input: {rtf_file}")
    print(f"Output: {html_file}")
    
    try:
        # Check if pandoc is available
        try:
            subprocess.run(['pandoc', '--version'], 
                          capture_output=True, 
                          check=True)
            print("Pandoc found")
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("WARNING: Pandoc not found, trying alternative method")
            return False
        
        # Convert RTF to HTML using pandoc
        cmd = [
            'pandoc',
            rtf_file,
            '-f', 'rtf',
            '-t', 'html',
            '--standalone',
            '-o', html_file
        ]
        
        print(f"Running command: {' '.join(cmd)}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True
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
    except Exception as e:
        print(f"ERROR: Pandoc conversion error: {e}")
        traceback.print_exc()
        return False

def convert_rtf_to_html_libreoffice(rtf_file, html_file):
    """
    Convert RTF to HTML using LibreOffice.
    
    Args:
        rtf_file (str): Path to input RTF file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting RTF to HTML conversion with LibreOffice...")
    print(f"Input: {rtf_file}")
    print(f"Output: {html_file}")
    
    try:
        # Get output directory
        output_dir = os.path.dirname(html_file)
        
        # Set environment variables for LibreOffice
        env = os.environ.copy()
        env['SAL_USE_VCLPLUGIN'] = 'svp'
        env['HOME'] = os.path.expanduser('~')
        
        # LibreOffice conversion command
        cmd = [
            'libreoffice',
            '--headless',
            '--invisible',
            '--nocrashreport',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
            '-env:UserInstallation=file:///tmp/libreoffice_user_profile',
            '--convert-to', 'html',
            '--outdir', output_dir,
            rtf_file
        ]
        
        print(f"Running command: {' '.join(cmd)}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60,
            env=env
        )
        
        if result.stdout:
            print(f"LibreOffice stdout: {result.stdout}")
        if result.stderr:
            print(f"LibreOffice stderr: {result.stderr}")
        
        # LibreOffice creates output with the input filename + .html extension
        # Need to check for the actual output file
        base_name = os.path.splitext(os.path.basename(rtf_file))[0]
        actual_output = os.path.join(output_dir, f"{base_name}.html")
        
        if os.path.exists(actual_output):
            # Rename to desired output filename if different
            if actual_output != html_file:
                os.rename(actual_output, html_file)
            
            file_size = os.path.getsize(html_file)
            print(f"HTML file created successfully: {file_size} bytes")
            return True
        else:
            print(f"ERROR: LibreOffice did not create output file: {actual_output}")
            return False
            
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out")
        return False
    except FileNotFoundError:
        print("WARNING: LibreOffice not found")
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion error: {e}")
        traceback.print_exc()
        return False

def convert_rtf_to_html_simple(rtf_file, html_file):
    """
    Simple RTF to HTML conversion (basic text extraction).
    Fallback method if Pandoc and LibreOffice are not available.
    
    Args:
        rtf_file (str): Path to input RTF file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Using simple RTF text extraction...")
    
    try:
        with open(rtf_file, 'r', encoding='utf-8', errors='ignore') as f:
            rtf_content = f.read()
        
        # Very basic RTF parsing - extract text content
        # This is a simplified approach and won't preserve all formatting
        import re
        
        # Remove RTF control words
        text = re.sub(r'\\[a-z]+(-?\d+)? ?', ' ', rtf_content)
        # Remove braces
        text = re.sub(r'[{}]', '', text)
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()
        
        # Create simple HTML
        html_content = f"""
        <h2>RTF Document Content</h2>
        <p><em>Note: This is a basic text extraction. Download the file for full formatting.</em></p>
        <div style="white-space: pre-wrap;">{text}</div>
        """
        
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Simple HTML file created: {os.path.getsize(html_file)} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: Simple conversion failed: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert RTF to HTML for web preview')
    parser.add_argument('rtf_file', help='Input RTF file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== RTF to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.rtf_file):
        print(f"ERROR: Input RTF file not found: {args.rtf_file}")
        sys.exit(1)
    
    # Try conversion methods in order of preference
    success = False
    
    # 1. Try Pandoc first (best quality)
    success = convert_rtf_to_html_pandoc(args.rtf_file, args.html_file)
    
    # 2. Try LibreOffice if Pandoc failed
    if not success:
        print("\nTrying LibreOffice conversion...")
        success = convert_rtf_to_html_libreoffice(args.rtf_file, args.html_file)
    
    # 3. Fall back to simple text extraction
    if not success:
        print("\nFalling back to simple text extraction...")
        success = convert_rtf_to_html_simple(args.rtf_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

