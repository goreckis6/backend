#!/usr/bin/env python3
"""
DOCX to HTML converter for web preview using Python + LibreOffice.
Converts DOCX files to HTML format for browser viewing.
"""

import argparse
import os
import sys
import subprocess
import traceback

def convert_docx_to_html_libreoffice(docx_file, html_file):
    """
    Convert DOCX to HTML using LibreOffice.
    
    Args:
        docx_file (str): Path to input DOCX file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting DOCX to HTML conversion with LibreOffice...")
    print(f"Input: {docx_file}")
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
            docx_file
        ]
        
        print(f"Running command: {' '.join(cmd)}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,  # 2 minutes for large files
            env=env
        )
        
        if result.stdout:
            print(f"LibreOffice stdout: {result.stdout}")
        if result.stderr:
            print(f"LibreOffice stderr: {result.stderr}")
        
        # LibreOffice creates output with the input filename + .html extension
        base_name = os.path.splitext(os.path.basename(docx_file))[0]
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
        print("ERROR: LibreOffice conversion timed out (>120s)")
        return False
    except FileNotFoundError:
        print("ERROR: LibreOffice not found. Please install LibreOffice.")
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion error: {e}")
        traceback.print_exc()
        return False

def convert_docx_to_html_pandoc(docx_file, html_file):
    """
    Convert DOCX to HTML using Pandoc.
    
    Args:
        docx_file (str): Path to input DOCX file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting DOCX to HTML conversion with Pandoc...")
    print(f"Input: {docx_file}")
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
        
        # Convert DOCX to HTML using pandoc
        cmd = [
            'pandoc',
            docx_file,
            '-f', 'docx',
            '-t', 'html',
            '--standalone',
            '-o', html_file
        ]
        
        print(f"Running command: {' '.join(cmd)}")
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=120  # 2 minutes for large files
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
        print("ERROR: Pandoc conversion timed out (>120s)")
        return False
    except Exception as e:
        print(f"ERROR: Pandoc conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert DOCX to HTML for web preview')
    parser.add_argument('docx_file', help='Input DOCX file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== DOCX to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.docx_file):
        print(f"ERROR: Input DOCX file not found: {args.docx_file}")
        sys.exit(1)
    
    # Try conversion methods in order of preference
    success = False
    
    # 1. Try LibreOffice first (best for DOCX files)
    success = convert_docx_to_html_libreoffice(args.docx_file, args.html_file)
    
    # 2. Try Pandoc if LibreOffice failed
    if not success:
        print("\nTrying Pandoc conversion...")
        success = convert_docx_to_html_pandoc(args.docx_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        print("ERROR: All conversion methods failed")
        print("Please ensure LibreOffice or Pandoc is installed")
        sys.exit(1)

if __name__ == "__main__":
    main()

