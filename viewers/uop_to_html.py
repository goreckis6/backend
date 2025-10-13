#!/usr/bin/env python3
"""
UOP (Uniform Office Presentation) to HTML converter for web preview.
Uses LibreOffice for conversion.
"""

import argparse
import os
import sys
import subprocess
import shutil
import traceback

def convert_uop_to_html_libreoffice(uop_file, html_file):
    """
    Convert UOP to HTML using LibreOffice.
    
    Args:
        uop_file (str): Path to input UOP file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting UOP to HTML conversion with LibreOffice...")
    
    try:
        # Ensure output directory exists
        output_dir = os.path.dirname(html_file)
        os.makedirs(output_dir, exist_ok=True)
        
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'html',
            '--outdir', output_dir,
            uop_file
        ]
        
        print(f"Executing LibreOffice command: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, timeout=120)
        print("LibreOffice stdout:", result.stdout)
        print("LibreOffice stderr:", result.stderr)
        
        # LibreOffice creates output with the input filename + .html extension
        base_name = os.path.splitext(os.path.basename(uop_file))[0]
        html_output_file = os.path.join(output_dir, f"{base_name}.html")
        
        if not os.path.exists(html_output_file):
            # Try to find any .html file in the output directory
            html_files = [f for f in os.listdir(output_dir) if f.endswith('.html')]
            if html_files:
                html_output_file = os.path.join(output_dir, html_files[0])
            else:
                raise FileNotFoundError(f"No HTML file produced by LibreOffice in {output_dir}")

        # Rename to desired output filename if different
        if html_output_file != html_file:
            shutil.move(html_output_file, html_file)
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        return True
        
    except FileNotFoundError:
        print("ERROR: LibreOffice not found. Please install LibreOffice.")
        return False
    except subprocess.CalledProcessError as e:
        print(f"ERROR: LibreOffice conversion failed with exit code {e.returncode}")
        print("Stdout:", e.stdout)
        print("Stderr:", e.stderr)
        return False
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out (>120s)")
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert UOP to HTML for web preview')
    parser.add_argument('uop_file', help='Input UOP file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== UOP to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.uop_file):
        print(f"ERROR: Input UOP file not found: {args.uop_file}")
        sys.exit(1)
    
    print("Detected format: UOP (Uniform Office Presentation)")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Try LibreOffice conversion
    success = convert_uop_to_html_libreoffice(args.uop_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        print("ERROR: LibreOffice conversion failed. Please ensure LibreOffice is installed.")
        sys.exit(1)

if __name__ == "__main__":
    main()

