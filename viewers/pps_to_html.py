#!/usr/bin/env python3
"""
PPS/PPSX (PowerPoint Slide Show) to HTML converter for web preview.
Uses LibreOffice for conversion. Supports both legacy PPS and modern PPSX formats.
"""

import argparse
import os
import sys
import subprocess
import shutil
import traceback

def convert_pps_to_html_libreoffice(pps_file, html_file):
    """
    Convert PPS/PPSX to HTML using LibreOffice.
    
    Args:
        pps_file (str): Path to input PPS/PPSX file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting PPS/PPSX to HTML conversion with LibreOffice...")
    
    try:
        # Ensure output directory exists
        output_dir = os.path.dirname(html_file)
        os.makedirs(output_dir, exist_ok=True)
        
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'html',
            '--outdir', output_dir,
            pps_file
        ]
        
        print(f"Executing LibreOffice command: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, timeout=120)
        print("LibreOffice stdout:", result.stdout)
        print("LibreOffice stderr:", result.stderr)
        
        # LibreOffice creates output with the input filename + .html extension
        base_name = os.path.splitext(os.path.basename(pps_file))[0]
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
    parser = argparse.ArgumentParser(description='Convert PPS/PPSX to HTML for web preview')
    parser.add_argument('pps_file', help='Input PPS/PPSX file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== PPS/PPSX to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.pps_file):
        print(f"ERROR: Input PPS/PPSX file not found: {args.pps_file}")
        sys.exit(1)
    
    # Determine file type
    file_ext = os.path.splitext(args.pps_file)[1].lower()
    if file_ext == '.pps':
        print("Detected format: PPS (PowerPoint 97-2003 Slide Show)")
    elif file_ext == '.ppsx':
        print("Detected format: PPSX (PowerPoint 2007+ Slide Show)")
    else:
        print(f"WARNING: Unknown extension '{file_ext}', will try to convert anyway")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Try LibreOffice conversion
    success = convert_pps_to_html_libreoffice(args.pps_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        print("ERROR: LibreOffice conversion failed. Please ensure LibreOffice is installed.")
        sys.exit(1)

if __name__ == "__main__":
    main()

