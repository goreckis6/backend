#!/usr/bin/env python3
"""
DOC to EPUB converter.
Converts Microsoft Word DOC files to EPUB format.
"""

import argparse
import os
import sys
import subprocess
import traceback
from pathlib import Path

def convert_doc_to_epub(doc_file, epub_file):
    """
    Convert DOC to EPUB using LibreOffice and Calibre.
    Strategy: DOC → HTML (LibreOffice) → EPUB (ebook-convert)
    
    Args:
        doc_file (str): Path to input DOC file
        epub_file (str): Path to output EPUB file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"=== Converting DOC to EPUB ===", flush=True)
    print(f"Input: {doc_file}", flush=True)
    print(f"Output: {epub_file}", flush=True)
    
    try:
        import tempfile
        
        # Check if input file exists
        if not os.path.exists(doc_file):
            raise FileNotFoundError(f"DOC file not found: {doc_file}")
        
        file_size = os.path.getsize(doc_file)
        print(f"DOC file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)", flush=True)
        
        # Create temporary directory for intermediate files
        with tempfile.TemporaryDirectory() as tmpdir:
            # Step 1: Convert DOC to HTML using LibreOffice
            print("Step 1/2: Converting DOC to HTML with LibreOffice...", flush=True)
            
            base_name = os.path.splitext(os.path.basename(doc_file))[0]
            html_file = os.path.join(tmpdir, f'{base_name}.html')
            
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
                '--convert-to', 'html',
                '--outdir', tmpdir,
                doc_file
            ]
            
            # Set LibreOffice environment
            env = os.environ.copy()
            env['SAL_USE_VCLPLUGIN'] = 'svp'
            env['HOME'] = '/tmp'
            env['LANG'] = 'en_US.UTF-8'
            env['LC_ALL'] = 'en_US.UTF-8'
            
            print(f"Executing: {' '.join(cmd)}", flush=True)
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
                env=env,
                encoding='utf-8',
                errors='replace'
            )
            
            if result.stdout:
                print(f"LibreOffice stdout: {result.stdout}", flush=True)
            if result.stderr:
                print(f"LibreOffice stderr: {result.stderr}", flush=True)
            
            # Check if HTML file was created
            if not os.path.exists(html_file):
                print(f"ERROR: HTML file not created at: {html_file}", flush=True)
                print(f"Directory contents: {os.listdir(tmpdir)}", flush=True)
                raise Exception("Failed to convert DOC to HTML with LibreOffice")
            
            print(f"HTML file created: {html_file}", flush=True)
            html_size = os.path.getsize(html_file)
            print(f"HTML size: {html_size:,} bytes ({html_size / 1024 / 1024:.2f} MB)", flush=True)
            
            # Step 2: Convert HTML to EPUB using ebook-convert (Calibre)
            print("Step 2/2: Converting HTML to EPUB with Calibre...", flush=True)
            
            # Create output directory if needed
            output_dir = os.path.dirname(epub_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            cmd = [
                'ebook-convert',
                html_file,
                epub_file,
                '--enable-heuristics',
                '--chapter', '/'
            ]
            
            print(f"Executing: {' '.join(cmd)}", flush=True)
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=180,
                encoding='utf-8',
                errors='replace'
            )
            
            if result.stdout:
                print(f"ebook-convert stdout: {result.stdout}", flush=True)
            if result.stderr:
                print(f"ebook-convert stderr: {result.stderr}", flush=True)
            
            if not os.path.exists(epub_file):
                raise Exception("Failed to create EPUB file with ebook-convert")
            
            output_size = os.path.getsize(epub_file)
            print(f"=== CONVERSION SUCCESSFUL ===", flush=True)
            print(f"EPUB file created: {epub_file}", flush=True)
            print(f"EPUB size: {output_size:,} bytes ({output_size / 1024 / 1024:.2f} MB)", flush=True)
            
            return True
        
    except subprocess.TimeoutExpired:
        print("ERROR: Conversion timed out", flush=True)
        return False
    except FileNotFoundError as e:
        print(f"ERROR: File not found: {e}", flush=True)
        return False
    except Exception as e:
        print(f"ERROR: Conversion failed: {e}", flush=True)
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert DOC to EPUB')
    parser.add_argument('doc_file', help='Input DOC file path')
    parser.add_argument('epub_file', help='Output EPUB file path')
    
    args = parser.parse_args()
    
    print("=== DOC to EPUB Converter ===", flush=True)
    print(f"Python version: {sys.version}", flush=True)
    
    # Check if input file exists
    if not os.path.exists(args.doc_file):
        print(f"ERROR: Input DOC file not found: {args.doc_file}", flush=True)
        sys.exit(1)
    
    # Convert
    success = convert_doc_to_epub(args.doc_file, args.epub_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===", flush=True)
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===", flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()

