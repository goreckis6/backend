#!/usr/bin/env python3
"""
DOCX to TXT Converter
Converts Microsoft Word DOCX files to plain text (TXT) format
Uses Pandoc for clean text output (recommended for TXT)
"""

import os
import sys
import argparse
import traceback
import subprocess
import tempfile
import shutil


def find_pandoc():
    """Find Pandoc binary"""
    pandoc_paths = [
        'pandoc',
        '/usr/bin/pandoc',
        '/usr/local/bin/pandoc',
        '/opt/local/bin/pandoc'
    ]
    
    for path in pandoc_paths:
        try:
            result = subprocess.run(
                [path, '--version'],
                capture_output=True,
                check=True,
                timeout=5
            )
            print(f"Found Pandoc at: {path}")
            return path
        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
            continue
    
    return None


def convert_docx_to_txt(docx_file, output_file, preserve_line_breaks=True, remove_formatting=True):
    """
    Convert DOCX file to TXT format using Pandoc
    
    Args:
        docx_file (str): Path to input DOCX file
        output_file (str): Path to output TXT file
        preserve_line_breaks (bool): Preserve line breaks from DOCX
        remove_formatting (bool): Remove all formatting (default for clean TXT)
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOCX to TXT conversion...")
    print(f"Input: {docx_file}")
    print(f"Output: {output_file}")
    print(f"Preserve line breaks: {preserve_line_breaks}")
    print(f"Remove formatting: {remove_formatting}")
    
    try:
        # Check if DOCX file exists
        if not os.path.exists(docx_file):
            print(f"ERROR: DOCX file does not exist: {docx_file}")
            return False
        
        file_size = os.path.getsize(docx_file)
        print(f"DOCX file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: Input file is empty")
            return False
        
        # Find Pandoc
        pandoc = find_pandoc()
        
        if not pandoc:
            print("ERROR: Pandoc not found. Please ensure Pandoc is installed.")
            return False
        
        # Create output directory if needed
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # Build Pandoc command for clean TXT output
        # Pandoc can convert DOCX to plain text very cleanly
        cmd = [
            pandoc,
            docx_file,
            '-f', 'docx',
            '-t', 'plain',  # Plain text format for clean TXT
            '-o', output_file
        ]
        
        # Add options based on settings
        if preserve_line_breaks:
            # Preserve paragraphs and line breaks
            cmd.extend(['--wrap=none'])  # Don't wrap, preserve original line breaks
        else:
            cmd.extend(['--wrap=preserve'])  # Preserve but may wrap long lines
        
        # Remove formatting is default for plain text, but we can ensure it
        if remove_formatting:
            # Plain text format already removes formatting
            pass
        
        print(f"Running Pandoc: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,  # 5 minute timeout
            encoding='utf-8',
            errors='replace'
        )
        
        if result.stdout:
            print(f"Pandoc stdout: {result.stdout}")
        if result.stderr:
            print(f"Pandoc stderr: {result.stderr}")
        
        # Verify output file exists
        if os.path.exists(output_file):
            output_size = os.path.getsize(output_file)
            print(f"TXT file created successfully: {output_size} bytes")
            return True
        else:
            print(f"ERROR: Pandoc did not create TXT file: {output_file}")
            return False
            
    except subprocess.TimeoutExpired:
        print("ERROR: Pandoc conversion timed out after 5 minutes")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DOCX to TXT: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert DOCX file to TXT format using Pandoc')
    parser.add_argument('docx_file', help='Path to input DOCX file')
    parser.add_argument('output_file', help='Path to output TXT file')
    parser.add_argument('--no-line-breaks', action='store_true',
                        help='Do not preserve line breaks')
    parser.add_argument('--keep-formatting', action='store_true',
                        help='Keep formatting (plain text format removes formatting by default)')
    
    args = parser.parse_args()
    
    print("=== DOCX to TXT Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_docx_to_txt(
        args.docx_file, 
        args.output_file,
        preserve_line_breaks=not args.no_line_breaks,
        remove_formatting=not args.keep_formatting
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()

