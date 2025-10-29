#!/usr/bin/env python3
"""
DOCX to ODT Converter
Converts Microsoft Word DOCX files to OpenDocument Text (ODT) format
Uses LibreOffice for conversion (best option for ODT)
"""

import os
import sys
import argparse
import traceback
import subprocess
import tempfile
import shutil


def find_libreoffice():
    """Find LibreOffice binary"""
    libreoffice_paths = [
        'libreoffice',
        '/usr/bin/libreoffice',
        '/usr/local/bin/libreoffice',
        '/opt/libreoffice/program/soffice',
        'soffice'
    ]
    
    for path in libreoffice_paths:
        try:
            result = subprocess.run(
                [path, '--version'],
                capture_output=True,
                check=True,
                timeout=5
            )
            print(f"Found LibreOffice at: {path}")
            return path
        except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
            continue
    
    return None


def convert_docx_to_odt(docx_file, output_file, preserve_formatting=True):
    """
    Convert DOCX file to ODT format using LibreOffice
    
    Args:
        docx_file (str): Path to input DOCX file
        output_file (str): Path to output ODT file
        preserve_formatting (bool): Preserve document formatting
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOCX to ODT conversion...")
    print(f"Input: {docx_file}")
    print(f"Output: {output_file}")
    print(f"Preserve formatting: {preserve_formatting}")
    
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
        
        # Find LibreOffice
        libreoffice = find_libreoffice()
        
        if not libreoffice:
            print("ERROR: LibreOffice not found. Please ensure LibreOffice is installed.")
            return False
        
        # Create output directory if needed
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # Build LibreOffice command
        # LibreOffice command for headless conversion
        cmd = [
            libreoffice,
            '--headless',
            '--invisible',
            '--nocrashreport',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
            '--convert-to', 'odt',
            '--outdir', output_dir if output_dir else '.',
            docx_file
        ]
        
        # Set LibreOffice environment
        env = os.environ.copy()
        env['SAL_USE_VCLPLUGIN'] = 'svp'
        env['HOME'] = '/tmp'
        env['LANG'] = 'en_US.UTF-8'
        env['LC_ALL'] = 'en_US.UTF-8'
        
        print(f"Running LibreOffice: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,  # 5 minute timeout
            env=env,
            encoding='utf-8',
            errors='replace'
        )
        
        if result.stdout:
            print(f"LibreOffice stdout: {result.stdout}")
        if result.stderr:
            print(f"LibreOffice stderr: {result.stderr}")
        
        # LibreOffice creates filename.odt in output directory
        # We need to find and rename/move it to the target location
        base_name = os.path.splitext(os.path.basename(docx_file))[0]
        actual_odt = os.path.join(output_dir if output_dir else '.', f"{base_name}.odt")
        
        # If output_dir is empty, use current directory
        if not output_dir:
            actual_odt = f"{base_name}.odt"
        
        if os.path.exists(actual_odt):
            # If the actual file is different from target, move/rename it
            if actual_odt != output_file:
                shutil.move(actual_odt, output_file)
                print(f"Renamed {actual_odt} to {output_file}")
            else:
                print(f"ODT file created: {output_file}")
            
            # Verify output file
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"ODT file created successfully: {output_size} bytes")
                return True
            else:
                print(f"ERROR: ODT file was not created at {output_file}")
                return False
        else:
            print(f"ERROR: LibreOffice did not create ODT file: {actual_odt}")
            # List directory contents for debugging
            if output_dir:
                print(f"Directory contents: {os.listdir(output_dir)}")
            else:
                print(f"Current directory contents: {os.listdir('.')}")
            return False
            
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out after 5 minutes")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DOCX to ODT: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert DOCX file to ODT format using LibreOffice')
    parser.add_argument('docx_file', help='Path to input DOCX file')
    parser.add_argument('output_file', help='Path to output ODT file')
    parser.add_argument('--no-formatting', action='store_true',
                        help='Do not preserve formatting (LibreOffice always preserves formatting by default)')
    
    args = parser.parse_args()
    
    print("=== DOCX to ODT Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_docx_to_odt(
        args.docx_file, 
        args.output_file,
        preserve_formatting=not args.no_formatting
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()

