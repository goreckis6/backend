#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOC to ODT Converter
Converts Microsoft Word DOC files to OpenDocument Text (ODT) format
Uses LibreOffice for conversion (DOC -> DOCX -> ODT)
"""

import os
import sys
import argparse
import traceback
import subprocess
import tempfile
import shutil
import io

# Ensure UTF-8 encoding for stdout/stderr
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
if sys.stderr.encoding != 'utf-8':
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


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


def convert_doc_to_docx_with_libreoffice(doc_file, output_dir):
    """Convert DOC to DOCX using LibreOffice"""
    libreoffice = find_libreoffice()
    if not libreoffice:
        return None
    
    # Use absolute paths to avoid issues with special characters
    doc_file = os.path.abspath(doc_file)
    output_dir = os.path.abspath(output_dir)
    
    base_name = os.path.splitext(os.path.basename(doc_file))[0]
    output_docx = os.path.join(output_dir, f"{base_name}.docx")
    
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
        '--convert-to', 'docx',
        '--outdir', output_dir,
        doc_file
    ]
    
    env = os.environ.copy()
    env['SAL_USE_VCLPLUGIN'] = 'svp'
    env['HOME'] = '/tmp'
    env['LANG'] = 'en_US.UTF-8'
    env['LC_ALL'] = 'en_US.UTF-8'
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            env=env,
            encoding='utf-8',
            errors='replace'
        )
        
        if os.path.exists(output_docx):
            return output_docx
        else:
            # LibreOffice might create the DOCX with a slightly different name
            # Try to find any .docx file in the output directory
            print(f"Expected DOCX not found at {output_docx}, searching output directory...")
            temp_files = os.listdir(output_dir)
            print(f"Contents of output directory {output_dir}: {temp_files}")
            
            # Look for any .docx file in the output directory
            docx_files = [f for f in temp_files if f.lower().endswith('.docx')]
            if docx_files:
                found_docx = os.path.join(output_dir, docx_files[0])
                print(f"Found DOCX file: {found_docx}")
                return found_docx
            else:
                print(f"LibreOffice conversion failed: {result.stderr}")
                return None
    except Exception as e:
        print(f"Error converting DOC to DOCX: {e}")
        return None


def convert_doc_to_odt(doc_file, output_file, preserve_formatting=True, include_images=True):
    """
    Convert DOC file to ODT format using LibreOffice
    Strategy: DOC -> DOCX (using LibreOffice) -> ODT (using LibreOffice)
    
    Args:
        doc_file (str): Path to input DOC file
        output_file (str): Path to output ODT file
        preserve_formatting (bool): Preserve document formatting
        include_images (bool): Include images in conversion
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOC to ODT conversion using LibreOffice...")
    print(f"Input: {doc_file}")
    print(f"Output: {output_file}")
    print(f"Preserve formatting: {preserve_formatting}")
    print(f"Include images: {include_images}")
    
    try:
        # Use absolute paths to avoid issues with special characters
        doc_file = os.path.abspath(doc_file)
        output_file = os.path.abspath(output_file)
        
        # Check if DOC file exists
        if not os.path.exists(doc_file):
            print(f"ERROR: DOC file does not exist: {doc_file}")
            return False
        
        file_size = os.path.getsize(doc_file)
        print(f"DOC file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: Input file is empty")
            return False
        
        # Find LibreOffice
        libreoffice = find_libreoffice()
        
        if not libreoffice:
            print("ERROR: LibreOffice not found. Please ensure LibreOffice is installed.")
            return False
        
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Create output directory if needed
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # Step 1: Convert DOC to DOCX using LibreOffice
            print("Step 1: Converting DOC to DOCX using LibreOffice...")
            intermediate_docx = convert_doc_to_docx_with_libreoffice(doc_file, temp_dir)
            
            if not intermediate_docx or not os.path.exists(intermediate_docx):
                print("ERROR: Failed to convert DOC to DOCX using LibreOffice")
                return False
            
            print(f"Step 1 complete: DOCX created at {intermediate_docx}")
            
            # Step 2: Convert DOCX to ODT using LibreOffice
            print("Step 2: Converting DOCX to ODT using LibreOffice...")
            
            # Use absolute paths to avoid issues with special characters
            intermediate_docx = os.path.abspath(intermediate_docx)
            output_file = os.path.abspath(output_file)
            output_dir_path = os.path.abspath(output_dir) if output_dir else os.path.dirname(output_file)
            
            # Build LibreOffice command
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
                '--outdir', output_dir_path,
                intermediate_docx
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
            base_name = os.path.splitext(os.path.basename(intermediate_docx))[0]
            actual_odt = os.path.join(output_dir_path, f"{base_name}.odt")
            
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
                if output_dir_path:
                    print(f"Directory contents: {os.listdir(output_dir_path)}")
                return False
                
        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"Warning: Could not clean up temp directory: {e}")
            
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out after 5 minutes")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DOC to ODT: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert DOC file to ODT format using LibreOffice')
    parser.add_argument('doc_file', help='Path to input DOC file')
    parser.add_argument('output_file', help='Path to output ODT file')
    parser.add_argument('--no-formatting', action='store_true',
                        help='Do not preserve formatting (LibreOffice always preserves formatting by default)')
    parser.add_argument('--no-images', action='store_true',
                        help='Exclude images from conversion')
    
    args = parser.parse_args()
    
    print("=== DOC to ODT Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_doc_to_odt(
        args.doc_file, 
        args.output_file,
        preserve_formatting=not args.no_formatting,
        include_images=not args.no_images
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()

















