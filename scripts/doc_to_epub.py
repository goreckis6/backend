#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOC to EPUB Converter
Converts Microsoft Word DOC files to EPUB format
Uses Pandoc for faster, simpler conversion with good text layout
Falls back to LibreOffice if Pandoc is not available
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
try:
    from docx import Document
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    HAS_DOCX_EPUB = True
except ImportError:
    HAS_DOCX_EPUB = False


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


def convert_doc_to_epub(doc_file, output_file, include_images=True, preserve_formatting=True, generate_toc=True):
    """
    Convert DOC file to EPUB format using Pandoc
    
    Args:
        doc_file (str): Path to input DOC file
        output_file (str): Path to output EPUB file
        include_images (bool): Include images in EPUB
        preserve_formatting (bool): Preserve text formatting
        generate_toc (bool): Generate table of contents
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOC to EPUB conversion using Pandoc...")
    print(f"Input: {doc_file}")
    print(f"Output: {output_file}")
    print(f"Include images: {include_images}")
    print(f"Preserve formatting: {preserve_formatting}")
    print(f"Generate TOC: {generate_toc}")
    
    try:
        # Check if DOC file exists
        if not os.path.exists(doc_file):
            print(f"ERROR: DOC file does not exist: {doc_file}")
            return False
        
        file_size = os.path.getsize(doc_file)
        print(f"DOC file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: Input file is empty")
            return False
        
        # Find Pandoc
        pandoc = find_pandoc()
        
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Create output directory if needed
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # Step 1: Convert DOC to DOCX using LibreOffice (Pandoc doesn't support DOC directly)
            if not pandoc:
                print("ERROR: Pandoc not found. Please ensure Pandoc is installed.")
                return False
            
            print("Step 1: Converting DOC to DOCX using LibreOffice...")
            intermediate_docx = convert_doc_to_docx_with_libreoffice(doc_file, temp_dir)
            
            if not intermediate_docx or not os.path.exists(intermediate_docx):
                print("ERROR: Failed to convert DOC to DOCX using LibreOffice")
                return False
            
            print(f"Step 1 complete: DOCX created at {intermediate_docx}")
            
            # Step 2: Convert DOCX to EPUB using Pandoc
            print("Step 2: Converting DOCX to EPUB using Pandoc...")
            
            cmd = [
                pandoc,
                intermediate_docx,
                '-f', 'docx',
                '-t', 'epub3',
                '-o', output_file
            ]
            
            # Add options based on settings
            if include_images:
                cmd.extend(['--extract-media', temp_dir])
            
            if generate_toc:
                cmd.extend(['--toc', '--toc-depth', '3'])
            
            if preserve_formatting:
                # Pandoc preserves formatting by default for EPUB
                pass
            else:
                cmd.extend(['--strip-comments'])
            
            # Set metadata
            base_title = os.path.splitext(os.path.basename(doc_file))[0]
            cmd.extend(['--metadata', f'title={base_title}'])
            
            # Only add epub-cover-image if we have an actual cover image file
            # (Not adding it here since we don't have a cover image to use)
            
            print(f"Running Pandoc: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300,  # 5 minute timeout
                encoding='utf-8',
                errors='replace',
                cwd=temp_dir
            )
            
            if result.stdout:
                print(f"Pandoc stdout: {result.stdout}")
            if result.stderr and 'Warning' not in result.stderr:
                print(f"Pandoc stderr: {result.stderr}")
            
            # Verify output file exists
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"EPUB file created successfully: {output_size} bytes")
                return True
            else:
                print(f"ERROR: Pandoc did not create EPUB file: {output_file}")
                return False
                
        finally:
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"Warning: Could not clean up temp directory: {e}")
            
    except subprocess.TimeoutExpired:
        print("ERROR: Pandoc conversion timed out after 5 minutes")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DOC to EPUB: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert DOC file to EPUB format using Pandoc')
    parser.add_argument('doc_file', help='Path to input DOC file')
    parser.add_argument('output_file', help='Path to output EPUB file')
    parser.add_argument('--no-images', action='store_true',
                        help='Exclude images from EPUB')
    parser.add_argument('--no-formatting', action='store_true',
                        help='Exclude formatting from EPUB')
    parser.add_argument('--no-toc', action='store_true',
                        help='Do not generate table of contents')
    
    args = parser.parse_args()
    
    print("=== DOC to EPUB Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_doc_to_epub(
        args.doc_file,
        args.output_file,
        include_images=not args.no_images,
        preserve_formatting=not args.no_formatting,
        generate_toc=not args.no_toc
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()
