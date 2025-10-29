#!/usr/bin/env python3
"""
DOC to TXT Converter
Converts Microsoft Word DOC files to plain text (TXT) format
Uses pypandoc (or pandoc) for clean text output
DOC -> DOCX (LibreOffice) -> TXT (Pandoc/pypandoc)
"""

import os
import sys
import argparse
import traceback
import subprocess
import tempfile
import shutil

try:
    import pypandoc
    HAS_PYPANDOC = True
except ImportError:
    HAS_PYPANDOC = False


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
            print(f"LibreOffice conversion failed: {result.stderr}")
            return None
    except Exception as e:
        print(f"Error converting DOC to DOCX: {e}")
        return None


def convert_docx_to_txt_pypandoc(docx_file, output_file, preserve_line_breaks=True, remove_formatting=True):
    """Convert DOCX to TXT using pypandoc library"""
    try:
        print("Using pypandoc to convert DOCX to TXT...")
        
        # pypandoc.convert_file(source_file, to, format=None, outputfile=None, extra_args=None)
        extra_args = []
        
        if preserve_line_breaks:
            extra_args.append('--wrap=none')
        else:
            extra_args.append('--wrap=preserve')
        
        # Convert DOCX to plain text
        pypandoc.convert_file(
            docx_file,
            'plain',
            format='docx',
            outputfile=output_file,
            extra_args=extra_args
        )
        
        if os.path.exists(output_file):
            output_size = os.path.getsize(output_file)
            print(f"TXT file created successfully using pypandoc: {output_size} bytes")
            return True
        else:
            print("ERROR: pypandoc did not create TXT file")
            return False
            
    except Exception as e:
        print(f"Error using pypandoc: {e}")
        traceback.print_exc()
        return False


def convert_docx_to_txt_pandoc(docx_file, output_file, preserve_line_breaks=True, remove_formatting=True):
    """Convert DOCX to TXT using Pandoc binary"""
    pandoc = find_pandoc()
    if not pandoc:
        return False
    
    try:
        # Create output directory if needed
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        cmd = [
            pandoc,
            docx_file,
            '-f', 'docx',
            '-t', 'plain',
            '-o', output_file
        ]
        
        if preserve_line_breaks:
            cmd.extend(['--wrap=none'])
        else:
            cmd.extend(['--wrap=preserve'])
        
        print(f"Running Pandoc: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            encoding='utf-8',
            errors='replace'
        )
        
        if result.stdout:
            print(f"Pandoc stdout: {result.stdout}")
        if result.stderr:
            print(f"Pandoc stderr: {result.stderr}")
        
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
        print(f"ERROR: Failed to convert DOCX to TXT with Pandoc: {e}")
        traceback.print_exc()
        return False


def convert_doc_to_txt(doc_file, output_file, preserve_line_breaks=True, remove_formatting=True):
    """
    Convert DOC file to TXT format
    Strategy: DOC -> DOCX (LibreOffice) -> TXT (pypandoc/pandoc)
    
    Args:
        doc_file (str): Path to input DOC file
        output_file (str): Path to output TXT file
        preserve_line_breaks (bool): Preserve line breaks from DOC
        remove_formatting (bool): Remove all formatting (default for clean TXT)
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOC to TXT conversion...")
    print(f"Input: {doc_file}")
    print(f"Output: {output_file}")
    print(f"Preserve line breaks: {preserve_line_breaks}")
    print(f"Remove formatting: {remove_formatting}")
    
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
        
        temp_dir = tempfile.mkdtemp()
        temp_docx = None
        
        try:
            # Step 1: Convert DOC to DOCX using LibreOffice
            print("Step 1: Converting DOC to DOCX using LibreOffice...")
            temp_docx = convert_doc_to_docx_with_libreoffice(doc_file, temp_dir)
            
            if not temp_docx or not os.path.exists(temp_docx):
                print("ERROR: Failed to convert DOC to DOCX using LibreOffice")
                return False
            
            print(f"Step 1 complete: DOCX created at {temp_docx}")
            
            # Step 2: Convert DOCX to TXT using pypandoc or pandoc
            print("Step 2: Converting DOCX to TXT...")
            
            # Create output directory if needed
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            # Try pypandoc first, then fall back to pandoc binary
            success = False
            
            if HAS_PYPANDOC:
                try:
                    print("Trying pypandoc...")
                    success = convert_docx_to_txt_pypandoc(temp_docx, output_file, preserve_line_breaks, remove_formatting)
                except Exception as e:
                    print(f"pypandoc failed: {e}, trying pandoc binary...")
                    success = False
            
            if not success:
                print("Trying pandoc binary...")
                success = convert_docx_to_txt_pandoc(temp_docx, output_file, preserve_line_breaks, remove_formatting)
            
            if success and os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"Step 2 complete: TXT created at {output_file} ({output_size} bytes)")
                return True
            else:
                print("ERROR: Failed to convert DOCX to TXT")
                return False
                
        finally:
            # Clean up temporary files
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.unlink(temp_docx)
                    print(f"Cleaned up temporary DOCX file: {temp_docx}")
                except Exception as e:
                    print(f"Warning: Failed to clean up temporary DOCX file {temp_docx}: {e}")
            
            if temp_dir and os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                    print(f"Cleaned up temporary directory: {temp_dir}")
                except Exception as e:
                    print(f"Warning: Failed to clean up temporary directory {temp_dir}: {e}")
                    
    except Exception as e:
        print(f"ERROR: Failed to convert DOC to TXT: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert DOC file to TXT format using pypandoc/pandoc')
    parser.add_argument('doc_file', help='Path to input DOC file')
    parser.add_argument('output_file', help='Path to output TXT file')
    parser.add_argument('--no-line-breaks', action='store_true',
                        help='Do not preserve line breaks')
    parser.add_argument('--keep-formatting', action='store_true',
                        help='Keep formatting (plain text format removes formatting by default)')
    
    args = parser.parse_args()
    
    print("=== DOC to TXT Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    print(f"pypandoc available: {HAS_PYPANDOC}")
    
    success = convert_doc_to_txt(
        args.doc_file,
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

