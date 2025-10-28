#!/usr/bin/env python3
"""
EPS to WebP Converter
Converts EPS (Encapsulated PostScript) files to WebP format
"""

import os
import sys
import argparse
import tempfile
import traceback
import subprocess
import shutil
from PIL import Image, ImageOps

def convert_eps_to_webp(eps_file, output_file, quality=80, lossless=False):
    """
    Convert EPS file to WebP format
    
    Args:
        eps_file (str): Path to input EPS file
        output_file (str): Path to output WebP file
        quality (int): WebP quality (0-100, higher is better)
        lossless (bool): Use lossless compression
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting EPS to WebP conversion...")
    print(f"Input: {eps_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Lossless: {lossless}")
    
    try:
        # Check if EPS file exists and is readable
        if not os.path.exists(eps_file):
            print(f"ERROR: EPS file does not exist: {eps_file}")
            return False
        
        file_size = os.path.getsize(eps_file)
        print(f"EPS file size: {file_size} bytes")
        
        # Check if Ghostscript is available
        gs_path = shutil.which('gs')
        if not gs_path:
            print("ERROR: Ghostscript (gs) command not found in PATH")
            print("Please ensure Ghostscript is installed and available in PATH")
            return False
        
        print(f"Using Ghostscript at: {gs_path}")
        
        # Create a temporary PNG file first
        temp_png = output_file.replace('.webp', '_temp.png')
        
        try:
            # Convert EPS to PNG using Ghostscript
            gs_cmd = [
                gs_path,
                '-dNOPAUSE',
                '-dBATCH',
                '-dSAFER',
                '-sDEVICE=png16m',
                '-r300',  # 300 DPI
                f'-sOutputFile={temp_png}',
                eps_file
            ]
            
            print(f"Running Ghostscript command: {' '.join(gs_cmd)}")
            result = subprocess.run(gs_cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode != 0:
                print(f"ERROR: Ghostscript failed: {result.stderr}")
                return False
            
            if not os.path.exists(temp_png):
                print("ERROR: Ghostscript did not create PNG file")
                return False
            
            print(f"Ghostscript created PNG: {os.path.getsize(temp_png)} bytes")
            
            # Now convert PNG to WebP using Pillow
            with Image.open(temp_png) as png_image:
                if png_image.mode != 'RGB':
                    png_image = png_image.convert('RGB')
                
                webp_options = {
                    'format': 'WebP',
                    'quality': quality,
                    'lossless': lossless,
                    'method': 6
                }
                
                if lossless:
                    print("Using lossless compression")
                    webp_options['quality'] = 100
                else:
                    print(f"Using lossy compression with quality {quality}")
                
                png_image.save(output_file, **webp_options)
                print("WebP file saved successfully via Ghostscript")
            
            # Clean up temporary PNG file
            os.remove(temp_png)
            
            # Verify the output file
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"WebP file created successfully: {file_size} bytes")
                
                # Verify it's a valid WebP file
                try:
                    with Image.open(output_file) as verify_img:
                        print(f"Verified WebP file: {verify_img.size}, mode: {verify_img.mode}")
                        return True
                except Exception as verify_error:
                    print(f"ERROR: Output file is not a valid WebP: {verify_error}")
                    return False
            else:
                print("ERROR: WebP file was not created")
                return False
                
        except Exception as gs_error:
            print(f"ERROR: Ghostscript conversion failed: {gs_error}")
            # Clean up temp file if it exists
            if os.path.exists(temp_png):
                os.remove(temp_png)
            return False
                    
    except Exception as e:
        print(f"ERROR: Failed to convert EPS to WebP: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert EPS file to WebP format')
    parser.add_argument('eps_file', help='Path to input EPS file')
    parser.add_argument('output_file', help='Path to output WebP file')
    parser.add_argument('--quality', type=int, default=80, choices=range(0, 101),
                        help='WebP quality 0-100 (default: 80)')
    parser.add_argument('--lossless', action='store_true',
                        help='Use lossless compression')
    
    args = parser.parse_args()
    
    print("=== EPS to WebP Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_eps_to_webp(args.eps_file, args.output_file, args.quality, args.lossless)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == '__main__':
    main()
