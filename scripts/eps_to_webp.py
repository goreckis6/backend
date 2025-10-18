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
        
        # Open EPS file with Pillow
        print("Opening EPS file with Pillow...")
        try:
            # Try to open EPS file - this may fail if Ghostscript is not available
            with Image.open(eps_file) as eps_image:
                print(f"EPS image info: {eps_image.size}, mode: {eps_image.mode}")
                
                # Convert to RGB if necessary
                if eps_image.mode != 'RGB':
                    print(f"Converting from {eps_image.mode} to RGB...")
                    eps_image = eps_image.convert('RGB')
                
                # Save as WebP
                print("Saving WebP file...")
                
                # Prepare WebP save options
                webp_options = {
                    'format': 'WebP',
                    'quality': quality,
                    'lossless': lossless,
                    'method': 6  # Best compression method
                }
                
                if lossless:
                    print("Using lossless compression")
                    webp_options['quality'] = 100
                else:
                    print(f"Using lossy compression with quality {quality}")
                
                eps_image.save(output_file, **webp_options)
                print("WebP file saved successfully")
                
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
                    
        except Exception as eps_error:
            print(f"ERROR: Failed to open EPS file: {eps_error}")
            print("This might not be a valid EPS file or the file is corrupted")
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
