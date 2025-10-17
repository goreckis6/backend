#!/usr/bin/env python3
"""
BMP to WebP converter using Python + Pillow
Converts BMP files to WebP format with various quality options
"""

import argparse
import os
import sys
import traceback
from PIL import Image

def convert_bmp_to_webp(bmp_file, output_file, quality=80, lossless=False, method=6):
    """
    Convert BMP file to WebP format using Pillow
    
    Args:
        bmp_file (str): Path to input BMP file
        output_file (str): Path to output WebP file
        quality (int): WebP quality (0-100, ignored if lossless=True)
        lossless (bool): Use lossless compression
        method (int): Compression method (0-6, higher = better compression)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting BMP to WebP conversion...")
    print(f"Input: {bmp_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Lossless: {lossless}")
    print(f"Method: {method}")
    
    try:
        # Check if input file exists and is readable
        if not os.path.exists(bmp_file):
            print(f"ERROR: Input BMP file not found: {bmp_file}")
            return False
            
        file_size = os.path.getsize(bmp_file)
        print(f"BMP file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: BMP file is empty")
            return False
        
        # Open and convert the image
        print("Opening BMP file with Pillow...")
        with Image.open(bmp_file) as img:
            print(f"Image format: {img.format}")
            print(f"Image mode: {img.mode}")
            print(f"Image size: {img.size}")
            
            # Convert to RGB if necessary (WebP doesn't support all modes)
            if img.mode in ('RGBA', 'LA', 'P'):
                print("Converting image to RGB mode...")
                # Create a white background for transparency
                if img.mode == 'RGBA':
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    background.paste(img, mask=img.split()[-1])  # Use alpha channel as mask
                    img = background
                else:
                    img = img.convert('RGB')
            elif img.mode != 'RGB':
                print(f"Converting from {img.mode} to RGB...")
                img = img.convert('RGB')
            
            print(f"Final image mode: {img.mode}")
            print(f"Final image size: {img.size}")
            
            # Save as WebP
            print("Saving as WebP...")
            save_kwargs = {
                'format': 'WebP',
                'method': method,
                'optimize': True
            }
            
            if lossless:
                save_kwargs['lossless'] = True
            else:
                save_kwargs['quality'] = quality
            
            img.save(output_file, **save_kwargs)
            
            # Verify the output file
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"WebP file created successfully: {output_size} bytes")
                print(f"Compression ratio: {file_size/output_size:.2f}x")
                return True
            else:
                print("ERROR: WebP file was not created")
                return False
                
    except Exception as e:
        print(f"ERROR: Failed to convert BMP to WebP: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert BMP to WebP format')
    parser.add_argument('bmp_file', help='Input BMP file path')
    parser.add_argument('output_file', help='Output WebP file path')
    parser.add_argument('--quality', type=int, default=80, help='WebP quality 0-100 (default: 80)')
    parser.add_argument('--lossless', action='store_true', help='Use lossless compression')
    parser.add_argument('--method', type=int, default=6, help='Compression method 0-6 (default: 6)')
    
    args = parser.parse_args()
    
    print("=== BMP to WebP Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    try:
        from PIL import Image
        print(f"Pillow version: {Image.__version__}")
    except ImportError as e:
        print(f"ERROR: Required library not available: {e}")
        print("Please install Pillow: pip install Pillow")
        sys.exit(1)
    
    # Validate quality range
    if not 0 <= args.quality <= 100:
        print("ERROR: Quality must be between 0 and 100")
        sys.exit(1)
    
    # Validate method range
    if not 0 <= args.method <= 6:
        print("ERROR: Method must be between 0 and 6")
        sys.exit(1)
    
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    success = convert_bmp_to_webp(
        args.bmp_file,
        args.output_file,
        quality=args.quality,
        lossless=args.lossless,
        method=args.method
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()
