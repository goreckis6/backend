#!/usr/bin/env python3
"""
TIFF to PNG preview converter using Python + Pillow
Converts TIFF files to PNG for web preview.
"""

import argparse
import os
import sys
import traceback
from PIL import Image

def convert_tiff_to_png(tiff_file, output_file, max_dimension=2048):
    """
    Convert TIFF file to PNG format for web preview.

    Args:
        tiff_file (str): Path to input TIFF file
        output_file (str): Path to output PNG file
        max_dimension (int): Maximum width/height for preview (default: 2048)

    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting TIFF to PNG preview conversion...")
    print(f"Input: {tiff_file}")
    print(f"Output: {output_file}")
    print(f"Max dimension: {max_dimension}")

    try:
        # Open TIFF file with Pillow
        print("Opening TIFF file...")
        with Image.open(tiff_file) as img:
            print(f"Image size: {img.size}")
            print(f"Image mode: {img.mode}")
            print(f"Image format: {img.format}")

            # Convert to RGB if necessary (for web compatibility)
            if img.mode not in ('RGB', 'RGBA', 'L'):
                print(f"Converting from {img.mode} to RGB...")
                if img.mode == 'CMYK':
                    # Convert CMYK to RGB
                    img = img.convert('RGB')
                elif img.mode == 'I' or img.mode == 'I;16':
                    # Convert 16-bit to 8-bit
                    img = img.convert('L')
                elif 'A' in img.mode:
                    img = img.convert('RGBA')
                else:
                    img = img.convert('RGB')
                print(f"Converted to: {img.mode}")

            # Resize if image is too large
            width, height = img.size
            if width > max_dimension or height > max_dimension:
                print(f"Resizing from {width}x{height}...")
                # Calculate new size maintaining aspect ratio
                if width > height:
                    new_width = max_dimension
                    new_height = int(height * (max_dimension / width))
                else:
                    new_height = max_dimension
                    new_width = int(width * (max_dimension / height))
                
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                print(f"Resized to: {img.size}")

            # Save as PNG
            print("Saving as PNG...")
            img.save(output_file, format='PNG', optimize=True)

            # Verify the output file
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"PNG file created successfully: {file_size} bytes")
                return True
            else:
                print("ERROR: PNG file was not created")
                return False

    except Exception as e:
        print(f"ERROR: Failed to convert TIFF to PNG: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert TIFF to PNG for web preview')
    parser.add_argument('tiff_file', help='Input TIFF file path')
    parser.add_argument('output_file', help='Output PNG file path')
    parser.add_argument('--max-dimension', type=int, default=2048,
                        help='Maximum width/height for preview (default: 2048)')

    args = parser.parse_args()

    print("=== TIFF to PNG Preview Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")

    # Check if input file exists
    if not os.path.exists(args.tiff_file):
        print(f"ERROR: Input TIFF file not found: {args.tiff_file}")
        sys.exit(1)

    # Check required libraries
    try:
        from PIL import Image
        print(f"Pillow version: {Image.__version__}")
    except ImportError as e:
        print(f"ERROR: Pillow not available: {e}")
        print("Please install Pillow: pip install Pillow")
        sys.exit(1)

    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Convert TIFF to PNG
    success = convert_tiff_to_png(
        args.tiff_file,
        args.output_file,
        max_dimension=args.max_dimension
    )

    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()


