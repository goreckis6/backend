#!/usr/bin/env python3
"""
HEIC to PNG preview converter using Python + pillow-heif + Pillow
Converts HEIC/HEIF files to PNG for web preview.
"""

import argparse
import os
import sys
import traceback
from PIL import Image

def convert_heic_to_png(heic_file, output_file, max_dimension=2048):
    """
    Convert HEIC file to PNG format for web preview.

    Args:
        heic_file (str): Path to input HEIC file
        output_file (str): Path to output PNG file
        max_dimension (int): Maximum width/height for preview (default: 2048)

    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting HEIC to PNG preview conversion...")
    print(f"Input: {heic_file}")
    print(f"Output: {output_file}")
    print(f"Max dimension: {max_dimension}")

    try:
        # Register HEIF opener with Pillow
        try:
            from pillow_heif import register_heif_opener
            register_heif_opener()
            print("pillow-heif registered successfully")
        except ImportError:
            print("WARNING: pillow-heif not available, trying direct Pillow open")

        # Open HEIC file with Pillow
        print("Opening HEIC file...")
        with Image.open(heic_file) as img:
            print(f"Image size: {img.size}")
            print(f"Image mode: {img.mode}")
            print(f"Image format: {img.format}")

            # Convert to RGB if necessary (for web compatibility)
            if img.mode not in ('RGB', 'RGBA'):
                print(f"Converting from {img.mode} to RGB...")
                if 'A' in img.mode or img.mode == 'P':
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
        print(f"ERROR: Failed to convert HEIC to PNG: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert HEIC to PNG for web preview')
    parser.add_argument('heic_file', help='Input HEIC file path')
    parser.add_argument('output_file', help='Output PNG file path')
    parser.add_argument('--max-dimension', type=int, default=2048,
                        help='Maximum width/height for preview (default: 2048)')

    args = parser.parse_args()

    print("=== HEIC to PNG Preview Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")

    # Check if input file exists
    if not os.path.exists(args.heic_file):
        print(f"ERROR: Input HEIC file not found: {args.heic_file}")
        sys.exit(1)

    # Check required libraries
    try:
        from PIL import Image
        print(f"Pillow version: {Image.__version__}")
    except ImportError as e:
        print(f"ERROR: Pillow not available: {e}")
        print("Please install Pillow: pip install Pillow")
        sys.exit(1)

    try:
        import pillow_heif
        print(f"pillow-heif available")
    except ImportError:
        print("WARNING: pillow-heif not available")
        print("Please install pillow-heif: pip install pillow-heif")

    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Convert HEIC to PNG
    success = convert_heic_to_png(
        args.heic_file,
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

