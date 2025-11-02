#!/usr/bin/env python3
"""
JPG/JPEG Image Compressor
Compresses JPG/JPEG images using Pillow with adjustable quality settings.
Supports optimization flags for better compression.
"""

import os
import sys
import argparse
import traceback

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def compress_jpg(input_file: str, output_file: str, quality: int = 85, optimize: bool = True) -> tuple[bool, int, int]:
    """
    Compress JPG/JPEG image.
    
    Args:
        input_file: Path to input JPG/JPEG file
        output_file: Path to output compressed JPG file
        quality: JPEG quality (1-100, higher = better quality but larger file)
        optimize: Apply additional optimization
    
    Returns:
        tuple: (success: bool, original_size: int, compressed_size: int)
    """
    print("Starting JPG/JPEG compression")
    print(f"Input: {input_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Optimize: {optimize}")

    if not HAS_PIL:
        print("ERROR: Pillow (PIL) not available")
        return False, 0, 0

    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        return False, 0, 0

    try:
        # Get original file size
        original_size = os.path.getsize(input_file)
        print(f"Original file size: {original_size} bytes ({original_size / 1024:.2f} KB)")

        # Open image
        img = Image.open(input_file)
        print(f"Opened image. Format={img.format}, Mode={img.mode}, Size={img.size}")

        # Verify it's a JPEG
        if img.format not in ('JPEG', 'JPG'):
            # Try to convert to RGB if needed
            if img.mode != 'RGB':
                print(f"Converting from {img.mode} to RGB")
                if img.mode in ('RGBA', 'LA') or 'transparency' in img.info:
                    # Flatten transparency on white background
                    bg = Image.new("RGB", img.size, (255, 255, 255))
                    if img.mode == 'RGBA':
                        bg.paste(img, mask=img.split()[3])
                    else:
                        bg.paste(img)
                    img = bg
                else:
                    img = img.convert("RGB")
        elif img.mode != 'RGB':
            # Ensure RGB mode for JPEG
            if img.mode in ('RGBA', 'LA'):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == 'RGBA':
                    bg.paste(img, mask=img.split()[3])
                else:
                    bg.paste(img)
                img = bg
            else:
                img = img.convert("RGB")

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save with compression
        # Clamp quality between 1 and 100
        quality_value = max(1, min(100, int(quality)))
        
        save_kwargs = {
            'format': 'JPEG',
            'quality': quality_value,
            'optimize': optimize
        }
        
        img.save(output_file, **save_kwargs)

        # Get compressed file size
        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            compressed_size = os.path.getsize(output_file)
            savings_percent = ((original_size - compressed_size) / original_size * 100) if original_size > 0 else 0
            print(f"Compressed file size: {compressed_size} bytes ({compressed_size / 1024:.2f} KB)")
            print(f"Compression savings: {savings_percent:.2f}% ({original_size - compressed_size} bytes saved)")
            print("JPG compressed successfully")
            return True, original_size, compressed_size
        
        print("ERROR: Compressed output not created")
        return False, original_size, 0
    except Exception as e:
        print(f"ERROR: Failed to compress JPG: {e}")
        traceback.print_exc()
        return False, 0, 0


def main():
    parser = argparse.ArgumentParser(description='Compress JPG/JPEG image')
    parser.add_argument('input_file', help='Path to input JPG/JPEG file')
    parser.add_argument('output_file', help='Path to output compressed JPG file')
    parser.add_argument('--quality', type=int, default=85, help='JPEG quality (1-100, default: 85)')
    parser.add_argument('--optimize', action='store_true', default=True, help='Apply optimization (default: True)')
    parser.add_argument('--no-optimize', action='store_false', dest='optimize', help='Disable optimization')
    args = parser.parse_args()

    print("=== JPG/JPEG Compressor ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    success, original_size, compressed_size = compress_jpg(
        args.input_file, 
        args.output_file, 
        args.quality, 
        args.optimize
    )
    
    if success:
        print(f"SUCCESS: Original={original_size} bytes, Compressed={compressed_size} bytes")
        sys.exit(0)
    else:
        print("FAILURE: Compression failed")
        sys.exit(1)


if __name__ == '__main__':
    main()

