#!/usr/bin/env python3
"""
PNG Image Compressor
Compresses PNG images using Pillow with adjustable quality settings.
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


def compress_png(input_file: str, output_file: str, quality: int = 85, optimize: bool = True) -> tuple[bool, int, int]:
    """
    Compress PNG image.
    
    Args:
        input_file: Path to input PNG file
        output_file: Path to output compressed PNG file
        quality: Compression quality (1-100, higher = better quality but larger file)
        optimize: Apply additional optimization
    
    Returns:
        tuple: (success: bool, original_size: int, compressed_size: int)
    """
    print("Starting PNG compression")
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

        # Verify it's a PNG
        if img.format != 'PNG':
            print(f"WARNING: Image format is {img.format}, converting to PNG")
        
        # PNG supports RGBA, RGB, P, L, LA, etc.
        # Preserve the original mode if it's PNG-compatible
        if img.mode not in ('RGB', 'RGBA', 'P', 'L', 'LA'):
            print(f"Converting from {img.mode} to RGB")
            if img.mode in ('RGBA', 'LA') or 'transparency' in img.info:
                # Keep transparency for RGBA
                if img.mode == 'RGBA':
                    pass  # Already RGBA
                elif img.mode == 'LA':
                    # Convert LA to RGBA
                    img = img.convert('RGBA')
                else:
                    # Convert to RGBA with transparency
                    img = img.convert('RGBA')
            else:
                img = img.convert("RGB")

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save with compression
        # Clamp quality between 1 and 100
        # For PNG, quality is interpreted as compression level (0-9, where 9 is highest compression)
        # We map 1-100 to 0-9: 1-10 -> 0, 11-20 -> 1, ..., 91-100 -> 9
        compression_level = min(9, max(0, (quality - 1) // 10))
        
        save_kwargs = {
            'format': 'PNG',
            'optimize': optimize,
            'compress_level': compression_level
        }
        
        img.save(output_file, **save_kwargs)

        # Get compressed file size
        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            compressed_size = os.path.getsize(output_file)
            savings_percent = ((original_size - compressed_size) / original_size * 100) if original_size > 0 else 0
            print(f"Compressed file size: {compressed_size} bytes ({compressed_size / 1024:.2f} KB)")
            print(f"Compression savings: {savings_percent:.2f}% ({original_size - compressed_size} bytes saved)")
            print("PNG compressed successfully")
            return True, original_size, compressed_size
        
        print("ERROR: Compressed output not created")
        return False, original_size, 0
    except Exception as e:
        print(f"ERROR: Failed to compress PNG: {e}")
        traceback.print_exc()
        return False, 0, 0


def main():
    parser = argparse.ArgumentParser(description='Compress PNG image')
    parser.add_argument('input_file', help='Path to input PNG file')
    parser.add_argument('output_file', help='Path to output compressed PNG file')
    parser.add_argument('--quality', type=int, default=85, help='PNG compression quality (1-100, default: 85)')
    parser.add_argument('--optimize', action='store_true', default=True, help='Apply optimization (default: True)')
    parser.add_argument('--no-optimize', action='store_false', dest='optimize', help='Disable optimization')
    args = parser.parse_args()

    print("=== PNG Compressor ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    success, original_size, compressed_size = compress_png(
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

