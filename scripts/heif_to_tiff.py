#!/usr/bin/env python3
"""
HEIF/HEIC to TIFF Converter
Reads HEIF/HEIC via pillow-heif and saves TIFF using Pillow.
Supports adjustable quality and max-dimension downscale.
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

try:
    from pillow_heif import register_heif_opener
    HAS_PILLOW_HEIF = True
except ImportError:
    HAS_PILLOW_HEIF = False


def convert_heif_to_tiff(heif_file: str, output_file: str, quality: int = 90, max_dimension: int = 4096) -> bool:
    print("Starting HEIF/HEIC to TIFF conversion")
    print(f"Input: {heif_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Max dimension: {max_dimension}")

    if not HAS_PIL:
        print("ERROR: Pillow (PIL) not available")
        return False

    if HAS_PILLOW_HEIF:
        try:
            register_heif_opener()
            print("HEIF opener registered")
        except Exception as e:
            print(f"Warning: Could not register HEIF opener: {e}")

    if not os.path.exists(heif_file):
        print(f"ERROR: Input file not found: {heif_file}")
        return False

    try:
        img = Image.open(heif_file)
        print(f"Opened image. Format={img.format}, Mode={img.mode}, Size={img.size}")

        # Downscale large images to speed up processing
        width, height = img.size
        if max(width, height) > max_dimension:
            if width > height:
                new_w = max_dimension
                new_h = int(height * (max_dimension / width))
            else:
                new_h = max_dimension
                new_w = int(width * (max_dimension / height))
            img = img.resize((new_w, new_h), Image.Resampling.BILINEAR)
            print(f"Resized to {new_w}x{new_h}")

        # TIFF supports transparency - preserve RGBA mode if present
        if img.mode in ("RGBA", "LA"):
            print("Preserving transparency (RGBA mode)")
        elif img.mode == 'P' and 'transparency' in img.info:
            print("Converting palette with transparency to RGBA")
            img = img.convert("RGBA")
        elif img.mode not in ("RGB", "RGBA"):
            print(f"Converting {img.mode} to RGB")
            img = img.convert("RGB")

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save as TIFF with LZW compression
        # TIFF quality is not directly supported, but we use compression
        compression = "tiff_lzw" if quality >= 80 else "tiff_deflate"
        img.save(output_file, format='TIFF', compression=compression)

        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("TIFF created successfully")
            return True
        print("ERROR: TIFF output not created")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert HEIF/HEIC to TIFF: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIF/HEIC image to TIFF')
    parser.add_argument('heif_file', help='Path to input HEIF/HEIC file')
    parser.add_argument('output_file', help='Path to output TIFF file')
    parser.add_argument('--quality', type=int, default=90, help='Quality setting (affects compression)')
    parser.add_argument('--max-dimension', type=int, default=4096, help='Max width or height for downscaling')
    args = parser.parse_args()

    print("=== HEIF/HEIC to TIFF Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heif_to_tiff(args.heif_file, args.output_file, args.quality, args.max_dimension)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
