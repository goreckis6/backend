#!/usr/bin/env python3
"""
HEIF/HEIC to AVIF Converter
Reads HEIF/HEIC via pillow-heif and saves AVIF using Pillow.
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


def convert_heif_to_avif(heif_file: str, output_file: str, quality: int = 90, max_dimension: int = 4096) -> bool:
    print("Starting HEIF/HEIC to AVIF conversion")
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

        # AVIF supports transparency - preserve RGBA mode if present
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

        # Save as AVIF with quality setting
        # AVIF uses 0-100 quality scale where higher is better
        img.save(output_file, format='AVIF', quality=max(1, min(100, int(quality))))

        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("AVIF created successfully")
            return True
        print("ERROR: AVIF output not created")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert HEIF/HEIC to AVIF: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIF/HEIC image to AVIF')
    parser.add_argument('heif_file', help='Path to input HEIF/HEIC file')
    parser.add_argument('output_file', help='Path to output AVIF file')
    parser.add_argument('--quality', type=int, default=90, help='AVIF quality (1-100)')
    parser.add_argument('--max-dimension', type=int, default=4096, help='Max width or height for downscaling')
    args = parser.parse_args()

    print("=== HEIF/HEIC to AVIF Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heif_to_avif(args.heif_file, args.output_file, args.quality, args.max_dimension)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
