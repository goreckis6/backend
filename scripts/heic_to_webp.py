#!/usr/bin/env python3
"""
HEIC to WEBP Converter
Reads HEIC/HEIF via pillow-heif and saves WEBP using Pillow
Supports adjustable quality, optional lossless, and max-dimension downscale
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


def convert_heic_to_webp(heic_file: str, output_file: str, quality: int = 90, lossless: bool = False, max_dimension: int = 4096) -> bool:
    print("Starting HEIC to WEBP conversion")
    print(f"Input: {heic_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Lossless: {lossless}")
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

    if not os.path.exists(heic_file):
        print(f"ERROR: Input file not found: {heic_file}")
        return False

    try:
        img = Image.open(heic_file)
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

        # Ensure webp-compatible mode
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA" if ('transparency' in img.info or img.mode in ("LA",)) else "RGB")

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        save_kwargs = {"format": "WEBP"}
        if lossless:
            save_kwargs.update({"lossless": True, "quality": 100, "method": 4})
        else:
            # quality 0-100; method 4 is a good speed/quality balance
            save_kwargs.update({"quality": max(1, min(100, int(quality))), "method": 4})

        img.save(output_file, **save_kwargs)

        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("WEBP created successfully")
            return True
        print("ERROR: WEBP output not created")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert HEIC to WEBP: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIC/HEIF image to WEBP')
    parser.add_argument('heic_file', help='Path to input HEIC/HEIF file')
    parser.add_argument('output_file', help='Path to output WEBP file')
    parser.add_argument('--quality', type=int, default=90, help='WEBP quality (1-100). Ignored when --lossless set')
    parser.add_argument('--lossless', action='store_true', help='Use lossless WEBP')
    parser.add_argument('--max-dimension', type=int, default=4096, help='Max width or height for downscaling')
    args = parser.parse_args()

    print("=== HEIC to WEBP Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heic_to_webp(args.heic_file, args.output_file, args.quality, args.lossless, args.max_dimension)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()


