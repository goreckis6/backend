#!/usr/bin/env python3
"""
HEIF/HEIC to JPG Converter
Reads HEIF/HEIC via pillow-heif and saves JPEG using Pillow.
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


def convert_heif_to_jpg(heif_file: str, output_file: str, quality: int = 90, max_dimension: int = 4096) -> bool:
    print("Starting HEIF/HEIC to JPG conversion")
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

        # JPEG does not support alpha channel; flatten on white background
        if img.mode in ("RGBA", "LA") or 'transparency' in img.info:
            print("Flattening transparency over white for JPEG")
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == 'RGBA':
                bg.paste(img, mask=img.split()[3])
            else:
                bg.paste(img)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        img.save(output_file, format='JPEG', quality=max(1, min(100, int(quality))), optimize=False)

        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("JPG created successfully")
            return True
        print("ERROR: JPG output not created")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert HEIF/HEIC to JPG: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIF/HEIC image to JPG')
    parser.add_argument('heif_file', help='Path to input HEIF/HEIC file')
    parser.add_argument('output_file', help='Path to output JPG file')
    parser.add_argument('--quality', type=int, default=90, help='JPEG quality (1-100)')
    parser.add_argument('--max-dimension', type=int, default=4096, help='Max width or height for downscaling')
    args = parser.parse_args()

    print("=== HEIF/HEIC to JPG Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heif_to_jpg(args.heif_file, args.output_file, args.quality, args.max_dimension)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()


