#!/usr/bin/env python3
"""
HEIC to JPG Converter (backend-friendly)
- Reads HEIC/HEIF via pillow-heif
- Saves JPEG using Pillow
- Optional downscale by max dimension
- Correct EXIF orientation
"""

import os
import sys
import argparse
import traceback

try:
    from PIL import Image, ImageOps
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from pillow_heif import register_heif_opener
    HAS_PILLOW_HEIF = True
except ImportError:
    HAS_PILLOW_HEIF = False


_HEIF_REGISTERED = False


def _ensure_heif():
    global _HEIF_REGISTERED
    if HAS_PILLOW_HEIF and not _HEIF_REGISTERED:
        register_heif_opener()
        _HEIF_REGISTERED = True


def convert_heic_to_jpg(
    heic_file: str,
    output_file: str,
    quality: int = 90,
    max_dimension: int = 4096
) -> bool:
    print("Starting HEIC → JPG conversion")

    if not HAS_PIL:
        print("ERROR: Pillow not available")
        return False

    if not os.path.exists(heic_file):
        print(f"ERROR: File not found: {heic_file}")
        return False

    try:
        _ensure_heif()

        img = Image.open(heic_file)
        print(f"Opened: format={img.format}, mode={img.mode}, size={img.size}")

        # Fix iPhone orientation
        img = ImageOps.exif_transpose(img)

        # Downscale if needed
        w, h = img.size
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            img = img.resize(
                (int(w * scale), int(h * scale)),
                Image.Resampling.BILINEAR
            )
            print(f"Resized to {img.size}")

        # JPG requires RGB
        if img.mode != "RGB":
            img = img.convert("RGB")

        # Ensure output dir
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save JPG
        img.save(
            output_file,
            format="JPEG",
            quality=quality,
            subsampling=0 if quality >= 90 else 2,
            optimize=True
        )

        return os.path.exists(output_file) and os.path.getsize(output_file) > 0

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert HEIC/HEIF to JPG")
    parser.add_argument("heic_file")
    parser.add_argument("output_file")
    parser.add_argument("--quality", type=int, default=90)
    parser.add_argument("--max-dimension", type=int, default=4096)
    args = parser.parse_args()

    ok = convert_heic_to_jpg(
        args.heic_file,
        args.output_file,
        args.quality,
        args.max_dimension
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
