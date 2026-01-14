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
    print(f"Input file: {heic_file}")
    print(f"Output file: {output_file}")

    if not HAS_PIL:
        print("ERROR: Pillow not available")
        return False

    if not HAS_PILLOW_HEIF:
        print("ERROR: pillow-heif not available")
        return False

    if not os.path.exists(heic_file):
        print(f"ERROR: File not found: {heic_file}")
        return False

    # Check file size
    file_size = os.path.getsize(heic_file)
    print(f"File size: {file_size} bytes")
    
    if file_size == 0:
        print("ERROR: File is empty")
        return False

    # Check file signature (HEIC files start with specific bytes)
    try:
        with open(heic_file, 'rb') as f:
            header = f.read(12)
            header_hex = header.hex()
            print(f"File header (hex): {header_hex}")
            
            # HEIC/HEIF files typically start with:
            # - ftyp box: bytes 4-8 contain "ftyp"
            # - Then brand identifier like "heic", "heif", "mif1", etc.
            if len(header) >= 8:
                ftyp_pos = header.find(b'ftyp')
                if ftyp_pos >= 0 and ftyp_pos < 8:
                    brand_start = ftyp_pos + 4
                    if len(header) > brand_start + 4:
                        brand = header[brand_start:brand_start+4]
                        print(f"Detected brand: {brand}")
                        if brand not in [b'heic', b'heif', b'mif1', b'msf1']:
                            print(f"WARNING: Unusual brand identifier: {brand}")
                else:
                    print("WARNING: 'ftyp' box not found in expected position")
    except Exception as read_error:
        print(f"WARNING: Could not read file header: {read_error}")

    try:
        _ensure_heif()
        print("pillow-heif registered successfully")

        # Try to open the image
        try:
            img = Image.open(heic_file)
            print(f"Opened: format={img.format}, mode={img.mode}, size={img.size}")
            
            # Verify it's actually a HEIC/HEIF image
            if img.format not in ['HEIC', 'HEIF']:
                print(f"WARNING: Image format is {img.format}, not HEIC/HEIF")
        except Exception as open_error:
            # Log detailed error to console (for server logs) but don't expose file paths in user-facing messages
            print(f"ERROR: Failed to open image: {open_error}")
            print(f"ERROR: File path: {heic_file}")
            print(f"ERROR: File exists: {os.path.exists(heic_file)}")
            # Print user-friendly error message (without file paths)
            print("ERROR: The file is corrupted or not a valid HEIC image")
            raise

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
