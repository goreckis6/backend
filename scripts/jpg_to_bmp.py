#!/usr/bin/env python3
"""
JPG to BMP Converter
- Reads JPG/JPEG using Pillow
- Saves BMP using Pillow
- Optional downscale by max dimension
- Preserves quality with uncompressed BMP format
- BMP is uncompressed, so files will be larger than JPG
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


def convert_jpg_to_bmp(
    jpg_file: str,
    output_file: str,
    max_dimension: int = 4096
) -> bool:
    """
    Convert JPG to BMP format (uncompressed)
    
    Args:
        jpg_file: Input JPG file path
        output_file: Output BMP file path
        max_dimension: Maximum width or height (downscale if larger)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print("Starting JPG → BMP conversion")
    print(f"Input file: {jpg_file}")
    print(f"Output file: {output_file}")

    if not HAS_PIL:
        print("ERROR: Pillow not available")
        return False

    if not os.path.exists(jpg_file):
        print(f"ERROR: File not found: {jpg_file}")
        return False

    # Check file size
    file_size = os.path.getsize(jpg_file)
    print(f"File size: {file_size} bytes")
    
    if file_size == 0:
        print("ERROR: File is empty")
        return False

    try:
        # Try to open the image
        try:
            img = Image.open(jpg_file)
            print(f"Opened: format={img.format}, mode={img.mode}, size={img.size}")
            
            # Verify it's actually a JPG/JPEG image
            if img.format not in ['JPEG', 'JPG']:
                print(f"WARNING: Image format is {img.format}, expected JPEG")
        except Exception as open_error:
            print(f"ERROR: Failed to open image: {open_error}")
            print(f"ERROR: File path: {jpg_file}")
            print(f"ERROR: File exists: {os.path.exists(jpg_file)}")
            print("ERROR: The file is corrupted or not a valid JPG image")
            raise

        # Fix orientation based on EXIF data
        img = ImageOps.exif_transpose(img)
        print(f"After orientation fix: mode={img.mode}, size={img.size}")

        # Downscale if needed
        w, h = img.size
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            new_width = int(w * scale)
            new_height = int(h * scale)
            img = img.resize(
                (new_width, new_height),
                Image.Resampling.LANCZOS  # High-quality resampling
            )
            print(f"Resized to {img.size}")

        # Convert to RGB mode (BMP typically uses RGB)
        if img.mode == "CMYK":
            # CMYK to RGB conversion
            img = img.convert("RGB")
            print("Converted CMYK to RGB")
        elif img.mode == "L":
            # Grayscale - keep as-is, BMP supports grayscale
            print("Grayscale mode preserved")
        elif img.mode in ["RGBA", "LA", "P"]:
            # Convert RGBA/LA/P to RGB (BMP doesn't support transparency well)
            # For RGBA, composite onto white background
            if img.mode == "RGBA":
                background = Image.new("RGB", img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[-1])  # Use alpha channel as mask
                img = background
                print("Converted RGBA to RGB (composited on white background)")
            else:
                img = img.convert("RGB")
                print(f"Converted {img.mode} to RGB")
        elif img.mode not in ["RGB", "L"]:
            # Convert other modes to RGB
            img = img.convert("RGB")
            print(f"Converted {img.mode} to RGB")

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save as BMP (uncompressed format)
        # BMP format doesn't have compression options
        img.save(output_file, format="BMP")
        
        output_size = os.path.getsize(output_file)
        print(f"BMP saved successfully: {output_size} bytes")
        print(f"Size ratio: {output_size / file_size:.2%}")
        print(f"Note: BMP files are uncompressed and will be larger than JPG files")

        return os.path.exists(output_file) and output_size > 0

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert JPG/JPEG to BMP")
    parser.add_argument("jpg_file", help="Input JPG file path")
    parser.add_argument("output_file", help="Output BMP file path")
    parser.add_argument("--max-dimension", type=int, default=4096,
                        help="Maximum width or height (default: 4096)")
    args = parser.parse_args()

    ok = convert_jpg_to_bmp(
        args.jpg_file,
        args.output_file,
        args.max_dimension
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
