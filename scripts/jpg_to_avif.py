#!/usr/bin/env python3
"""
JPG to AVIF Converter
- Reads JPG/JPEG using Pillow
- Saves AVIF using Pillow with AVIF plugin
- Optional downscale by max dimension
- Supports quality and effort settings
- Optimized for web use with superior compression
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
    from pillow_avif import register_avif_opener
    register_avif_opener()
    HAS_AVIF = True
except ImportError:
    HAS_AVIF = False


def convert_jpg_to_avif(
    jpg_file: str,
    output_file: str,
    quality: int = 50,
    max_dimension: int = 4096,
    effort: int = 4
) -> bool:
    """
    Convert JPG to AVIF format optimized for web use
    
    Args:
        jpg_file: Input JPG file path
        output_file: Output AVIF file path
        quality: AVIF quality (0-100, default: 50 for good compression)
        max_dimension: Maximum width or height (downscale if larger)
        effort: AVIF effort (0-9, default: 4 for balance between speed and compression)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print("Starting JPG → AVIF conversion")
    print(f"Input file: {jpg_file}")
    print(f"Output file: {output_file}")
    print(f"Quality: {quality}, Max dimension: {max_dimension}, Effort: {effort}")

    if not HAS_PIL:
        print("ERROR: Pillow not available")
        return False

    if not HAS_AVIF:
        print("ERROR: pillow-avif-plugin not available")
        print("Please install: pip install pillow-avif-plugin")
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
                Image.Resampling.LANCZOS  # High-quality resampling for AVIF
            )
            print(f"Resized to {img.size}")

        # Convert color modes for AVIF
        # AVIF supports RGB, RGBA, L (grayscale), LA (grayscale with alpha)
        if img.mode == "CMYK":
            # CMYK to RGB conversion
            img = img.convert("RGB")
            print("Converted CMYK to RGB")
        elif img.mode == "L":
            # Grayscale - keep as-is, AVIF supports grayscale
            print("Grayscale mode preserved")
        elif img.mode == "P":
            # Palette mode - convert to RGB
            img = img.convert("RGB")
            print("Converted palette to RGB")
        elif img.mode not in ["RGB", "RGBA", "L", "LA"]:
            # Convert other modes to RGB
            img = img.convert("RGB")
            print(f"Converted {img.mode} to RGB")

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save as AVIF with optimized settings
        # AVIF quality: 0-100 (default: 50 for good compression)
        # AVIF effort: 0-9 (default: 4 for balance)
        save_kwargs = {
            "format": "AVIF",
            "quality": quality,
            "speed": 9 - effort,  # speed is inverse of effort (0-9, where 0 is slowest/best)
        }
        
        print(f"Using AVIF compression with quality {quality}, effort {effort}")
        
        img.save(output_file, **save_kwargs)
        
        output_size = os.path.getsize(output_file)
        print(f"AVIF saved successfully: {output_size} bytes")
        print(f"Size ratio: {output_size / file_size:.2%}")
        print(f"Compression: {((1 - output_size / file_size) * 100):.1f}% smaller")

        return os.path.exists(output_file) and output_size > 0

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert JPG/JPEG to AVIF")
    parser.add_argument("jpg_file", help="Input JPG file path")
    parser.add_argument("output_file", help="Output AVIF file path")
    parser.add_argument("--quality", type=int, default=50, choices=range(0, 101),
                        help="AVIF quality 0-100 (default: 50)")
    parser.add_argument("--max-dimension", type=int, default=4096,
                        help="Maximum width or height (default: 4096)")
    parser.add_argument("--effort", type=int, default=4, choices=range(0, 10),
                        help="AVIF effort 0-9 (default: 4 for balance)")
    args = parser.parse_args()

    ok = convert_jpg_to_avif(
        args.jpg_file,
        args.output_file,
        args.quality,
        args.max_dimension,
        args.effort
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
