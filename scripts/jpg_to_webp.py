#!/usr/bin/env python3
"""
JPG to WebP Converter
- Reads JPG/JPEG using Pillow
- Saves WebP using Pillow with optimized compression
- Optional downscale by max dimension
- Supports lossy and lossless compression
- Optimized for web use with quality/compression balance
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


def convert_jpg_to_webp(
    jpg_file: str,
    output_file: str,
    quality: int = 90,
    lossless: bool = False,
    max_dimension: int = 4096,
    method: int = 6
) -> bool:
    """
    Convert JPG to WebP format optimized for web use
    
    Args:
        jpg_file: Input JPG file path
        output_file: Output WebP file path
        quality: WebP quality (0-100), ignored if lossless=True
        lossless: Use lossless compression (default: False for web optimization)
        max_dimension: Maximum width or height (downscale if larger)
        method: WebP compression method (0-6, default 6 for best compression)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print("Starting JPG → WebP conversion")
    print(f"Input file: {jpg_file}")
    print(f"Output file: {output_file}")
    print(f"Quality: {quality}, Lossless: {lossless}, Max dimension: {max_dimension}")

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
                Image.Resampling.LANCZOS  # High-quality resampling for WebP
            )
            print(f"Resized to {img.size}")

        # Convert color modes for WebP
        # WebP supports RGB, RGBA, L (grayscale), LA (grayscale with alpha)
        if img.mode == "CMYK":
            # CMYK to RGB conversion
            img = img.convert("RGB")
            print("Converted CMYK to RGB")
        elif img.mode == "L":
            # Grayscale - keep as-is, WebP supports grayscale
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

        # Save as WebP with optimized settings
        # WebP method: 0 (fastest) to 6 (best compression, slower)
        # Quality: 0-100 (for lossy), ignored if lossless=True
        save_kwargs = {
            "format": "WEBP",
            "method": method,  # Compression method (0-6)
            "optimize": True,  # Optimize WebP compression
        }
        
        if lossless:
            save_kwargs["lossless"] = True
            print("Using lossless WebP compression")
        else:
            save_kwargs["quality"] = quality
            print(f"Using lossy WebP compression with quality {quality}")
        
        img.save(output_file, **save_kwargs)
        
        output_size = os.path.getsize(output_file)
        print(f"WebP saved successfully: {output_size} bytes")
        print(f"Size ratio: {output_size / file_size:.2%}")
        print(f"Compression: {((1 - output_size / file_size) * 100):.1f}% smaller")

        return os.path.exists(output_file) and output_size > 0

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert JPG/JPEG to WebP")
    parser.add_argument("jpg_file", help="Input JPG file path")
    parser.add_argument("output_file", help="Output WebP file path")
    parser.add_argument("--quality", type=int, default=90, choices=range(0, 101),
                        help="WebP quality 0-100 (default: 90, ignored if --lossless)")
    parser.add_argument("--lossless", action="store_true",
                        help="Use lossless WebP compression")
    parser.add_argument("--max-dimension", type=int, default=4096,
                        help="Maximum width or height (default: 4096)")
    parser.add_argument("--method", type=int, default=6, choices=range(0, 7),
                        help="WebP compression method 0-6 (default: 6 for best compression)")
    args = parser.parse_args()

    ok = convert_jpg_to_webp(
        args.jpg_file,
        args.output_file,
        args.quality,
        args.lossless,
        args.max_dimension,
        args.method
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
