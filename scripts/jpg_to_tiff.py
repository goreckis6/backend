#!/usr/bin/env python3
"""
JPG to TIFF Converter
- Reads JPG/JPEG using Pillow
- Saves TIFF using Pillow with lossless compression
- Optional downscale by max dimension
- Supports lossless compression (TIFF is lossless by default)
- Optimized for professional printing and archival purposes
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


def convert_jpg_to_tiff(
    jpg_file: str,
    output_file: str,
    compression: str = "lzw",
    lossless: bool = True,
    max_dimension: int = 4096
) -> bool:
    """
    Convert JPG to TIFF format optimized for professional printing and archival
    
    Args:
        jpg_file: Input JPG file path
        output_file: Output TIFF file path
        compression: TIFF compression method ('lzw', 'zip', 'jpeg', 'none', default: 'lzw')
        lossless: Use lossless compression (default: True, TIFF is lossless by default)
        max_dimension: Maximum width or height (downscale if larger)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print("Starting JPG → TIFF conversion")
    print(f"Input file: {jpg_file}")
    print(f"Output file: {output_file}")
    print(f"Compression: {compression}, Lossless: {lossless}, Max dimension: {max_dimension}")

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
                Image.Resampling.LANCZOS  # High-quality resampling for TIFF
            )
            print(f"Resized to {img.size}")

        # Convert color modes for TIFF
        # TIFF supports RGB, RGBA, L (grayscale), LA (grayscale with alpha), CMYK, LAB
        if img.mode == "CMYK":
            # CMYK to RGB conversion
            img = img.convert("RGB")
            print("Converted CMYK to RGB")
        elif img.mode == "L":
            # Grayscale - keep as-is, TIFF supports grayscale
            print("Grayscale mode preserved")
        elif img.mode == "P":
            # Palette mode - convert to RGB for better compatibility
            img = img.convert("RGB")
            print("Converted palette to RGB")
        elif img.mode == "CMYK":
            # CMYK - keep as-is for professional printing (TIFF supports CMYK)
            print("CMYK mode preserved for professional printing")
        elif img.mode not in ["RGB", "RGBA", "L", "LA", "CMYK", "LAB"]:
            # Convert other modes to RGB
            img = img.convert("RGB")
            print(f"Converted {img.mode} to RGB")

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save as TIFF with optimized settings
        # TIFF compression: 'lzw' (lossless, good compression), 'zip' (lossless, better compression),
        # 'jpeg' (lossy, smaller files), 'none' (no compression, largest files)
        save_kwargs = {
            "format": "TIFF",
        }
        
        # Set compression method
        if compression == "lzw":
            save_kwargs["compression"] = "tiff_lzw"
            print("Using LZW compression (lossless, good compression)")
        elif compression == "zip":
            save_kwargs["compression"] = "tiff_adobe_deflate"
            print("Using ZIP/Deflate compression (lossless, better compression)")
        elif compression == "jpeg":
            save_kwargs["compression"] = "tiff_jpeg"
            print("Using JPEG compression (lossy, smaller files)")
        elif compression == "none":
            save_kwargs["compression"] = "tiff_lzw"  # Use LZW as default even if "none" requested
            print("Using LZW compression (TIFF default)")
        else:
            save_kwargs["compression"] = "tiff_lzw"
            print(f"Using LZW compression (default, requested '{compression}' not recognized)")
        
        img.save(output_file, **save_kwargs)
        
        output_size = os.path.getsize(output_file)
        print(f"TIFF saved successfully: {output_size} bytes")
        print(f"Size ratio: {output_size / file_size:.2%}")
        if output_size < file_size:
            print(f"Compression: {((1 - output_size / file_size) * 100):.1f}% smaller")
        else:
            print(f"Size increase: {((output_size / file_size - 1) * 100):.1f}% larger (lossless quality)")

        return os.path.exists(output_file) and output_size > 0

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert JPG/JPEG to TIFF")
    parser.add_argument("jpg_file", help="Input JPG file path")
    parser.add_argument("output_file", help="Output TIFF file path")
    parser.add_argument("--compression", type=str, default="lzw",
                        choices=["lzw", "zip", "jpeg", "none"],
                        help="TIFF compression method: 'lzw' (lossless, default), 'zip' (lossless, better), 'jpeg' (lossy), 'none' (no compression)")
    parser.add_argument("--lossless", action="store_true", default=True,
                        help="Use lossless compression (default: True, TIFF is lossless by default)")
    parser.add_argument("--max-dimension", type=int, default=4096,
                        help="Maximum width or height (default: 4096)")
    args = parser.parse_args()

    ok = convert_jpg_to_tiff(
        args.jpg_file,
        args.output_file,
        args.compression,
        args.lossless,
        args.max_dimension
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
