#!/usr/bin/env python3
"""
HEIC to PNG Converter
Reads HEIC/HEIF via pillow-heif and saves PNG using Pillow
Optimized to avoid heavy processing; supports optional max-dimension downscale
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


def convert_heic_to_png(heic_file: str, output_file: str, quality: int = 95, max_dimension: int = 4096) -> bool:
    print(f"Starting HEIC to PNG conversion")
    print(f"Input: {heic_file}")
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

        # PNG supports alpha; ensure mode is RGBA or RGB
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA" if (img.mode in ("LA",) or 'transparency' in img.info) else "RGB")

        # Optional palette quantization to reduce size (lossy, good for non-photo or when transparency not critical)
        # Apply for medium/low quality levels and only when image is RGB (no alpha) to avoid poor alpha results
        try:
            if img.mode == "RGB" and quality <= 85:
                # Number of colors based on quality hint
                colors = 256 if quality > 75 else 128
                print(f"Applying palette quantization to {colors} colors for size reduction")
                img = img.quantize(colors=colors, method=Image.MEDIANCUT, dither=Image.FLOYDSTEINBERG)
        except Exception as qerr:
            print(f"Warning: quantization skipped due to error: {qerr}")

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Map quality (0-100) to PNG compression (0-9); higher compression for lower quality
        if quality >= 90:
            compression_level = 3
        elif quality >= 80:
            compression_level = 6
        else:
            compression_level = 9
        print(f"Using PNG compress_level={compression_level} optimize=True")

        # Save PNG with optimization enabled
        img.save(output_file, format='PNG', optimize=True, compress_level=compression_level)

        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("PNG created successfully")
            return True
        print("ERROR: PNG output not created")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert HEIC to PNG: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIC/HEIF image to PNG')
    parser.add_argument('heic_file', help='Path to input HEIC/HEIF file')
    parser.add_argument('output_file', help='Path to output PNG file')
    parser.add_argument('--quality', type=int, default=95, help='Quality (0-100, maps to PNG compress_level)')
    parser.add_argument('--max-dimension', type=int, default=4096, help='Max width or height for downscaling')
    args = parser.parse_args()

    print("=== HEIC to PNG Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heic_to_png(args.heic_file, args.output_file, args.quality, args.max_dimension)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()


