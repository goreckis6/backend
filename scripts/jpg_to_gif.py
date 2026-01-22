#!/usr/bin/env python3
"""
JPG → GIF converter (backend-ready)

- Reads JPG/JPEG using Pillow
- Saves GIF using Pillow with palette optimization
- EXIF orientation fix
- Optional max-dimension downscale
- Stateless and safe for workers / API
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


def convert_jpg_to_gif(
    input_file: str,
    output_file: str,
    quality: int = 85,
    max_dimension: int = 4096,
    optimize: bool = True,
    palette_size: int = 256
) -> bool:
    """
    Convert JPG/JPEG image to GIF format.
    
    Args:
        input_file: Path to input JPG/JPEG file
        output_file: Path to output GIF file
        quality: Quality parameter (0-100, default: 85). Note: GIF uses palette quantization.
        max_dimension: Maximum width or height (will downscale if exceeded)
        optimize: Use GIF optimization (default: True)
        palette_size: Number of colors in palette (default: 256, max: 256)
    
    Returns:
        True if conversion successful, False otherwise
    """
    if not HAS_PIL:
        print("ERROR: Pillow not available. Please install with: pip install Pillow", file=sys.stderr)
        return False

    if not os.path.isfile(input_file):
        print(f"ERROR: File not found: {input_file}", file=sys.stderr)
        return False

    try:
        # Validate quality range
        quality = max(0, min(100, quality))
        
        # Validate palette size
        palette_size = max(2, min(256, palette_size))

        img = Image.open(input_file)
        
        # Convert to RGB if necessary (GIF requires RGB)
        if img.mode not in ('RGB', 'RGBA', 'L', 'P'):
            img = img.convert('RGB')
        
        # Fix EXIF orientation
        img = ImageOps.exif_transpose(img)
        
        # Convert RGBA to RGB if needed (GIF doesn't support alpha well)
        if img.mode == 'RGBA':
            # Create white background
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])  # Use alpha channel as mask
            img = background
        elif img.mode == 'RGBA':
            img = img.convert('RGB')

        # Downscale if needed
        w, h = img.size
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            new_width = int(w * scale)
            new_height = int(h * scale)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

        # Convert to palette mode for GIF
        # Use quality to determine quantization method
        if quality >= 80:
            # High quality: use adaptive palette
            img = img.quantize(colors=palette_size, method=Image.Quantize.MEDIANCUT)
        elif quality >= 50:
            # Medium quality: use fast octree
            img = img.quantize(colors=palette_size, method=Image.Quantize.FASTOCTREE)
        else:
            # Low quality: use fast octree with fewer colors
            reduced_palette = max(128, int(palette_size * (quality / 50)))
            img = img.quantize(colors=reduced_palette, method=Image.Quantize.FASTOCTREE)

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save as GIF
        save_kwargs = {
            "format": "GIF",
            "optimize": optimize,
        }

        img.save(output_file, **save_kwargs)

        output_size = os.path.getsize(output_file)
        input_size = os.path.getsize(input_file)
        print(f"GIF saved successfully: {output_size} bytes")
        print(f"Size ratio: {output_size / input_size:.2%}")

        return os.path.exists(output_file) and output_size > 0

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert JPG/JPEG to GIF")
    parser.add_argument("input_file", help="Input JPG/JPEG file path")
    parser.add_argument("output_file", help="Output GIF file path")
    parser.add_argument("--quality", type=int, default=85, choices=range(0, 101),
                        help="Quality 0-100 (default: 85, affects palette quantization)")
    parser.add_argument("--max-dimension", type=int, default=4096,
                        help="Maximum width or height (default: 4096)")
    parser.add_argument("--no-optimize", action="store_true",
                        help="Disable GIF optimization")
    parser.add_argument("--palette-size", type=int, default=256, choices=range(2, 257),
                        help="Number of colors in palette (default: 256, max: 256)")
    args = parser.parse_args()

    ok = convert_jpg_to_gif(
        args.input_file,
        args.output_file,
        args.quality,
        args.max_dimension,
        not args.no_optimize,
        args.palette_size,
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
