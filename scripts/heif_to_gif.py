#!/usr/bin/env python3
"""
HEIF/HEIC to GIF Converter
Reads HEIF/HEIC via pillow-heif and saves GIF using Pillow.
Supports adjustable quality and max-dimension downscale.
Optimized for social media sharing.
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
    """Register HEIF opener once per process"""
    global _HEIF_REGISTERED
    if HAS_PILLOW_HEIF and not _HEIF_REGISTERED:
        register_heif_opener()
        _HEIF_REGISTERED = True
    elif not HAS_PILLOW_HEIF:
        raise ImportError("pillow-heif is not installed. Please install it with: pip install pillow-heif")


def convert_heif_to_gif(
    heif_file: str, 
    output_file: str, 
    quality: int = 85, 
    max_dimension: int = 4096,
    optimize: bool = True,
    palette_size: int = 256
) -> bool:
    """
    Convert HEIF/HEIC image to GIF format.
    
    Args:
        heif_file: Path to input HEIF/HEIC file
        output_file: Path to output GIF file
        quality: Quality parameter (0-100, default: 85). Influences color palette size.
        max_dimension: Maximum width or height (will downscale if exceeded)
        optimize: Use GIF optimization (default: True)
        palette_size: Number of colors in palette (default: 256, max: 256)
    
    Returns:
        True if conversion successful, False otherwise
    """
    print("Starting HEIF/HEIC to GIF conversion")
    print(f"Input: {heif_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Max dimension: {max_dimension}")

    if not HAS_PIL:
        print("ERROR: Pillow (PIL) not available")
        return False

    if not os.path.exists(heif_file):
        print(f"ERROR: Input file not found: {heif_file}")
        return False

    try:
        _ensure_heif()

        # Validate quality range
        quality = max(0, min(100, quality))
        
        # Validate palette size
        palette_size = max(2, min(256, palette_size))

        img = Image.open(heif_file)
        print(f"Opened image. Format={img.format}, Mode={img.mode}, Size={img.size}")

        # Fix EXIF orientation
        img = ImageOps.exif_transpose(img) if img is not None else img
        print("Fixed EXIF orientation")

        # Convert to RGB if necessary (GIF requires RGB)
        if img.mode not in ('RGB', 'RGBA', 'L', 'P'):
            print(f"Converting {img.mode} to RGB")
            img = img.convert('RGB')
        
        # Convert RGBA to RGB (GIF doesn't support alpha well)
        if img.mode == 'RGBA':
            print("Converting RGBA to RGB with white background")
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])  # Use alpha channel as mask
            img = background

        # Downscale large images
        width, height = img.size
        if max(width, height) > max_dimension:
            if width > height:
                new_w = max_dimension
                new_h = int(height * (max_dimension / width))
            else:
                new_h = max_dimension
                new_w = int(width * (max_dimension / height))
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            print(f"Resized to {new_w}x{new_h}")

        # Convert to palette mode (GIF requirement)
        if img.mode != 'P':
            # Use quality to determine number of colors
            num_colors = int(palette_size * (quality / 100.0))
            num_colors = max(2, min(256, num_colors))
            print(f"Quantizing to {num_colors} colors")
            
            # Convert to palette mode using median cut
            img = img.quantize(colors=num_colors, method=Image.Quantize.MEDIANCUT)

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Save as GIF
        save_kwargs = {
            'format': 'GIF',
            'optimize': optimize,
        }
        
        img.save(output_file, **save_kwargs)

        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            print("GIF created successfully")
            return True
        print("ERROR: GIF output not created")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert HEIF/HEIC to GIF: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIF/HEIC image to GIF')
    parser.add_argument('heif_file', help='Path to input HEIF/HEIC file')
    parser.add_argument('output_file', help='Path to output GIF file')
    parser.add_argument('--quality', type=int, default=85, help='Quality (0-100, affects palette size)')
    parser.add_argument('--max-dimension', type=int, default=4096, help='Max width or height for downscaling')
    parser.add_argument('--no-optimize', action='store_true', help='Disable GIF optimization')
    parser.add_argument('--palette-size', type=int, default=256, help='Max colors in palette (2-256)')
    args = parser.parse_args()

    print("=== HEIF/HEIC to GIF Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heif_to_gif(
        args.heif_file, 
        args.output_file, 
        args.quality, 
        args.max_dimension,
        optimize=not args.no_optimize,
        palette_size=args.palette_size
    )
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
