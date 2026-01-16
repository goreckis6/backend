#!/usr/bin/env python3
"""
HEIC → GIF converter (backend-ready)

- Reads HEIC/HEIF via pillow-heif
- Saves GIF via PIL/Pillow
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


def convert_heic_to_gif(
    input_file: str,
    output_file: str,
    quality: int = 85,
    max_dimension: int = 4096,
    optimize: bool = True,
    palette_size: int = 256
) -> bool:
    """
    Convert HEIC/HEIF image to GIF format.
    
    Args:
        input_file: Path to input HEIC/HEIF file
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
        _ensure_heif()

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
            new_w = int(w * scale)
            new_h = int(h * scale)
            img = img.resize(
                (new_w, new_h),
                Image.Resampling.LANCZOS if scale < 0.5 else Image.Resampling.BILINEAR
            )

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        # Convert to palette mode (GIF requirement)
        # Use quality to determine quantization method
        # Higher quality = more colors preserved
        if img.mode != 'P':
            # Quantize to palette
            # PIL's quantize method uses number of colors
            # We'll use the quality parameter to influence quantization
            # Lower quality = fewer colors (faster, smaller file)
            # Higher quality = more colors (slower, larger file but better quality)
            num_colors = int(palette_size * (quality / 100.0))
            num_colors = max(2, min(256, num_colors))  # Ensure valid range
            
            # Convert to palette mode using median cut
            img = img.quantize(colors=num_colors, method=Image.Quantize.MEDIANCUT)
        
        # Save as GIF
        save_kwargs = {
            'format': 'GIF',
            'optimize': optimize,
        }
        
        img.save(output_file, **save_kwargs)

        # Verify output file was created and has content
        if not os.path.exists(output_file):
            print(f"ERROR: Output file was not created: {output_file}", file=sys.stderr)
            return False
        
        if os.path.getsize(output_file) == 0:
            print(f"ERROR: Output file is empty: {output_file}", file=sys.stderr)
            return False

        return True

    except ImportError as e:
        print(f"ERROR: Missing dependency: {e}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"ERROR: HEIC → GIF conversion failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Convert HEIC/HEIF images to GIF format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.heic output.gif
  %(prog)s input.heic output.gif --quality 85
  %(prog)s input.heic output.gif --max-dimension 2048
  %(prog)s input.heic output.gif --quality 90 --no-optimize
        """
    )
    parser.add_argument("input_file", help="Input HEIC/HEIF file path")
    parser.add_argument("output_file", help="Output GIF file path")
    parser.add_argument(
        "--quality",
        type=int,
        default=85,
        help="Quality parameter (0-100, default: 85). Influences color palette size."
    )
    parser.add_argument(
        "--max-dimension",
        type=int,
        default=4096,
        help="Maximum width or height in pixels (default: 4096). Images larger than this will be downscaled."
    )
    parser.add_argument(
        "--no-optimize",
        action="store_true",
        help="Disable GIF optimization (smaller file but slower encoding)"
    )
    parser.add_argument(
        "--palette-size",
        type=int,
        default=256,
        choices=range(2, 257),
        metavar="[2-256]",
        help="Maximum number of colors in palette (default: 256)"
    )
    args = parser.parse_args()

    ok = convert_heic_to_gif(
        args.input_file,
        args.output_file,
        args.quality,
        args.max_dimension,
        optimize=not args.no_optimize,
        palette_size=args.palette_size
    )

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
