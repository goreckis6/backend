#!/usr/bin/env python3
"""
HEIC → AVIF converter (backend-ready)

- Reads HEIC/HEIF via pillow-heif
- Saves AVIF via pillow-avif-plugin
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

# required side-effect import
try:
    import pillow_avif  # noqa: F401
    HAS_AVIF = True
except ImportError:
    HAS_AVIF = False


_HEIF_REGISTERED = False


def _ensure_heif():
    """Register HEIF opener once per process"""
    global _HEIF_REGISTERED
    if HAS_PILLOW_HEIF and not _HEIF_REGISTERED:
        register_heif_opener()
        _HEIF_REGISTERED = True
    elif not HAS_PILLOW_HEIF:
        raise ImportError("pillow-heif is not installed. Please install it with: pip install pillow-heif")


def convert_heic_to_avif(
    input_file: str,
    output_file: str,
    quality: int = 45,
    max_dimension: int = 4096,
    lossless: bool = False,
    speed: int = 6
) -> bool:
    """
    Convert HEIC/HEIF image to AVIF format.
    
    Args:
        input_file: Path to input HEIC/HEIF file
        output_file: Path to output AVIF file
        quality: AVIF quality (30-60 recommended, default: 45)
        max_dimension: Maximum width or height (will downscale if exceeded)
        lossless: Use lossless AVIF compression
        speed: AVIF encoding speed (0-10, default: 6). Higher = faster but larger files
    
    Returns:
        True if conversion successful, False otherwise
    """
    if not HAS_PIL:
        print("ERROR: Pillow not available. Please install with: pip install Pillow", file=sys.stderr)
        return False

    if not HAS_AVIF:
        print("ERROR: pillow-avif-plugin not available. Please install with: pip install pillow-avif-plugin", file=sys.stderr)
        return False

    if not os.path.isfile(input_file):
        print(f"ERROR: File not found: {input_file}", file=sys.stderr)
        return False

    try:
        _ensure_heif()

        # Validate quality range (AVIF typically uses 30-60)
        if not lossless:
            quality = max(0, min(100, quality))
        
        # Validate speed range
        speed = max(0, min(10, speed))

        img = Image.open(input_file)
        
        # Fix EXIF orientation
        img = ImageOps.exif_transpose(img)

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

        # Save as AVIF
        img.save(
            output_file,
            format="AVIF",
            quality=quality,
            lossless=lossless,
            speed=speed
        )

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
        print(f"ERROR: HEIC → AVIF conversion failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Convert HEIC/HEIF images to AVIF format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.heic output.avif
  %(prog)s input.heic output.avif --quality 50
  %(prog)s input.heic output.avif --max-dimension 2048 --lossless
  %(prog)s input.heic output.avif --quality 45 --speed 8
        """
    )
    parser.add_argument("input_file", help="Input HEIC/HEIF file path")
    parser.add_argument("output_file", help="Output AVIF file path")
    parser.add_argument(
        "--quality",
        type=int,
        default=45,
        help="AVIF quality (0-100, default: 45). Recommended range: 30-60. Ignored if --lossless is used."
    )
    parser.add_argument(
        "--max-dimension",
        type=int,
        default=4096,
        help="Maximum width or height in pixels (default: 4096). Images larger than this will be downscaled."
    )
    parser.add_argument(
        "--lossless",
        action="store_true",
        help="Use lossless AVIF compression (quality parameter will be ignored)"
    )
    parser.add_argument(
        "--speed",
        type=int,
        default=6,
        choices=range(0, 11),
        metavar="[0-10]",
        help="AVIF encoding speed (0-10, default: 6). Higher values encode faster but produce larger files."
    )
    args = parser.parse_args()

    ok = convert_heic_to_avif(
        args.input_file,
        args.output_file,
        args.quality,
        args.max_dimension,
        args.lossless,
        args.speed
    )

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
