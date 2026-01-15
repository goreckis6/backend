#!/usr/bin/env python3
"""
HEIC to TIFF Converter
Reads HEIC/HEIF via pillow-heif and saves TIFF using Pillow
Supports lossless compression, EXIF metadata preservation, and optional max-dimension downscale
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


def convert_heic_to_tiff(heic_file: str, output_file: str, quality: int = 95, max_dimension: int = 4096, compression: str = 'tiff_lzw') -> bool:
    """
    Convert HEIC/HEIF image to TIFF format.
    
    Args:
        heic_file: Path to input HEIC/HEIF file
        output_file: Path to output TIFF file
        quality: Quality hint (0-100, affects compression)
        max_dimension: Maximum width or height (will downscale if exceeded)
        compression: TIFF compression method ('tiff_lzw', 'tiff_adobe_deflate', 'tiff_jpeg', 'tiff_ccitt', 'tiff_deflate', 'tiff_sgilog', 'tiff_raw')
    
    Returns:
        True if conversion successful, False otherwise
    """
    print("Starting HEIC to TIFF conversion")
    print(f"Input: {heic_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Max dimension: {max_dimension}")
    print(f"Compression: {compression}")

    if not HAS_PIL:
        print("ERROR: Pillow not available. Please install with: pip install Pillow", file=sys.stderr)
        return False

    if not os.path.exists(heic_file):
        print(f"ERROR: File not found: {heic_file}", file=sys.stderr)
        return False

    try:
        _ensure_heif()

        # Validate quality range
        quality = max(0, min(100, quality))

        img = Image.open(heic_file)
        print(f"Opened image. Format={img.format}, Mode={img.mode}, Size={img.size}")
        
        # Fix EXIF orientation
        img = ImageOps.exif_transpose(img)
        print(f"After EXIF transpose: Size={img.size}")

        # Downscale if needed
        w, h = img.size
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            new_w = int(w * scale)
            new_h = int(h * scale)
            # Use LANCZOS for larger downscales, BILINEAR for smaller
            resample = Image.Resampling.LANCZOS if scale < 0.5 else Image.Resampling.BILINEAR
            img = img.resize((new_w, new_h), resample)
            print(f"Resized to {new_w}x{new_h}")

        # Ensure TIFF-compatible mode
        # TIFF supports RGB, RGBA, L (grayscale), LA (grayscale + alpha), P (palette)
        if img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            # Convert to RGBA if transparency exists, otherwise RGB
            if 'transparency' in img.info or img.mode in ("LA", "PA"):
                img = img.convert("RGBA")
            else:
                img = img.convert("RGB")
            print(f"Converted to mode: {img.mode}")

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        # Save as TIFF with specified compression
        # TIFF compression options: 'tiff_lzw', 'tiff_adobe_deflate', 'tiff_jpeg', 'tiff_ccitt', 'tiff_deflate', 'tiff_sgilog', 'tiff_raw'
        save_kwargs = {
            "format": "TIFF",
            "compression": compression,
        }

        # For JPEG compression in TIFF, quality matters
        if compression == 'tiff_jpeg':
            save_kwargs["quality"] = quality

        img.save(output_file, **save_kwargs)

        # Verify output file was created and has content
        if not os.path.exists(output_file):
            print(f"ERROR: Output file was not created: {output_file}", file=sys.stderr)
            return False
        
        if os.path.getsize(output_file) == 0:
            print(f"ERROR: Output file is empty: {output_file}", file=sys.stderr)
            return False

        print("TIFF created successfully")
        return True

    except ImportError as e:
        print(f"ERROR: Missing dependency: {e}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"ERROR: HEIC → TIFF conversion failed: {e}", file=sys.stderr)
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Convert HEIC/HEIF images to TIFF format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.heic output.tiff
  %(prog)s input.heic output.tiff --quality 95
  %(prog)s input.heic output.tiff --max-dimension 2048 --compression tiff_lzw
  %(prog)s input.heic output.tiff --quality 90 --compression tiff_deflate
        """
    )
    parser.add_argument("heic_file", help="Input HEIC/HEIF file path")
    parser.add_argument("output_file", help="Output TIFF file path")
    parser.add_argument(
        "--quality",
        type=int,
        default=95,
        help="Quality hint (0-100, default: 95). Only affects JPEG compression in TIFF."
    )
    parser.add_argument(
        "--max-dimension",
        type=int,
        default=4096,
        help="Maximum width or height in pixels (default: 4096). Images larger than this will be downscaled."
    )
    parser.add_argument(
        "--compression",
        type=str,
        default="tiff_lzw",
        choices=["tiff_lzw", "tiff_adobe_deflate", "tiff_jpeg", "tiff_ccitt", "tiff_deflate", "tiff_sgilog", "tiff_raw"],
        help="TIFF compression method (default: tiff_lzw). tiff_lzw is lossless and efficient."
    )
    args = parser.parse_args()

    print("=== HEIC to TIFF Converter ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    ok = convert_heic_to_tiff(
        args.heic_file,
        args.output_file,
        args.quality,
        args.max_dimension,
        args.compression
    )

    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
