#!/usr/bin/env python3
"""
X3F (Sigma RAW) to JPEG preview generator
Uses rawpy (LibRaw) to read X3F and outputs a JPEG preview, resized for web.
"""

import os
import sys
import argparse
import traceback

try:
    import rawpy
    import numpy as np
    HAS_RAWPY = True
except Exception:
    HAS_RAWPY = False

try:
    from PIL import Image
    HAS_PIL = True
except Exception:
    HAS_PIL = False


def generate_preview(x3f_file: str, output_file: str, max_dimension: int = 2048, quality: int = 85) -> bool:
    if not HAS_RAWPY or not HAS_PIL:
        print("ERROR: rawpy or Pillow not available")
        return False

    if not os.path.exists(x3f_file):
        print(f"ERROR: input file not found: {x3f_file}")
        return False

    try:
        with rawpy.imread(x3f_file) as raw:
            # fast postprocess for preview
            rgb = raw.postprocess(use_auto_wb=True, no_auto_bright=False, half_size=True, gamma=(2.222, 4.5))
            img = Image.fromarray(rgb)
            w, h = img.size
            scale = 1.0
            if max(w, h) > max_dimension:
                if w >= h:
                    scale = max_dimension / float(w)
                else:
                    scale = max_dimension / float(h)
                new_w = max(1, int(w * scale))
                new_h = max(1, int(h * scale))
                img = img.resize((new_w, new_h), Image.Resampling.BILINEAR)
            img = img.convert('RGB')
            out_dir = os.path.dirname(output_file)
            if out_dir:
                os.makedirs(out_dir, exist_ok=True)
            img.save(output_file, format='JPEG', quality=max(1, min(100, int(quality))), optimize=False)
        return os.path.exists(output_file) and os.path.getsize(output_file) > 0
    except Exception as e:
        print(f"ERROR: Failed to generate preview: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Generate JPEG preview from X3F (Sigma RAW)')
    parser.add_argument('x3f_file', help='Path to input X3F file')
    parser.add_argument('output_file', help='Path to output JPEG file')
    parser.add_argument('--max-dimension', type=int, default=2048, help='Max width/height for preview (default 2048)')
    parser.add_argument('--quality', type=int, default=85, help='JPEG quality 1-100 (default 85)')
    args = parser.parse_args()

    ok = generate_preview(args.x3f_file, args.output_file, args.max_dimension, args.quality)
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()


