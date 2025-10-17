#!/usr/bin/env python3
"""
CR2 to WebP converter using Python + rawpy + Pillow
Converts Canon CR2 RAW files to WebP format with various quality options
"""

import argparse
import os
import sys
import traceback
import rawpy
from PIL import Image
import numpy as np

def convert_cr2_to_webp(cr2_file, output_file, quality=80, lossless=False, method=6):
    """
    Convert CR2 file to WebP format using rawpy + Pillow
    
    Args:
        cr2_file (str): Path to input CR2 file
        output_file (str): Path to output WebP file
        quality (int): WebP quality (0-100, ignored if lossless=True)
        lossless (bool): Use lossless compression
        method (int): Compression method (0-6, higher = better compression)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CR2 to WebP conversion...")
    print(f"Input: {cr2_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Lossless: {lossless}")
    print(f"Method: {method}")
    
    try:
        # Check if input file exists and is readable
        if not os.path.exists(cr2_file):
            print(f"ERROR: Input CR2 file not found: {cr2_file}")
            return False
            
        file_size = os.path.getsize(cr2_file)
        print(f"CR2 file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: CR2 file is empty")
            return False
        
        # Process CR2 file with rawpy
        print("Processing CR2 file with rawpy...")
        with rawpy.imread(cr2_file) as raw:
            print(f"Raw image shape: {raw.raw_image.shape}")
            print(f"Raw image type: {raw.raw_image.dtype}")
            print(f"Color matrix: {raw.color_matrix}")
            print(f"White balance: {raw.camera_whitebalance}")
            
            # Process the raw image
            print("Processing raw image...")
            rgb = raw.postprocess(
                use_camera_wb=True,
                half_size=False,
                no_auto_bright=True,
                output_bps=8
            )
            
            print(f"Processed RGB shape: {rgb.shape}")
            print(f"Processed RGB type: {rgb.dtype}")
            print(f"RGB range: {rgb.min()} - {rgb.max()}")
            
            # Convert to PIL Image
            print("Converting to PIL Image...")
            if rgb.shape[2] == 3:  # RGB
                pil_image = Image.fromarray(rgb, 'RGB')
            elif rgb.shape[2] == 4:  # RGBA
                pil_image = Image.fromarray(rgb, 'RGBA')
            else:
                print(f"ERROR: Unexpected number of channels: {rgb.shape[2]}")
                return False
            
            print(f"PIL Image mode: {pil_image.mode}")
            print(f"PIL Image size: {pil_image.size}")
            
            # Convert to RGB if necessary (WebP works best with RGB)
            if pil_image.mode in ('RGBA', 'LA', 'P'):
                print("Converting image to RGB mode...")
                # Create a white background for transparency
                if pil_image.mode == 'RGBA':
                    background = Image.new('RGB', pil_image.size, (255, 255, 255))
                    background.paste(pil_image, mask=pil_image.split()[-1])  # Use alpha channel as mask
                    pil_image = background
                else:
                    pil_image = pil_image.convert('RGB')
            elif pil_image.mode != 'RGB':
                print(f"Converting from {pil_image.mode} to RGB...")
                pil_image = pil_image.convert('RGB')
            
            print(f"Final image mode: {pil_image.mode}")
            print(f"Final image size: {pil_image.size}")
            
            # Save as WebP
            print("Saving as WebP...")
            save_kwargs = {
                'format': 'WebP',
                'method': method,
                'optimize': True
            }
            
            if lossless:
                save_kwargs['lossless'] = True
            else:
                save_kwargs['quality'] = quality
            
            pil_image.save(output_file, **save_kwargs)
            
            # Verify the output file
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"WebP file created successfully: {output_size} bytes")
                print(f"Compression ratio: {file_size/output_size:.2f}x")
                return True
            else:
                print("ERROR: WebP file was not created")
                return False
                
    except Exception as e:
        print(f"ERROR: Failed to convert CR2 to WebP: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CR2 to WebP format')
    parser.add_argument('cr2_file', help='Input CR2 file path')
    parser.add_argument('output_file', help='Output WebP file path')
    parser.add_argument('--quality', type=int, default=80, help='WebP quality 0-100 (default: 80)')
    parser.add_argument('--lossless', action='store_true', help='Use lossless compression')
    parser.add_argument('--method', type=int, default=6, help='Compression method 0-6 (default: 6)')
    
    args = parser.parse_args()
    
    print("=== CR2 to WebP Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    try:
        import rawpy
        from PIL import Image
        import numpy as np
        print(f"rawpy version: {rawpy.__version__}")
        print(f"Pillow version: {Image.__version__}")
        print(f"NumPy version: {np.__version__}")
    except ImportError as e:
        print(f"ERROR: Required library not available: {e}")
        print("Please install required libraries: pip install rawpy Pillow numpy")
        sys.exit(1)
    
    # Validate quality range
    if not 0 <= args.quality <= 100:
        print("ERROR: Quality must be between 0 and 100")
        sys.exit(1)
    
    # Validate method range
    if not 0 <= args.method <= 6:
        print("ERROR: Method must be between 0 and 6")
        sys.exit(1)
    
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    success = convert_cr2_to_webp(
        args.cr2_file,
        args.output_file,
        quality=args.quality,
        lossless=args.lossless,
        method=args.method
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()
