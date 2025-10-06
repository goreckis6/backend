#!/usr/bin/env python3
"""
DNG to WebP converter using Python + Pillow + rawpy
Converts DNG (Digital Negative) files to WebP format with high quality
"""

import argparse
import os
import sys
import traceback
from datetime import datetime
from PIL import Image
import rawpy
import numpy as np

def convert_dng_to_webp(dng_file, output_file, quality=95, lossless=False, width=None, height=None):
    """
    Convert DNG file to WebP format using rawpy and Pillow
    
    Args:
        dng_file (str): Path to input DNG file
        output_file (str): Path to output WebP file
        quality (int): WebP quality (1-100, ignored if lossless=True)
        lossless (bool): Use lossless WebP compression
        width (int): Target width (optional, maintains aspect ratio)
        height (int): Target height (optional, maintains aspect ratio)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DNG to WebP conversion...")
    print(f"Input: {dng_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Lossless: {lossless}")
    print(f"Target size: {width}x{height}" if width or height else "Original size")
    
    try:
        # Open DNG file with rawpy
        print("Opening DNG file with rawpy...")
        with rawpy.imread(dng_file) as raw:
            print(f"Raw image info: {raw.sizes}")
            print(f"Color description: {raw.color_description}")
            print(f"White balance: {raw.camera_whitebalance}")
            print(f"Color matrix: {raw.color_matrix}")
            
            # Process the raw image
            print("Processing raw image...")
            # Use default processing parameters for best quality
            rgb = raw.postprocess(
                use_camera_wb=True,  # Use camera white balance
                half_size=False,     # Full resolution
                no_auto_bright=True, # Don't auto-adjust brightness
                output_bps=16        # 16-bit output for better quality
            )
            
            print(f"Processed image shape: {rgb.shape}")
            print(f"Data type: {rgb.dtype}")
            print(f"Value range: {rgb.min()} - {rgb.max()}")
            
            # Convert to 8-bit for WebP
            if rgb.dtype == np.uint16:
                print("Converting from 16-bit to 8-bit...")
                rgb = (rgb / 256).astype(np.uint8)
            elif rgb.dtype != np.uint8:
                print("Converting to 8-bit...")
                rgb = rgb.astype(np.uint8)
            
            # Convert numpy array to PIL Image
            print("Converting to PIL Image...")
            if rgb.shape[2] == 3:
                # RGB image
                pil_image = Image.fromarray(rgb, 'RGB')
            elif rgb.shape[2] == 4:
                # RGBA image
                pil_image = Image.fromarray(rgb, 'RGBA')
            else:
                raise ValueError(f"Unsupported number of channels: {rgb.shape[2]}")
            
            print(f"PIL Image size: {pil_image.size}")
            print(f"PIL Image mode: {pil_image.mode}")
            
            # Resize if requested
            if width or height:
                print("Resizing image...")
                original_width, original_height = pil_image.size
                
                if width and height:
                    # Both dimensions specified
                    new_size = (width, height)
                elif width:
                    # Width specified, calculate height maintaining aspect ratio
                    new_height = int((width * original_height) / original_width)
                    new_size = (width, new_height)
                else:
                    # Height specified, calculate width maintaining aspect ratio
                    new_width = int((height * original_width) / original_height)
                    new_size = (new_width, height)
                
                print(f"Resizing from {original_width}x{original_height} to {new_size[0]}x{new_size[1]}")
                pil_image = pil_image.resize(new_size, Image.Resampling.LANCZOS)
            
            # Convert to WebP
            print("Converting to WebP...")
            webp_options = {
                'quality': quality,
                'lossless': lossless,
                'method': 6,  # Best compression method
                'near_lossless': 80 if not lossless else 0  # Near-lossless quality
            }
            
            # Remove quality if lossless
            if lossless:
                webp_options.pop('quality', None)
                webp_options.pop('near_lossless', None)
            
            print(f"WebP options: {webp_options}")
            pil_image.save(output_file, 'WebP', **webp_options)
            
            # Verify the output file
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"WebP file created successfully: {file_size} bytes")
                
                # Verify it's a valid WebP file
                try:
                    with Image.open(output_file) as verify_img:
                        print(f"Verified WebP file: {verify_img.size}, mode: {verify_img.mode}")
                        return True
                except Exception as verify_error:
                    print(f"ERROR: Output file is not a valid WebP: {verify_error}")
                    return False
            else:
                print("ERROR: WebP file was not created")
                return False
                
    except rawpy.LibRawFileUnsupportedError as e:
        print(f"ERROR: Unsupported DNG file format: {e}")
        return False
    except rawpy.LibRawIOError as e:
        print(f"ERROR: DNG file I/O error: {e}")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DNG to WebP: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert DNG to WebP format')
    parser.add_argument('dng_file', help='Input DNG file path')
    parser.add_argument('output_file', help='Output WebP file path')
    parser.add_argument('--quality', type=int, default=95, help='WebP quality (1-100, default: 95)')
    parser.add_argument('--lossless', action='store_true', help='Use lossless WebP compression')
    parser.add_argument('--width', type=int, help='Target width (maintains aspect ratio)')
    parser.add_argument('--height', type=int, help='Target height (maintains aspect ratio)')
    
    args = parser.parse_args()
    
    print("=== DNG to WebP Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.dng_file):
        print(f"ERROR: Input DNG file not found: {args.dng_file}")
        sys.exit(1)
    
    # Check required libraries
    try:
        import rawpy
        print(f"rawpy version: {rawpy.__version__}")
    except ImportError as e:
        print(f"ERROR: rawpy not available: {e}")
        print("Please install rawpy: pip install rawpy")
        sys.exit(1)
    
    try:
        from PIL import Image
        print(f"Pillow version: {Image.__version__}")
    except ImportError as e:
        print(f"ERROR: Pillow not available: {e}")
        print("Please install Pillow: pip install Pillow")
        sys.exit(1)
    
    # Validate quality parameter
    if not args.lossless and (args.quality < 1 or args.quality > 100):
        print(f"ERROR: Quality must be between 1 and 100, got: {args.quality}")
        sys.exit(1)
    
    # Validate size parameters
    if args.width and args.width < 1:
        print(f"ERROR: Width must be positive, got: {args.width}")
        sys.exit(1)
    
    if args.height and args.height < 1:
        print(f"ERROR: Height must be positive, got: {args.height}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert DNG to WebP
    success = convert_dng_to_webp(
        args.dng_file,
        args.output_file,
        quality=args.quality,
        lossless=args.lossless,
        width=args.width,
        height=args.height
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

