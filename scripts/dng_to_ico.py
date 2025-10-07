#!/usr/bin/env python3
"""
DNG to ICO converter using Python + Pillow + rawpy
Converts DNG (Digital Negative) files to ICO format with multiple sizes
"""

import argparse
import os
import sys
import traceback
from PIL import Image
import rawpy
import numpy as np

def convert_dng_to_ico(dng_file, output_file, sizes=[16, 32, 48, 64, 128, 256], quality='high'):
    """
    Convert DNG file to ICO format using rawpy and Pillow
    
    Args:
        dng_file (str): Path to input DNG file
        output_file (str): Path to output ICO file
        sizes (list): List of icon sizes to include
        quality (str): Quality level ('high', 'medium', 'low')
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DNG to ICO conversion...")
    print(f"Input: {dng_file}")
    print(f"Output: {output_file}")
    print(f"Sizes: {sizes}")
    print(f"Quality: {quality}")
    
    try:
        # Open DNG file with rawpy
        print("Opening DNG file with rawpy...")
        with rawpy.imread(dng_file) as raw:
            print(f"Raw image info: {raw.sizes}")
            
            # Try to print optional attributes if they exist
            try:
                print(f"White balance: {raw.camera_whitebalance}")
            except (AttributeError, ValueError):
                print("White balance: Not available")
            
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
            
            # Convert to 8-bit for ICO
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
            
            # Create multiple sizes for ICO
            print(f"Creating ICO with sizes: {sizes}")
            icon_images = []
            
            for size in sizes:
                print(f"Creating {size}x{size} version...")
                resized = pil_image.resize((size, size), Image.Resampling.LANCZOS)
                icon_images.append(resized)
            
            # Save as ICO with multiple sizes
            print("Saving ICO file...")
            
            # Set quality parameters based on quality level
            if quality == 'high':
                # Use PNG compression for high quality
                pil_image.save(output_file, format='ICO', sizes=[(s, s) for s in sizes])
            elif quality == 'medium':
                # Balanced quality
                pil_image.save(output_file, format='ICO', sizes=[(s, s) for s in sizes])
            else:
                # Lower quality, smaller file size
                pil_image.save(output_file, format='ICO', sizes=[(s, s) for s in sizes])
            
            # Verify the output file
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"ICO file created successfully: {file_size} bytes")
                
                # Verify it's a valid ICO file
                try:
                    with Image.open(output_file) as verify_img:
                        print(f"Verified ICO file: {verify_img.size}, mode: {verify_img.mode}")
                        return True
                except Exception as verify_error:
                    print(f"ERROR: Output file is not a valid ICO: {verify_error}")
                    return False
            else:
                print("ERROR: ICO file was not created")
                return False
                
    except rawpy.LibRawFileUnsupportedError as e:
        print(f"ERROR: Unsupported DNG file format: {e}")
        return False
    except rawpy.LibRawIOError as e:
        print(f"ERROR: DNG file I/O error: {e}")
        return False
    except Exception as e:
        print(f"ERROR: Failed to convert DNG to ICO: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert DNG to ICO format')
    parser.add_argument('dng_file', help='Input DNG file path')
    parser.add_argument('output_file', help='Output ICO file path')
    parser.add_argument('--sizes', type=int, nargs='+', default=[16, 32, 48, 64, 128, 256],
                        help='Icon sizes to include (default: 16 32 48 64 128 256)')
    parser.add_argument('--quality', choices=['high', 'medium', 'low'], default='high',
                        help='Quality level (default: high)')
    
    args = parser.parse_args()
    
    print("=== DNG to ICO Converter ===")
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
    
    # Validate sizes
    for size in args.sizes:
        if size < 16 or size > 256:
            print(f"ERROR: Icon size must be between 16 and 256, got: {size}")
            sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert DNG to ICO
    success = convert_dng_to_ico(
        args.dng_file,
        args.output_file,
        sizes=args.sizes,
        quality=args.quality
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

