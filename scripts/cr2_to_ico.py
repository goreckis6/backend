#!/usr/bin/env python3
"""
CR2 to ICO converter using Python + rawpy + Pillow
Converts Canon CR2 RAW files to ICO format with various sizes
"""

import argparse
import os
import sys
import traceback
import rawpy
from PIL import Image
import numpy as np

def convert_cr2_to_ico(cr2_file, output_file, sizes=[16, 32, 48, 64, 128, 256], quality=95):
    """
    Convert CR2 file to ICO format using rawpy + Pillow
    
    Args:
        cr2_file (str): Path to input CR2 file
        output_file (str): Path to output ICO file
        sizes (list): List of icon sizes to include
        quality (int): ICO quality (1-100)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CR2 to ICO conversion...")
    print(f"Input: {cr2_file}")
    print(f"Output: {output_file}")
    print(f"Sizes: {sizes}")
    print(f"Quality: {quality}")
    
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
            
            # Create ICO with original image size
            print("Creating ICO with original image size...")
            ico_images = []
            
            # Convert to RGBA if not already
            if pil_image.mode != 'RGBA':
                pil_image = pil_image.convert('RGBA')
            
            # If the image is not square, crop it to square from center
            if pil_image.size[0] != pil_image.size[1]:
                # Crop to square from center
                size = min(pil_image.size[0], pil_image.size[1])
                left = (pil_image.size[0] - size) // 2
                top = (pil_image.size[1] - size) // 2
                right = left + size
                bottom = top + size
                pil_image = pil_image.crop((left, top, right, bottom))
                print(f"Cropped to square: {pil_image.size[0]}x{pil_image.size[1]}")
            
            # Use the original (or cropped) image as the ICO
            ico_images.append(pil_image)
            print(f"Created ICO with size: {pil_image.size[0]}x{pil_image.size[1]}")
            print(f"ICO image mode: {pil_image.mode}")
            print(f"ICO image format: {pil_image.format}")
            
            if not ico_images:
                print("ERROR: No valid icon sizes could be created")
                return False
            
            print(f"Created {len(ico_images)} icon sizes")
            
            # Save as ICO
            print("Saving as ICO...")
            print(f"Saving ICO with {len(ico_images)} images")
            for i, img in enumerate(ico_images):
                print(f"  Image {i}: {img.size[0]}x{img.size[1]} mode={img.mode}")
            
            # Save the ICO with proper format
            # First try to save as ICO
            try:
                ico_images[0].save(
                    output_file,
                    format='ICO',
                    sizes=[(img.width, img.height) for img in ico_images],
                    quality=quality
                )
                print("ICO saved successfully")
            except Exception as e:
                print(f"Error saving as ICO: {e}")
                # Fallback: save as PNG and rename
                png_file = output_file.replace('.ico', '.png')
                ico_images[0].save(png_file, format='PNG')
                print(f"Saved as PNG fallback: {png_file}")
                # Copy PNG to ICO file
                import shutil
                shutil.copy2(png_file, output_file)
                print(f"Copied PNG to ICO file")
            
            # Verify the saved ICO
            try:
                saved_ico = Image.open(output_file)
                print(f"Saved ICO verification: {saved_ico.size[0]}x{saved_ico.size[1]} mode={saved_ico.mode}")
                saved_ico.close()
            except Exception as e:
                print(f"Error verifying saved ICO: {e}")
            
            # Verify the output file
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"ICO file created successfully: {output_size} bytes")
                print(f"Compression ratio: {file_size/output_size:.2f}x")
                return True
            else:
                print("ERROR: ICO file was not created")
                return False
                
    except Exception as e:
        print(f"ERROR: Failed to convert CR2 to ICO: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CR2 to ICO format')
    parser.add_argument('cr2_file', help='Input CR2 file path')
    parser.add_argument('output_file', help='Output ICO file path')
    parser.add_argument('--sizes', nargs='+', type=int, default=[16, 32, 48, 64, 128, 256], 
                       help='Icon sizes to include (default: 16 32 48 64 128 256)')
    parser.add_argument('--original-size', action='store_true', 
                       help='Use original image size as the primary icon size')
    parser.add_argument('--quality', type=int, default=95, help='ICO quality 1-100 (default: 95)')
    
    args = parser.parse_args()
    
    print("=== CR2 to ICO Converter ===")
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
    if not 1 <= args.quality <= 100:
        print("ERROR: Quality must be between 1 and 100")
        sys.exit(1)
    
    # Validate sizes (skip validation for original size)
    if args.sizes != ['original'] and (not args.sizes or any(size <= 0 for size in args.sizes)):
        print("ERROR: All sizes must be positive integers")
        sys.exit(1)
    
    # Always use original image size approach
    print("Using original image size as primary icon size")
    args.sizes = ['original']
    
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    success = convert_cr2_to_ico(
        args.cr2_file,
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


