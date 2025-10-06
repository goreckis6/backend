#!/usr/bin/env python3
"""
BMP to ICO Converter
Converts BMP files to ICO format using Pillow (PIL) with multi-size icons.
"""

from PIL import Image
import argparse
import os
import sys
import io
from datetime import datetime
import traceback

def create_ico_from_bmp(bmp_file, output_file, sizes=None, include_alpha=True):
    """
    Convert BMP file to ICO format with multiple sizes.
    
    Args:
        bmp_file (str): Path to input BMP file
        output_file (str): Path to output ICO file
        sizes (list): List of icon sizes to include (default: [16, 24, 32, 48, 64, 128, 256])
    """
    if sizes is None:
        sizes = [16, 24, 32, 48, 64, 128, 256]
    
    print(f"Starting BMP to ICO conversion...")
    print(f"Input: {bmp_file}")
    print(f"Output: {output_file}")
    print(f"Target sizes: {sizes}")
    
    try:
        # Check if file exists and get file info
        if not os.path.exists(bmp_file):
            raise Exception(f"Input file does not exist: {bmp_file}")
        
        file_size = os.path.getsize(bmp_file)
        print(f"Input file size: {file_size} bytes")
        
        if file_size == 0:
            raise Exception("Input file is empty")
        
        # Read first few bytes to check file signature
        with open(bmp_file, 'rb') as f:
            header = f.read(10)
            print(f"File header (first 10 bytes): {header.hex()}")
            
            # Check BMP signature
            if header[:2] == b'BM':
                print("BMP signature detected")
            else:
                print(f"Warning: File does not start with BMP signature (BM), got: {header[:2]}")
        
        # Open the image file
        print("Opening image file with Pillow...")
        try:
            with Image.open(bmp_file) as img:
                print(f"Successfully opened image: {img.size[0]}x{img.size[1]} pixels, mode: {img.mode}, format: {img.format}")
                
                # Verify it's a valid image format
                if img.format not in ['BMP', 'PNG', 'JPEG', 'JPG', 'GIF', 'TIFF', 'WEBP']:
                    print(f"Warning: Unsupported image format: {img.format}")
                    # Try to convert anyway
                    print("Attempting to process despite unsupported format...")
        except Exception as open_error:
            print(f"Error opening image with Pillow: {open_error}")
            # Try to load as raw data and convert
            print("Attempting to load as raw BMP data...")
            try:
                # Try to load as raw BMP
                with open(bmp_file, 'rb') as f:
                    data = f.read()
                
                # Try different approaches to load the image
                img = None
                
                # Method 1: Try with explicit format
                try:
                    img = Image.open(io.BytesIO(data), formats=['BMP'])
                    print("Successfully loaded as BMP with explicit format")
                except Exception as e1:
                    print(f"Method 1 failed: {e1}")
                
                # Method 2: Try without format specification
                if img is None:
                    try:
                        img = Image.open(io.BytesIO(data))
                        print("Successfully loaded without format specification")
                    except Exception as e2:
                        print(f"Method 2 failed: {e2}")
                
                # Method 3: Try with different mode
                if img is None:
                    try:
                        img = Image.open(io.BytesIO(data))
                        img.load()  # Force load
                        print("Successfully loaded with force load")
                    except Exception as e3:
                        print(f"Method 3 failed: {e3}")
                
                if img is None:
                    raise Exception("Could not load image with any method")
                
                print(f"Successfully loaded image: {img.size[0]}x{img.size[1]} pixels, mode: {img.mode}, format: {img.format}")
                
            except Exception as raw_error:
                print(f"Error loading as raw data: {raw_error}")
                raise Exception(f"Cannot open image file: {open_error}. Raw loading also failed: {raw_error}")
        
        # Convert to RGBA if not already (for alpha support) and alpha is requested
        if include_alpha and img.mode != 'RGBA':
            print("Converting to RGBA for alpha support...")
            try:
                img = img.convert('RGBA')
                print("Successfully converted to RGBA")
            except Exception as convert_error:
                print(f"Error converting to RGBA: {convert_error}")
                # Try converting to RGB first, then to RGBA
                try:
                    print("Trying RGB intermediate conversion...")
                    img = img.convert('RGB').convert('RGBA')
                    print("Successfully converted via RGB intermediate step")
                except Exception as rgb_error:
                    print(f"Error converting via RGB: {rgb_error}")
                    # Last resort: convert to RGB only
                    try:
                        img = img.convert('RGB')
                        print("Converted to RGB only (no alpha support)")
                    except Exception as final_error:
                        print(f"Final conversion attempt failed: {final_error}")
                        raise Exception(f"Cannot process image: {final_error}")
        elif not include_alpha and img.mode != 'RGB':
            print("Converting to RGB (no alpha support requested)...")
            try:
                img = img.convert('RGB')
                print("Successfully converted to RGB")
            except Exception as convert_error:
                print(f"Error converting to RGB: {convert_error}")
                raise Exception(f"Cannot process image: {convert_error}")
        else:
            print(f"Image already in {'RGBA' if include_alpha else 'RGB'} mode")
        
        # Create list of resized images for different icon sizes
        icon_images = []
        
        for size in sizes:
            print(f"Creating {size}x{size} icon...")
            
            # Resize image maintaining aspect ratio
            resized = img.resize((size, size), Image.Resampling.LANCZOS)
            
            # Ensure the image is exactly the target size
            if resized.size != (size, size):
                # Create a new image with the exact size and paste the resized image
                new_img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
                new_img.paste(resized, (0, 0))
                resized = new_img
            
            icon_images.append(resized)
            print(f"  Created {size}x{size} icon successfully")
        
        # Save as ICO with multiple sizes
        print(f"Saving ICO file with {len(icon_images)} sizes...")
        try:
            icon_images[0].save(
                output_file,
                format='ICO',
                sizes=[(img.width, img.height) for img in icon_images],
                append_images=icon_images[1:]
            )
            print("ICO file saved successfully")
        except Exception as save_error:
            print(f"Error saving ICO file: {save_error}")
            # Try saving with a simpler approach
            try:
                print("Trying alternative ICO save method...")
                icon_images[0].save(output_file, format='ICO')
                print("ICO file saved with alternative method")
            except Exception as alt_error:
                print(f"Alternative save method also failed: {alt_error}")
                raise Exception(f"Failed to save ICO file: {save_error}")
        
        # Verify the file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"ICO file created successfully: {file_size} bytes")
            print(f"Included sizes: {[img.size for img in icon_images]}")
            return True
        else:
            print("ERROR: ICO file was not created")
            return False
                
    except Exception as e:
        print(f"ERROR: Failed to create ICO from BMP: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert BMP to ICO format')
    parser.add_argument('bmp_file', help='Input BMP file path')
    parser.add_argument('output_file', help='Output ICO file path')
    parser.add_argument('--sizes', nargs='+', type=int, default=[16, 24, 32, 48, 64, 128, 256],
                       help='Icon sizes to include (default: 16 24 32 48 64 128 256)')
    parser.add_argument('--alpha', type=str, default='true',
                       help='Include alpha transparency (default: true)')
    
    args = parser.parse_args()
    
    print("=== BMP to ICO Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.bmp_file):
        print(f"ERROR: Input BMP file not found: {args.bmp_file}")
        sys.exit(1)
    
    # Check Pillow availability
    try:
        print(f"Pillow version: {Image.__version__}")
    except Exception as e:
        print(f"ERROR: Pillow not available: {e}")
        sys.exit(1)
    
    # Validate sizes
    valid_sizes = [16, 24, 32, 48, 64, 128, 256]
    invalid_sizes = [s for s in args.sizes if s not in valid_sizes]
    if invalid_sizes:
        print(f"WARNING: Invalid sizes detected: {invalid_sizes}")
        print(f"Valid sizes are: {valid_sizes}")
        args.sizes = [s for s in args.sizes if s in valid_sizes]
        if not args.sizes:
            print("ERROR: No valid sizes provided")
            sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Parse alpha setting
    include_alpha = args.alpha.lower() in ['true', '1', 'yes', 'on']
    
    # Convert BMP to ICO
    success = create_ico_from_bmp(
        args.bmp_file,
        args.output_file,
        args.sizes,
        include_alpha
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()
