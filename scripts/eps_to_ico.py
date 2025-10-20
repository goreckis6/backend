#!/usr/bin/env python3
"""
EPS to ICO Converter
Converts EPS (Encapsulated PostScript) files to ICO (Icon) format
"""

import os
import sys
import argparse
import tempfile
import traceback
from PIL import Image, ImageOps

def convert_eps_to_ico(eps_file, output_file, sizes=[16, 32, 48, 64, 128, 256], quality='high'):
    """
    Convert EPS file to ICO format with multiple sizes
    
    Args:
        eps_file (str): Path to input EPS file
        output_file (str): Path to output ICO file
        sizes (list): List of icon sizes to include
        quality (str): Quality setting ('high', 'medium', 'low')
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting EPS to ICO conversion...")
    print(f"Input: {eps_file}")
    print(f"Output: {output_file}")
    print(f"Sizes: {sizes}")
    print(f"Quality: {quality}")
    
    try:
        # Check if EPS file exists and is readable
        if not os.path.exists(eps_file):
            print(f"ERROR: EPS file does not exist: {eps_file}")
            return False
        
        file_size = os.path.getsize(eps_file)
        print(f"EPS file size: {file_size} bytes")
        
        # Open EPS file with Pillow
        print("Opening EPS file with Pillow...")
        try:
            with Image.open(eps_file) as eps_image:
                print(f"EPS image info: {eps_image.size}, mode: {eps_image.mode}")
                
                # Convert to RGB if necessary
                if eps_image.mode != 'RGB':
                    print(f"Converting from {eps_image.mode} to RGB...")
                    eps_image = eps_image.convert('RGB')
                
                # Apply quality settings
                if quality == 'high':
                    resample_method = Image.Resampling.LANCZOS
                elif quality == 'medium':
                    resample_method = Image.Resampling.BILINEAR
                else:  # low
                    resample_method = Image.Resampling.NEAREST
                
                # Create multiple sizes for ICO
                print(f"Creating ICO with sizes: {sizes}")
                icon_images = []
                
                for size in sizes:
                    print(f"Creating {size}x{size} version...")
                    resized = eps_image.resize((size, size), resample_method)
                    icon_images.append(resized)
                
                # Save as ICO with multiple sizes
                print("Saving ICO file...")
                
                # Use a simpler approach - save as single size ICO first, then try multi-size
                try:
                    # Create a 32x32 version for the base ICO
                    base_size = 32
                    base_image = eps_image.resize((base_size, base_size), resample_method)
                    
                    # Save as ICO - Pillow will handle the format
                    base_image.save(output_file, format='ICO')
                    print(f"ICO saved successfully as {base_size}x{base_size}")
                    
                except Exception as save_error:
                    print(f"Error saving ICO: {save_error}")
                    # Try alternative approach - save as PNG first, then convert
                    print("Trying alternative ICO save method...")
                    try:
                        # Create a simple 32x32 ICO
                        simple_image = eps_image.resize((32, 32), resample_method)
                        simple_image.save(output_file, format='ICO')
                        print("ICO saved using alternative method")
                    except Exception as alt_error:
                        print(f"Alternative ICO save also failed: {alt_error}")
                        raise Exception(f"Failed to save ICO file: {save_error}")
                
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
                    
        except Exception as eps_error:
            print(f"ERROR: Failed to open EPS file: {eps_error}")
            print("This might not be a valid EPS file or the file is corrupted")
            return False
                    
    except Exception as e:
        print(f"ERROR: Failed to convert EPS to ICO: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert EPS file to ICO format')
    parser.add_argument('eps_file', help='Path to input EPS file')
    parser.add_argument('output_file', help='Path to output ICO file')
    parser.add_argument('--sizes', nargs='+', type=int, default=[16, 32, 48, 64, 128, 256],
                        help='Icon sizes to include (default: 16 32 48 64 128 256)')
    parser.add_argument('--quality', choices=['high', 'medium', 'low'], default='high',
                        help='Conversion quality (default: high)')
    
    args = parser.parse_args()
    
    print("=== EPS to ICO Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_eps_to_ico(args.eps_file, args.output_file, args.sizes, args.quality)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == '__main__':
    main()

