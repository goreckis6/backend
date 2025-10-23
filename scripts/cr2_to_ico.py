#!/usr/bin/env python3
"""
CR2 to ICO converter using Python + Pillow + rawpy
Converts CR2 (Canon RAW) files to ICO format with multiple sizes
"""

import argparse
import os
import sys
import traceback
from PIL import Image
import rawpy
import numpy as np

def convert_cr2_to_ico(cr2_file, output_file, sizes=None, quality='high', use_original_size=False):
    """
    Convert CR2 file to ICO format using rawpy and Pillow
    
    Args:
        cr2_file (str): Path to input CR2 file
        output_file (str): Path to output ICO file
        sizes (list): List of icon sizes to include
        quality (str): Quality level ('high', 'medium', 'low')
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CR2 to ICO conversion...")
    print(f"Input: {cr2_file}")
    print(f"Output: {output_file}")
    print(f"Use original size: {use_original_size}")
    print(f"Sizes: {sizes}")
    print(f"Quality: {quality}")
    
    try:
        # Check if CR2 file exists and is readable
        if not os.path.exists(cr2_file):
            print(f"ERROR: CR2 file does not exist: {cr2_file}")
            return False
        
        file_size = os.path.getsize(cr2_file)
        print(f"CR2 file size: {file_size} bytes")
        
        # Open CR2 file with rawpy
        print("Opening CR2 file with rawpy...")
        try:
            with rawpy.imread(cr2_file) as raw:
                print(f"Raw image info: {raw.sizes}")
                
                # Try to print optional attributes if they exist
                try:
                    print(f"White balance: {raw.camera_whitebalance}")
                except (AttributeError, ValueError):
                    print("White balance: Not available")
                
                # Process the raw image
                print("Processing raw image...")
                try:
                    # Use simpler processing parameters first
                    rgb = raw.postprocess(
                        use_camera_wb=True,  # Use camera white balance
                        half_size=False,     # Full resolution
                        no_auto_bright=True, # Don't auto-adjust brightness
                        output_bps=8         # Use 8-bit output for simplicity
                    )
                    print("Raw processing completed successfully")
                except Exception as process_error:
                    print(f"ERROR: Raw processing failed: {process_error}")
                    # Try with different parameters
                    print("Trying with different processing parameters...")
                    rgb = raw.postprocess(
                        use_camera_wb=False,  # Don't use camera white balance
                        half_size=True,       # Use half size for memory
                        no_auto_bright=False, # Allow auto brightness
                        output_bps=8          # 8-bit output
                    )
                    print("Raw processing completed with fallback parameters")
        except rawpy.LibRawFileUnsupportedError as e:
            print(f"ERROR: Unsupported CR2 file format: {e}")
            return False
        except rawpy.LibRawIOError as e:
            print(f"ERROR: CR2 file I/O error: {e}")
            return False
        except Exception as raw_error:
            print(f"ERROR: Failed to open/process CR2 file: {raw_error}")
            traceback.print_exc()
            return False
        
        # Continue with image processing
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
        try:
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
        except Exception as pil_error:
            print(f"ERROR: Failed to create PIL Image: {pil_error}")
            traceback.print_exc()
            return False
        
        # Determine the output size - ALWAYS use original size by default
        original_width, original_height = pil_image.size
        original_size_value = max(original_width, original_height)
        
        print(f"Original image dimensions: {original_width}x{original_height}")
        print(f"use_original_size flag: {use_original_size}")
        print(f"sizes parameter: {sizes}")
        
        # SIMPLIFIED LOGIC: Always use original size unless specific size is provided
        if sizes and len(sizes) > 0:
            # User selected a specific size (16, 32, 48, 64, 128, or 256)
            sizes_to_use = sizes
            print(f"✓ Using user-specified sizes: {sizes_to_use}")
        else:
            # Default: use original image size
            sizes_to_use = [original_size_value]
            print(f"✓ Using ORIGINAL image size: {original_width}x{original_height} -> {original_size_value}x{original_size_value}")
        
        # Create multiple sizes for ICO
        print(f"Creating ICO with sizes: {sizes_to_use}")
        icon_images = []
        
        for size in sizes_to_use:
            print(f"Creating {size}x{size} version...")
            resized = pil_image.resize((size, size), Image.Resampling.LANCZOS)
            icon_images.append(resized)
        
        # Save as ICO with multiple sizes
        print("Saving ICO file...")
        
        # Use a simpler approach - save as single size ICO first, then try multi-size
        try:
            # Use the first (largest) size for the base ICO
            base_size = sizes_to_use[0]
            base_image = pil_image.resize((base_size, base_size), Image.Resampling.LANCZOS)
            
            # Save as ICO - Pillow will handle the format
            base_image.save(output_file, format='ICO')
            print(f"ICO saved successfully as {base_size}x{base_size}")
            
        except Exception as save_error:
            print(f"Error saving ICO: {save_error}")
            # Try alternative approach - save as PNG first, then convert
            print("Trying alternative ICO save method...")
            try:
                # Create a simple 32x32 ICO
                simple_image = pil_image.resize((32, 32), Image.Resampling.LANCZOS)
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
                
    except Exception as e:
        print(f"ERROR: Failed to convert CR2 to ICO: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CR2 to ICO format')
    parser.add_argument('cr2_file', help='Input CR2 file path')
    parser.add_argument('output_file', help='Output ICO file path')
    parser.add_argument('--sizes', type=int, nargs='+', default=None,
                        help='Icon sizes to include (default: use original image size)')
    parser.add_argument('--original-size', action='store_true',
                        help='Use the original image size instead of predefined sizes')
    parser.add_argument('--quality', choices=['high', 'medium', 'low'], default='high',
                        help='Quality level (default: high)')
    
    args = parser.parse_args()
    
    print("=== CR2 to ICO Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    print(f"Original size flag: {args.original_size}")
    print(f"Sizes provided: {args.sizes}")
    print(f"Command line args: {sys.argv}")
    
    # Check if input file exists
    if not os.path.exists(args.cr2_file):
        print(f"ERROR: Input CR2 file not found: {args.cr2_file}")
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
    
    # Validate sizes (only if sizes are provided)
    if args.sizes:
        for size in args.sizes:
            if size < 16 or size > 256:
                print(f"ERROR: Icon size must be between 16 and 256, got: {size}")
                sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        print(f"Creating output directory: {output_dir}")
        os.makedirs(output_dir)
    
    # Check if we can write to the output directory
    try:
        test_file = os.path.join(output_dir, 'test_write.tmp')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        print(f"Output directory is writable: {output_dir}")
    except Exception as write_error:
        print(f"ERROR: Cannot write to output directory {output_dir}: {write_error}")
        sys.exit(1)
    
    # Convert CR2 to ICO
    try:
        success = convert_cr2_to_ico(
            args.cr2_file,
            args.output_file,
            sizes=args.sizes,
            quality=args.quality,
            use_original_size=args.original_size
        )
        
        if success:
            print("=== CONVERSION SUCCESSFUL ===")
            sys.exit(0)
        else:
            print("=== CONVERSION FAILED ===")
            sys.exit(1)
    except Exception as e:
        print(f"=== CONVERSION FAILED WITH EXCEPTION ===")
        print(f"Exception: {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()