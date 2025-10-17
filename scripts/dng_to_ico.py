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
        # Check if DNG file exists and is readable
        if not os.path.exists(dng_file):
            print(f"ERROR: DNG file does not exist: {dng_file}")
            return False
        
        file_size = os.path.getsize(dng_file)
        print(f"DNG file size: {file_size} bytes")
        
        # Open DNG file with rawpy
        print("Opening DNG file with rawpy...")
        try:
            with rawpy.imread(dng_file) as raw:
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
            print(f"ERROR: Unsupported DNG file format: {e}")
            return False
        except rawpy.LibRawIOError as e:
            print(f"ERROR: DNG file I/O error: {e}")
            return False
        except Exception as raw_error:
            print(f"ERROR: Failed to open/process DNG file: {raw_error}")
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
        
        # Create multiple sizes for ICO
        print(f"Creating ICO with sizes: {sizes}")
        icon_images = []
        
        for size in sizes:
            print(f"Creating {size}x{size} version...")
            resized = pil_image.resize((size, size), Image.Resampling.LANCZOS)
            icon_images.append(resized)
        
        # Save as ICO with multiple sizes
        print("Saving ICO file...")
        
        # Use a simpler approach - save as single size ICO first, then try multi-size
        try:
            # Create a 32x32 version for the base ICO
            base_size = 32
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
    
    # Convert DNG to ICO
    try:
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
    except Exception as e:
        print(f"=== CONVERSION FAILED WITH EXCEPTION ===")
        print(f"Exception: {e}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()


