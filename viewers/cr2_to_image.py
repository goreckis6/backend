#!/usr/bin/env python3
"""
CR2 (Canon RAW) to Image converter for web preview.
Uses rawpy (LibRaw), Pillow for rendering, and exiftool for fast embedded JPEG extraction.
"""

import argparse
import os
import sys
import subprocess
import traceback
from PIL import Image
import io

def extract_embedded_jpeg(cr2_file, output_file):
    """
    Fast preview using exiftool to extract embedded JPEG (no demosaic).
    
    Args:
        cr2_file (str): Path to input CR2 file
        output_file (str): Path to output image file
    
    Returns:
        bool: True if extraction successful, False otherwise
    """
    print(f"Attempting fast preview with exiftool (embedded JPEG)...")
    
    try:
        # Try to extract embedded JPEG preview with exiftool
        cmd = [
            'exiftool',
            '-b',
            '-PreviewImage',
            cr2_file
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=30
        )
        
        if result.returncode == 0 and len(result.stdout) > 1000:
            # Valid JPEG data extracted
            print(f"Embedded JPEG extracted: {len(result.stdout)} bytes")
            
            # Open with Pillow to resize if needed
            img = Image.open(io.BytesIO(result.stdout))
            print(f"Embedded JPEG size: {img.size}")
            
            # Resize if too large for web (max 2048px)
            max_dimension = 2048
            if img.width > max_dimension or img.height > max_dimension:
                print(f"Resizing to fit within {max_dimension}px...")
                img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
                print(f"Resized to: {img.size}")
            
            # Save as JPEG
            img.save(output_file, 'JPEG', quality=92, optimize=True)
            
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"Fast preview created: {file_size} bytes")
                return True
        
        return False
        
    except FileNotFoundError:
        print("WARNING: exiftool not found, will use rawpy instead")
        return False
    except Exception as e:
        print(f"WARNING: exiftool extraction failed: {e}")
        return False

def convert_cr2_with_rawpy(cr2_file, output_file, max_dimension=2048):
    """
    Full RAW processing with rawpy (LibRaw) and Pillow.
    
    Args:
        cr2_file (str): Path to input CR2 file
        output_file (str): Path to output image file
        max_dimension (int): Maximum width or height for output (default: 2048)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Processing CR2 with rawpy (LibRaw)...")
    
    try:
        import rawpy
        
        # Open CR2 file
        print("Opening CR2 file with rawpy...")
        with rawpy.imread(cr2_file) as raw:
            # Get RAW parameters
            print(f"RAW size: {raw.sizes.raw_width}x{raw.sizes.raw_height}")
            print(f"Output size: {raw.sizes.width}x{raw.sizes.height}")
            
            # Process RAW with default settings (high quality)
            print("Processing RAW data (demosaic, white balance, color correction)...")
            rgb = raw.postprocess(
                use_camera_wb=True,          # Use camera white balance
                half_size=False,              # Full resolution
                no_auto_bright=False,         # Auto brightness
                output_bps=8,                 # 8-bit output for web
                gamma=(2.222, 4.5),          # Standard gamma curve
                output_color=rawpy.ColorSpace.sRGB,  # sRGB color space
                user_flip=0                   # No rotation
            )
            
            print(f"Processed RGB array shape: {rgb.shape}")
            
            # Convert to PIL Image
            img = Image.fromarray(rgb)
            print(f"PIL Image size: {img.size}, mode: {img.mode}")
            
            # Resize if too large
            if img.width > max_dimension or img.height > max_dimension:
                print(f"Resizing to fit within {max_dimension}px...")
                img.thumbnail((max_dimension, max_dimension), Image.Resampling.LANCZOS)
                print(f"Resized to: {img.size}")
            
            # Save as JPEG
            print("Saving as JPEG...")
            img.save(output_file, 'JPEG', quality=92, optimize=True)
            
            # Verify output
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"RAW processed image created: {file_size} bytes")
                return True
            else:
                print("ERROR: Output file was not created")
                return False
                
    except ImportError as e:
        print(f"ERROR: rawpy not available: {e}")
        print("Please install: pip install rawpy")
        return False
    except Exception as e:
        print(f"ERROR: rawpy processing failed: {e}")
        traceback.print_exc()
        return False

def convert_cr2_to_image(cr2_file, output_file, fast=True, max_dimension=2048):
    """
    Convert CR2 to web-viewable image.
    
    Args:
        cr2_file (str): Path to input CR2 file
        output_file (str): Path to output image file
        fast (bool): Try fast preview first (embedded JPEG)
        max_dimension (int): Maximum dimension for output
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting CR2 to Image conversion...")
    print(f"Input: {cr2_file}")
    print(f"Output: {output_file}")
    print(f"Fast mode: {fast}")
    print(f"Max dimension: {max_dimension}")
    
    # Check if input file exists
    if not os.path.exists(cr2_file):
        print(f"ERROR: Input CR2 file not found: {cr2_file}")
        return False
    
    file_size = os.path.getsize(cr2_file)
    print(f"CR2 file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
    
    # Try fast preview first if enabled
    if fast:
        if extract_embedded_jpeg(cr2_file, output_file):
            print("Fast preview successful!")
            return True
        else:
            print("Fast preview not available, falling back to full RAW processing...")
    
    # Fall back to full RAW processing
    return convert_cr2_with_rawpy(cr2_file, output_file, max_dimension)

def main():
    parser = argparse.ArgumentParser(description='Convert CR2 (Canon RAW) to web-viewable image')
    parser.add_argument('cr2_file', help='Input CR2 file path')
    parser.add_argument('output_file', help='Output image file path (JPEG)')
    parser.add_argument('--no-fast', action='store_true',
                        help='Skip fast preview, use full RAW processing')
    parser.add_argument('--max-dimension', type=int, default=2048,
                        help='Maximum width or height for output (default: 2048)')
    
    args = parser.parse_args()
    
    print("=== CR2 to Image Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check required libraries
    try:
        import rawpy
        print(f"rawpy available: {rawpy.__version__}")
    except ImportError:
        print("WARNING: rawpy not available (pip install rawpy)")
    
    try:
        from PIL import Image
        print(f"Pillow available: {Image.__version__}")
    except ImportError:
        print("ERROR: Pillow not available (pip install Pillow)")
        sys.exit(1)
    
    # Check exiftool
    try:
        result = subprocess.run(['exiftool', '-ver'], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            print(f"exiftool available: {result.stdout.strip()}")
    except:
        print("WARNING: exiftool not available")
    
    # Create output directory if needed
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CR2
    success = convert_cr2_to_image(
        args.cr2_file,
        args.output_file,
        fast=not args.no_fast,
        max_dimension=args.max_dimension
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()

