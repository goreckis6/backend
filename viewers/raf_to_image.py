#!/usr/bin/env python3
"""
RAF (Fujifilm RAW) to Image converter for web preview.
Uses rawpy (LibRaw), Pillow for rendering, and exiftool for fast embedded JPEG extraction.
Handles X-Trans sensor and GFX medium format.
"""

import argparse
import os
import sys
import subprocess
import traceback
import json
from PIL import Image
import io
from datetime import datetime

def extract_exif_metadata(raf_file):
    """Extract EXIF metadata from RAF file using exiftool."""
    metadata = {
        'dateTaken': 'N/A', 'dimensions': 'N/A', 'fileSize': 'N/A',
        'iso': 'N/A', 'camera': 'N/A', 'exposure': 'N/A'
    }
    try:
        file_size = os.path.getsize(raf_file)
        metadata['fileSize'] = f"{file_size / 1024 / 1024:.2f} MB"
        
        cmd = ['exiftool', '-json', '-DateTimeOriginal', '-ImageWidth', '-ImageHeight',
               '-ISO', '-Model', '-Make', '-ExposureTime', '-FNumber', '-FocalLength', raf_file]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            data = json.loads(result.stdout)[0]
            if 'DateTimeOriginal' in data:
                try:
                    dt = datetime.strptime(data['DateTimeOriginal'], '%Y:%m:%d %H:%M:%S')
                    metadata['dateTaken'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                except:
                    metadata['dateTaken'] = data['DateTimeOriginal']
            if 'ImageWidth' in data and 'ImageHeight' in data:
                metadata['dimensions'] = f"{data['ImageWidth']} × {data['ImageHeight']} px"
            if 'ISO' in data:
                metadata['iso'] = str(data['ISO'])
            make = data.get('Make', '')
            model = data.get('Model', '')
            if make and model:
                metadata['camera'] = f"{make} {model}"
            elif model:
                metadata['camera'] = model
            exposure_parts = []
            if 'ExposureTime' in data:
                exposure_parts.append(f"{data['ExposureTime']}s")
            if 'FNumber' in data:
                exposure_parts.append(f"f/{data['FNumber']}")
            if 'FocalLength' in data:
                focal = data['FocalLength']
                if 'mm' not in str(focal):
                    focal = f"{focal}mm"
                exposure_parts.append(str(focal))
            if exposure_parts:
                metadata['exposure'] = ' • '.join(exposure_parts)
        print(f"Extracted metadata: {metadata}")
    except Exception as e:
        print(f"WARNING: Could not extract metadata: {e}")
    return metadata

def extract_embedded_jpeg(raf_file, output_file):
    """
    Fast preview using exiftool to extract embedded JPEG (no demosaic).
    
    Args:
        raf_file (str): Path to input RAF file
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
            raf_file
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

def convert_raf_with_rawpy(raf_file, output_file, max_dimension=2048):
    """
    Full RAW processing with rawpy (LibRaw) and Pillow.
    Handles Fujifilm X-Trans sensor and GFX medium format.
    
    Args:
        raf_file (str): Path to input RAF file
        output_file (str): Path to output image file
        max_dimension (int): Maximum width or height for output (default: 2048)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Processing RAF with rawpy (LibRaw)...")
    
    try:
        import rawpy
        
        # Open RAF file
        print("Opening RAF file with rawpy...")
        with rawpy.imread(raf_file) as raw:
            # Get RAW parameters
            print(f"RAW size: {raw.sizes.raw_width}x{raw.sizes.raw_height}")
            print(f"Output size: {raw.sizes.width}x{raw.sizes.height}")
            
            # Process RAW with settings optimized for X-Trans
            print("Processing X-Trans RAW data (demosaic, white balance, color correction)...")
            rgb = raw.postprocess(
                use_camera_wb=True,          # Use camera white balance (Fujifilm color science)
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
                print(f"X-Trans processed image created: {file_size} bytes")
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

def convert_raf_to_image(raf_file, output_file, fast=True, max_dimension=2048):
    """
    Convert RAF to web-viewable image.
    
    Args:
        raf_file (str): Path to input RAF file
        output_file (str): Path to output image file
        fast (bool): Try fast preview first (embedded JPEG)
        max_dimension (int): Maximum dimension for output
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting RAF to Image conversion...")
    print(f"Input: {raf_file}")
    print(f"Output: {output_file}")
    print(f"Fast mode: {fast}")
    print(f"Max dimension: {max_dimension}")
    
    # Check if input file exists
    if not os.path.exists(raf_file):
        print(f"ERROR: Input RAF file not found: {raf_file}")
        return False
    
    file_size = os.path.getsize(raf_file)
    print(f"RAF file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
    
    # Try fast preview first if enabled
    if fast:
        if extract_embedded_jpeg(raf_file, output_file):
            print("Fast preview successful!")
            return True
        else:
            print("Fast preview not available, falling back to full RAW processing...")
    
    # Fall back to full RAW processing
    return convert_raf_with_rawpy(raf_file, output_file, max_dimension)

def main():
    parser = argparse.ArgumentParser(description='Convert RAF (Fujifilm RAW) to web-viewable image')
    parser.add_argument('raf_file', help='Input RAF file path')
    parser.add_argument('output_file', help='Output image file path (JPEG)')
    parser.add_argument('metadata_file', help='Output metadata file path (JSON)')
    parser.add_argument('--no-fast', action='store_true',
                        help='Skip fast preview, use full RAW processing')
    parser.add_argument('--max-dimension', type=int, default=2048,
                        help='Maximum width or height for output (default: 2048)')
    
    args = parser.parse_args()
    
    print("=== RAF to Image Converter ===")
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
    
    # Extract metadata
    print("Extracting EXIF metadata...")
    metadata = extract_exif_metadata(args.raf_file)
    
    # Save metadata to JSON file
    with open(args.metadata_file, 'w') as f:
        json.dump(metadata, f, indent=2)
    print(f"Metadata saved to: {args.metadata_file}")
    
    # Convert RAF
    success = convert_raf_to_image(
        args.raf_file,
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


