#!/usr/bin/env python3
"""
HEIC to SVG Converter
Converts HEIC (High Efficiency Image Container) files to SVG format
Uses pillow-heif to read HEIC, PIL for image processing, and wand (ImageMagick) for SVG conversion
"""

import os
import sys
import argparse
import traceback
import io
import base64

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from pillow_heif import register_heif_opener
    HAS_PILLOW_HEIF = True
except ImportError:
    HAS_PILLOW_HEIF = False

def convert_heic_to_svg(heic_file, output_file, quality=95, preserve_transparency=True, max_dimension=8192):
    """
    Convert HEIC file to SVG format (optimized for speed)
    
    Args:
        heic_file (str): Path to input HEIC file
        output_file (str): Path to output SVG file
        quality (int): Quality for embedded image (0-100)
        preserve_transparency (bool): Preserve transparency if available
        max_dimension (int): Maximum width or height to limit processing time (default: 8192)
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting HEIC to SVG conversion (optimized)...")
    print(f"Input: {heic_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Preserve transparency: {preserve_transparency}")
    print(f"Max dimension: {max_dimension}")
    
    try:
        # Check if HEIC file exists
        if not os.path.exists(heic_file):
            print(f"ERROR: HEIC file does not exist: {heic_file}")
            return False
        
        file_size = os.path.getsize(heic_file)
        print(f"HEIC file size: {file_size} bytes")
        
        if file_size == 0:
            print("ERROR: Input file is empty")
            return False
        
        # Register HEIF opener if pillow-heif is available
        if HAS_PILLOW_HEIF:
            try:
                register_heif_opener()
                print("HEIF opener registered successfully")
            except Exception as e:
                print(f"Warning: Could not register HEIF opener: {e}")
        
        # Read HEIC file using PIL
        if not HAS_PIL:
            print("ERROR: PIL (Pillow) is required but not available")
            return False
        
        print("Opening HEIC file with PIL...")
        try:
            pil_image = Image.open(heic_file)
            original_size = pil_image.size
            print(f"Image opened successfully. Format: {pil_image.format}, Mode: {pil_image.mode}, Size: {original_size}")
            
            # OPTIMIZATION: Resize if image is too large (speed optimization)
            width, height = original_size
            if max(width, height) > max_dimension:
                print(f"Resizing large image from {width}x{height} to max {max_dimension}px for faster processing...")
                # Maintain aspect ratio
                if width > height:
                    new_width = max_dimension
                    new_height = int(height * (max_dimension / width))
                else:
                    new_height = max_dimension
                    new_width = int(width * (max_dimension / height))
                
                # Use fast resampling for speed (LANCZOS is slower but better quality)
                # Using NEAREST is fastest, but BILINEAR is a good balance
                pil_image = pil_image.resize((new_width, new_height), Image.Resampling.BILINEAR)
                print(f"Resized to: {new_width}x{new_height}")
            
            # Convert image mode if needed
            has_alpha = pil_image.mode in ('RGBA', 'LA') or 'transparency' in pil_image.info
            
            if has_alpha and preserve_transparency:
                # Keep RGBA for transparency
                if pil_image.mode != 'RGBA' and pil_image.mode != 'LA':
                    if pil_image.mode == 'P':
                        pil_image = pil_image.convert('RGBA')
                    else:
                        pil_image = pil_image.convert('RGBA')
            else:
                # Convert to RGB (faster processing, smaller file)
                if pil_image.mode != 'RGB':
                    pil_image = pil_image.convert('RGB')
        except Exception as e:
            print(f"ERROR: Failed to open HEIC file: {e}")
            traceback.print_exc()
            return False
        
        # Create output directory if needed
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # OPTIMIZED METHOD: Direct PIL-to-base64 without temp file
        print("Creating SVG with optimized direct conversion...")
        try:
            # Use in-memory buffer instead of temp file (much faster)
            img_buffer = io.BytesIO()
            
            # OPTIMIZATION: Use JPEG for RGB images (much smaller and faster than PNG)
            # PNG only when transparency is needed
            if has_alpha and preserve_transparency:
                # Save as PNG with minimal compression for speed
                try:
                    pil_image.save(img_buffer, format='PNG', optimize=False, compress_level=1)
                except TypeError:
                    # Fallback if compress_level not supported
                    pil_image.save(img_buffer, format='PNG', optimize=False)
                mime_type = 'image/png'
                print("Saved as PNG (with transparency)")
            else:
                # Save as JPEG (much faster and smaller - 3-5x faster than PNG)
                pil_image.save(img_buffer, format='JPEG', quality=quality, optimize=False)
                mime_type = 'image/jpeg'
                print(f"Saved as JPEG (quality={quality})")
            
            # Get image data from buffer
            img_buffer.seek(0)
            image_data = img_buffer.getvalue()
            img_buffer.close()
            
            # OPTIMIZATION: Use efficient base64 encoding
            base64_data = base64.b64encode(image_data).decode('utf-8')
            
            # Get final image dimensions
            width, height = pil_image.size
            
            # Create SVG with embedded image (minimal XML for speed)
            svg_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
<image width="{width}" height="{height}" xlink:href="data:{mime_type};base64,{base64_data}"/>
</svg>'''
            
            # Write SVG file directly
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            # Verify output
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"SVG file created successfully: {output_size} bytes")
                print(f"Compression ratio: {file_size/float(output_size):.2f}x")
                return True
            else:
                print("ERROR: SVG file was not created")
                return False
                
        except Exception as conversion_error:
            print(f"ERROR: Direct conversion failed: {conversion_error}")
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to convert HEIC to SVG: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIC file to SVG format')
    parser.add_argument('heic_file', help='Path to input HEIC file')
    parser.add_argument('output_file', help='Path to output SVG file')
    parser.add_argument('--quality', type=int, default=95,
                        help='Quality for embedded image (0-100, default: 95)')
    parser.add_argument('--no-transparency', action='store_true',
                        help='Do not preserve transparency')
    parser.add_argument('--max-dimension', type=int, default=8192,
                        help='Maximum width or height in pixels (default: 8192, use lower for faster conversion)')
    
    args = parser.parse_args()
    
    print("=== HEIC to SVG Converter (Optimized) ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    print(f"PIL (Pillow) available: {HAS_PIL}")
    print(f"pillow-heif available: {HAS_PILLOW_HEIF}")
    
    success = convert_heic_to_svg(
        args.heic_file,
        args.output_file,
        quality=args.quality,
        preserve_transparency=not args.no_transparency,
        max_dimension=args.max_dimension
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()

