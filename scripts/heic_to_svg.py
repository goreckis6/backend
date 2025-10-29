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
import tempfile

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

try:
    from wand.image import Image as WandImage
    from wand.color import Color
    HAS_WAND = True
except ImportError:
    HAS_WAND = False


def convert_heic_to_svg(heic_file, output_file, quality=95, preserve_transparency=True):
    """
    Convert HEIC file to SVG format
    
    Args:
        heic_file (str): Path to input HEIC file
        output_file (str): Path to output SVG file
        quality (int): Quality for embedded image (0-100)
        preserve_transparency (bool): Preserve transparency if available
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting HEIC to SVG conversion...")
    print(f"Input: {heic_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Preserve transparency: {preserve_transparency}")
    
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
            print(f"Image opened successfully. Format: {pil_image.format}, Mode: {pil_image.mode}, Size: {pil_image.size}")
            
            # Convert image mode if needed
            if pil_image.mode == 'RGBA' and not preserve_transparency:
                print("Converting RGBA to RGB...")
                pil_image = pil_image.convert('RGB')
            elif pil_image.mode != 'RGB' and pil_image.mode != 'RGBA':
                print(f"Converting image mode from {pil_image.mode} to RGB...")
                # Create RGB background for transparency
                if pil_image.mode in ('P', 'LA', 'L'):
                    pil_image = pil_image.convert('RGB')
                else:
                    pil_image = pil_image.convert('RGB')
        except Exception as e:
            print(f"ERROR: Failed to open HEIC file: {e}")
            traceback.print_exc()
            return False
        
        # Create output directory if needed
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # Method 1: Try using wand (ImageMagick) for SVG conversion
        if HAS_WAND:
            print("Using wand (ImageMagick) for SVG conversion...")
            try:
                # Save PIL image to temporary PNG first
                temp_dir = tempfile.mkdtemp()
                temp_png = os.path.join(temp_dir, 'temp_image.png')
                
                # Save PIL image as PNG
                pil_image.save(temp_png, 'PNG')
                print(f"Saved temporary PNG: {temp_png}")
                
                # Convert PNG to SVG using wand
                with WandImage(filename=temp_png) as wand_img:
                    # Convert to SVG format (ImageMagick will embed the raster image in SVG)
                    wand_img.format = 'svg'
                    wand_img.save(filename=output_file)
                
                print(f"SVG file created using wand: {output_file}")
                
                # Clean up temp file
                try:
                    os.unlink(temp_png)
                    os.rmdir(temp_dir)
                except:
                    pass
                
                # Verify output
                if os.path.exists(output_file):
                    output_size = os.path.getsize(output_file)
                    print(f"SVG file created successfully: {output_size} bytes")
                    return True
                else:
                    print("ERROR: wand did not create SVG file")
                    return False
                    
            except Exception as wand_error:
                print(f"Warning: wand conversion failed: {wand_error}")
                print("Falling back to manual SVG creation...")
                traceback.print_exc()
                # Fall through to manual method
        
        # Method 2: Manual SVG creation with embedded base64 image
        print("Using manual SVG creation with embedded image...")
        try:
            # Save PIL image to temporary PNG with quality settings
            temp_dir = tempfile.mkdtemp()
            temp_png = os.path.join(temp_dir, 'temp_image.png')
            
            # Determine if we have transparency
            has_alpha = pil_image.mode in ('RGBA', 'LA') or 'transparency' in pil_image.info
            
            if has_alpha and preserve_transparency:
                pil_image.save(temp_png, 'PNG', optimize=True)
                image_format = 'png'
            else:
                # Convert to RGB if not already
                if pil_image.mode != 'RGB':
                    pil_image = pil_image.convert('RGB')
                pil_image.save(temp_png, 'PNG', quality=quality, optimize=True)
                image_format = 'png'
            
            # Read the PNG as base64
            import base64
            with open(temp_png, 'rb') as f:
                image_data = f.read()
            
            base64_data = base64.b64encode(image_data).decode('utf-8')
            mime_type = 'image/png' if image_format == 'png' else 'image/jpeg'
            
            # Get image dimensions
            width, height = pil_image.size
            
            # Create SVG with embedded image
            svg_content = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
  <image x="0" y="0" width="{width}" height="{height}" xlink:href="data:{mime_type};base64,{base64_data}"/>
</svg>'''
            
            # Write SVG file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            # Clean up temp file
            try:
                os.unlink(temp_png)
                os.rmdir(temp_dir)
            except:
                pass
            
            # Verify output
            if os.path.exists(output_file):
                output_size = os.path.getsize(output_file)
                print(f"SVG file created successfully: {output_size} bytes")
                return True
            else:
                print("ERROR: SVG file was not created")
                return False
                
        except Exception as manual_error:
            print(f"ERROR: Manual SVG creation failed: {manual_error}")
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
    
    args = parser.parse_args()
    
    print("=== HEIC to SVG Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    print(f"PIL (Pillow) available: {HAS_PIL}")
    print(f"pillow-heif available: {HAS_PILLOW_HEIF}")
    print(f"wand available: {HAS_WAND}")
    
    success = convert_heic_to_svg(
        args.heic_file,
        args.output_file,
        quality=args.quality,
        preserve_transparency=not args.no_transparency
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)


if __name__ == '__main__':
    main()

