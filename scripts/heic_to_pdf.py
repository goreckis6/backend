#!/usr/bin/env python3
"""
HEIC to PDF Converter
Converts HEIC (High Efficiency Image Container) files to PDF format
Uses pillow-heif to read HEIC, PIL for image processing, and reportlab/wand for PDF conversion
"""

import os
import sys
import argparse
import traceback
import io

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
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.utils import ImageReader
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

try:
    from wand.image import Image as WandImage
    HAS_WAND = True
except ImportError:
    HAS_WAND = False


def convert_heic_to_pdf(heic_file, output_file, quality=95, page_size='auto', fit_to_page=True, max_dimension=8192):
    """
    Convert HEIC file to PDF format (optimized for speed)
    
    Args:
        heic_file (str): Path to input HEIC file
        output_file (str): Path to output PDF file
        quality (int): Quality for embedded image (0-100)
        page_size (str): Page size ('auto', 'letter', 'a4')
        fit_to_page (bool): Fit image to page size
        max_dimension (int): Maximum width or height to limit processing time (default: 8192)
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting HEIC to PDF conversion (optimized)...")
    print(f"Input: {heic_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Page size: {page_size}")
    print(f"Fit to page: {fit_to_page}")
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
        except Exception as open_error:
            # Log detailed error to console (for server logs) but don't expose file paths in user-facing messages
            print(f"ERROR: Failed to open image: {open_error}")
            print(f"ERROR: File path: {heic_file}")
            # Print user-friendly error message (without file paths)
            print("ERROR: The file is corrupted or not a valid HEIC image")
            raise
            
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
                
                # Use fast resampling for speed
                pil_image = pil_image.resize((new_width, new_height), Image.Resampling.BILINEAR)
                width, height = new_width, new_height
                print(f"Resized to: {width}x{height}")
            
            # Convert to RGB (PDF doesn't support alpha channel)
            if pil_image.mode != 'RGB':
                if pil_image.mode in ('RGBA', 'LA'):
                    # Create white background for transparency
                    rgb_image = Image.new('RGB', pil_image.size, (255, 255, 255))
                    if pil_image.mode == 'RGBA':
                        rgb_image.paste(pil_image, mask=pil_image.split()[3])  # Use alpha channel as mask
                    else:
                        rgb_image.paste(pil_image)
                    pil_image = rgb_image
                else:
                    pil_image = pil_image.convert('RGB')
                print(f"Converted to RGB mode")
        except Exception as e:
            print(f"ERROR: Failed to open HEIC file: {e}")
            traceback.print_exc()
            return False
        
        # Create output directory if needed
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # Method 1: Try using wand (ImageMagick) - faster and better quality
        if HAS_WAND:
            print("Using wand (ImageMagick) for PDF conversion...")
            try:
                # Save PIL image to temporary JPEG first
                img_buffer = io.BytesIO()
                pil_image.save(img_buffer, format='JPEG', quality=quality, optimize=False)
                img_buffer.seek(0)
                
                # Convert to PDF using wand
                with WandImage(blob=img_buffer.getvalue()) as wand_img:
                    wand_img.format = 'pdf'
                    wand_img.save(filename=output_file)
                
                img_buffer.close()
                
                print(f"PDF file created using wand: {output_file}")
                
                # Verify output
                if os.path.exists(output_file):
                    output_size = os.path.getsize(output_file)
                    print(f"PDF file created successfully: {output_size} bytes")
                    return True
                else:
                    print("ERROR: wand did not create PDF file")
                    return False
                    
            except Exception as wand_error:
                print(f"Warning: wand conversion failed: {wand_error}")
                print("Falling back to reportlab method...")
                traceback.print_exc()
                # Fall through to reportlab method
        
        # Method 2: Use reportlab (fallback)
        if HAS_REPORTLAB:
            print("Using reportlab for PDF conversion...")
            try:
                # Determine page size
                if page_size == 'auto':
                    # Use image dimensions, but with reasonable limits
                    img_width, img_height = pil_image.size
                    # Use A4 as base if image is reasonable size
                    if img_width <= 2480 and img_height <= 3508:  # A4 at 300 DPI
                        pdf_width, pdf_height = A4
                    else:
                        # Use letter size for larger images
                        pdf_width, pdf_height = letter
                elif page_size == 'letter':
                    pdf_width, pdf_height = letter
                elif page_size == 'a4':
                    pdf_width, pdf_height = A4
                else:
                    pdf_width, pdf_height = A4
                
                # Create PDF
                c = canvas.Canvas(output_file, pagesize=(pdf_width, pdf_height))
                
                # Calculate image placement
                img_width, img_height = pil_image.size
                if fit_to_page:
                    # Scale to fit page while maintaining aspect ratio
                    scale_w = pdf_width / img_width
                    scale_h = pdf_height / img_height
                    scale = min(scale_w, scale_h) * 0.95  # 95% to add margins
                    scaled_width = img_width * scale
                    scaled_height = img_height * scale
                    x = (pdf_width - scaled_width) / 2
                    y = (pdf_height - scaled_height) / 2
                else:
                    # Use original size, centered
                    scaled_width = min(img_width, pdf_width)
                    scaled_height = min(img_height, pdf_height)
                    x = (pdf_width - scaled_width) / 2
                    y = (pdf_height - scaled_height) / 2
                
                # Save PIL image to buffer for reportlab
                img_buffer = io.BytesIO()
                pil_image.save(img_buffer, format='JPEG', quality=quality, optimize=False)
                img_buffer.seek(0)
                
                # Draw image on PDF
                img_reader = ImageReader(img_buffer)
                c.drawImage(img_reader, x, y, width=scaled_width, height=scaled_height, preserveAspectRatio=True)
                c.save()
                
                img_buffer.close()
                
                print(f"PDF file created using reportlab: {output_file}")
                
                # Verify output
                if os.path.exists(output_file):
                    output_size = os.path.getsize(output_file)
                    print(f"PDF file created successfully: {output_size} bytes")
                    return True
                else:
                    print("ERROR: reportlab did not create PDF file")
                    return False
                    
            except Exception as reportlab_error:
                print(f"ERROR: reportlab conversion failed: {reportlab_error}")
                traceback.print_exc()
                return False
        
        # If neither method is available
        print("ERROR: Neither wand nor reportlab is available. Please install one of them.")
        return False
            
    except Exception as e:
        print(f"ERROR: Failed to convert HEIC to PDF: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description='Convert HEIC file to PDF format')
    parser.add_argument('heic_file', help='Path to input HEIC file')
    parser.add_argument('output_file', help='Path to output PDF file')
    parser.add_argument('--quality', type=int, default=95,
                        help='Quality for embedded image (0-100, default: 95)')
    parser.add_argument('--page-size', choices=['auto', 'letter', 'a4'], default='auto',
                        help='PDF page size (default: auto)')
    parser.add_argument('--no-fit-to-page', action='store_true',
                        help='Do not fit image to page size')
    parser.add_argument('--max-dimension', type=int, default=8192,
                        help='Maximum width or height in pixels (default: 8192, use lower for faster conversion)')
    
    args = parser.parse_args()
    
    print("=== HEIC to PDF Converter (Optimized) ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    print(f"PIL (Pillow) available: {HAS_PIL}")
    print(f"pillow-heif available: {HAS_PILLOW_HEIF}")
    print(f"reportlab available: {HAS_REPORTLAB}")
    print(f"wand available: {HAS_WAND}")
    
    success = convert_heic_to_pdf(
        args.heic_file,
        args.output_file,
        quality=args.quality,
        page_size=args.page_size,
        fit_to_page=not args.no_fit_to_page,
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

