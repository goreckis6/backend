#!/usr/bin/env python3
"""
JPG to PDF Converter
- Reads JPG/JPEG using Pillow
- Converts to PDF using pdfkit (optimized for quality and file size)
- Optional downscale by max dimension
- Supports page size options and quality settings
"""

import os
import sys
import argparse
import traceback
import io
import base64

try:
    from PIL import Image, ImageOps
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import pdfkit
    HAS_PDFKIT = True
    # Configure pdfkit to use system wkhtmltopdf
    # Try common paths for wkhtmltopdf
    import shutil
    wkhtmltopdf_path = shutil.which('wkhtmltopdf')
    if wkhtmltopdf_path:
        pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        print(f"pdfkit configured to use: {wkhtmltopdf_path}")
    else:
        # Try default system paths
        for path in ['/usr/bin/wkhtmltopdf', '/usr/local/bin/wkhtmltopdf']:
            if os.path.exists(path):
                pdfkit.configuration(wkhtmltopdf=path)
                print(f"pdfkit configured to use: {path}")
                break
except ImportError:
    HAS_PDFKIT = False

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.lib.utils import ImageReader
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


def convert_jpg_to_pdf(
    jpg_file: str,
    output_file: str,
    quality: int = 95,
    max_dimension: int = 4096,
    page_size: str = "auto",
    fit_to_page: bool = True
) -> bool:
    """
    Convert JPG to PDF format optimized for quality and file size
    
    Args:
        jpg_file: Input JPG file path
        output_file: Output PDF file path
        quality: Image quality (0-100, default: 95)
        max_dimension: Maximum width or height (downscale if larger)
        page_size: PDF page size ('auto', 'letter', 'a4')
        fit_to_page: Fit image to page size (default: True)
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print("Starting JPG → PDF conversion")
    print(f"Input file: {jpg_file}")
    print(f"Output file: {output_file}")
    print(f"Quality: {quality}, Max dimension: {max_dimension}, Page size: {page_size}")

    if not HAS_PIL:
        print("ERROR: Pillow not available")
        return False

    if not os.path.exists(jpg_file):
        print(f"ERROR: File not found: {jpg_file}")
        return False

    # Check file size
    file_size = os.path.getsize(jpg_file)
    print(f"File size: {file_size} bytes")
    
    if file_size == 0:
        print("ERROR: File is empty")
        return False

    try:
        # Try to open the image
        try:
            img = Image.open(jpg_file)
            print(f"Opened: format={img.format}, mode={img.mode}, size={img.size}")
            
            # Verify it's actually a JPG/JPEG image
            if img.format not in ['JPEG', 'JPG']:
                print(f"WARNING: Image format is {img.format}, expected JPEG")
        except Exception as open_error:
            print(f"ERROR: Failed to open image: {open_error}")
            print(f"ERROR: File path: {jpg_file}")
            print(f"ERROR: File exists: {os.path.exists(jpg_file)}")
            print("ERROR: The file is corrupted or not a valid JPG image")
            raise

        # Fix orientation based on EXIF data
        img = ImageOps.exif_transpose(img)
        print(f"After orientation fix: mode={img.mode}, size={img.size}")

        # Downscale if needed
        w, h = img.size
        if max(w, h) > max_dimension:
            scale = max_dimension / max(w, h)
            new_width = int(w * scale)
            new_height = int(h * scale)
            img = img.resize(
                (new_width, new_height),
                Image.Resampling.LANCZOS  # High-quality resampling
            )
            print(f"Resized to {img.size}")

        # Convert color modes for PDF
        # PDF supports RGB, RGBA (with alpha), L (grayscale)
        if img.mode == "CMYK":
            # CMYK to RGB conversion
            img = img.convert("RGB")
            print("Converted CMYK to RGB")
        elif img.mode == "P":
            # Palette mode - convert to RGB
            img = img.convert("RGB")
            print("Converted palette to RGB")
        elif img.mode not in ["RGB", "RGBA", "L", "LA"]:
            # Convert other modes to RGB
            img = img.convert("RGB")
            print(f"Converted {img.mode} to RGB")
        elif img.mode in ["RGBA", "LA"]:
            # For images with alpha, create white background
            if img.mode == "RGBA":
                rgb_img = Image.new("RGB", img.size, (255, 255, 255))
                rgb_img.paste(img, mask=img.split()[3])  # Use alpha channel as mask
                img = rgb_img
                print("Converted RGBA to RGB with white background")
            else:  # LA
                rgb_img = Image.new("RGB", img.size, (255, 255, 255))
                rgb_img.paste(img, mask=img.split()[1])  # Use alpha channel as mask
                img = rgb_img
                print("Converted LA to RGB with white background")

        # Ensure output directory exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Method 1: Try using pdfkit (if available)
        if HAS_PDFKIT:
            print("Using pdfkit for PDF conversion...")
            try:
                # Save PIL image to temporary JPEG buffer with quality
                img_buffer = io.BytesIO()
                img.save(img_buffer, format='JPEG', quality=quality, optimize=True)
                img_buffer.seek(0)
                
                # Convert image to base64 for embedding in HTML
                img_base64 = base64.b64encode(img_buffer.getvalue()).decode('utf-8')
                img_buffer.close()
                
                # Determine page size
                if page_size == 'auto':
                    img_width, img_height = img.size
                    # Use A4 if image is reasonable size, otherwise use letter
                    if img_width <= 2480 and img_height <= 3508:  # A4 at 300 DPI
                        page_width, page_height = "210mm", "297mm"  # A4
                    else:
                        page_width, page_height = "216mm", "279mm"  # Letter
                elif page_size == 'letter':
                    page_width, page_height = "216mm", "279mm"
                elif page_size == 'a4':
                    page_width, page_height = "210mm", "297mm"
                else:
                    page_width, page_height = "210mm", "297mm"  # Default A4
                
                # Create HTML with embedded image
                html_content = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                        @page {{
                            size: {page_width} {page_height};
                            margin: 0;
                        }}
                        body {{
                            margin: 0;
                            padding: 0;
                            display: flex;
                            justify-content: center;
                            align-items: center;
                            min-height: 100vh;
                        }}
                        img {{
                            max-width: 100%;
                            max-height: 100%;
                            {'width: 100%; height: auto;' if fit_to_page else ''}
                        }}
                    </style>
                </head>
                <body>
                    <img src="data:image/jpeg;base64,{img_base64}" />
                </body>
                </html>
                """
                
                # PDF options for optimization
                options = {
                    'page-size': page_size if page_size != 'auto' else 'A4',
                    'margin-top': '0mm',
                    'margin-right': '0mm',
                    'margin-bottom': '0mm',
                    'margin-left': '0mm',
                    'encoding': "UTF-8",
                    'no-outline': None,
                    'enable-local-file-access': None,
                    'quiet': '',
                }
                
                # Convert HTML to PDF using pdfkit
                pdfkit.from_string(html_content, output_file, options=options)
                
                print(f"PDF file created using pdfkit: {output_file}")
                
                # Verify output
                if os.path.exists(output_file):
                    output_size = os.path.getsize(output_file)
                    print(f"PDF file created successfully: {output_size} bytes")
                    print(f"Size ratio: {output_size / file_size:.2%}")
                    return True
                else:
                    print("ERROR: pdfkit did not create PDF file")
                    return False
                    
            except Exception as pdfkit_error:
                print(f"Warning: pdfkit conversion failed: {pdfkit_error}")
                print("Falling back to reportlab method...")
                traceback.print_exc()
                # Fall through to reportlab method
        
        # Method 2: Use reportlab (fallback)
        if HAS_REPORTLAB:
            print("Using reportlab for PDF conversion...")
            try:
                # Determine page size
                if page_size == 'auto':
                    img_width, img_height = img.size
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
                img_width, img_height = img.size
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
                img.save(img_buffer, format='JPEG', quality=quality, optimize=True)
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
                    print(f"Size ratio: {output_size / file_size:.2%}")
                    return True
                else:
                    print("ERROR: reportlab did not create PDF file")
                    return False
                    
            except Exception as reportlab_error:
                print(f"ERROR: reportlab conversion failed: {reportlab_error}")
                traceback.print_exc()
                return False
        
        # If neither method is available
        print("ERROR: Neither pdfkit nor reportlab is available. Please install one of them.")
        return False

    except Exception as e:
        print(f"ERROR: Conversion failed: {e}")
        traceback.print_exc()
        return False


def main():
    parser = argparse.ArgumentParser(description="Convert JPG/JPEG to PDF")
    parser.add_argument("jpg_file", help="Input JPG file path")
    parser.add_argument("output_file", help="Output PDF file path")
    parser.add_argument("--quality", type=int, default=95, choices=range(0, 101),
                        help="JPEG quality 0-100 (default: 95)")
    parser.add_argument("--max-dimension", type=int, default=4096,
                        help="Maximum width or height (default: 4096)")
    parser.add_argument("--page-size", choices=['auto', 'letter', 'a4'], default='auto',
                        help="PDF page size (default: auto)")
    parser.add_argument("--no-fit-to-page", action="store_true",
                        help="Do not fit image to page size")
    args = parser.parse_args()

    ok = convert_jpg_to_pdf(
        args.jpg_file,
        args.output_file,
        args.quality,
        args.max_dimension,
        args.page_size,
        not args.no_fit_to_page
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
