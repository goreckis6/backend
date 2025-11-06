#!/usr/bin/env python3
"""
PDF Compressor
Compresses PDF files using PyPDF2/pypdf with adjustable compression settings.
Supports optimization for better file size reduction.
"""

import os
import sys
import argparse
import traceback

try:
    from pypdf import PdfWriter, PdfReader
    HAS_PYPDF = True
except ImportError:
    try:
        from PyPDF2 import PdfWriter, PdfReader
        HAS_PYPDF = True
    except ImportError:
        HAS_PYPDF = False


def compress_pdf(input_file: str, output_file: str, quality: int = 85, optimize: bool = True) -> tuple[bool, int, int]:
    """
    Compress PDF file.
    
    Args:
        input_file: Path to input PDF file
        output_file: Path to output compressed PDF file
        quality: Compression quality (1-100, higher = better quality but larger file)
        optimize: Apply additional optimization
    
    Returns:
        tuple: (success: bool, original_size: int, compressed_size: int)
    """
    print("Starting PDF compression")
    print(f"Input: {input_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Optimize: {optimize}")

    if not HAS_PYPDF:
        print("ERROR: PyPDF2/pypdf not available")
        return False, 0, 0

    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        return False, 0, 0

    try:
        # Get original file size
        original_size = os.path.getsize(input_file)
        print(f"Original file size: {original_size} bytes ({original_size / 1024:.2f} KB)")

        # Read PDF
        reader = PdfReader(input_file)
        writer = PdfWriter()

        print(f"PDF has {len(reader.pages)} pages")

        # Add all pages with compression
        for page_num, page in enumerate(reader.pages):
            # Compress page content streams based on quality
            # Lower quality = more compression
            if quality < 50:
                # High compression - compress all streams
                page.compress_content_streams()
            elif quality < 80:
                # Medium compression - compress large streams
                if hasattr(page, 'compress_content_streams'):
                    page.compress_content_streams()
            
            writer.add_page(page)
            print(f"Processed page {page_num + 1}/{len(reader.pages)}")

        # Handle metadata - pypdf 5.0.0+ requires all keys to start with '/'
        if reader.metadata:
            if optimize:
                # Keep minimal metadata with proper '/'-prefixed keys
                metadata = {}
                if '/Title' in reader.metadata:
                    metadata['/Title'] = reader.metadata.get('/Title', '')
                if '/Author' in reader.metadata:
                    metadata['/Author'] = reader.metadata.get('/Author', '')
                if '/Creator' in reader.metadata:
                    metadata['/Creator'] = reader.metadata.get('/Creator', '')
                metadata['/Producer'] = 'PDF Compressor'
                if metadata:
                    writer.add_metadata(metadata)
            else:
                # Copy all metadata, ensuring keys start with '/'
                metadata = {}
                for key, value in reader.metadata.items():
                    # Ensure key starts with '/'
                    normalized_key = key if key.startswith('/') else f'/{key}'
                    metadata[normalized_key] = value
                if metadata:
                    writer.add_metadata(metadata)

        # Ensure output dir exists
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # Write compressed PDF
        with open(output_file, 'wb') as output_pdf:
            writer.write(output_pdf)

        # Get compressed file size
        if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
            compressed_size = os.path.getsize(output_file)
            savings_percent = ((original_size - compressed_size) / original_size * 100) if original_size > 0 else 0
            print(f"Compressed file size: {compressed_size} bytes ({compressed_size / 1024:.2f} KB)")
            print(f"Compression savings: {savings_percent:.2f}% ({original_size - compressed_size} bytes saved)")
            print("PDF compressed successfully")
            return True, original_size, compressed_size
        
        print("ERROR: Compressed output not created")
        return False, original_size, 0
    except Exception as e:
        print(f"ERROR: Failed to compress PDF: {e}")
        traceback.print_exc()
        return False, 0, 0


def main():
    parser = argparse.ArgumentParser(description='Compress PDF file')
    parser.add_argument('input_file', help='Path to input PDF file')
    parser.add_argument('output_file', help='Path to output compressed PDF file')
    parser.add_argument('--quality', type=int, default=85, help='PDF compression quality (1-100, default: 85)')
    parser.add_argument('--optimize', action='store_true', default=True, help='Apply optimization (default: True)')
    parser.add_argument('--no-optimize', action='store_false', dest='optimize', help='Disable optimization')
    args = parser.parse_args()

    print("=== PDF Compressor ===")
    print(f"Python: {sys.version}")
    print(f"Args: {vars(args)}")

    success, original_size, compressed_size = compress_pdf(
        args.input_file, 
        args.output_file, 
        args.quality, 
        args.optimize
    )
    
    if success:
        print(f"SUCCESS: Original={original_size} bytes, Compressed={compressed_size} bytes")
        sys.exit(0)
    else:
        print("FAILURE: Compression failed")
        sys.exit(1)


if __name__ == '__main__':
    main()



