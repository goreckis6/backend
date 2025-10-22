#!/usr/bin/env python3
"""
CSV to MOBI Converter
Converts CSV files to MOBI format using pandas, ebooklib, and Calibre
"""

import sys
import os
import pandas as pd
import argparse
from pathlib import Path
import tempfile
import subprocess
import logging
from ebooklib import epub

# Try to import Calibre Python API
try:
    from calibre.ebooks.conversion import convert
    from calibre.ebooks.conversion.cli import main as calibre_main
    CALIBRE_AVAILABLE = True
except ImportError:
    CALIBRE_AVAILABLE = False

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def convert_epub_to_mobi_python_api(epub_path, mobi_path, book_title, author):
    """
    Convert EPUB to MOBI using Calibre Python API
    """
    try:
        logger.info("Using Calibre Python API for conversion...")
        
        # Set up conversion options
        options = {
            'mobi_file_type': 'old',
            'disable_font_rescaling': True,
            'title': book_title,
            'authors': [author],
            'language': 'en'
        }
        
        # Convert using Calibre Python API
        convert(epub_path, mobi_path, options)
        logger.info("EPUB to MOBI conversion completed using Python API")
        return True
        
    except Exception as e:
        logger.error(f"Calibre Python API conversion failed: {e}")
        return False


def convert_csv_to_mobi(csv_path, output_path, book_title=None, author=None, include_headers=True, chunk_size=1000):
    """
    Convert CSV file to MOBI format using the correct approach:
    1. Read CSV with pandas
    2. Generate EPUB with ebooklib
    3. Convert EPUB → MOBI using ebook-convert (from Calibre)
    
    Args:
        csv_path (str): Path to input CSV file
        output_path (str): Path to output MOBI file
        book_title (str): Title for the MOBI book
        author (str): Author for the MOBI book
        include_headers (bool): Whether to include CSV headers
        chunk_size (int): Number of rows to process at a time
    """
    try:
        logger.info(f"Starting CSV to MOBI conversion: {csv_path} -> {output_path}")
        
        # Step 1: Read CSV with pandas
        logger.info("Step 1: Reading CSV file with pandas...")
        df = pd.read_csv(csv_path)
        
        if df.empty:
            raise ValueError("CSV file is empty")
        
        logger.info(f"CSV file contains {len(df)} rows and {len(df.columns)} columns")
        
        # Set default values
        if not book_title:
            book_title = Path(csv_path).stem.replace('_', ' ').title()
        if not author:
            author = "CSV Converter"
        
        # Step 2: Generate EPUB with ebooklib
        logger.info("Step 2: Generating EPUB with ebooklib...")
        epub_path = output_path.replace('.mobi', '.epub')
        
        # Create EPUB book
        book = epub.EpubBook()
        book.set_identifier(f"csv-{Path(csv_path).stem}")
        book.set_title(book_title)
        book.set_language('en')
        book.add_author(author)
        
        # Create table of contents
        toc = []
        
        # Create main content chapter
        chapter = epub.EpubHtml(title='Data Table', file_name='data_table.xhtml', lang='en')
        
        # Generate HTML content for the table
        html_content = f"""
        <html>
        <head>
            <title>{book_title}</title>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; color: #333; }}
                h1 {{ color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }}
                h2 {{ color: #34495e; margin-top: 30px; }}
                table {{ width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; font-weight: bold; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
            </style>
        </head>
        <body>
            <h1>{book_title}</h1>
            <p><em>by {author}</em></p>
            <h2>Data Table</h2>
            <table>
"""
        
        # Add table headers if requested
        if include_headers and not df.columns.empty:
            html_content += "                <thead>\n                    <tr>\n"
            for col in df.columns:
                html_content += f"                        <th>{col}</th>\n"
            html_content += "                    </tr>\n                </thead>\n"
        
        # Add table body
        html_content += "                <tbody>\n"
        
        # Process data in chunks for large files
        total_rows = len(df)
        processed_rows = 0
        
        for start_idx in range(0, total_rows, chunk_size):
            end_idx = min(start_idx + chunk_size, total_rows)
            chunk_df = df.iloc[start_idx:end_idx]
            
            for _, row in chunk_df.iterrows():
                html_content += "                    <tr>\n"
                for value in row:
                    # Handle NaN values and escape HTML
                    if pd.isna(value):
                        cell_value = ""
                    else:
                        cell_value = str(value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    html_content += f"                        <td>{cell_value}</td>\n"
                html_content += "                    </tr>\n"
            
            processed_rows += len(chunk_df)
            logger.info(f"Processed {processed_rows}/{total_rows} rows")
        
        html_content += """                </tbody>
            </table>
        </body>
        </html>"""
        
        chapter.content = html_content
        book.add_item(chapter)
        toc.append(chapter)
        
        # Add table of contents
        book.toc = toc
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        
        # Write EPUB file
        epub.write_epub(epub_path, book, {})
        logger.info(f"Created EPUB file: {epub_path}")
        
        # Step 3: Convert EPUB → MOBI using Calibre
        logger.info("Step 3: Converting EPUB to MOBI using Calibre...")
        
        # Try Python API first if available
        if CALIBRE_AVAILABLE:
            logger.info("Calibre Python API is available, trying that first...")
            if convert_epub_to_mobi_python_api(epub_path, output_path, book_title, author):
                logger.info("MOBI conversion completed successfully using Python API")
            else:
                logger.warning("Python API failed, falling back to command line...")
                CALIBRE_AVAILABLE = False  # Force fallback
        
        # Fallback to command line if Python API failed or not available
        if not CALIBRE_AVAILABLE:
            logger.info("Using command line ebook-convert...")
            
            # Try to find ebook-convert
            ebook_convert_paths = [
                'ebook-convert',
                '/usr/bin/ebook-convert',
                '/usr/local/bin/ebook-convert',
                '/opt/calibre/bin/ebook-convert',
                '/usr/bin/calibre-ebook-convert',
                '/usr/local/bin/calibre-ebook-convert',
                '/usr/bin/calibre',
                '/usr/local/bin/calibre'
            ]
            
            ebook_convert = None
            for path in ebook_convert_paths:
                try:
                    result = subprocess.run([path, '--version'], capture_output=True, check=True, timeout=10)
                    logger.info(f"Found ebook-convert at: {path}")
                    ebook_convert = path
                    break
                except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired):
                    logger.debug(f"ebook-convert not found at: {path}")
                    continue
            
            if not ebook_convert:
                # Check if calibre is installed but ebook-convert is not in PATH
                logger.info("ebook-convert not found in standard paths, searching for Calibre installation...")
                
                # Check common Calibre installation locations
                calibre_locations = [
                    '/usr/bin/calibre',
                    '/usr/local/bin/calibre',
                    '/opt/calibre/bin/calibre',
                    '/usr/share/calibre/bin/calibre'
                ]
                
                for calibre_path in calibre_locations:
                    if os.path.exists(calibre_path):
                        logger.info(f"Calibre found at: {calibre_path}")
                        # Try to find ebook-convert in calibre directory
                        calibre_dir = os.path.dirname(calibre_path)
                        ebook_convert_candidates = [
                            os.path.join(calibre_dir, 'ebook-convert'),
                            os.path.join(calibre_dir, 'calibre-ebook-convert'),
                            os.path.join(os.path.dirname(calibre_dir), 'bin', 'ebook-convert'),
                            os.path.join(calibre_dir, '..', 'bin', 'ebook-convert')
                        ]
                        for candidate in ebook_convert_candidates:
                            if os.path.exists(candidate):
                                ebook_convert = candidate
                                logger.info(f"Found ebook-convert at: {candidate}")
                                break
                        if ebook_convert:
                            break
                
                # Also try using 'which' command as fallback
                if not ebook_convert:
                    try:
                        result = subprocess.run(['which', 'calibre'], capture_output=True, text=True)
                        if result.returncode == 0:
                            calibre_path = result.stdout.strip()
                            logger.info(f"Calibre found via which at: {calibre_path}")
                            # Try to find ebook-convert in calibre directory
                            calibre_dir = os.path.dirname(calibre_path)
                            ebook_convert_candidates = [
                                os.path.join(calibre_dir, 'ebook-convert'),
                                os.path.join(calibre_dir, 'calibre-ebook-convert'),
                                os.path.join(os.path.dirname(calibre_dir), 'bin', 'ebook-convert')
                            ]
                            for candidate in ebook_convert_candidates:
                                if os.path.exists(candidate):
                                    ebook_convert = candidate
                                    logger.info(f"Found ebook-convert at: {candidate}")
                                    break
                    except Exception as e:
                        logger.debug(f"Error checking calibre installation with which: {e}")
            
            if not ebook_convert:
                # Debug: List what's available in common directories
                logger.error("ebook-convert not found. Debugging system...")
                debug_dirs = ['/usr/bin', '/usr/local/bin', '/opt/calibre/bin', '/usr/share/calibre/bin']
                for debug_dir in debug_dirs:
                    if os.path.exists(debug_dir):
                        try:
                            files = os.listdir(debug_dir)
                            calibre_files = [f for f in files if 'calibre' in f.lower() or 'ebook' in f.lower()]
                            if calibre_files:
                                logger.info(f"Files in {debug_dir}: {calibre_files}")
                        except Exception as e:
                            logger.debug(f"Could not list {debug_dir}: {e}")
                
                logger.error("ebook-convert not found. Cannot convert to MOBI format.")
                raise RuntimeError("ebook-convert not found. Please ensure Calibre is installed and available on the PATH.")
            
            # Convert EPUB to MOBI
            convert_cmd = [
                ebook_convert,
                epub_path,
                output_path,
                '--title', book_title,
                '--authors', author,
                '--language', 'en',
                '--mobi-file-type', 'old',
                '--disable-font-rescaling'
            ]
            
            logger.info(f"Running command: {' '.join(convert_cmd)}")
            result = subprocess.run(convert_cmd, capture_output=True, text=True, timeout=300)
            
            if result.returncode != 0:
                logger.error(f"ebook-convert failed with return code {result.returncode}")
                logger.error(f"stderr: {result.stderr}")
                raise RuntimeError(f"ebook-convert failed: {result.stderr}")
            
            logger.info("MOBI conversion completed successfully using command line")
        
        # Clean up temporary EPUB file
        try:
            os.unlink(epub_path)
            logger.info("Cleaned up temporary EPUB file")
        except OSError:
            logger.warning("Failed to clean up temporary EPUB file")
        
        # Verify output file was created
        if not os.path.exists(output_path):
            raise RuntimeError("MOBI file was not created")
        
        file_size = os.path.getsize(output_path)
        logger.info(f"MOBI file created successfully: {output_path} ({file_size} bytes)")
        
        return True
        
    except Exception as e:
        logger.error(f"Error converting CSV to MOBI: {str(e)}")
        raise

def main():
    parser = argparse.ArgumentParser(description='Convert CSV file to MOBI format')
    parser.add_argument('input_csv', help='Input CSV file path')
    parser.add_argument('output_mobi', help='Output MOBI file path')
    parser.add_argument('--title', help='Book title (default: filename)')
    parser.add_argument('--author', help='Book author (default: "CSV Converter")')
    parser.add_argument('--no-headers', action='store_true', help='Do not include CSV headers')
    parser.add_argument('--chunk-size', type=int, default=1000, help='Number of rows to process at a time')
    
    args = parser.parse_args()
    
    try:
        success = convert_csv_to_mobi(
            csv_path=args.input_csv,
            output_path=args.output_mobi,
            book_title=args.title,
            author=args.author,
            include_headers=not args.no_headers,
            chunk_size=args.chunk_size
        )
        
        if success or success is None:
            # success is None when function completes without explicit return
            print(f"Successfully converted {args.input_csv} to {args.output_mobi}")
            sys.exit(0)
        else:
            print("Conversion failed")
            sys.exit(1)
            
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()
