#!/usr/bin/env python3
"""
EPUB to CSV Converter
Extracts text content from EPUB files and converts to CSV format
"""

import os
import sys
import argparse
import csv
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup
import re
import traceback

def clean_text(text):
    """Clean and normalize text content"""
    if not text:
        return ""
    
    # Remove excessive whitespace
    text = re.sub(r'\s+', ' ', text)
    # Remove leading/trailing whitespace
    text = text.strip()
    return text

def extract_epub_content(epub_file):
    """
    Extract content from EPUB file and return structured data
    
    Returns:
        list: List of dictionaries containing chapter/section data
    """
    print(f"Reading EPUB file: {epub_file}")
    
    try:
        book = epub.read_epub(epub_file)
    except Exception as e:
        print(f"ERROR: Failed to read EPUB file: {e}")
        return None
    
    # Get metadata
    title = book.get_metadata('DC', 'title')
    author = book.get_metadata('DC', 'creator')
    
    book_title = title[0][0] if title else "Unknown"
    book_author = author[0][0] if author else "Unknown"
    
    print(f"Book Title: {book_title}")
    print(f"Book Author: {book_author}")
    
    content_data = []
    chapter_num = 0
    
    # Extract content from all document items
    for item in book.get_items():
        if item.get_type() == ebooklib.ITEM_DOCUMENT:
            chapter_num += 1
            
            # Parse HTML content
            soup = BeautifulSoup(item.get_content(), 'html.parser')
            
            # Extract chapter title (try h1, h2, or title tag)
            chapter_title = None
            for tag in ['h1', 'h2', 'h3', 'title']:
                heading = soup.find(tag)
                if heading:
                    chapter_title = clean_text(heading.get_text())
                    break
            
            if not chapter_title:
                chapter_title = f"Chapter {chapter_num}"
            
            # Extract all text content
            text_content = clean_text(soup.get_text())
            
            # Split into paragraphs
            paragraphs = [p.strip() for p in text_content.split('\n') if p.strip()]
            
            # Add each paragraph as a separate row
            for para_num, paragraph in enumerate(paragraphs, 1):
                if paragraph and len(paragraph) > 0:
                    content_data.append({
                        'book_title': book_title,
                        'author': book_author,
                        'chapter_number': chapter_num,
                        'chapter_title': chapter_title,
                        'paragraph_number': para_num,
                        'content': paragraph
                    })
    
    print(f"Extracted {len(content_data)} paragraphs from {chapter_num} chapters")
    return content_data

def convert_epub_to_csv(epub_file, output_file, include_metadata=True, delimiter=','):
    """
    Convert EPUB file to CSV format
    
    Args:
        epub_file (str): Path to input EPUB file
        output_file (str): Path to output CSV file
        include_metadata (bool): Include book metadata in CSV
        delimiter (str): CSV delimiter character
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting EPUB to CSV conversion...")
    print(f"Input: {epub_file}")
    print(f"Output: {output_file}")
    
    try:
        # Check if EPUB file exists
        if not os.path.exists(epub_file):
            print(f"ERROR: EPUB file does not exist: {epub_file}")
            return False
        
        file_size = os.path.getsize(epub_file)
        print(f"EPUB file size: {file_size} bytes")
        
        # Extract content from EPUB
        content_data = extract_epub_content(epub_file)
        
        if not content_data:
            print("ERROR: No content extracted from EPUB file")
            return False
        
        if len(content_data) == 0:
            print("WARNING: EPUB file appears to be empty")
            return False
        
        # Write to CSV file
        print(f"Writing CSV file with {len(content_data)} rows...")
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            if include_metadata:
                fieldnames = ['book_title', 'author', 'chapter_number', 'chapter_title', 'paragraph_number', 'content']
            else:
                fieldnames = ['chapter_number', 'chapter_title', 'paragraph_number', 'content']
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=delimiter, quoting=csv.QUOTE_ALL)
            writer.writeheader()
            
            for row in content_data:
                if not include_metadata:
                    # Remove metadata fields
                    row = {k: v for k, v in row.items() if k in fieldnames}
                writer.writerow(row)
        
        # Verify output file
        if os.path.exists(output_file):
            output_size = os.path.getsize(output_file)
            print(f"CSV file created successfully: {output_size} bytes")
            
            # Verify it's a valid CSV
            try:
                with open(output_file, 'r', encoding='utf-8') as f:
                    csv.reader(f)
                    print(f"Verified CSV file with {len(content_data)} data rows")
                    return True
            except Exception as verify_error:
                print(f"ERROR: Output file is not a valid CSV: {verify_error}")
                return False
        else:
            print("ERROR: CSV file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to convert EPUB to CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert EPUB file to CSV format')
    parser.add_argument('epub_file', help='Path to input EPUB file')
    parser.add_argument('output_file', help='Path to output CSV file')
    parser.add_argument('--no-metadata', action='store_true',
                        help='Exclude book metadata (title, author) from CSV')
    parser.add_argument('--delimiter', default=',',
                        help='CSV delimiter character (default: comma)')
    
    args = parser.parse_args()
    
    print("=== EPUB to CSV Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_epub_to_csv(
        args.epub_file, 
        args.output_file,
        include_metadata=not args.no_metadata,
        delimiter=args.delimiter
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == '__main__':
    main()

