#!/usr/bin/env python3
"""
DOC to CSV Converter
Extracts text content from DOC files and converts to CSV format
"""

import os
import sys
import argparse
import csv
from docx import Document
import traceback

def extract_text_from_doc(doc_file):
    """
    Extract text content from DOC file using python-docx
    
    Args:
        doc_file (str): Path to the DOC file
        
    Returns:
        list: List of strings containing paragraphs
    """
    print(f"Reading DOC file: {doc_file}")
    
    # Check if file exists
    if not os.path.exists(doc_file):
        print(f"ERROR: Input file does not exist: {doc_file}")
        return None
    
    # Check file size
    file_size = os.path.getsize(doc_file)
    print(f"DOC file size: {file_size} bytes")
    
    if file_size == 0:
        print("ERROR: Input file is empty")
        return None
    
    try:
        # Try to open with python-docx (works for both DOC and DOCX)
        doc = Document(doc_file)
        
        # Extract all paragraphs
        paragraphs = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:  # Only add non-empty paragraphs
                paragraphs.append(text)
        
        print(f"Extracted {len(paragraphs)} paragraphs")
        return paragraphs
        
    except Exception as e:
        print(f"ERROR: Failed to read DOC file: {e}")
        print("Note: python-docx can only read DOCX files. For DOC files, try converting to DOCX first.")
        traceback.print_exc()
        return None

def paragraphs_to_csv_data(paragraphs):
    """
    Convert list of paragraphs into CSV-friendly data structure
    
    Args:
        paragraphs (list): List of text strings
        
    Returns:
        list: List of dictionaries suitable for CSV output
    """
    if not paragraphs:
        return []
    
    # Create CSV data with paragraph number and content
    csv_data = []
    for idx, para in enumerate(paragraphs, 1):
        csv_data.append({
            'paragraph_number': idx,
            'content': para
        })
    
    return csv_data

def convert_doc_to_csv(doc_file, output_file, include_metadata=True):
    """
    Convert DOC file to CSV format
    
    Args:
        doc_file (str): Path to input DOC file
        output_file (str): Path to output CSV file
        include_metadata (bool): Include paragraph numbers in CSV
        
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Starting DOC to CSV conversion...")
    print(f"Input: {doc_file}")
    print(f"Output: {output_file}")
    
    try:
        # Extract text from DOC file
        paragraphs = extract_text_from_doc(doc_file)
        
        if not paragraphs:
            print("ERROR: No content extracted from DOC file")
            return False
        
        if len(paragraphs) == 0:
            print("WARNING: DOC file appears to be empty")
            return False
        
        # Convert paragraphs to CSV data
        csv_data = paragraphs_to_csv_data(paragraphs)
        
        # Write to CSV file
        print(f"Writing CSV file with {len(csv_data)} rows...")
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            if include_metadata:
                fieldnames = ['paragraph_number', 'content']
            else:
                fieldnames = ['content']
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
            writer.writeheader()
            
            for row in csv_data:
                if not include_metadata:
                    # Remove paragraph_number field
                    row = {k: v for k, v in row.items() if k in fieldnames}
                writer.writerow(row)
        
        # Verify output file
        if os.path.exists(output_file):
            output_size = os.path.getsize(output_file)
            print(f"CSV file created successfully: {output_size} bytes")
            
            # Verify it's a valid CSV
            try:
                with open(output_file, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    row_count = sum(1 for row in reader) - 1  # Subtract header
                    print(f"Verified CSV file with {row_count} data rows")
                    return True
            except Exception as verify_error:
                print(f"ERROR: Output file is not a valid CSV: {verify_error}")
                return False
        else:
            print("ERROR: CSV file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to convert DOC to CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert DOC file to CSV format')
    parser.add_argument('doc_file', help='Path to input DOC file')
    parser.add_argument('output_file', help='Path to output CSV file')
    parser.add_argument('--no-metadata', action='store_true',
                        help='Exclude paragraph numbers from CSV')
    
    args = parser.parse_args()
    
    print("=== DOC to CSV Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    success = convert_doc_to_csv(
        args.doc_file, 
        args.output_file,
        include_metadata=not args.no_metadata
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == '__main__':
    main()
