#!/usr/bin/env python3
"""
CSV to PDF Converter
Converts CSV files to PDF format using pandas and reportlab.
"""

import pandas as pd
import argparse
import os
import sys
from datetime import datetime
import traceback

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
except ImportError as e:
    print(f"ERROR: Required ReportLab library not available: {e}")
    print("Please install reportlab: pip install reportlab")
    sys.exit(1)

def create_pdf_from_csv(csv_file, output_file, title="CSV Data", author="Unknown"):
    """
    Convert CSV file to PDF format.
    
    Args:
        csv_file (str): Path to input CSV file
        output_file (str): Path to output PDF file
        title (str): Document title
        author (str): Document author
    """
    print(f"Starting CSV to PDF conversion...")
    print(f"Input: {csv_file}")
    print(f"Output: {output_file}")
    print(f"Title: {title}")
    print(f"Author: {author}")
    print("Processing all data (no row limits)")
    
    try:
        # Read CSV file
        print("Reading CSV file...")
        df = pd.read_csv(csv_file)
        print(f"CSV loaded successfully: {len(df)} rows, {len(df.columns)} columns")
        
        # Process all rows - no limits
        print(f"Processing all {len(df)} rows (including any repeated data)")
        
        # Create PDF document
        print("Creating PDF document...")
        doc = SimpleDocTemplate(
            output_file,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=18
        )
        
        # Get styles
        styles = getSampleStyleSheet()
        
        # Create custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.darkblue
        )
        
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=12,
            alignment=TA_LEFT,
            textColor=colors.darkblue
        )
        
        # Build content
        story = []
        
        # Add title
        print("Adding document title...")
        story.append(Paragraph(title, title_style))
        story.append(Spacer(1, 12))
        
        # Add author info
        story.append(Paragraph(f"<b>Author:</b> {author}", header_style))
        story.append(Paragraph(f"<b>Created:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", header_style))
        story.append(Spacer(1, 20))
        
        # Prepare table data
        print("Preparing table data...")
        table_data = []
        
        # Add header row
        header_row = [str(col) for col in df.columns]
        table_data.append(header_row)
        
        # Add data rows
        print(f"Adding {len(df)} data rows...")
        for idx, row in df.iterrows():
            if idx % 1000 == 0:
                print(f"Processing row {idx + 1}/{len(df)}")
            
            # Handle NaN values and convert to string
            data_row = [str(value) if pd.notna(value) else "" for value in row]
            table_data.append(data_row)
        
        # Create table
        print("Creating PDF table...")
        table = Table(table_data)
        
        # Style the table
        table_style = TableStyle([
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Data rows styling
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ])
        
        # Apply alternating row colors
        for i in range(1, len(table_data)):
            if i % 2 == 0:
                table_style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
        
        table.setStyle(table_style)
        story.append(table)
        
        # Add summary
        print("Adding document summary...")
        story.append(Spacer(1, 20))
        summary_style = ParagraphStyle(
            'Summary',
            parent=styles['Normal'],
            fontSize=10,
            alignment=TA_LEFT,
            textColor=colors.darkblue
        )
        story.append(Paragraph(f"<b>Summary:</b> {len(df)} rows, {len(df.columns)} columns", summary_style))
        
        # Build PDF
        print(f"Building PDF document...")
        doc.build(story)
        
        # Verify file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"PDF file created successfully: {file_size} bytes")
            return True
        else:
            print("ERROR: PDF file was not created")
            return False
            
    except Exception as e:
        print(f"ERROR: Failed to create PDF from CSV: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to PDF format')
    parser.add_argument('csv_file', help='Input CSV file path')
    parser.add_argument('output_file', help='Output PDF file path')
    parser.add_argument('--title', default='CSV Data', help='Document title')
    parser.add_argument('--author', default='Unknown', help='Document author')
    
    args = parser.parse_args()
    
    print("=== CSV to PDF Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.csv_file):
        print(f"ERROR: Input CSV file not found: {args.csv_file}")
        sys.exit(1)
    
    # Check pandas availability
    try:
        print(f"Pandas version: {pd.__version__}")
    except Exception as e:
        print(f"ERROR: Pandas not available: {e}")
        sys.exit(1)
    
    # Check reportlab availability
    try:
        from reportlab import __version__ as reportlab_version
        print(f"ReportLab version: {reportlab_version}")
    except Exception as e:
        print(f"ERROR: ReportLab not available: {e}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert CSV to PDF
    success = create_pdf_from_csv(
        args.csv_file,
        args.output_file,
        args.title,
        args.author
    )
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()


