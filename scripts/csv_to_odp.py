#!/usr/bin/env python3
"""
CSV to ODP Converter using Python
Optimized for large files with pagination and streaming
"""

import sys
import os
import pandas as pd
import argparse
from pathlib import Path
from odf.opendocument import OpenDocumentPresentation
from odf import draw, text, style
from odf.draw import Page, Frame, TextBox
from odf.text import P, H
from odf.style import Style, PageLayout, PageLayoutProperties, MasterPage, TextProperties, ParagraphProperties
import tempfile

def escape_html(text):
    """Escape HTML special characters"""
    if not isinstance(text, str):
        text = str(text)
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#39;'))

def create_odp_from_csv(csv_file, output_file, title="CSV Data", author="Unknown", max_rows_per_slide=50):
    """Convert CSV to ODP with pagination for large files"""
    try:
        print(f"Reading CSV file: {csv_file}")
        
        # Read CSV in chunks to handle large files
        chunk_size = 1000
        df_chunks = []
        total_rows = 0
        
        for chunk in pd.read_csv(csv_file, chunksize=chunk_size):
            df_chunks.append(chunk)
            total_rows += len(chunk)
            print(f"Processed {total_rows} rows so far...")
        
        # Combine chunks
        df = pd.concat(df_chunks, ignore_index=True)
        print(f"Total rows: {len(df)}, Columns: {len(df.columns)}")
        
        # Create ODP document
        doc = OpenDocumentPresentation()
        
        # Set document properties
        doc.meta.addElement(text.Title(text=title))
        doc.meta.addElement(text.Creator(text=author))
        
        # Create styles
        create_odp_styles(doc)
        
        # Get columns
        columns = df.columns.tolist()
        
        # Calculate number of slides needed
        rows_per_slide = min(max_rows_per_slide, len(df))
        num_slides = (len(df) + rows_per_slide - 1) // rows_per_slide
        
        print(f"Creating {num_slides} slides with {rows_per_slide} rows each")
        
        # Create slides
        for slide_num in range(num_slides):
            start_row = slide_num * rows_per_slide
            end_row = min(start_row + rows_per_slide, len(df))
            slide_data = df.iloc[start_row:end_row]
            
            print(f"Creating slide {slide_num + 1}/{num_slides} (rows {start_row}-{end_row-1})")
            
            # Create slide
            slide = create_odp_slide(doc, slide_num, slide_data, columns, title, slide_num + 1, num_slides)
            doc.presentation.addElement(slide)
        
        # Save ODP file
        doc.save(output_file)
        print(f"Successfully created ODP: {output_file}")
        
        return True, f"Successfully converted to ODP: {len(df)} rows, {len(columns)} columns, {num_slides} slides"
        
    except Exception as e:
        print(f"Error creating ODP: {str(e)}")
        return False, f"Error converting to ODP: {str(e)}"

def create_odp_styles(doc):
    """Create styles for the ODP document"""
    
    # Page layout
    pagelayout = PageLayout(name="PL1")
    pagelayout.addElement(PageLayoutProperties(
        pagewidth="28cm",
        pageheight="21cm",
        margin="2cm",
        marginleft="2cm",
        marginright="2cm",
        margintop="2cm",
        marginbottom="2cm",
        printorientation="landscape"
    ))
    doc.automaticstyles.addElement(pagelayout)
    
    # Master page
    masterpage = MasterPage(name="Standard", pagelayoutname="PL1")
    doc.masterstyles.addElement(masterpage)
    
    # Title style
    title_style = Style(name="TitleStyle", family="paragraph")
    title_style.addElement(ParagraphProperties(textalign="center"))
    title_style.addElement(TextProperties(fontsize="24pt", fontweight="bold"))
    doc.styles.addElement(title_style)
    
    # Header style
    header_style = Style(name="HeaderStyle", family="paragraph")
    header_style.addElement(ParagraphProperties(textalign="center"))
    header_style.addElement(TextProperties(fontsize="18pt", fontweight="bold", color="#2E86AB"))
    doc.styles.addElement(header_style)
    
    # Table header style
    table_header_style = Style(name="TableHeaderStyle", family="paragraph")
    table_header_style.addElement(ParagraphProperties(textalign="center"))
    table_header_style.addElement(TextProperties(fontsize="12pt", fontweight="bold", color="#FFFFFF"))
    doc.styles.addElement(table_header_style)
    
    # Table cell style
    table_cell_style = Style(name="TableCellStyle", family="paragraph")
    table_cell_style.addElement(ParagraphProperties(textalign="left"))
    table_cell_style.addElement(TextProperties(fontsize="10pt"))
    doc.styles.addElement(table_cell_style)

def create_odp_slide(doc, slide_num, slide_data, columns, title, slide_number, total_slides):
    """Create a single ODP slide with table data"""
    
    # Create slide
    slide = Page(name=f"Slide{slide_num + 1}", masterpagename="Standard")
    
    # Add title
    title_frame = Frame(
        width="26cm",
        height="2cm",
        x="1cm",
        y="0.5cm"
    )
    title_text = TextBox()
    title_text.addElement(P(stylename="TitleStyle", text=f"{title} - Slide {slide_number} of {total_slides}"))
    title_frame.addElement(title_text)
    slide.addElement(title_frame)
    
    # Add table
    table_frame = Frame(
        width="26cm",
        height="16cm",
        x="1cm",
        y="3cm"
    )
    
    # Create table content
    table_text = TextBox()
    
    # Add table header
    header_text = " | ".join([escape_html(str(col)) for col in columns])
    table_text.addElement(P(stylename="HeaderStyle", text=header_text))
    table_text.addElement(P(text="-" * len(header_text)))
    
    # Add table rows
    for _, row in slide_data.iterrows():
        row_text = " | ".join([escape_html(str(cell)) for cell in row.values])
        table_text.addElement(P(stylename="TableCellStyle", text=row_text))
    
    table_frame.addElement(table_text)
    slide.addElement(table_frame)
    
    return slide

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to ODP')
    parser.add_argument('csv_file', help='Input CSV file')
    parser.add_argument('output_file', help='Output ODP file')
    parser.add_argument('--title', default='CSV Data', help='Document title')
    parser.add_argument('--author', default='Unknown', help='Document author')
    parser.add_argument('--max-rows-per-slide', type=int, default=50, help='Maximum rows per slide')
    
    args = parser.parse_args()
    
    print(f"Converting {args.csv_file} to {args.output_file}")
    print(f"Title: {args.title}, Author: {args.author}")
    print(f"Max rows per slide: {args.max_rows_per_slide}")
    
    success, message = create_odp_from_csv(
        args.csv_file,
        args.output_file,
        args.title,
        args.author,
        args.max_rows_per_slide
    )
    
    if success:
        print(f"SUCCESS: {message}")
        sys.exit(0)
    else:
        print(f"ERROR: {message}")
        sys.exit(1)

if __name__ == "__main__":
    main()