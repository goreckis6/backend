#!/usr/bin/env python3
"""
CSV to ODP Converter
Converts CSV data to OpenDocument Presentation (ODP) format
"""

import argparse
import logging
import pandas as pd
import tempfile
import os
import sys
from pathlib import Path
import traceback

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

try:
    from odf.opendocument import OpenDocumentPresentation
    from odf.style import Style, TextProperties, ParagraphProperties
    from odf.table import Table, TableColumn, TableRow, TableCell
    from odf.text import P, Span
    from odf.draw import Page, Frame, TextBox
except ImportError as e:
    logger.error(f"Required ODF library not found: {e}")
    logger.error("Please install odfpy: pip install odfpy")
    sys.exit(1)

def create_odp_from_csv(csv_path, output_path, title="CSV Data", author="CSV Converter", 
                       slide_layout="table", include_headers=True, chunk_size=1000):
    """
    Convert CSV data to ODP presentation
    
    Args:
        csv_path: Path to input CSV file
        output_path: Path to output ODP file
        title: Presentation title
        author: Presentation author
        slide_layout: Layout type ('table', 'chart', 'mixed')
        include_headers: Whether to include headers
        chunk_size: Number of rows per slide for large files
    """
    
    logger.info(f"Starting CSV to ODP conversion: {csv_path} -> {output_path}")
    
    try:
        # Read CSV file
        logger.info("Reading CSV file...")
        df = pd.read_csv(csv_path)
        
        logger.info(f"CSV file contains {len(df)} rows and {len(df.columns)} columns")
        
        # Create ODP document
        doc = OpenDocumentPresentation()
        
        # Set document metadata using DOM manipulation
        logger.info("Setting document metadata")
        try:
            # Find and update existing meta elements
            meta_elements = doc.meta.getElementsByTagName("title")
            if meta_elements:
                meta_elements[0].addText(title)
            else:
                logger.warning("Could not set document title")
            
            meta_elements = doc.meta.getElementsByTagName("initial-creator")
            if meta_elements:
                meta_elements[0].addText(author)
            else:
                logger.warning("Could not set document author")
        except Exception as e:
            logger.warning(f"Could not set document metadata: {e}")
        
        # Create styles
        title_style = Style(name="TitleStyle", family="paragraph")
        title_props = TextProperties(fontsize="24pt", fontweight="bold")
        title_style.addElement(title_props)
        doc.styles.addElement(title_style)
        
        header_style = Style(name="HeaderStyle", family="table-cell")
        header_cell_props = TextProperties(fontweight="bold", backgroundcolor="#4472C4", color="#FFFFFF")
        header_style.addElement(header_cell_props)
        doc.styles.addElement(header_style)
        
        # Get column names
        columns = df.columns.tolist()
        
        if slide_layout == "table":
            # Create table-based slides
            create_table_slides(doc, df, columns, include_headers, chunk_size)
        elif slide_layout == "chart":
            # Create chart-based slides (simplified as tables for now)
            create_chart_slides(doc, df, columns, include_headers, chunk_size)
        elif slide_layout == "mixed":
            # Create mixed layout slides
            create_mixed_slides(doc, df, columns, include_headers, chunk_size)
        
        # Save document
        logger.info(f"Saving ODP file: {output_path}")
        doc.save(output_path)
        
        logger.info("CSV to ODP conversion completed successfully")
        
    except Exception as e:
        logger.error(f"Error converting CSV to ODP: {e}")
        logger.error(traceback.format_exc())
        raise

def create_table_slides(doc, df, columns, include_headers, chunk_size):
    """Create slides with table layout"""
    
    # Process data in chunks
    total_rows = len(df)
    slide_num = 1
    
    for start_idx in range(0, total_rows, chunk_size):
        end_idx = min(start_idx + chunk_size, total_rows)
        chunk_df = df.iloc[start_idx:end_idx]
        
        # Create slide with required master page name
        slide = Page(name=f"Slide{slide_num}", masterpagename="Standard")
        
        # Add title frame
        title_frame = Frame(
            width="25cm", height="2cm",
            x="1cm", y="1cm"
        )
        title_textbox = TextBox()
        title_p = P(stylename="TitleStyle")
        title_p.addText(f"Data Overview - Part {slide_num}")
        title_textbox.addElement(title_p)
        title_frame.addElement(title_textbox)
        slide.addElement(title_frame)
        
        # Create table
        table = Table()
        
        # Add columns
        for _ in columns:
            table.addElement(TableColumn())
        
        # Add header row if requested
        if include_headers:
            header_row = TableRow()
            for col in columns:
                cell = TableCell(stylename="HeaderStyle")
                cell_p = P()
                cell_p.addText(str(col))
                cell.addElement(cell_p)
                header_row.addElement(cell)
            table.addElement(header_row)
        
        # Add data rows
        for _, row in chunk_df.iterrows():
            data_row = TableRow()
            for col in columns:
                cell = TableCell()
                cell_p = P()
                cell_p.addText(str(row[col]) if pd.notna(row[col]) else "")
                cell.addElement(cell_p)
                data_row.addElement(cell)
            table.addElement(data_row)
        
        # Add table frame
        table_frame = Frame(
            width="23cm", height="15cm",
            x="1cm", y="3.5cm"
        )
        table_frame.addElement(table)
        slide.addElement(table_frame)
        
        # Add slide to presentation - use presentation namespace
        try:
            # Get or create presentation element
            presentation_elements = doc.body.getElementsByTagName("presentation")
            if presentation_elements:
                presentation = presentation_elements[0]
            else:
                # Create presentation element
                from odf.element import Element
                from odf.namespaces import PRESENTATIONNS
                presentation = Element(qname=(PRESENTATIONNS, "presentation"))
                doc.body.addElement(presentation)
            
            presentation.addElement(slide)
        except Exception as e:
            logger.warning(f"Could not add slide to presentation: {e}")
            # Fallback: try direct body addition
            # Add slide to presentation - use presentation namespace
    try:
        # Get or create presentation element
        presentation_elements = doc.body.getElementsByTagName("presentation")
        if presentation_elements:
            presentation = presentation_elements[0]
        else:
            # Create presentation element
            from odf.element import Element
            from odf.namespaces import PRESENTATIONNS
            presentation = Element(qname=(PRESENTATIONNS, "presentation"))
            doc.body.addElement(presentation)
        
        presentation.addElement(slide)
    except Exception as e:
        logger.warning(f"Could not add slide to presentation: {e}")
        # Fallback: try direct body addition
        doc.body.addElement(slide)
        slide_num += 1
        
        logger.info(f"Created slide {slide_num - 1} with {len(chunk_df)} rows")

def create_chart_slides(doc, df, columns, include_headers, chunk_size):
    """Create slides with chart layout (simplified as summary tables)"""
    
    # Create summary slide
    slide = Page(name="Summary", masterpagename="Standard")
    
    # Add title
    title_frame = Frame(
        width="25cm", height="2cm",
        x="1cm", y="1cm"
    )
    title_textbox = TextBox()
    title_p = P(stylename="TitleStyle")
    title_p.addText("Data Summary")
    title_textbox.addElement(title_p)
    title_frame.addElement(title_textbox)
    slide.addElement(title_frame)
    
    # Create summary table
    table = Table()
    
    # Add summary columns
    summary_cols = ["Metric", "Value"]
    for _ in summary_cols:
        table.addElement(TableColumn())
    
    # Add header row
    header_row = TableRow()
    for col in summary_cols:
        cell = TableCell(stylename="HeaderStyle")
        cell_p = P()
        cell_p.addText(col)
        cell.addElement(cell_p)
        header_row.addElement(cell)
    table.addElement(header_row)
    
    # Add summary data
    summary_data = [
        ("Total Rows", len(df)),
        ("Total Columns", len(columns)),
        ("Columns", ", ".join(columns[:5]) + ("..." if len(columns) > 5 else ""))
    ]
    
    for metric, value in summary_data:
        data_row = TableRow()
        for item in [metric, str(value)]:
            cell = TableCell()
            cell_p = P()
            cell_p.addText(str(item))
            cell.addElement(cell_p)
            data_row.addElement(cell)
        table.addElement(data_row)
    
    # Add table frame
    table_frame = Frame(
        width="23cm", height="10cm",
        x="1cm", y="3.5cm"
    )
    table_frame.addElement(table)
    slide.addElement(table_frame)
    
    # Add slide to presentation - use presentation namespace
    try:
        # Get or create presentation element
        presentation_elements = doc.body.getElementsByTagName("presentation")
        if presentation_elements:
            presentation = presentation_elements[0]
        else:
            # Create presentation element
            from odf.element import Element
            from odf.namespaces import PRESENTATIONNS
            presentation = Element(qname=(PRESENTATIONNS, "presentation"))
            doc.body.addElement(presentation)
        
        presentation.addElement(slide)
    except Exception as e:
        logger.warning(f"Could not add slide to presentation: {e}")
        # Fallback: try direct body addition
        doc.body.addElement(slide)
    
    # Create data slides with table layout
    create_table_slides(doc, df, columns, include_headers, chunk_size)

def create_mixed_slides(doc, df, columns, include_headers, chunk_size):
    """Create slides with mixed layout"""
    
    # Create overview slide
    slide = Page(name="Overview", masterpagename="Standard")
    
    # Add title
    title_frame = Frame(
        width="25cm", height="2cm",
        x="1cm", y="1cm"
    )
    title_textbox = TextBox()
    title_p = P(stylename="TitleStyle")
    title_p.addText("Data Overview")
    title_textbox.addElement(title_p)
    title_frame.addElement(title_textbox)
    slide.addElement(title_frame)
    
    # Add description
    desc_frame = Frame(
        width="23cm", height="3cm",
        x="1cm", y="3.5cm"
    )
    desc_textbox = TextBox()
    desc_p = P()
    desc_p.addText(f"This presentation contains {len(df)} rows of data with {len(columns)} columns.")
    desc_textbox.addElement(desc_p)
    desc_frame.addElement(desc_textbox)
    slide.addElement(desc_frame)
    
    # Add slide to presentation - use presentation namespace
    try:
        # Get or create presentation element
        presentation_elements = doc.body.getElementsByTagName("presentation")
        if presentation_elements:
            presentation = presentation_elements[0]
        else:
            # Create presentation element
            from odf.element import Element
            from odf.namespaces import PRESENTATIONNS
            presentation = Element(qname=(PRESENTATIONNS, "presentation"))
            doc.body.addElement(presentation)
        
        presentation.addElement(slide)
    except Exception as e:
        logger.warning(f"Could not add slide to presentation: {e}")
        # Fallback: try direct body addition
        doc.body.addElement(slide)
    
    # Create data slides with table layout
    create_table_slides(doc, df, columns, include_headers, chunk_size)

def main():
    parser = argparse.ArgumentParser(description='Convert CSV to ODP presentation')
    parser.add_argument('csv_path', help='Path to input CSV file')
    parser.add_argument('output_path', help='Path to output ODP file')
    parser.add_argument('--title', default='CSV Data', help='Presentation title')
    parser.add_argument('--author', default='CSV Converter', help='Presentation author')
    parser.add_argument('--slide-layout', choices=['table', 'chart', 'mixed'], 
                       default='table', help='Slide layout type')
    parser.add_argument('--no-headers', action='store_true', help='Exclude headers from table')
    parser.add_argument('--chunk-size', type=int, default=1000, 
                       help='Number of rows per slide for large files')
    
    args = parser.parse_args()
    
    try:
        create_odp_from_csv(
            csv_path=args.csv_path,
            output_path=args.output_path,
            title=args.title,
            author=args.author,
            slide_layout=args.slide_layout,
            include_headers=not args.no_headers,
            chunk_size=args.chunk_size
        )
        
        print(f"Successfully converted {args.csv_path} to {args.output_path}")
        
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()