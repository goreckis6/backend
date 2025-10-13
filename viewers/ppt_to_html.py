#!/usr/bin/env python3
"""
PPT/PPTX (PowerPoint Presentation) to HTML converter for web preview.
Uses LibreOffice for conversion. Supports both legacy PPT and modern PPTX formats.
"""

import argparse
import os
import sys
import subprocess
import shutil
import traceback

def convert_ppt_to_html_libreoffice(ppt_file, html_file):
    """
    Convert PPT/PPTX to HTML using LibreOffice.
    
    Args:
        ppt_file (str): Path to input PPT/PPTX file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting PPT/PPTX to HTML conversion with LibreOffice...")
    
    try:
        # Ensure output directory exists
        output_dir = os.path.dirname(html_file)
        os.makedirs(output_dir, exist_ok=True)
        
        # Try multiple LibreOffice command variations
        # Method 1: Direct conversion with filter
        cmd = [
            'libreoffice',
            '--headless',
            '--invisible',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
            '--convert-to', 'html:HTML:EmbedImages',
            '--outdir', output_dir,
            ppt_file
        ]
        
        print(f"Executing LibreOffice command: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        print("LibreOffice stdout:", result.stdout)
        print("LibreOffice stderr:", result.stderr)
        print("LibreOffice return code:", result.returncode)
        
        # LibreOffice creates output with the input filename + .html extension
        base_name = os.path.splitext(os.path.basename(ppt_file))[0]
        html_output_file = os.path.join(output_dir, f"{base_name}.html")
        
        if not os.path.exists(html_output_file):
            # Try to find any .html file in the output directory
            html_files = [f for f in os.listdir(output_dir) if f.endswith('.html')]
            if html_files:
                html_output_file = os.path.join(output_dir, html_files[0])
            else:
                print(f"ERROR: No HTML file found in {output_dir}")
                print(f"Directory contents: {os.listdir(output_dir)}")
                raise FileNotFoundError(f"No HTML file produced by LibreOffice in {output_dir}")

        # Rename to desired output filename if different
        if html_output_file != html_file:
            shutil.move(html_output_file, html_file)
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        
        # Enhance HTML with header bar for print/close
        enhance_html_with_header(html_file)
        
        return True
        
    except FileNotFoundError as e:
        print(f"ERROR: LibreOffice not found or file error: {e}")
        return False
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out (>120s)")
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion error: {e}")
        traceback.print_exc()
        return False

def enhance_html_with_header(html_file):
    """Add header bar with print and close buttons to HTML."""
    try:
        with open(html_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        header_html = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PowerPoint Preview</title>
    <style>
        body { margin: 0; padding: 0; font-family: Arial, sans-serif; }
        #header-bar {
            position: sticky;
            top: 0;
            background: linear-gradient(135deg, #0078d4 0%, #00a4ef 100%);
            color: white;
            padding: 12px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            z-index: 1000;
        }
        #header-bar h1 {
            margin: 0;
            font-size: 18px;
            font-weight: 600;
        }
        .header-actions {
            display: flex;
            gap: 10px;
        }
        .header-btn {
            background: rgba(255,255,255,0.2);
            border: 1px solid rgba(255,255,255,0.3);
            color: white;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.2s;
        }
        .header-btn:hover {
            background: rgba(255,255,255,0.3);
            transform: translateY(-1px);
        }
        #content-area {
            padding: 20px;
            max-width: 1200px;
            margin: 0 auto;
        }
        @media print {
            #header-bar { display: none; }
            #content-area { padding: 0; max-width: none; }
        }
    </style>
</head>
<body>
    <div id="header-bar">
        <h1>📊 PowerPoint Preview</h1>
        <div class="header-actions">
            <button class="header-btn" onclick="window.print()">🖨️ Print</button>
            <button class="header-btn" onclick="window.close()">✖️ Close</button>
        </div>
    </div>
    <div id="content-area">
"""
        
        footer_html = """
    </div>
</body>
</html>
"""
        
        # Extract body content if it exists
        if '<body' in content.lower():
            body_start = content.lower().find('<body')
            body_end = content.lower().find('</body>')
            if body_start != -1 and body_end != -1:
                body_content = content[content.find('>', body_start) + 1:body_end]
                enhanced_content = header_html + body_content + footer_html
            else:
                enhanced_content = header_html + content + footer_html
        else:
            enhanced_content = header_html + content + footer_html
        
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(enhanced_content)
        
        print("HTML enhanced with header bar")
    except Exception as e:
        print(f"Warning: Could not enhance HTML: {e}")

def main():
    parser = argparse.ArgumentParser(description='Convert PPT/PPTX to HTML for web preview')
    parser.add_argument('ppt_file', help='Input PPT/PPTX file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== PPT/PPTX to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.ppt_file):
        print(f"ERROR: Input PPT/PPTX file not found: {args.ppt_file}")
        sys.exit(1)
    
    # Determine file type
    file_ext = os.path.splitext(args.ppt_file)[1].lower()
    if file_ext == '.ppt':
        print("Detected format: PPT (PowerPoint 97-2003)")
    elif file_ext == '.pptx':
        print("Detected format: PPTX (PowerPoint 2007+)")
    else:
        print(f"WARNING: Unknown extension '{file_ext}', will try to convert anyway")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Try LibreOffice conversion
    success = convert_ppt_to_html_libreoffice(args.ppt_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        print("ERROR: LibreOffice conversion failed. Please ensure LibreOffice is installed.")
        sys.exit(1)

if __name__ == "__main__":
    main()

