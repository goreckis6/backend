#!/usr/bin/env python3
"""
SDD (StarOffice Presentation) to HTML converter for web preview.
Uses LibreOffice for conversion. Handles legacy StarOffice/StarImpress format.
"""

import argparse
import os
import sys
import subprocess
import shutil
import traceback

def convert_sdd_to_html_libreoffice(sdd_file, html_file):
    """
    Convert SDD to HTML using LibreOffice.
    
    Args:
        sdd_file (str): Path to input SDD file
        html_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Attempting SDD to HTML conversion with LibreOffice...")
    
    try:
        # Ensure output directory exists
        output_dir = os.path.dirname(html_file)
        os.makedirs(output_dir, exist_ok=True)
        
        # Set environment variables for LibreOffice
        env = os.environ.copy()
        env['SAL_USE_VCLPLUGIN'] = 'svp'
        env['HOME'] = os.path.expanduser('~')
        
        cmd = [
            'libreoffice',
            '--headless',
            '--invisible',
            '--nocrashreport',
            '--nodefault',
            '--nofirststartwizard',
            '--nolockcheck',
            '--nologo',
            '--norestore',
            '-env:UserInstallation=file:///tmp/libreoffice_user_profile',
            '--convert-to', 'html:impress_html_Export',
            '--outdir', output_dir,
            sdd_file
        ]
        
        print(f"Executing LibreOffice command: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120, env=env)
        print("LibreOffice stdout:", result.stdout)
        print("LibreOffice stderr:", result.stderr)
        print("LibreOffice return code:", result.returncode)
        
        # LibreOffice creates output with the input filename + .html extension
        base_name = os.path.splitext(os.path.basename(sdd_file))[0]
        html_output_file = os.path.join(output_dir, f"{base_name}.html")
        
        if not os.path.exists(html_output_file):
            # Try to find any .html file in the output directory
            html_files = [f for f in os.listdir(output_dir) if f.endswith('.html')]
            if html_files:
                html_output_file = os.path.join(output_dir, html_files[0])
            else:
                raise FileNotFoundError(f"No HTML file produced by LibreOffice in {output_dir}")

        # Rename to desired output filename if different
        if html_output_file != html_file:
            shutil.move(html_output_file, html_file)
        
        file_size = os.path.getsize(html_file)
        print(f"HTML file created successfully: {file_size} bytes")
        return True
        
    except FileNotFoundError:
        print("ERROR: LibreOffice not found. Please install LibreOffice.")
        return False
    except subprocess.CalledProcessError as e:
        print(f"ERROR: LibreOffice conversion failed with exit code {e.returncode}")
        print("Stdout:", e.stdout)
        print("Stderr:", e.stderr)
        return False
    except subprocess.TimeoutExpired:
        print("ERROR: LibreOffice conversion timed out (>120s)")
        return False
    except Exception as e:
        print(f"ERROR: LibreOffice conversion error: {e}")
        traceback.print_exc()
        return False

def create_legacy_format_error_html(html_file, format_name="StarOffice Presentation"):
    """Create an informative error page for legacy formats."""
    try:
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{format_name} Preview - Legacy Format</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 50px auto;
            padding: 20px;
            background: #f5f5f5;
        }}
        .info-box {{
            background: white;
            border-left: 4px solid #ff9800;
            padding: 20px;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #ff9800;
            margin-top: 0;
        }}
        .legacy-notice {{
            background: #fff3e0;
            padding: 15px;
            border-radius: 4px;
            margin: 15px 0;
        }}
        .suggestion {{
            background: #e3f2fd;
            padding: 15px;
            border-radius: 4px;
            border-left: 4px solid #2196f3;
        }}
        .download-section {{
            background: #e8f5e9;
            padding: 15px;
            border-radius: 4px;
            margin: 15px 0;
            border-left: 4px solid #4caf50;
        }}
    </style>
</head>
<body>
    <div class="info-box">
        <h1>📁 {format_name} Preview</h1>
        <p>This is a legacy file format that requires special handling.</p>
        
        <div class="legacy-notice">
            <h3>⚠️ Legacy Format Notice</h3>
            <p>StarOffice/StarImpress files (.sdd) are legacy presentation formats from the 1990s-2000s. 
            These files require LibreOffice for conversion, which may not be fully supported in the current environment.</p>
        </div>
        
        <div class="download-section">
            <h3>✅ Recommended Action</h3>
            <p><strong>Download the file</strong> and open it with:</p>
            <ul>
                <li><strong>LibreOffice Impress</strong> (Free, open-source)</li>
                <li><strong>Apache OpenOffice</strong> (Free, open-source)</li>
                <li>Convert to modern formats (.pptx, .odp) for better compatibility</li>
            </ul>
        </div>
        
        <div class="suggestion">
            <h3>💡 About This Format</h3>
            <p><strong>StarOffice Presentation (.sdd)</strong> was used by Sun Microsystems' StarOffice suite 
            (1996-2005) before it evolved into OpenOffice and later LibreOffice.</p>
            <p>For modern usage, we recommend converting these files to:</p>
            <ul>
                <li>OpenDocument Presentation (.odp) - Open standard</li>
                <li>Microsoft PowerPoint (.pptx) - Widely supported</li>
            </ul>
        </div>
    </div>
</body>
</html>
"""
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"Created legacy format info page: {html_file}")
        return True
    except Exception as e:
        print(f"ERROR: Could not create info page: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert SDD to HTML for web preview')
    parser.add_argument('sdd_file', help='Input SDD file path')
    parser.add_argument('html_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== SDD to HTML Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.sdd_file):
        print(f"ERROR: Input SDD file not found: {args.sdd_file}")
        sys.exit(1)
    
    print("Detected format: SDD (StarOffice Presentation - Legacy)")
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.html_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Try LibreOffice conversion
    success = convert_sdd_to_html_libreoffice(args.sdd_file, args.html_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        print("LibreOffice conversion failed. Creating informative legacy format page...")
        
        # Create an informative error page instead of failing completely
        if create_legacy_format_error_html(args.html_file, "StarOffice Presentation (.sdd)"):
            print("Created legacy format info page for user")
            sys.exit(0)  # Exit successfully so the info page is shown
        else:
            print("ERROR: Could not create info page")
            sys.exit(1)

if __name__ == "__main__":
    main()

