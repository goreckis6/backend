#!/usr/bin/env python3
"""
PDF to HTML converter for web preview using PDF.js integration.
Converts PDF files to HTML with embedded PDF.js viewer for cross-browser compatibility.
"""

import argparse
import os
import sys
import base64
import traceback

def convert_pdf_to_html(pdf_file, output_file):
    """
    Convert PDF to HTML with embedded PDF.js viewer.
    
    Args:
        pdf_file (str): Path to input PDF file
        output_file (str): Path to output HTML file
    
    Returns:
        bool: True if conversion successful, False otherwise
    """
    print(f"Converting PDF to HTML viewer...")
    print(f"Input: {pdf_file}")
    print(f"Output: {output_file}")
    
    try:
        # Get file size
        file_size = os.path.getsize(pdf_file)
        print(f"PDF file size: {file_size:,} bytes ({file_size / 1024 / 1024:.2f} MB)")
        
        # Read PDF file and encode to base64
        print("Reading and encoding PDF...")
        with open(pdf_file, 'rb') as f:
            pdf_data = f.read()
        
        pdf_base64 = base64.b64encode(pdf_data).decode('utf-8')
        
        # Create HTML with PDF.js viewer
        html_content = f'''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF Preview</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            background: #f3f4f6;
            overflow: hidden;
        }}
        
        .header-bar {{
            background: linear-gradient(to right, #dc2626, #ec4899);
            color: white;
            padding: 15px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 1000;
        }}
        
        .header-title {{
            font-size: 18px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        
        .header-controls {{
            display: flex;
            gap: 15px;
            align-items: center;
        }}
        
        .page-info {{
            background: rgba(255,255,255,0.2);
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        
        .nav-btn {{
            background: rgba(255,255,255,0.2);
            border: 1px solid rgba(255,255,255,0.3);
            color: white;
            padding: 6px 12px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s;
        }}
        
        .nav-btn:hover {{
            background: rgba(255,255,255,0.3);
            transform: scale(1.05);
        }}
        
        .nav-btn:disabled {{
            opacity: 0.5;
            cursor: not-allowed;
        }}
        
        .zoom-btn {{
            background: rgba(255,255,255,0.2);
            border: 1px solid rgba(255,255,255,0.3);
            color: white;
            padding: 6px 10px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            min-width: 35px;
            transition: all 0.2s;
        }}
        
        .zoom-btn:hover {{
            background: rgba(255,255,255,0.3);
            transform: scale(1.05);
        }}
        
        .btn {{
            padding: 8px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s;
        }}
        
        .btn-print {{
            background: white;
            color: #dc2626;
        }}
        
        .btn-print:hover {{
            background: #fee2e2;
            transform: scale(1.05);
        }}
        
        .btn-close {{
            background: rgba(255,255,255,0.2);
            color: white;
            border: 1px solid rgba(255,255,255,0.3);
        }}
        
        .btn-close:hover {{
            background: rgba(255,255,255,0.3);
            transform: scale(1.05);
        }}
        
        .canvas-container {{
            position: fixed;
            top: 70px;
            left: 0;
            right: 0;
            bottom: 0;
            overflow: auto;
            background: #525252;
            display: flex;
            justify-content: center;
            padding: 20px;
            scroll-behavior: smooth;
        }}
        
        .canvas-container::-webkit-scrollbar {{
            width: 12px;
            height: 12px;
        }}
        
        .canvas-container::-webkit-scrollbar-track {{
            background: #404040;
            border-radius: 6px;
        }}
        
        .canvas-container::-webkit-scrollbar-thumb {{
            background: #666;
            border-radius: 6px;
            border: 2px solid #404040;
        }}
        
        .canvas-container::-webkit-scrollbar-thumb:hover {{
            background: #777;
        }}
        
        #pdf-canvas {{
            background: white;
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
            max-width: 100%;
            height: auto;
        }}
        
        .loading {{
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            text-align: center;
            color: white;
        }}
        
        .spinner {{
            border: 4px solid rgba(255,255,255,0.3);
            border-top: 4px solid white;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }}
        
        @keyframes spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
        
        @media print {{
            .header-bar {{
                display: none;
            }}
            .canvas-container {{
                top: 0;
                background: white;
            }}
        }}
    </style>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
</head>
<body>
    <div class="header-bar">
        <div class="header-title">
            <span>📄</span>
            <span>PDF Viewer</span>
        </div>
        <div class="header-controls">
            <div class="page-info">
                <span>Page</span>
                <span id="page-num">1</span>
                <span>/</span>
                <span id="page-count">-</span>
            </div>
            <button id="prev-btn" class="nav-btn">◀ Previous</button>
            <button id="next-btn" class="nav-btn">Next ▶</button>
            <button id="zoom-out" class="zoom-btn">-</button>
            <span style="color: white; font-size: 14px;" id="zoom-level">150%</span>
            <button id="zoom-in" class="zoom-btn">+</button>
            <button id="zoom-reset" class="zoom-btn" style="min-width: 45px;">100%</button>
            <button id="zoom-fit-width" class="zoom-btn" style="min-width: 60px;">Fit Width</button>
            <button onclick="window.print()" class="btn btn-print">
                🖨️ Print
            </button>
            <button onclick="window.close()" class="btn btn-close">
                ✖️ Close
            </button>
        </div>
    </div>
    
    <div class="canvas-container">
        <canvas id="pdf-canvas"></canvas>
    </div>
    
    <div id="scroll-indicator" style="position: fixed; bottom: 20px; right: 20px; background: rgba(0,0,0,0.7); color: white; padding: 8px 12px; border-radius: 6px; font-size: 12px; z-index: 1001; display: none;">
        Scroll to change pages, Ctrl+Scroll to zoom<br>
        Use +/- buttons or F key for fit-to-width
    </div>
    
    <div id="loading" class="loading">
        <div class="spinner"></div>
        <p>Loading PDF...</p>
    </div>
    
    <script>
        // PDF.js setup
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        
        const pdfData = atob('{pdf_base64}');
        const loadingTask = pdfjsLib.getDocument({{data: pdfData}});
        
        let pdfDoc = null;
        let pageNum = 1;
        let pageRendering = false;
        let pageNumPending = null;
        let scale = 1.5; // Good balance for text quality and performance
        
        const canvas = document.getElementById('pdf-canvas');
        const ctx = canvas.getContext('2d', {{ alpha: false }});
        
        // Set up high-quality rendering
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        
        function renderPage(num) {{
            pageRendering = true;
            
            pdfDoc.getPage(num).then(function(page) {{
                const viewport = page.getViewport({{scale: scale}});
                
                // Get device pixel ratio for crisp rendering
                const devicePixelRatio = window.devicePixelRatio || 1;
                
                // Set canvas size with device pixel ratio
                canvas.height = viewport.height * devicePixelRatio;
                canvas.width = viewport.width * devicePixelRatio;
                canvas.style.height = viewport.height + 'px';
                canvas.style.width = viewport.width + 'px';
                
                // Scale context to match device pixel ratio
                ctx.scale(devicePixelRatio, devicePixelRatio);
                
                // Clear canvas
                ctx.clearRect(0, 0, viewport.width, viewport.height);
                
                const renderContext = {{
                    canvasContext: ctx,
                    viewport: viewport,
                    enableWebGL: false,
                    renderInteractiveForms: false,
                    intent: 'display'
                }};
                
                const renderTask = page.render(renderContext);
                
                renderTask.promise.then(function() {{
                    pageRendering = false;
                    if (pageNumPending !== null) {{
                        renderPage(pageNumPending);
                        pageNumPending = null;
                    }}
                }});
            }});
            
            document.getElementById('page-num').textContent = num;
        }}
        
        function queueRenderPage(num) {{
            if (pageRendering) {{
                pageNumPending = num;
            }} else {{
                renderPage(num);
            }}
        }}
        
        function onPrevPage() {{
            if (pageNum <= 1) return;
            pageNum--;
            queueRenderPage(pageNum);
            updateButtons();
        }}
        
        function onNextPage() {{
            if (pageNum >= pdfDoc.numPages) return;
            pageNum++;
            queueRenderPage(pageNum);
            updateButtons();
        }}
        
        function updateButtons() {{
            document.getElementById('prev-btn').disabled = (pageNum <= 1);
            document.getElementById('next-btn').disabled = (pageNum >= pdfDoc.numPages);
        }}
        
        function zoomIn() {{
            scale = Math.min(scale + 0.25, 5.0); // Smaller increments, higher max zoom
            document.getElementById('zoom-level').textContent = Math.round(scale * 100) + '%';
            queueRenderPage(pageNum);
        }}
        
        function zoomOut() {{
            scale = Math.max(scale - 0.25, 0.25); // Smaller increments, lower min zoom
            document.getElementById('zoom-level').textContent = Math.round(scale * 100) + '%';
            queueRenderPage(pageNum);
        }}
        
        function resetZoom() {{
            scale = 1.0;
            document.getElementById('zoom-level').textContent = Math.round(scale * 100) + '%';
            queueRenderPage(pageNum);
        }}
        
        function fitToWidth() {{
            if (!pdfDoc) return;
            
            pdfDoc.getPage(pageNum).then(function(page) {{
                const container = document.querySelector('.canvas-container');
                const containerWidth = container.clientWidth - 40; // Account for padding
                
                const viewport = page.getViewport({{scale: 1}});
                scale = containerWidth / viewport.width;
                
                // Ensure scale is within reasonable bounds
                scale = Math.max(0.25, Math.min(5.0, scale));
                
                document.getElementById('zoom-level').textContent = Math.round(scale * 100) + '%';
                queueRenderPage(pageNum);
            }});
        }}
        
        document.getElementById('prev-btn').addEventListener('click', onPrevPage);
        document.getElementById('next-btn').addEventListener('click', onNextPage);
        document.getElementById('zoom-in').addEventListener('click', zoomIn);
        document.getElementById('zoom-out').addEventListener('click', zoomOut);
        document.getElementById('zoom-reset').addEventListener('click', resetZoom);
        document.getElementById('zoom-fit-width').addEventListener('click', fitToWidth);
        
        // Keyboard shortcuts
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'ArrowLeft') onPrevPage();
            if (e.key === 'ArrowRight') onNextPage();
            if (e.key === '+' || e.key === '=') zoomIn();
            if (e.key === '-') zoomOut();
            if (e.key === '0') resetZoom();
            if (e.key === 'f' || e.key === 'F') fitToWidth();
        }});
        
        // Mouse wheel handling
        document.addEventListener('wheel', function(e) {{
            if (e.ctrlKey || e.metaKey) {{
                // Ctrl+Scroll = Zoom
                e.preventDefault();
                const zoomFactor = 0.2; // More responsive zoom factor
                if (e.deltaY < 0) {{
                    scale = Math.min(scale + zoomFactor, 5.0); // Match button max zoom
                }} else {{
                    scale = Math.max(scale - zoomFactor, 0.25); // Match button min zoom
                }}
                document.getElementById('zoom-level').textContent = Math.round(scale * 100) + '%';
                queueRenderPage(pageNum);
            }} else {{
                // Regular scroll = Change pages
                e.preventDefault();
                if (e.deltaY > 0) {{
                    // Scroll down = Next page
                    onNextPage();
                }} else {{
                    // Scroll up = Previous page
                    onPrevPage();
                }}
            }}
        }}, {{ passive: false }});
        
        loadingTask.promise.then(function(pdfDoc_) {{
            pdfDoc = pdfDoc_;
            document.getElementById('page-count').textContent = pdfDoc.numPages;
            document.getElementById('loading').style.display = 'none';
            
            renderPage(pageNum);
            updateButtons();
            
            // Show scroll indicator briefly
            const scrollIndicator = document.getElementById('scroll-indicator');
            scrollIndicator.style.display = 'block';
            setTimeout(function() {{
                scrollIndicator.style.display = 'none';
            }}, 3000);
        }}).catch(function(error) {{
            console.error('Error loading PDF:', error);
            document.getElementById('loading').innerHTML = '<p>Error loading PDF: ' + error.message + '</p>';
        }});
    </script>
</body>
</html>'''
        
        # Write HTML file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        output_size = os.path.getsize(output_file)
        print(f"HTML viewer created successfully: {output_size:,} bytes")
        return True
        
    except Exception as e:
        print(f"ERROR: PDF to HTML conversion error: {e}")
        traceback.print_exc()
        return False

def main():
    parser = argparse.ArgumentParser(description='Convert PDF to HTML viewer')
    parser.add_argument('pdf_file', help='Input PDF file path')
    parser.add_argument('output_file', help='Output HTML file path')
    
    args = parser.parse_args()
    
    print("=== PDF to HTML Viewer Converter ===")
    print(f"Python version: {sys.version}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Arguments: {vars(args)}")
    
    # Check if input file exists
    if not os.path.exists(args.pdf_file):
        print(f"ERROR: Input PDF file not found: {args.pdf_file}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(args.output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Convert PDF
    success = convert_pdf_to_html(args.pdf_file, args.output_file)
    
    if success:
        print("=== CONVERSION SUCCESSFUL ===")
        sys.exit(0)
    else:
        print("=== CONVERSION FAILED ===")
        sys.exit(1)

if __name__ == "__main__":
    main()



