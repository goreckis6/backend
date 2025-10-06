#!/usr/bin/env python3
"""
Simple test for BMP to ICO conversion
"""

import sys
import os
from PIL import Image

def test_simple_ico():
    """Create a simple ICO file for testing"""
    try:
        print("Creating simple test ICO...")
        
        # Create a simple test image
        img = Image.new('RGB', (32, 32), (255, 0, 0))  # Red square
        
        # Save as ICO
        output_path = '/tmp/test_output.ico'
        img.save(output_path, format='ICO')
        
        if os.path.exists(output_path):
            size = os.path.getsize(output_path)
            print(f"✅ Test ICO created successfully: {size} bytes")
            return True
        else:
            print("❌ Test ICO was not created")
            return False
            
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

if __name__ == "__main__":
    success = test_simple_ico()
    sys.exit(0 if success else 1)

