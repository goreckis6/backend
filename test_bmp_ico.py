#!/usr/bin/env python3
"""
Test script for BMP to ICO conversion
"""

import sys
import os
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(CURRENT_DIR, 'scripts')
if SCRIPTS_DIR not in sys.path:
    sys.path.append(SCRIPTS_DIR)

from bmp_to_ico import create_ico_from_bmp

def test_bmp_conversion():
    """Test BMP to ICO conversion with a simple test"""
    print("Testing BMP to ICO conversion...")
    
    # Create a simple test BMP file
    from PIL import Image
    import tempfile
    
    # Create a simple test image
    test_img = Image.new('RGB', (32, 32), (255, 0, 0))  # Red square
    
    with tempfile.NamedTemporaryFile(suffix='.bmp', delete=False) as tmp_bmp:
        test_img.save(tmp_bmp.name, 'BMP')
        bmp_file = tmp_bmp.name
    
    with tempfile.NamedTemporaryFile(suffix='.ico', delete=False) as tmp_ico:
        ico_file = tmp_ico.name
    
    try:
        print(f"Test BMP file: {bmp_file}")
        print(f"Test ICO file: {ico_file}")
        
        # Test the conversion
        result = create_ico_from_bmp(bmp_file, ico_file, [16, 32], True)
        
        if result:
            print("✅ BMP to ICO conversion test PASSED")
            print(f"Output file size: {os.path.getsize(ico_file)} bytes")
            return True
        else:
            print("❌ BMP to ICO conversion test FAILED")
            return False
            
    except Exception as e:
        print(f"❌ BMP to ICO conversion test FAILED with error: {e}")
        return False
    finally:
        # Cleanup
        try:
            os.unlink(bmp_file)
            os.unlink(ico_file)
        except:
            pass

if __name__ == "__main__":
    success = test_bmp_conversion()
    sys.exit(0 if success else 1)



