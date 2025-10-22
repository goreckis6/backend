#!/usr/bin/env python3
"""
Test the hybrid EPUB converter
"""

import os
import sys
import pandas as pd
import tempfile

# Add scripts directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from csv_to_epub_hybrid import create_epub_from_csv_hybrid

def create_test_csv():
    """Create a test CSV file"""
    print("Creating test CSV file...")
    
    # Create sample data
    data = {
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown', 'Charlie Wilson'],
        'Age': [25, 30, 35, 28, 42],
        'City': ['New York', 'Los Angeles', 'Chicago', 'Houston', 'Phoenix'],
        'Salary': [50000, 60000, 70000, 55000, 80000],
        'Department': ['IT', 'HR', 'Finance', 'Marketing', 'Sales']
    }
    
    df = pd.DataFrame(data)
    
    # Create test CSV
    test_csv = 'test_data_hybrid.csv'
    df.to_csv(test_csv, index=False)
    print(f"Test CSV created: {test_csv} ({len(df)} rows)")
    
    return test_csv

def test_hybrid_epub():
    """Test hybrid EPUB generation"""
    print("=" * 50)
    print("TESTING HYBRID EPUB GENERATION")
    print("=" * 50)
    
    # Create test CSV
    csv_file = create_test_csv()
    
    try:
        # Test hybrid EPUB generation
        epub_file = 'test_hybrid_output.epub'
        success = create_epub_from_csv_hybrid(
            csv_file,
            epub_file,
            'Test Hybrid CSV Data',
            'Test Author'
        )
        
        if success:
            print(f"✅ Hybrid EPUB generation successful: {epub_file}")
            
            # Check file size
            if os.path.exists(epub_file):
                file_size = os.path.getsize(epub_file)
                print(f"File size: {file_size} bytes ({file_size / 1024:.1f} KB)")
                
                # Test ZIP structure
                import zipfile
                try:
                    with zipfile.ZipFile(epub_file, 'r') as zip_file:
                        file_list = zip_file.namelist()
                        print(f"ZIP contains {len(file_list)} files:")
                        for file in file_list:
                            print(f"  - {file}")
                        
                        # Check for required EPUB files
                        required_files = ['mimetype', 'META-INF/container.xml', 'OEBPS/content.opf', 'OEBPS/content.xhtml']
                        missing_files = [f for f in required_files if f not in file_list]
                        if missing_files:
                            print(f"❌ Missing required files: {missing_files}")
                            return False
                        else:
                            print("✅ All required EPUB files present")
                            
                except Exception as e:
                    print(f"❌ ZIP validation failed: {e}")
                    return False
                
                print("✅ Hybrid EPUB file appears to be valid")
                return True
            else:
                print("❌ EPUB file was not created")
                return False
        else:
            print("❌ Hybrid EPUB generation failed")
            return False
            
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # Clean up
        for file in [csv_file, 'test_hybrid_output.epub']:
            if os.path.exists(file):
                os.remove(file)
                print(f"Cleaned up: {file}")

if __name__ == "__main__":
    success = test_hybrid_epub()
    if success:
        print("\n🎉 Hybrid EPUB test passed!")
        sys.exit(0)
    else:
        print("\n💥 Hybrid EPUB test failed!")
        sys.exit(1)
