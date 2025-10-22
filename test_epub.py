#!/usr/bin/env python3
"""
Test script for EPUB generation
"""

import os
import sys
import pandas as pd
import tempfile

# Add scripts directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from csv_to_epub_manual import create_manual_epub_from_csv

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
    test_csv = 'test_data.csv'
    df.to_csv(test_csv, index=False)
    print(f"Test CSV created: {test_csv} ({len(df)} rows)")
    
    return test_csv

def test_epub_generation():
    """Test EPUB generation"""
    print("=" * 50)
    print("TESTING EPUB GENERATION")
    print("=" * 50)
    
    # Create test CSV
    csv_file = create_test_csv()
    
    try:
        # Test EPUB generation
        epub_file = 'test_output.epub'
        success = create_manual_epub_from_csv(
            csv_file,
            epub_file,
            'Test CSV Data',
            'Test Author'
        )
        
        if success:
            print(f"✅ EPUB generation successful: {epub_file}")
            
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
                except Exception as e:
                    print(f"❌ ZIP validation failed: {e}")
                    return False
                
                print("✅ EPUB file appears to be valid")
                return True
            else:
                print("❌ EPUB file was not created")
                return False
        else:
            print("❌ EPUB generation failed")
            return False
            
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        return False
    finally:
        # Clean up
        for file in [csv_file, 'test_output.epub']:
            if os.path.exists(file):
                os.remove(file)
                print(f"Cleaned up: {file}")

if __name__ == "__main__":
    success = test_epub_generation()
    if success:
        print("\n🎉 All tests passed!")
        sys.exit(0)
    else:
        print("\n💥 Tests failed!")
        sys.exit(1)
