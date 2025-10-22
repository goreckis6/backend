#!/usr/bin/env python3
"""
Test script to check Calibre installation
"""

import os
import subprocess
import sys

def test_calibre():
    print("=== Calibre Installation Test ===")
    
    # Check common paths
    paths_to_check = [
        '/usr/bin/calibre',
        '/usr/bin/ebook-convert',
        '/usr/local/bin/calibre',
        '/usr/local/bin/ebook-convert',
        '/opt/calibre/bin/calibre',
        '/opt/calibre/bin/ebook-convert',
        '/usr/share/calibre/bin/calibre',
        '/usr/share/calibre/bin/ebook-convert'
    ]
    
    print("Checking paths:")
    for path in paths_to_check:
        exists = os.path.exists(path)
        print(f"  {path}: {'✅' if exists else '❌'}")
        if exists:
            try:
                result = subprocess.run([path, '--version'], capture_output=True, text=True, timeout=5)
                print(f"    Version: {result.stdout.strip()}")
            except Exception as e:
                print(f"    Error running: {e}")
    
    # Check directories
    print("\nChecking directories:")
    dirs_to_check = ['/usr/bin', '/usr/local/bin', '/opt/calibre/bin', '/usr/share/calibre/bin']
    for dir_path in dirs_to_check:
        if os.path.exists(dir_path):
            try:
                files = os.listdir(dir_path)
                calibre_files = [f for f in files if 'calibre' in f.lower() or 'ebook' in f.lower()]
                if calibre_files:
                    print(f"  {dir_path}: {calibre_files}")
                else:
                    print(f"  {dir_path}: (no calibre files)")
            except Exception as e:
                print(f"  {dir_path}: Error listing - {e}")
        else:
            print(f"  {dir_path}: Does not exist")
    
    # Try which command
    print("\nTrying 'which' command:")
    try:
        result = subprocess.run(['which', 'calibre'], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"  calibre found at: {result.stdout.strip()}")
        else:
            print("  calibre not found in PATH")
    except Exception as e:
        print(f"  Error running which: {e}")
    
    try:
        result = subprocess.run(['which', 'ebook-convert'], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"  ebook-convert found at: {result.stdout.strip()}")
        else:
            print("  ebook-convert not found in PATH")
    except Exception as e:
        print(f"  Error running which: {e}")

if __name__ == "__main__":
    test_calibre()
