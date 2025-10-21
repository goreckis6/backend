#!/usr/bin/env python3
"""
Test Python environment and basic functionality
"""

import sys
import os

print("=" * 50)
print("PYTHON ENVIRONMENT TEST")
print("=" * 50)
print(f"Python version: {sys.version}")
print(f"Python executable: {sys.executable}")
print(f"Current working directory: {os.getcwd()}")
print(f"Python path: {sys.path}")

# Test basic imports
print("\nTesting basic imports...")
try:
    import pandas as pd
    print(f"✅ pandas: {pd.__version__}")
except Exception as e:
    print(f"❌ pandas: {e}")

try:
    import psutil
    print(f"✅ psutil: {psutil.__version__}")
    print(f"   CPU cores: {psutil.cpu_count(logical=True)}")
except Exception as e:
    print(f"❌ psutil: {e}")

try:
    from docx import Document
    print("✅ python-docx: OK")
except Exception as e:
    print(f"❌ python-docx: {e}")

try:
    import multiprocessing as mp
    print(f"✅ multiprocessing: OK (processes: {mp.cpu_count()})")
except Exception as e:
    print(f"❌ multiprocessing: {e}")

print("=" * 50)
print("ENVIRONMENT TEST COMPLETE")
print("=" * 50)
