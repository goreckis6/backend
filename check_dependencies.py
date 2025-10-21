#!/usr/bin/env python3
"""
Dependency checker for CSV to DOC conversion
"""

print("=" * 50)
print("CHECKING PYTHON DEPENDENCIES")
print("=" * 50)

# Check built-in modules
try:
    import os
    import sys
    import multiprocessing
    from concurrent.futures import ProcessPoolExecutor
    print("✅ Built-in modules: OK")
except ImportError as e:
    print(f"❌ Built-in modules: FAILED - {e}")

# Check pandas
try:
    import pandas as pd
    print(f"✅ pandas: OK (version: {pd.__version__})")
except ImportError as e:
    print(f"❌ pandas: MISSING - {e}")

# Check psutil
try:
    import psutil
    print(f"✅ psutil: OK (version: {psutil.__version__})")
    print(f"   CPU cores: {psutil.cpu_count(logical=True)}")
except ImportError as e:
    print(f"❌ psutil: MISSING - {e}")

# Check python-docx
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    print("✅ python-docx: OK")
except ImportError as e:
    print(f"❌ python-docx: MISSING - {e}")

# Check openpyxl (for Excel support)
try:
    import openpyxl
    print(f"✅ openpyxl: OK (version: {openpyxl.__version__})")
except ImportError as e:
    print(f"❌ openpyxl: MISSING - {e}")

# Check xlsxwriter (for Excel support)
try:
    import xlsxwriter
    print(f"✅ xlsxwriter: OK (version: {xlsxwriter.__version__})")
except ImportError as e:
    print(f"❌ xlsxwriter: MISSING - {e}")

print("=" * 50)
print("DEPENDENCY CHECK COMPLETE")
print("=" * 50)
