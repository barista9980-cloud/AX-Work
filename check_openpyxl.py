import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

try:
    import openpyxl
    print("openpyxl is installed!")
except ImportError:
    print("openpyxl is NOT installed!")
