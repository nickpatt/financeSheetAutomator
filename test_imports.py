#!/usr/bin/env python3
"""
Test script to verify all imports work correctly
"""

def test_imports():
    print("Testing imports...")
    
    try:
        # Test GUI imports
        import tkinter
        from tkinter import ttk, messagebox, filedialog
        print("✓ Tkinter imports working")
        
        # Test core module import
        from daily_summary_generator import generate_summary
        print("✓ Core module import working")
        print(f"✓ generate_summary function found: {generate_summary}")
        
        # Test data processing imports
        import pandas as pd
        print("✓ Pandas import working")
        
        import openpyxl
        print("✓ OpenPyXL import working")
        
        from docx import Document
        print("✓ Python-docx import working")
        
        # Test other imports
        import threading
        import datetime
        import os
        import sys
        import argparse
        print("✓ Standard library imports working")
        
        print("\n🎉 All imports successful!")
        return True
        
    except ImportError as e:
        print(f"❌ Import test failed: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

if __name__ == "__main__":
    success = test_imports()
    if success:
        print("\n✅ Ready to build executable!")
    else:
        print("\n❌ Fix import issues before building") 