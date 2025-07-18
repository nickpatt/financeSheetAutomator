#!/usr/bin/env python3
"""
Build script to create standalone .exe from the GUI version
Fixed to properly handle module imports
"""

import subprocess
import sys
import os
import shutil

def install_pyinstaller():
    """Install PyInstaller if not already installed"""
    try:
        import PyInstaller
        print("✓ PyInstaller is already installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller installed successfully")

def check_dependencies():
    """Check if all required dependencies are installed"""
    required_packages = [
        ('pandas', 'pandas'),
        ('openpyxl', 'openpyxl'),
        ('docx', 'python-docx'),
        ('PIL', 'pillow')
    ]
    
    missing_packages = []
    for import_name, package_name in required_packages:
        try:
            __import__(import_name)
            print(f"✓ {package_name} is installed")
        except ImportError:
            missing_packages.append(package_name)
            print(f"❌ {package_name} is missing")
    
    if missing_packages:
        print(f"\nInstalling missing dependencies: {', '.join(missing_packages)}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            print("✓ All dependencies installed successfully")
        except subprocess.CalledProcessError:
            print("❌ Failed to install dependencies")
            return False
    
    return True

def create_spec_file():
    """Create a custom spec file for better control"""
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['daily_summary_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('quarterly sheets', 'quarterly sheets')] if os.path.exists('quarterly sheets') else [],
    hiddenimports=[
        'daily_summary_generator',
        'pandas',
        'pandas.core.common',
        'pandas.core.ops',
        'pandas.tseries.offsets',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timezones',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.colors',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.styles.borders',
        'openpyxl.styles.alignment',
        'openpyxl.utils',
        'openpyxl.utils.cell',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'docx',
        'docx.shared',
        'docx.enum',
        'docx.enum.table',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.oxml.shared',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'threading',
        'datetime',
        'argparse',
        'io',
        'contextlib',
        'subprocess'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DailySummaryGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
"""
    
    with open('DailySummaryGenerator.spec', 'w') as f:
        f.write(spec_content)
    
    print("✓ Created custom spec file")

def build_exe():
    """Build the executable using the spec file"""
    print("Building Daily Summary Generator .exe...")
    
    # Create the spec file
    create_spec_file()
    
    # Build using the spec file
    cmd = [
        "pyinstaller",
        "--clean",
        "--noconfirm",
        "DailySummaryGenerator.spec"
    ]
    
    try:
        # Run PyInstaller
        print("Running PyInstaller with custom spec...")
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✓ Build completed successfully!")
        
        # Move the exe to the root directory
        exe_source = os.path.join("dist", "DailySummaryGenerator.exe")
        exe_dest = "DailySummaryGenerator.exe"
        
        if os.path.exists(exe_source):
            if os.path.exists(exe_dest):
                os.remove(exe_dest)
            shutil.move(exe_source, exe_dest)
            print(f"✓ Executable moved to: {exe_dest}")
            
            # Get file size
            file_size = os.path.getsize(exe_dest) / (1024 * 1024)  # MB
            print(f"✓ File size: {file_size:.2f} MB")
            
            # Clean up build files
            cleanup_build_files()
            
            print("\n" + "="*50)
            print("SUCCESS: DailySummaryGenerator.exe created!")
            print("="*50)
            print(f"File location: {os.path.abspath(exe_dest)}")
            print("You can now distribute this single .exe file.")
            print("No Python installation required on target computers.")
            
            return True
            
        else:
            print("❌ Error: Executable not found in dist folder")
            if result.stdout:
                print("STDOUT:", result.stdout)
            if result.stderr:
                print("STDERR:", result.stderr)
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        if hasattr(e, 'stdout') and e.stdout:
            print("STDOUT:", e.stdout)
        if hasattr(e, 'stderr') and e.stderr:
            print("STDERR:", e.stderr)
        return False

def cleanup_build_files():
    """Clean up temporary build files"""
    dirs_to_remove = ["build", "dist", "__pycache__"]
    files_to_remove = ["DailySummaryGenerator.spec"]
    
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print(f"✓ Cleaned up {dir_name}")
            except Exception as e:
                print(f"Warning: Could not remove {dir_name}: {e}")
            
    for file_name in files_to_remove:
        if os.path.exists(file_name):
            try:
                os.remove(file_name)
                print(f"✓ Cleaned up {file_name}")
            except Exception as e:
                print(f"Warning: Could not remove {file_name}: {e}")

def create_icon():
    """Create a simple icon file if it doesn't exist"""
    if not os.path.exists("icon.ico"):
        print("Creating a simple icon file...")
        try:
            # Try to create a simple icon using PIL if available
            from PIL import Image, ImageDraw
            
            # Create a simple 32x32 icon
            img = Image.new('RGBA', (32, 32), (0, 100, 200, 255))
            draw = ImageDraw.Draw(img)
            
            # Draw a simple document icon
            draw.rectangle([6, 4, 26, 28], fill=(255, 255, 255, 255), outline=(0, 0, 0, 255))
            draw.polygon([20, 4, 26, 10, 20, 10], fill=(200, 200, 200, 255))
            draw.line([10, 14, 22, 14], fill=(0, 0, 0, 255), width=1)
            draw.line([10, 18, 22, 18], fill=(0, 0, 0, 255), width=1)
            draw.line([10, 22, 22, 22], fill=(0, 0, 0, 255), width=1)
            
            img.save("icon.ico", format='ICO')
            print("✓ Created icon.ico")
            
        except ImportError:
            print("Note: PIL not available, skipping icon creation")
        except Exception as e:
            print(f"Note: Could not create icon: {e}")

def test_imports():
    """Test if all imports work correctly"""
    print("Testing imports...")
    
    try:
        # Test GUI imports
        import tkinter
        from tkinter import ttk, messagebox, filedialog
        print("✓ Tkinter imports working")
        
        # Test core module import
        from daily_summary_generator import generate_summary
        print("✓ Core module import working")
        
        # Test data processing imports
        import pandas as pd
        import openpyxl
        from docx import Document
        print("✓ Data processing imports working")
        
        return True
        
    except ImportError as e:
        print(f"❌ Import test failed: {e}")
        return False

def main():
    """Main build function"""
    print("Daily Summary Generator - Build Script")
    print("="*40)
    
    # Check if required files exist
    if not os.path.exists("daily_summary_gui.py"):
        print("❌ Error: daily_summary_gui.py not found!")
        print("Make sure you're running this script from the correct directory.")
        return
        
    if not os.path.exists("daily_summary_generator.py"):
        print("❌ Error: daily_summary_generator.py not found!")
        print("Make sure you're running this script from the correct directory.")
        return
    
    # Test imports first
    if not test_imports():
        print("❌ Import test failed. Please check your Python environment.")
        return
    
    # Check and install dependencies
    if not check_dependencies():
        print("❌ Error: Failed to install required dependencies")
        return
    
    # Install PyInstaller
    install_pyinstaller()
    
    # Create icon if needed
    create_icon()
    
    # Build the executable
    success = build_exe()
    
    if success:
        print("\n🎉 Build completed successfully!")
        print("You can now run DailySummaryGenerator.exe")
        print("\nTo test the executable:")
        print("1. Double-click DailySummaryGenerator.exe")
        print("2. The GUI should open without any console windows")
        print("3. Test generating a summary to ensure it works")
    else:
        print("\n❌ Build failed. Check the error messages above.")
        print("Common solutions:")
        print("1. Ensure all Python packages are installed")
        print("2. Try running: pip install --upgrade pyinstaller")
        print("3. Check that both .py files are in the same directory")

if __name__ == "__main__":
    main() 