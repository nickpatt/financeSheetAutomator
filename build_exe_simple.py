#!/usr/bin/env python3
"""
Simple build script to create .exe from GUI version
Uses direct PyInstaller commands instead of spec file
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

def build_exe_simple():
    """Build the executable using simple PyInstaller command"""
    print("Building Daily Summary Generator .exe...")
    
    # Build command with minimal options
    cmd = [
        "pyinstaller",
        "--onefile",                    # Single file
        "--windowed",                   # No console
        "--name=DailySummaryGenerator", # Name
        "--clean",                      # Clean build
        "--noconfirm",                  # Don't ask for confirmation
        # Only essential hidden imports
        "--hidden-import=daily_summary_generator",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=docx",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.messagebox",
        "--hidden-import=tkinter.filedialog",
        "daily_summary_gui.py"
    ]
    
    # Add icon if available
    if os.path.exists("icon.ico"):
        cmd.insert(-1, "--icon=icon.ico")
    
    # Add data folder if available
    if os.path.exists("quarterly sheets"):
        cmd.insert(-1, "--add-data=quarterly sheets;quarterly sheets")
    
    try:
        print("Running PyInstaller (this may take a few minutes)...")
        print("Command:", " ".join(cmd))
        
        # Run without capturing output so we can see progress
        result = subprocess.run(cmd, check=True)
        print("✓ Build completed successfully!")
        
        # Move exe to root
        exe_source = os.path.join("dist", "DailySummaryGenerator.exe")
        exe_dest = "DailySummaryGenerator.exe"
        
        if os.path.exists(exe_source):
            if os.path.exists(exe_dest):
                os.remove(exe_dest)
            shutil.move(exe_source, exe_dest)
            print(f"✓ Executable created: {exe_dest}")
            
            # Get file size
            file_size = os.path.getsize(exe_dest) / (1024 * 1024)
            print(f"✓ File size: {file_size:.2f} MB")
            
            # Clean up
            cleanup_build_files()
            
            print("\n" + "="*50)
            print("SUCCESS: DailySummaryGenerator.exe created!")
            print("="*50)
            print(f"Location: {os.path.abspath(exe_dest)}")
            
            return True
        else:
            print("❌ Error: Executable not found after build")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed with error code: {e.returncode}")
        return False
    except KeyboardInterrupt:
        print("\n❌ Build interrupted by user")
        return False

def cleanup_build_files():
    """Clean up build files"""
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

def main():
    """Main function"""
    print("Daily Summary Generator - Simple Build Script")
    print("="*45)
    
    # Check files exist
    if not os.path.exists("daily_summary_gui.py"):
        print("❌ Error: daily_summary_gui.py not found!")
        return
    
    if not os.path.exists("daily_summary_generator.py"):
        print("❌ Error: daily_summary_generator.py not found!")
        return
    
    # Test import
    try:
        from daily_summary_generator import generate_summary
        print("✓ Import test passed")
    except ImportError as e:
        print(f"❌ Import test failed: {e}")
        return
    
    # Install PyInstaller
    install_pyinstaller()
    
    # Build
    success = build_exe_simple()
    
    if success:
        print("\n🎉 Build completed successfully!")
        print("Double-click DailySummaryGenerator.exe to test")
    else:
        print("\n❌ Build failed")

if __name__ == "__main__":
    main() 