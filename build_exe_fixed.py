#!/usr/bin/env python3
"""
Fixed build script that handles existing exe files properly
"""

import subprocess
import sys
import os
import shutil
import time

def check_if_exe_running():
    """Check if the exe is currently running"""
    try:
        # Try to rename the file - if it's running, this will fail
        if os.path.exists("DailySummaryGenerator.exe"):
            temp_name = f"DailySummaryGenerator_backup_{int(time.time())}.exe"
            os.rename("DailySummaryGenerator.exe", temp_name)
            os.rename(temp_name, "DailySummaryGenerator.exe")
            return False
    except OSError:
        return True
    return False

def build_exe():
    """Build the executable with proper error handling"""
    print("Daily Summary Generator - Fixed Build Script")
    print("=" * 45)
    
    # Check if files exist
    if not os.path.exists("daily_summary_gui.py"):
        print("❌ Error: daily_summary_gui.py not found!")
        return False
    
    if not os.path.exists("daily_summary_generator.py"):
        print("❌ Error: daily_summary_generator.py not found!")
        return False
    
    # Check if exe is running
    if check_if_exe_running():
        print("❌ Error: DailySummaryGenerator.exe is currently running!")
        print("Please close the application and try again.")
        return False
    
    # Install PyInstaller if needed
    try:
        import PyInstaller
        print("✓ PyInstaller is available")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Clean up any existing build files
    print("Cleaning up previous build files...")
    for item in ["build", "dist", "__pycache__", "*.spec"]:
        if os.path.exists(item):
            try:
                if os.path.isdir(item):
                    shutil.rmtree(item)
                else:
                    os.remove(item)
                print(f"✓ Removed {item}")
            except:
                pass
    
    # Remove existing exe if possible
    if os.path.exists("DailySummaryGenerator.exe"):
        try:
            os.remove("DailySummaryGenerator.exe")
            print("✓ Removed existing executable")
        except PermissionError:
            backup_name = f"DailySummaryGenerator_old_{int(time.time())}.exe"
            try:
                os.rename("DailySummaryGenerator.exe", backup_name)
                print(f"✓ Backed up existing exe to {backup_name}")
            except:
                print("❌ Cannot remove or backup existing exe. Please close it and try again.")
                return False
    
    # Simple PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed", 
        "--name=DailySummaryGenerator",
        "--noconfirm",
        "--clean",
        "--hidden-import=daily_summary_generator",
        "daily_summary_gui.py"
    ]
    
    # Add icon if available
    if os.path.exists("icon.ico"):
        cmd.insert(-1, "--icon=icon.ico")
    
    try:
        print("Building executable (this takes a few minutes)...")
        print("Please wait...")
        
        # Run PyInstaller and show output
        result = subprocess.run(cmd, check=False, text=True, capture_output=True)
        
        if result.returncode == 0:
            print("✓ PyInstaller completed successfully")
        else:
            print("❌ PyInstaller failed:")
            print("STDERR:", result.stderr[-500:])  # Last 500 chars
            return False
        
        # Check if exe was created
        exe_path = os.path.join("dist", "DailySummaryGenerator.exe")
        if os.path.exists(exe_path):
            # Move to main directory
            shutil.move(exe_path, "DailySummaryGenerator.exe")
            
            file_size = os.path.getsize("DailySummaryGenerator.exe") / (1024 * 1024)
            print(f"✓ Created DailySummaryGenerator.exe ({file_size:.1f} MB)")
            
            # Clean up
            shutil.rmtree("build", ignore_errors=True)
            shutil.rmtree("dist", ignore_errors=True)
            
            print("\n" + "="*50)
            print("🎉 SUCCESS!")
            print("="*50)
            print("DailySummaryGenerator.exe has been created!")
            print("You can now run it by double-clicking the file.")
            return True
        else:
            print("❌ Executable was not created")
            return False
            
    except Exception as e:
        print(f"❌ Build failed: {e}")
        return False

if __name__ == "__main__":
    success = build_exe()
    if not success:
        print("\nPress Enter to exit...")
        input() 