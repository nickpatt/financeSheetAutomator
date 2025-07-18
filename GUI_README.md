# Daily Summary Generator - GUI Version

This directory contains both command-line and GUI versions of the Daily Summary Generator, plus tools to create a standalone Windows executable.

## Files Overview

### Main Applications
- **`daily_summary_generator.py`** - Original command-line version
- **`daily_summary_gui.py`** - New GUI version with Windows interface
- **`DailySummaryGenerator.exe`** - Standalone executable (after building)

### Build Tools
- **`build_exe.py`** - Python script to create the .exe file
- **`build_exe.bat`** - Windows batch file to automate the build process
- **`requirements.txt`** - Python package dependencies

## Quick Start - Using the GUI

### Option 1: Run the Standalone .exe (Recommended)
1. Double-click `DailySummaryGenerator.exe`
2. The GUI will open with a user-friendly interface
3. Set your target date and output directory
4. Click "Generate Summary"

### Option 2: Run the Python GUI
1. Ensure Python is installed
2. Install dependencies: `pip install -r requirements.txt`
3. Run: `python daily_summary_gui.py`

## Building the .exe File

### Method 1: Using the Batch File (Easiest)
1. Double-click `build_exe.bat`
2. The script will:
   - Check for Python installation
   - Install required packages
   - Build the .exe file
   - Clean up temporary files

### Method 2: Manual Build
1. Install requirements: `pip install -r requirements.txt`
2. Run build script: `python build_exe.py`

## GUI Features

### User Interface
- **Date Selection**: Choose target date with date picker
- **Output Directory**: Browse and select output folder
- **Progress Tracking**: Real-time progress bar and status updates
- **Output Log**: View detailed processing information
- **Auto-Open Results**: Option to automatically open output folder when complete

### Configuration Display
- Shows primary data directory (N: drive)
- Lists fallback directories (quarterly sheets, reports)
- Displays supported file formats (.xlsx, .xlsm)

### Error Handling
- Input validation for dates and directories
- Dependency checking (ensures required packages are installed)
- Detailed error messages with troubleshooting information
- Graceful handling of missing files or network issues

## File Location Logic

The application searches for Project List files in this order:

1. **N:\Project List\[year] Project List\** (Network drive)
2. **quarterly sheets\** (Local backup folder)
3. **reports\** (Additional backup folder)

For each location, it tries both `.xlsx` and `.xlsm` extensions.

## Data Processing

### Receivables Calculation
- Reads from Column M of the "Totals" row in each year's Project List
- Sums amounts from 2023, 2024, and 2025 files

### Vendor Payments
- **2023 & 2024**: Uses Column V
- **2025 & later**: Uses Column W
- **Color Filter**: Only includes cells with cyan color `rgb(3,255,255)` or similar
- Sums all matching colored cells across all years

### Output Files
- **Word Document**: Summary with key totals and period comparisons
- **Excel File**: Detailed tables with invoice data and breakdowns

## Troubleshooting

### Common Issues

**"Missing Dependencies" Error**
- Solution: Run `pip install -r requirements.txt`

**"File Not Found" Errors**
- Check that Project List files exist in quarterly sheets folder
- Ensure files have correct names (e.g., "2025 Project List.xlsx")

**Build Fails**
- Ensure Python 3.7+ is installed
- Run `pip install --upgrade pip pyinstaller`
- Try building from command line: `python build_exe.py`

**Network Drive Issues**
- Application automatically falls back to local folders
- Copy files to "quarterly sheets" folder if N: drive unavailable

### Support File Structure
```
AllproCGI YTD automated script/
├── daily_summary_gui.py           # GUI application
├── daily_summary_generator.py     # Core functionality
├── build_exe.py                   # Build script
├── build_exe.bat                  # Build automation
├── requirements.txt               # Dependencies
├── quarterly sheets/              # Local data files
│   ├── 2023 Project List.xlsm
│   ├── 2024 Project List.xlsm
│   └── 2025 Project List.xlsx
└── reports/                       # Output directory
    ├── daily_summary_YYYYMMDD.docx
    └── daily_summary_tables_YYYYMMDD.xlsx
```

## Distribution

The standalone `DailySummaryGenerator.exe` file:
- Contains all required Python libraries
- Includes the quarterly sheets folder
- Can be run on any Windows computer without Python installation
- Is completely self-contained (no additional files needed)

## Version History

- **v1.0** - Command-line version
- **v2.0** - Added GUI interface
- **v2.1** - Added .exe build capability
- **v2.2** - Enhanced color filtering for vendor payments
- **v2.3** - Added support for both .xlsx and .xlsm files

For technical support or feature requests, please refer to the main README.md file. 