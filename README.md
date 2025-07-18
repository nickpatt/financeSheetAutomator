# Daily Summary Generator

A comprehensive tool for generating daily invoicing summaries and updating quarterly YTD tracking sheets. This tool processes Excel files containing project and financial data, and outputs formatted Word and Excel reports with key financial summaries.

## Features

- **Daily Summary Reports**: Generate detailed daily invoicing summaries
- **Quarterly YTD Updates**: Automatically update quarterly YTD tracking sheets
- **Multiple Input Sources**: Support for multiple years of project data
- **Professional Output**: Formatted Word documents and Excel spreadsheets
- **GUI Interface**: User-friendly Windows interface with date pickers and progress tracking
- **Command Line Support**: Batch processing and automation capabilities

## Quick Start Guide

### Step 1: Build the Executable

1. **Prerequisites**: Ensure you have Python 3.7+ installed on your system
2. **Run the build script**: Double-click `build_exe.bat` or run it from command prompt
3. **Wait for completion**: The script will install dependencies and create `DailySummaryGenerator.exe`
4. **Verify**: You should see `DailySummaryGenerator.exe` in the same folder

### Step 2: Prepare Required Files

Before running the executable, ensure you have the following files in the correct locations:

#### Required Project List Files

The tool looks for Project List files in this order:
1. **Primary Location**: `N:\Project List\[year] Project List\` (Network drive)
2. **Fallback Location**: `quarterly sheets\` folder (local backup)
3. **Secondary Fallback**: `reports\` folder

**Required Files:**
- `2023 Project List.xlsx` (or `.xlsm`)
- `2024 Project List.xlsx` (or `.xlsm`) 
- `2025 Project List.xlsx` (or `.xlsm`)

**File Structure:**
```
N:\Project List\
├── 2023 Project List\
│   └── 2023 Project List.xlsx
├── 2024 Project List\
│   └── 2024 Project List.xlsx
└── 2025 Project List\
    └── 2025 Project List.xlsx
```

**OR (if using local folders):**
```
quarterly sheets\
├── 2023 Project List.xlsx
├── 2024 Project List.xlsx
└── 2025 Project List.xlsx
```

#### Required YTD Sheet Files

The tool automatically updates quarterly YTD sheets. These should be named:
- `2023 1st Quarter YTD.xlsx`
- `2023 2nd Quarter YTD.xlsx`
- `2023 3rd Quarter YTD.xlsx`
- `2023 4th Quarter YTD.xlsx`
- `2024 1st Quarter YTD.xlsx`
- `2024 2nd Quarter YTD.xlsx`
- `2024 3rd Quarter YTD.xlsx`
- `2024 4th Quarter YTD.xlsx`
- `2025 1st Quarter YTD.xlsx`
- `2025 2nd Quarter YTD.xlsx`
- `2025 3rd Quarter YTD.xlsx`
- `2025 4th Quarter YTD.xlsx`

**Location**: Same as Project List files (N: drive or local folders)

### Step 3: Run the Application

1. **Double-click** `DailySummaryGenerator.exe`
2. **Select Date**: Choose the target date for your summary
3. **Choose Output Directory**: Select where to save the generated reports
4. **Click Generate**: The tool will process your data and create reports

## What the Tool Does

### Input Processing
- Reads Project List Excel files for the specified years (2023-2025)
- Extracts invoice data, completion dates, and financial information
- Processes vendor payment data from colored cells in the spreadsheets
- Calculates receivables and payment totals

### Output Generation
1. **Daily Summary Word Document**: 
   - Today's totals
   - Week-to-date totals
   - Month-to-date totals
   - Receivables vs. vendor payments
   - Net receivables calculation

2. **Excel Tables File**:
   - Detailed invoice table for the selected date
   - Receivables vs. vendors table by year
   - Year-specific detail tables (2023, 2024, 2025)

3. **YTD Sheet Updates**:
   - Automatically updates the appropriate quarterly YTD sheet
   - Adds or replaces daily invoice tables
   - Maintains chronological order by date
   - Updates monthly totals

## File Requirements

### Project List Files Must Contain:
- **Sheet Names**: Year numbers (2023, 2024, 2025)
- **Header Row**: Row 5 (data starts at row 6)
- **Required Columns**:
  - ACGI # (Project/Invoice number)
  - Dept (Department)
  - Project Number/Name
  - Type
  - Client / PO #
  - Line #
  - PO Date
  - Amount
  - Invoice Date
  - Amount Invoiced
  - Completion Date

### YTD Sheet Format:
- **Row 1**: Month names (Jan, Feb, Mar, etc.)
- **Row 2**: Monthly totals
- **Row 5+**: Daily tables with date headers and invoice data

## Troubleshooting

### Common Issues:

1. **"Project List file not found"**
   - Ensure files are in the correct locations
   - Check file names match exactly (including case)
   - Verify both .xlsx and .xlsm extensions are supported

2. **"YTD sheet not found"**
   - Create the appropriate quarterly YTD file
   - Ensure proper naming convention: `YYYY Q# Quarter YTD.xlsx`

3. **"Permission denied" errors**
   - Ensure you have write access to the output directory
   - Close any open Excel files that might be locked

4. **Build errors**
   - Ensure Python 3.7+ is installed
   - Run `build_exe.bat` as administrator if needed
   - Check internet connection for dependency downloads

### Debug Mode:
The tool includes extensive logging. Check the console output for detailed information about:
- File locations being searched
- Data processing steps
- YTD sheet update operations
- Any errors or warnings

## Advanced Usage

### Command Line Options:
```bash
# Generate for specific date
DailySummaryGenerator.exe --date 2025-01-15

# Specify output directory
DailySummaryGenerator.exe --output-dir "C:\Reports"

# Process specific years only
DailySummaryGenerator.exe --years 2024 2025

# Interactive mode
DailySummaryGenerator.exe --interactive
```

### Batch Processing:
Use the included batch files:
- `run_summary.bat`: Run daily summary generation
- `run_quarterly_ytd.bat`: Update quarterly YTD files

## Support

For issues or questions:
1. Check the console output for error messages
2. Verify all required files are in the correct locations
3. Ensure file formats match the expected structure
4. Check file permissions and network access

## Version History

- **v1.0**: Initial release with basic daily summary generation
- **v1.1**: Added quarterly YTD sheet updates
- **v1.2**: Enhanced GUI interface and error handling
- **v1.3**: Improved file location fallback logic
- **v1.4**: Added comprehensive logging and debugging 