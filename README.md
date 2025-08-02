# Daily Summary Generator

A comprehensive tool for generating daily invoicing summaries and updating quarterly YTD tracking sheets. This tool processes Excel files containing project and financial data, and outputs formatted Word and Excel reports with key financial summaries.

## Quick Start Guide

### Step 1: Build the Executable

1. **Prerequisites**: Ensure you have Python 3.7+ installed on your system
2. **Run the build script**: Double-click `build_exe.bat` - it may take up to 5 minutes to finish.
3. **Wait for completion**: The script will create `DailySummaryGenerator.exe`
4. **Verify**: You should see `DailySummaryGenerator.exe` in the same folder

### Step 3: Run the Application

1. **Double-click** `DailySummaryGenerator.exe`
2. **Select Date**: Choose the target date for your summary
3. **Choose Output Directory**: Select where to save the generated reports
4. **Click Generate**: The tool will process your data and create reports


-----------------------------------------------------------------------------





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