# Excel Rank Processor

## Overview
This application processes Excel files containing mutual fund data and generates a new file with:
- Rank calculations for each fund
- Rank change indicators with arrows (↑ for improvement, ↓ for decline, ■ for no change)
- Formatted output with color-coding and highlighting

## System Requirements
- Python 3.8 or higher
- Windows, macOS, or Linux
- Minimum 4GB RAM

## Installation Instructions
1. Clone or download this repository
2. Open a terminal/command prompt in the project directory
3. Run the following command to install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Running the Application
### Windows
1. Double-click `run_app.bat`
2. OR run this command in Command Prompt:
   ```
   .\run_app.bat
   ```

### macOS/Linux
1. Make the script executable:
   ```
   chmod +x run_app.sh
   ```
2. Run the application:
   ```
   ./run_app.sh
   ```

## Using the Application
1. The web interface will open in your default browser
2. Upload your input Excel file (must contain 'MutualFund Name' column)
3. Optionally upload a reference file for rank comparisons
4. Click 'Process Files'
5. Download the processed output file when complete

## Troubleshooting
### Permission Errors
If you encounter permission issues:
```
pip install --user -r requirements.txt
```

### File Access Errors
Ensure no other programs are using the Excel files during processing

## File Formats
### Input File
- Must contain a column named 'MutualFund Name'
- Should have numerical data columns to rank

### Reference File (Optional)
- Used for rank change calculations
- Should have 'Rank' and 'Mutual Fund' columns

## License
MIT License - Free for personal and commercial use
