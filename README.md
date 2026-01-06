# ExcelLookup

A comprehensive Excel Lookup utility built in Python to match data from one file against another and provide detailed reconciliation reports.

## Features

- **Dual-Panel Interface**: Load and compare two Excel files side by side
- **Sheet Selection**: Choose specific sheets from each Excel file
- **Column Comparison**: Select which columns to compare between files
- **Data Preview**: View sample data before performing comparison
- **Comprehensive Reconciliation**: Get detailed reports showing:
  - Matching values between files
  - Values only in left file
  - Values only in right file
  - Statistical summaries
- **Export Results**: Save reconciliation results to Excel with multiple sheets

## Installation

1. Clone this repository:
```bash
git clone https://github.com/arehman626/ExcelLookup.git
cd ExcelLookup
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Mode (Recommended)

Run the graphical application:
```bash
python excel_lookup.py
```

### Command Line Mode

For batch processing or environments without GUI support:

```bash
# List sheets in a file
python excel_lookup_cli.py list-sheets file.xlsx

# List columns in a sheet
python excel_lookup_cli.py list-columns file.xlsx SheetName

# Compare two files
python excel_lookup_cli.py compare \
  --left-file data1.xlsx --left-sheet Sheet1 --left-col ID \
  --right-file data2.xlsx --right-sheet Sheet2 --right-col Code \
  --output results.xlsx
```

### Step-by-Step Guide

1. **Select Excel Files**:
   - Click "Select Excel File" in the left panel to choose your first file
   - Click "Select Excel File" in the right panel to choose your second file

2. **Choose Sheets**:
   - Select the sheet you want to compare from the dropdown in each panel
   - Preview of the data will appear below

3. **Select Compare Columns**:
   - Choose which column from the left file you want to use for comparison
   - Choose which column from the right file you want to use for comparison

4. **Run Comparison**:
   - Click "Compare & Reconcile" button
   - Results will appear in the bottom panel

5. **Export Results** (Optional):
   - Click "Export Results to Excel" to save the reconciliation report
   - Choose location and filename for the output

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.0.0
- tkinter (usually included with Python)

## Repository Structure

- `excel_lookup.py` - Main GUI application
- `excel_lookup_cli.py` - Command-line interface
- `create_sample_data.py` - Generate sample Excel files for testing
- `test_functionality.py` - Automated tests for core functionality
- `requirements.txt` - Python dependencies
- `QUICKSTART.md` - Detailed usage guide with examples
- `README.md` - This file

## Use Cases

- **Data Reconciliation**: Compare two versions of a dataset to find differences
- **Data Validation**: Verify that data exists in both sources
- **List Comparison**: Check membership between two lists
- **Quality Assurance**: Ensure data consistency across systems
- **Migration Verification**: Confirm data transfer completeness

## Output

The reconciliation report includes:
- Summary statistics (total values, matches, differences)
- List of values only in the left file
- List of values only in the right file
- List of matching values
- Exportable Excel file with separate sheets for each category

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
