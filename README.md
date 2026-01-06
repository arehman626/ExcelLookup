# ExcelLookup - Comprehensive Excel Data Comparison Tool

A comprehensive Excel Lookup utility to match and reconcile data between two Excel files with visual highlighting and detailed results.

## Features

- **Dual-Panel Interface**: Side-by-side comparison of two Excel files
- **Sheet Selection**: Choose specific sheets from each Excel file
- **Column Selection**: Select key columns for comparison and additional columns to include
- **VLOOKUP-style Comparison**: Matches data based on selected key columns
- **Visual Highlighting**: 
  - Green background for matched records
  - Red background for non-matched records (left-only or right-only)
- **Timestamped Output**: Results saved with timestamp and original filenames
- **Comprehensive Results**: Shows all data from both files with match status

## Quick Start

1. **Install dependencies**:
```bash
pip install -r requirements.txt
```

2. **Run the application**:
```bash
python run.py
```
or
```bash
python excel_lookup_gui.py
```

3. **Use the GUI**:
   - Select Excel files in left and right panels
   - Choose sheets and columns
   - Click "Compare Data"
   - Save results with highlighting

## Documentation

- **[Installation Guide](INSTALL.md)** - Detailed installation instructions and troubleshooting
- **[User Guide](USER_GUIDE.md)** - Complete guide to using the application
- **[Examples](EXAMPLES.md)** - Step-by-step examples and use cases
- **[GUI Layout](GUI_LAYOUT.md)** - Visual reference for the application interface

## Requirements

- Python 3.7 or higher
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- xlrd >= 2.0.0
- tkinter (usually comes with Python)

## Installation

### Standard Installation
```bash
pip install -r requirements.txt
```

### Virtual Environment (Recommended)
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

For detailed installation instructions and troubleshooting, see [INSTALL.md](INSTALL.md).

## Usage

### Basic Workflow

1. **Select Files**:
   - Click "Browse..." in the left panel to select the source Excel file
   - Click "Browse..." in the right panel to select the comparison Excel file

2. **Select Sheets**:
   - Choose the sheet from the left file dropdown
   - Choose the sheet from the right file dropdown

3. **Select Columns**:
   - Select the key column from the left file (used for matching)
   - Select the key column from the right file (used for matching)
   - Optionally, select additional columns to include in the results

4. **Compare**:
   - Click "Compare Data" button
   - View the summary showing matches and non-matches

5. **Save Results**:
   - Click "Save Results" button
   - Choose location and filename (default includes both filenames and timestamp)
   - Results will be saved with color highlighting

For detailed usage instructions, see [USER_GUIDE.md](USER_GUIDE.md).

## Output Format

The output Excel file includes:
- All rows from both files
- Key column used for matching
- Selected columns from left file (prefixed with "Left_")
- Selected columns from right file (prefixed with "Right_")
- Match_Status column showing:
  - "Match": Record found in both files (green background)
  - "Left Only": Record only in left file (red background)
  - "Right Only": Record only in right file (red background)

## Example Use Cases

- Reconciling financial data between two systems
- Comparing inventory lists from different sources
- Validating data migration results
- Matching customer records across databases
- Auditing data consistency

See [EXAMPLES.md](EXAMPLES.md) for detailed examples with sample data.

## Testing

Run the included test script to verify installation:

```bash
python test_excel_lookup.py
```

This will:
- Create sample Excel files
- Perform a comparison
- Generate a test output file
- Verify the results

## Error Handling

The application includes comprehensive error handling for:
- Invalid file formats
- Missing sheets
- Empty dataframes
- File access issues
- Column selection errors
- Data type mismatches

## Project Structure

```
ExcelLookup/
├── excel_lookup_gui.py    # Main application
├── run.py                 # Quick start script
├── requirements.txt       # Python dependencies
├── test_excel_lookup.py   # Test script
├── README.md              # This file
├── INSTALL.md             # Installation guide
├── USER_GUIDE.md          # User manual
├── EXAMPLES.md            # Usage examples
├── GUI_LAYOUT.md          # Interface reference
└── .gitignore             # Git ignore rules
```

## Performance

- **Small files** (< 1,000 rows): Near instant
- **Medium files** (1,000 - 10,000 rows): 2-5 seconds
- **Large files** (10,000 - 100,000 rows): 5-30 seconds
- **Very large files** (> 100,000 rows): May take longer, consider filtering data first

## Troubleshooting

### Common Issues

1. **"No module named 'tkinter'"**:
   - Linux: `sudo apt-get install python3-tk`
   - See [INSTALL.md](INSTALL.md) for platform-specific solutions

2. **"Failed to load file"**:
   - Ensure file is not open in Excel
   - Verify file is valid Excel format (.xlsx or .xls)

3. **Unexpected comparison results**:
   - Verify key columns contain unique identifiers
   - Check for leading/trailing spaces in data

For more troubleshooting help, see [USER_GUIDE.md](USER_GUIDE.md).

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the MIT License.

## Support

For issues, questions, or suggestions:
1. Check the documentation files (INSTALL.md, USER_GUIDE.md, EXAMPLES.md)
2. Review common issues in the User Guide
3. Create an issue on GitHub with detailed information

## Author

Created as a comprehensive Excel data comparison and reconciliation tool.

## Version

1.0.0 - Initial release with full GUI and comparison features 
