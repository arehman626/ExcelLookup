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

## Installation

1. Clone this repository:
```bash
git clone https://github.com/arehman626/ExcelLookup.git
cd ExcelLookup
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python excel_lookup_gui.py
```

2. **Select Files**:
   - Click "Browse..." in the left panel to select the source Excel file
   - Click "Browse..." in the right panel to select the comparison Excel file

3. **Select Sheets**:
   - Choose the sheet from the left file dropdown
   - Choose the sheet from the right file dropdown

4. **Select Columns**:
   - Select the key column from the left file (used for matching)
   - Select the key column from the right file (used for matching)
   - Optionally, select additional columns to include in the results

5. **Compare**:
   - Click "Compare Data" button
   - View the summary showing matches and non-matches

6. **Save Results**:
   - Click "Save Results" button
   - Choose location and filename (default includes both filenames and timestamp)
   - Results will be saved with color highlighting

## Requirements

- Python 3.7 or higher
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- xlrd >= 2.0.0
- tkinter (usually comes with Python)

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

## Error Handling

The application includes comprehensive error handling for:
- Invalid file formats
- Missing sheets
- Empty dataframes
- File access issues
- Column selection errors

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is open source and available under the MIT License. 
