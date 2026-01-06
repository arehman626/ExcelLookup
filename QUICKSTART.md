# Excel Lookup Utility - Quick Start Guide

## Overview
The Excel Lookup utility is a comprehensive tool for comparing and reconciling data between two Excel files. It provides a user-friendly GUI for selecting files, sheets, and columns to compare.

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Steps
1. Clone or download this repository
2. Open a terminal/command prompt in the project directory
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

### Standard GUI Mode
```bash
python excel_lookup.py
```

This will launch the graphical interface where you can:
1. Select Excel files for both left and right panels
2. Choose sheets from each file
3. Select columns to compare
4. View reconciliation results
5. Export results to Excel

## Using the Application

### Step 1: Load Files
- Click "Select Excel File" in the left panel
- Browse and select your first Excel file
- Click "Select Excel File" in the right panel
- Browse and select your second Excel file

### Step 2: Choose Sheets
- Select the sheet you want to compare from the "Select Sheet" dropdown in the left panel
- Select the sheet you want to compare from the "Select Sheet" dropdown in the right panel
- Data preview will appear showing the first 5 rows

### Step 3: Select Compare Columns
- Choose the column from the left file's sheet that contains values to compare
- Choose the column from the right file's sheet that contains values to compare

### Step 4: Run Comparison
- Click the "Compare & Reconcile" button
- Results will appear in the "Reconciliation Results" text area

### Step 5: Review Results
The reconciliation report shows:
- Summary statistics (total values, matches, differences)
- List of matching values (present in both files)
- List of values only in the left file
- List of values only in the right file
- Match percentage

### Step 6: Export Results (Optional)
- Click "Export Results to Excel" button
- Choose a location and filename
- The exported file will contain multiple sheets:
  - Summary: High-level statistics
  - Matches: All matching values
  - Only in Left: Values present only in left file
  - Only in Right: Values present only in right file

## Example Use Case

### Scenario: Product Inventory Reconciliation
You have two files:
- `warehouse_inventory.xlsx` with a sheet "Stock" containing column "SKU"
- `online_catalog.xlsx` with a sheet "Products" containing column "Product_Code"

### Steps:
1. Load `warehouse_inventory.xlsx` in the left panel
2. Load `online_catalog.xlsx` in the right panel
3. Select "Stock" sheet in the left panel
4. Select "Products" sheet in the right panel
5. Select "SKU" column in the left panel
6. Select "Product_Code" column in the right panel
7. Click "Compare & Reconcile"
8. View which products are:
   - In both systems (matching)
   - Only in warehouse (not in online catalog)
   - Only in online catalog (not in warehouse)
9. Export results for further analysis

## Testing the Application

### Using Sample Data
Run the sample data generator:
```bash
python create_sample_data.py
```

This creates two sample Excel files in the `test_data` directory:
- `sample_left.xlsx` (Products and Employees sheets)
- `sample_right.xlsx` (Inventory and Staff sheets)

Use these files to test the application:
1. Load `sample_left.xlsx` and `sample_right.xlsx`
2. Compare Products sheet (Product_ID) vs Inventory sheet (Item_Code)
3. Expected results: 3 matches, 4 only in left, 3 only in right

### Running Automated Tests
```bash
python test_functionality.py
```

This verifies the core comparison logic is working correctly.

## Troubleshooting

### Issue: "No module named 'tkinter'"
**Solution**: tkinter comes pre-installed with Python, but on some Linux systems you may need to install it:
- Ubuntu/Debian: `sudo apt-get install python3-tk`
- Fedora: `sudo dnf install python3-tkinter`
- macOS: tkinter should be included with Python

### Issue: "Failed to load file"
**Solution**: Ensure the file is a valid Excel file (.xlsx or .xls) and not corrupted

### Issue: "No results to export"
**Solution**: Run the comparison first before trying to export results

## Features Summary

✓ Dual-panel interface for side-by-side comparison
✓ Support for multiple sheets in Excel files
✓ Flexible column selection
✓ Data preview before comparison
✓ Comprehensive reconciliation reports
✓ Statistical summaries with match percentages
✓ Export results to structured Excel file
✓ User-friendly error messages
✓ Clean, intuitive GUI

## Support

For issues or questions, please open an issue on the GitHub repository.

## License

MIT License - See LICENSE file for details
