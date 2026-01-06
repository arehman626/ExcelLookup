# User Guide - Excel Lookup Tool

## Quick Start

1. Launch the application
2. Select two Excel files to compare
3. Choose sheets and columns
4. Compare and save results

## Detailed Instructions

### Step 1: Launch the Application

Run the application using one of these methods:
```bash
python run.py
# or
python excel_lookup_gui.py
```

### Step 2: Select Files

#### Left Panel (Source File)
1. Click "Browse..." in the Left File section
2. Navigate to and select your source Excel file
3. The file name will appear, and available sheets will load

#### Right Panel (Comparison File)
1. Click "Browse..." in the Right File section
2. Navigate to and select your comparison Excel file
3. The file name will appear, and available sheets will load

**Tip**: You can select the same file in both panels if you want to compare different sheets within one file.

### Step 3: Select Sheets

1. **Left Sheet**: Use the dropdown to select the sheet from the left file
2. **Right Sheet**: Use the dropdown to select the sheet from the right file

After selection, you'll see:
- Number of rows and columns in each sheet
- Available columns populate in the column selection area

### Step 4: Choose Comparison Columns

#### Key Columns (Required)
These are the columns used to match records between files:

1. **Left Column**: Select the key column from the left file (e.g., "ID", "Employee Number")
2. **Right Column**: Select the key column from the right file

**Note**: The key columns don't need to have the same name, but they should contain comparable values.

#### Additional Columns (Optional)
Select extra columns to include in the results:

1. **Left Additional Columns**: 
   - Click/Ctrl+Click to select multiple columns from the left file
   - These will appear in the results with "Left_" prefix

2. **Right Additional Columns**:
   - Click/Ctrl+Click to select multiple columns from the right file
   - These will appear in the results with "Right_" prefix

**Tip**: You can include all columns by selecting all items in the listboxes.

### Step 5: Compare Data

1. Click the "Compare Data" button
2. A progress message will appear in the status bar
3. After completion, a summary dialog shows:
   - **Matches**: Records found in both files
   - **Left Only**: Records only in the left file
   - **Right Only**: Records only in the right file

### Step 6: Save Results

1. Click the "Save Results" button
2. Choose a location and filename:
   - Default name includes both source filenames and timestamp
   - Example: `Comparison_employees_vs_contractors_20240106_153045.xlsx`
3. Click "Save"

The saved file contains:
- All matched and unmatched records
- **Green background**: Matched records (found in both files)
- **Red background**: Unmatched records (found in only one file)
- Match status column indicating "Match", "Left Only", or "Right Only"

### Step 7: Review Results

Open the saved Excel file to review:

1. **Key Column**: Shows the matching values
2. **Left Columns**: Data from the left file (prefixed with "Left_")
3. **Right Columns**: Data from the right file (prefixed with "Right_")
4. **Match_Status**: Indicates the match status
5. **Color Coding**: Green for matches, red for non-matches

## Common Use Cases

### Use Case 1: Employee Reconciliation
- **Left**: Current employee roster
- **Right**: Payroll system export
- **Key Column**: Employee ID
- **Result**: Identify employees in one system but not the other

### Use Case 2: Inventory Comparison
- **Left**: Expected inventory
- **Right**: Actual inventory count
- **Key Column**: Product SKU
- **Result**: Find discrepancies in inventory

### Use Case 3: Data Migration Validation
- **Left**: Original database export
- **Right**: New system export
- **Key Column**: Record ID
- **Result**: Verify all records migrated successfully

### Use Case 4: Financial Reconciliation
- **Left**: Bank statement
- **Right**: Accounting records
- **Key Column**: Transaction ID
- **Result**: Identify missing or extra transactions

## Tips and Best Practices

### Data Preparation
1. **Clean your data**: Remove duplicate key values in each file
2. **Consistent formatting**: Ensure key columns use the same format (e.g., no leading zeros in one file)
3. **No merged cells**: Excel files should have proper tabular data

### Performance
1. **Large files**: For files with >100,000 rows, consider:
   - Filtering data before comparison
   - Using a more powerful computer
   - Closing other applications
2. **Multiple comparisons**: Use "Reset" button between comparisons

### File Organization
1. **Backup originals**: Keep original files unchanged
2. **Timestamp results**: Default naming includes timestamp for versioning
3. **Descriptive names**: Rename saved files to describe the comparison

## Keyboard Shortcuts

- **Ctrl/Cmd + A**: Select all in listboxes (when focused)
- **Ctrl/Cmd + Click**: Multi-select items in listboxes
- **Tab**: Navigate between fields

## Troubleshooting

### "Failed to load file" Error
- Verify file is not open in Excel
- Check file is a valid Excel format (.xlsx or .xls)
- Ensure you have read permissions

### "No results to save" Warning
- Click "Compare Data" before "Save Results"
- Ensure both files are loaded

### Comparison produces unexpected results
- Verify key columns contain unique identifiers
- Check for leading/trailing spaces in data
- Ensure data types match (e.g., both numeric or both text)

### Application freezes during comparison
- Wait for large files to process
- Check system resources (RAM)
- Consider breaking large files into smaller chunks

## Advanced Features

### Multiple Column Selection
Hold Ctrl (Windows/Linux) or Cmd (macOS) while clicking to select multiple columns in the additional columns listboxes.

### Reset Function
Click "Reset" to clear all selections and start fresh without closing the application.

## Limitations

- Maximum file size depends on available RAM
- Excel files only (.xlsx, .xls)
- Cannot compare more than two files simultaneously
- Color highlighting only in output file (not in preview)

## Next Steps

- Review the [README.md](README.md) for more information
- Check [INSTALL.md](INSTALL.md) for installation help
- Report issues on GitHub
