# ExcelLookup Examples

This directory contains example scripts and sample files demonstrating how to use the ExcelLookup library.

## Files

- `example_comparison.py` - Complete example showing file comparison workflow
- `sample_file1.xlsx` - Sample Excel file (generated automatically)
- `sample_file2.xlsx` - Sample Excel file (generated automatically)
- `comparison_report.xlsx` - Generated comparison report (created after running example)

## Running the Example

```bash
# From the project root directory
cd examples
python example_comparison.py
```

This will:
1. Create sample Excel files with employee data
2. Compare the files using ExcelLookup
3. Generate a detailed comparison report
4. Display summary statistics

## What the Example Demonstrates

- Reading Excel files
- Configuring comparison settings
- Performing data comparison
- Generating reports
- Handling different data types (strings, numbers, dates)
- Working with key columns and ignored columns

## Sample Data Structure

The example uses employee data with the following columns:
- ID (key column)
- Name
- Department  
- Salary
- StartDate
- LastUpdated (ignored in comparison)

The second file contains intentional differences:
- Salary changes for some employees
- Department changes
- Name variations
- One employee removed, one added