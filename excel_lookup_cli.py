#!/usr/bin/env python3
"""
Excel Lookup Utility - Command Line Interface
For environments without GUI support or for batch processing
"""

import argparse
import pandas as pd
from pathlib import Path
import sys


# Display limits for results
MAX_DISPLAY_ITEMS = 50
MAX_DISPLAY_MATCHES = 30


def perform_comparison(left_file, left_sheet, left_col, right_file, right_sheet, right_col, output_file=None):
    """
    Perform Excel comparison from command line
    
    Args:
        left_file: Path to left Excel file
        left_sheet: Sheet name in left file
        left_col: Column name to compare in left file
        right_file: Path to right Excel file
        right_sheet: Sheet name in right file
        right_col: Column name to compare in right file
        output_file: Optional path for output Excel file
    """
    try:
        # Load data
        print(f"Loading {left_file}...")
        left_df = pd.read_excel(left_file, sheet_name=left_sheet)
        
        print(f"Loading {right_file}...")
        right_df = pd.read_excel(right_file, sheet_name=right_sheet)
        
        # Validate columns exist
        if left_col not in left_df.columns:
            print(f"Error: Column '{left_col}' not found in left file sheet '{left_sheet}'")
            print(f"Available columns: {', '.join(left_df.columns)}")
            return False
            
        if right_col not in right_df.columns:
            print(f"Error: Column '{right_col}' not found in right file sheet '{right_sheet}'")
            print(f"Available columns: {', '.join(right_df.columns)}")
            return False
        
        # Perform comparison
        print("\nPerforming comparison...")
        left_values = set(left_df[left_col].dropna().astype(str))
        right_values = set(right_df[right_col].dropna().astype(str))
        
        matches = left_values & right_values
        only_in_left = left_values - right_values
        only_in_right = right_values - left_values
        
        # Display results
        print("\n" + "=" * 80)
        print("EXCEL LOOKUP & RECONCILIATION REPORT")
        print("=" * 80)
        print()
        print(f"Left File:  {Path(left_file).name}")
        print(f"  Sheet:    {left_sheet}")
        print(f"  Column:   {left_col}")
        print(f"  Rows:     {len(left_df)}")
        print()
        print(f"Right File: {Path(right_file).name}")
        print(f"  Sheet:    {right_sheet}")
        print(f"  Column:   {right_col}")
        print(f"  Rows:     {len(right_df)}")
        print()
        print("-" * 80)
        print("SUMMARY")
        print("-" * 80)
        print(f"Total unique values in left:  {len(left_values)}")
        print(f"Total unique values in right: {len(right_values)}")
        print(f"Matching values:              {len(matches)} ({len(matches)/max(len(left_values), 1)*100:.1f}%)")
        print(f"Only in left:                 {len(only_in_left)}")
        print(f"Only in right:                {len(only_in_right)}")
        print()
        
        if only_in_left:
            print("-" * 80)
            print(f"VALUES ONLY IN LEFT FILE ({len(only_in_left)} items)")
            print("-" * 80)
            for i, value in enumerate(sorted(only_in_left)[:MAX_DISPLAY_ITEMS], 1):
                print(f"{i}. {value}")
            if len(only_in_left) > MAX_DISPLAY_ITEMS:
                print(f"... and {len(only_in_left) - MAX_DISPLAY_ITEMS} more")
            print()
        
        if only_in_right:
            print("-" * 80)
            print(f"VALUES ONLY IN RIGHT FILE ({len(only_in_right)} items)")
            print("-" * 80)
            for i, value in enumerate(sorted(only_in_right)[:MAX_DISPLAY_ITEMS], 1):
                print(f"{i}. {value}")
            if len(only_in_right) > MAX_DISPLAY_ITEMS:
                print(f"... and {len(only_in_right) - MAX_DISPLAY_ITEMS} more")
            print()
        
        if matches:
            print("-" * 80)
            print(f"MATCHING VALUES (showing first {MAX_DISPLAY_MATCHES} of {len(matches)})")
            print("-" * 80)
            for i, value in enumerate(sorted(matches)[:MAX_DISPLAY_MATCHES], 1):
                print(f"{i}. {value}")
            if len(matches) > MAX_DISPLAY_MATCHES:
                print(f"... and {len(matches) - MAX_DISPLAY_MATCHES} more")
            print()
        
        print("=" * 80)
        print("END OF REPORT")
        print("=" * 80)
        
        # Export if requested
        if output_file:
            print(f"\nExporting results to {output_file}...")
            
            matches_df = pd.DataFrame(sorted(matches), columns=["Matching Values"])
            only_left_df = pd.DataFrame(sorted(only_in_left), 
                                       columns=[f"Only in Left ({left_col})"])
            only_right_df = pd.DataFrame(sorted(only_in_right), 
                                        columns=[f"Only in Right ({right_col})"])
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                matches_df.to_excel(writer, sheet_name='Matches', index=False)
                only_left_df.to_excel(writer, sheet_name='Only in Left', index=False)
                only_right_df.to_excel(writer, sheet_name='Only in Right', index=False)
                
                summary_data = {
                    'Metric': ['Total Matches', 'Only in Left', 'Only in Right', 
                              'Left Column', 'Right Column'],
                    'Value': [len(matches), len(only_in_left), len(only_in_right),
                             left_col, right_col]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            print(f"Results exported successfully!")
        
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False


def list_sheets(file_path):
    """List all sheets in an Excel file"""
    try:
        excel_file = pd.ExcelFile(file_path)
        print(f"\nSheets in {Path(file_path).name}:")
        for i, sheet in enumerate(excel_file.sheet_names, 1):
            print(f"  {i}. {sheet}")
        return True
    except Exception as e:
        print(f"Error reading file: {e}")
        return False


def list_columns(file_path, sheet_name):
    """List all columns in a specific sheet"""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"\nColumns in '{sheet_name}' sheet of {Path(file_path).name}:")
        for i, col in enumerate(df.columns, 1):
            print(f"  {i}. {col}")
        print(f"\nTotal rows: {len(df)}")
        return True
    except Exception as e:
        print(f"Error reading sheet: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(
        description='Excel Lookup & Reconciliation Tool - Command Line Interface',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Compare two files
  %(prog)s compare \\
    --left-file data1.xlsx --left-sheet Sheet1 --left-col ID \\
    --right-file data2.xlsx --right-sheet Sheet1 --right-col Code \\
    --output results.xlsx

  # List sheets in a file
  %(prog)s list-sheets data.xlsx

  # List columns in a specific sheet
  %(prog)s list-columns data.xlsx Sheet1
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Command to execute')
    
    # Compare command
    compare_parser = subparsers.add_parser('compare', help='Compare two Excel files')
    compare_parser.add_argument('--left-file', required=True, help='Path to left Excel file')
    compare_parser.add_argument('--left-sheet', required=True, help='Sheet name in left file')
    compare_parser.add_argument('--left-col', required=True, help='Column name in left sheet')
    compare_parser.add_argument('--right-file', required=True, help='Path to right Excel file')
    compare_parser.add_argument('--right-sheet', required=True, help='Sheet name in right file')
    compare_parser.add_argument('--right-col', required=True, help='Column name in right sheet')
    compare_parser.add_argument('--output', '-o', help='Output Excel file path (optional)')
    
    # List sheets command
    sheets_parser = subparsers.add_parser('list-sheets', help='List all sheets in an Excel file')
    sheets_parser.add_argument('file', help='Path to Excel file')
    
    # List columns command
    cols_parser = subparsers.add_parser('list-columns', help='List all columns in a sheet')
    cols_parser.add_argument('file', help='Path to Excel file')
    cols_parser.add_argument('sheet', help='Sheet name')
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    if args.command == 'compare':
        success = perform_comparison(
            args.left_file, args.left_sheet, args.left_col,
            args.right_file, args.right_sheet, args.right_col,
            args.output
        )
        return 0 if success else 1
        
    elif args.command == 'list-sheets':
        success = list_sheets(args.file)
        return 0 if success else 1
        
    elif args.command == 'list-columns':
        success = list_columns(args.file, args.sheet)
        return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
