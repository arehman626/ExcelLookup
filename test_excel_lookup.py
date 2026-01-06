"""
Test script for Excel Lookup GUI functionality
Tests the core comparison logic without GUI interaction
"""

import pandas as pd
import os
import traceback
from datetime import datetime


def test_comparison_logic():
    """Test the comparison logic"""
    print("Testing Excel Lookup comparison logic...")
    
    # Load test files
    left_df = pd.read_excel('test_left.xlsx', sheet_name='Employees')
    right_df = pd.read_excel('test_right.xlsx', sheet_name='Employees')
    
    print(f"✓ Loaded left file: {len(left_df)} rows")
    print(f"✓ Loaded right file: {len(right_df)} rows")
    
    # Set comparison columns
    left_key_col = 'ID'
    right_key_col = 'ID'
    
    # Get additional columns
    left_additional = ['Name', 'Department', 'Salary']
    right_additional = ['Name', 'Location', 'Status']
    
    # Prepare columns
    left_cols = [left_key_col] + left_additional
    right_cols = [right_key_col] + right_additional
    
    # Create subsets
    left_subset = left_df[left_cols].copy()
    right_subset = right_df[right_cols].copy()
    
    # Rename key columns for merging
    left_subset = left_subset.rename(columns={left_key_col: 'key_column'})
    right_subset = right_subset.rename(columns={right_key_col: 'key_column'})
    
    # Add suffixes to other columns
    left_subset.columns = ['key_column'] + [f'Left_{col}' for col in left_subset.columns[1:]]
    right_subset.columns = ['key_column'] + [f'Right_{col}' for col in right_subset.columns[1:]]
    
    print(f"✓ Prepared left columns: {list(left_subset.columns)}")
    print(f"✓ Prepared right columns: {list(right_subset.columns)}")
    
    # Perform merge
    result_df = pd.merge(left_subset, right_subset, on='key_column', 
                        how='outer', indicator=True)
    
    # Rename merge indicator
    result_df = result_df.rename(columns={'_merge': 'Match_Status'})
    
    # Map status
    status_map = {
        'both': 'Match',
        'left_only': 'Left Only',
        'right_only': 'Right Only'
    }
    result_df['Match_Status'] = result_df['Match_Status'].map(status_map)
    
    # Rename key_column back
    result_df = result_df.rename(columns={'key_column': f'Key_({left_key_col}={right_key_col})'})
    
    # Count results
    matches = len(result_df[result_df['Match_Status'] == 'Match'])
    left_only = len(result_df[result_df['Match_Status'] == 'Left Only'])
    right_only = len(result_df[result_df['Match_Status'] == 'Right Only'])
    
    print(f"\n✓ Comparison complete!")
    print(f"  - Matches: {matches}")
    print(f"  - Left Only: {left_only}")
    print(f"  - Right Only: {right_only}")
    print(f"  - Total rows: {len(result_df)}")
    
    # Verify expected results
    assert matches == 3, f"Expected 3 matches, got {matches}"
    assert left_only == 2, f"Expected 2 left only, got {left_only}"
    assert right_only == 2, f"Expected 2 right only, got {right_only}"
    
    print("\n✓ All assertions passed!")
    
    # Test save functionality
    print("\nTesting save functionality...")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    test_output = f"test_output_{timestamp}.xlsx"
    
    from openpyxl.styles import PatternFill
    
    with pd.ExcelWriter(test_output, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='Comparison_Results', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Comparison_Results']
        
        green_fill = PatternFill(start_color='90EE90', end_color='90EE90', 
                                fill_type='solid')
        red_fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', 
                              fill_type='solid')
        
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, 
                                                          max_row=len(result_df) + 1), 
                                     start=2):
            status = result_df.iloc[row_idx - 2]['Match_Status']
            
            for cell in row:
                if status == 'Match':
                    cell.fill = green_fill
                else:
                    cell.fill = red_fill
    
    print(f"✓ Saved test output to: {test_output}")
    
    # Verify file was created
    assert os.path.exists(test_output), "Output file not created"
    
    # Load and verify
    verify_df = pd.read_excel(test_output, sheet_name='Comparison_Results')
    assert len(verify_df) == len(result_df), "Row count mismatch"
    
    print(f"✓ Verified output file: {len(verify_df)} rows")
    
    print("\n" + "="*50)
    print("ALL TESTS PASSED! ✓")
    print("="*50)
    
    return True


if __name__ == "__main__":
    try:
        test_comparison_logic()
    except Exception as e:
        print(f"\n✗ TEST FAILED: {str(e)}")
        traceback.print_exc()
        exit(1)
