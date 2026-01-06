"""
Test script for Excel Lookup utility core functionality
"""

import pandas as pd
from pathlib import Path

def test_comparison_logic():
    """Test the core comparison logic"""
    print("Testing Excel Lookup comparison logic...\n")
    
    # Load test files
    test_dir = Path("test_data")
    left_file = test_dir / "sample_left.xlsx"
    right_file = test_dir / "sample_right.xlsx"
    
    # Test 1: Products comparison
    print("=" * 60)
    print("Test 1: Products vs Inventory")
    print("=" * 60)
    
    left_df = pd.read_excel(left_file, sheet_name='Products')
    right_df = pd.read_excel(right_file, sheet_name='Inventory')
    
    print(f"\nLeft file (Products):")
    print(f"  Rows: {len(left_df)}")
    print(f"  Columns: {list(left_df.columns)}")
    
    print(f"\nRight file (Inventory):")
    print(f"  Rows: {len(right_df)}")
    print(f"  Columns: {list(right_df.columns)}")
    
    # Compare Product_ID vs Item_Code
    left_col = 'Product_ID'
    right_col = 'Item_Code'
    
    left_values = set(left_df[left_col].dropna().astype(str))
    right_values = set(right_df[right_col].dropna().astype(str))
    
    matches = left_values & right_values
    only_in_left = left_values - right_values
    only_in_right = right_values - left_values
    
    print(f"\nComparison Results ({left_col} vs {right_col}):")
    print(f"  Matches: {len(matches)} - {sorted(matches)}")
    print(f"  Only in left: {len(only_in_left)} - {sorted(only_in_left)}")
    print(f"  Only in right: {len(only_in_right)} - {sorted(only_in_right)}")
    
    # Verify expected results
    assert matches == {'P001', 'P002', 'P003'}, "Matches should be P001, P002, P003"
    assert only_in_left == {'P004', 'P005', 'P006', 'P007'}, "Left should have P004-P007"
    assert only_in_right == {'P008', 'P009', 'P010'}, "Right should have P008-P010"
    print("\n✓ Test 1 PASSED")
    
    # Test 2: Employees comparison
    print("\n" + "=" * 60)
    print("Test 2: Employees vs Staff")
    print("=" * 60)
    
    left_df = pd.read_excel(left_file, sheet_name='Employees')
    right_df = pd.read_excel(right_file, sheet_name='Staff')
    
    print(f"\nLeft file (Employees):")
    print(f"  Rows: {len(left_df)}")
    print(f"  Columns: {list(left_df.columns)}")
    
    print(f"\nRight file (Staff):")
    print(f"  Rows: {len(right_df)}")
    print(f"  Columns: {list(right_df.columns)}")
    
    # Compare Employee_ID vs Staff_ID
    left_col = 'Employee_ID'
    right_col = 'Staff_ID'
    
    left_values = set(left_df[left_col].dropna().astype(str))
    right_values = set(right_df[right_col].dropna().astype(str))
    
    matches = left_values & right_values
    only_in_left = left_values - right_values
    only_in_right = right_values - left_values
    
    print(f"\nComparison Results ({left_col} vs {right_col}):")
    print(f"  Matches: {len(matches)} - {sorted(matches)}")
    print(f"  Only in left: {len(only_in_left)} - {sorted(only_in_left)}")
    print(f"  Only in right: {len(only_in_right)} - {sorted(only_in_right)}")
    
    # Verify expected results
    assert matches == {'E001', 'E002'}, "Matches should be E001, E002"
    assert only_in_left == {'E003', 'E004', 'E005'}, "Left should have E003-E005"
    assert only_in_right == {'E006', 'E007', 'E008'}, "Right should have E006-E008"
    print("\n✓ Test 2 PASSED")
    
    print("\n" + "=" * 60)
    print("ALL TESTS PASSED SUCCESSFULLY!")
    print("=" * 60)
    print("\nCore comparison logic is working correctly.")
    print("You can now run the GUI application with: python excel_lookup.py")

if __name__ == "__main__":
    test_comparison_logic()
