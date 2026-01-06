"""
Create sample Excel files for testing the Excel Lookup utility
"""

import pandas as pd
from pathlib import Path

# Create test data directory
test_dir = Path("test_data")
test_dir.mkdir(exist_ok=True)

# Sample data for left file
left_data_sheet1 = {
    'Product_ID': ['P001', 'P002', 'P003', 'P004', 'P005', 'P006', 'P007'],
    'Product_Name': ['Laptop', 'Mouse', 'Keyboard', 'Monitor', 'Webcam', 'Headset', 'Cable'],
    'Price': [1200, 25, 75, 350, 80, 120, 15],
    'Category': ['Electronics', 'Accessories', 'Accessories', 'Electronics', 'Accessories', 'Accessories', 'Accessories']
}

left_data_sheet2 = {
    'Employee_ID': ['E001', 'E002', 'E003', 'E004', 'E005'],
    'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Williams', 'Charlie Brown'],
    'Department': ['IT', 'HR', 'IT', 'Finance', 'Marketing'],
    'Salary': [75000, 60000, 72000, 68000, 55000]
}

# Sample data for right file (with some overlap and differences)
right_data_sheet1 = {
    'Item_Code': ['P001', 'P002', 'P003', 'P008', 'P009', 'P010'],
    'Item_Name': ['Laptop', 'Mouse', 'Keyboard', 'Speaker', 'Microphone', 'Desk Lamp'],
    'Cost': [1200, 25, 75, 95, 60, 40],
    'Type': ['Electronics', 'Accessories', 'Accessories', 'Accessories', 'Accessories', 'Furniture']
}

right_data_sheet2 = {
    'Staff_ID': ['E001', 'E002', 'E006', 'E007', 'E008'],
    'Full_Name': ['John Doe', 'Jane Smith', 'David Lee', 'Emma Wilson', 'Frank Miller'],
    'Division': ['IT', 'HR', 'Sales', 'IT', 'Operations'],
    'Pay': [75000, 60000, 58000, 70000, 62000]
}

# Create left Excel file
with pd.ExcelWriter(test_dir / "sample_left.xlsx", engine='openpyxl') as writer:
    pd.DataFrame(left_data_sheet1).to_excel(writer, sheet_name='Products', index=False)
    pd.DataFrame(left_data_sheet2).to_excel(writer, sheet_name='Employees', index=False)

# Create right Excel file
with pd.ExcelWriter(test_dir / "sample_right.xlsx", engine='openpyxl') as writer:
    pd.DataFrame(right_data_sheet1).to_excel(writer, sheet_name='Inventory', index=False)
    pd.DataFrame(right_data_sheet2).to_excel(writer, sheet_name='Staff', index=False)

print("Sample Excel files created successfully!")
print(f"\nTest files created in: {test_dir.absolute()}")
print("- sample_left.xlsx (2 sheets: Products, Employees)")
print("- sample_right.xlsx (2 sheets: Inventory, Staff)")
print("\nYou can use these files to test the Excel Lookup utility.")
print("\nExpected comparison results:")
print("- Products vs Inventory (Product_ID vs Item_Code):")
print("  * Matches: P001, P002, P003")
print("  * Only in left: P004, P005, P006, P007")
print("  * Only in right: P008, P009, P010")
print("\n- Employees vs Staff (Employee_ID vs Staff_ID):")
print("  * Matches: E001, E002")
print("  * Only in left: E003, E004, E005")
print("  * Only in right: E006, E007, E008")
