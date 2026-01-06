# Examples and Demonstrations

This document provides step-by-step examples of using the Excel Lookup tool.

## Example 1: Basic Employee Comparison

### Scenario
You have two Excel files:
- `employees_hr.xlsx`: HR system with employee details
- `employees_payroll.xlsx`: Payroll system with salary information

You want to find which employees are in one system but not the other.

### Sample Data

**employees_hr.xlsx (HR System)**
| ID  | Name    | Department | Hire_Date  |
|-----|---------|------------|------------|
| 101 | Alice   | Sales      | 2020-01-15 |
| 102 | Bob     | IT         | 2019-03-22 |
| 103 | Charlie | HR         | 2021-06-10 |
| 104 | David   | Sales      | 2020-11-05 |

**employees_payroll.xlsx (Payroll System)**
| Emp_ID | Full_Name | Annual_Salary | Status |
|--------|-----------|---------------|--------|
| 101    | Alice     | 50000         | Active |
| 102    | Bob       | 60000         | Active |
| 105    | Eve       | 55000         | Active |

### Steps

1. **Load Files**:
   - Left File: `employees_hr.xlsx`
   - Right File: `employees_payroll.xlsx`

2. **Select Sheets**:
   - Left Sheet: "Employees" (default)
   - Right Sheet: "Employees" (default)

3. **Choose Comparison Columns**:
   - Left Column: `ID`
   - Right Column: `Emp_ID`

4. **Select Additional Columns**:
   - Left: `Name`, `Department`, `Hire_Date`
   - Right: `Full_Name`, `Annual_Salary`, `Status`

5. **Compare**

### Expected Results

The comparison will produce 5 rows:

| Key_(ID=Emp_ID) | Left_Name | Left_Department | Left_Hire_Date | Right_Full_Name | Right_Annual_Salary | Right_Status | Match_Status |
|-----------------|-----------|-----------------|----------------|-----------------|---------------------|--------------|--------------|
| 101             | Alice     | Sales           | 2020-01-15     | Alice           | 50000               | Active       | Match ✓      |
| 102             | Bob       | IT              | 2019-03-22     | Bob             | 60000               | Active       | Match ✓      |
| 103             | Charlie   | HR              | 2021-06-10     |                 |                     |              | Left Only ✗  |
| 104             | David     | Sales           | 2020-11-05     |                 |                     |              | Left Only ✗  |
| 105             |           |                 |                | Eve             | 55000               | Active       | Right Only ✗ |

**Interpretation**:
- ✓ Green (Match): Alice and Bob are in both systems
- ✗ Red (Left Only): Charlie and David are in HR but not in Payroll
- ✗ Red (Right Only): Eve is in Payroll but not in HR

**Action Items**:
- Add Charlie and David to payroll system
- Investigate why Eve is in payroll without HR record

---

## Example 2: Inventory Reconciliation

### Scenario
Compare expected inventory vs. actual count.

### Sample Data

**inventory_expected.xlsx**
| SKU   | Product_Name     | Expected_Qty | Location |
|-------|------------------|--------------|----------|
| A001  | Widget A         | 100          | Warehouse|
| A002  | Widget B         | 150          | Store    |
| A003  | Widget C         | 75           | Warehouse|

**inventory_actual.xlsx**
| SKU   | Description      | Counted_Qty | Last_Count_Date |
|-------|------------------|-------------|-----------------|
| A001  | Widget A         | 98          | 2024-01-05      |
| A002  | Widget B         | 150         | 2024-01-05      |
| A004  | Widget D         | 25          | 2024-01-05      |

### Steps

1. **Load Files**:
   - Left: `inventory_expected.xlsx`
   - Right: `inventory_actual.xlsx`

2. **Choose Comparison Columns**:
   - Left: `SKU`
   - Right: `SKU`

3. **Select Additional Columns**:
   - Left: `Product_Name`, `Expected_Qty`, `Location`
   - Right: `Description`, `Counted_Qty`, `Last_Count_Date`

### Results Interpretation

| SKU  | Match_Status | Interpretation                                |
|------|--------------|-----------------------------------------------|
| A001 | Match ✓      | Found in both; compare quantities (100 vs 98) |
| A002 | Match ✓      | Found in both; quantities match (150 = 150)   |
| A003 | Left Only ✗  | Expected but not counted - missing item       |
| A004 | Right Only ✗ | Counted but not in expected - extra item      |

**Action Items**:
- A001: Investigate 2-unit discrepancy
- A003: Search for missing Widget C (75 units)
- A004: Identify source of unexpected Widget D

---

## Example 3: Financial Transaction Reconciliation

### Scenario
Compare bank statement with accounting records.

### Sample Data

**bank_statement.xlsx**
| Transaction_ID | Date       | Amount   | Description      |
|----------------|------------|----------|------------------|
| TXN001         | 2024-01-02 | 1500.00  | Customer Payment |
| TXN002         | 2024-01-03 | -500.00  | Vendor Payment   |
| TXN003         | 2024-01-04 | 2000.00  | Customer Payment |
| TXN005         | 2024-01-06 | -150.00  | Bank Fees        |

**accounting_records.xlsx**
| Ref_Number | Entry_Date | Debit   | Credit  | Account |
|------------|------------|---------|---------|---------|
| TXN001     | 2024-01-02 |         | 1500.00 | AR      |
| TXN002     | 2024-01-03 | 500.00  |         | AP      |
| TXN004     | 2024-01-05 | 300.00  |         | Expense |

### Steps

1. **Choose Comparison Columns**:
   - Left: `Transaction_ID`
   - Right: `Ref_Number`

2. **Include All Other Columns**

### Results Interpretation

| ID     | Match_Status | Issue                                    |
|--------|--------------|------------------------------------------|
| TXN001 | Match ✓      | Reconciled                               |
| TXN002 | Match ✓      | Reconciled                               |
| TXN003 | Left Only ✗  | In bank but not recorded - missing entry |
| TXN004 | Right Only ✗ | Recorded but not in bank - not cleared   |
| TXN005 | Left Only ✗  | In bank but not recorded - bank fees     |

---

## Example 4: Same File, Different Sheets

### Scenario
Compare two versions of data in the same file (e.g., "Current" vs "Previous" sheets).

### Steps

1. **Load Same File in Both Panels**:
   - Left File: `quarterly_data.xlsx`
   - Right File: `quarterly_data.xlsx`

2. **Select Different Sheets**:
   - Left Sheet: "Q1_2024"
   - Right Sheet: "Q4_2023"

3. **Compare** to see changes between quarters

---

## Tips for Interpreting Results

### Understanding Color Coding

1. **Green (Match)**:
   - Key value exists in both files
   - Check if other columns match
   - Potential for data comparison

2. **Red (Left Only)**:
   - Record only in source/left file
   - May indicate:
     - New items not yet in right file
     - Deletions from right file
     - Data migration issues

3. **Red (Right Only)**:
   - Record only in comparison/right file
   - May indicate:
     - New items in right file
     - Deletions from left file
     - Extra or duplicate entries

### Filtering in Excel

After saving results:

1. Apply AutoFilter (Data → Filter)
2. Filter by `Match_Status`:
   - Show only "Left Only" items
   - Show only "Right Only" items
   - Show only "Match" items
3. Sort by key column or any other column

### Creating Reports

Use the Excel results to create reports:

1. **Summary Table**:
   - Count of Matches
   - Count of Left Only
   - Count of Right Only

2. **Pivot Tables**:
   - Group by department/category
   - Analyze patterns in mismatches

3. **Charts**:
   - Pie chart of match distribution
   - Bar chart by category

---

## Common Patterns

### Pattern 1: One-to-Many Relationships

If a key appears multiple times:
- Use combined keys (concatenate multiple columns)
- Or handle duplicates before comparison

### Pattern 2: Case Sensitivity

Keys are case-sensitive:
- "ABC" ≠ "abc"
- Clean data to consistent case before comparison

### Pattern 3: Leading/Trailing Spaces

- " 123" ≠ "123"
- Use Excel TRIM function on data before comparison

### Pattern 4: Numeric vs Text

- Ensure key columns have the same data type
- "001" (text) ≠ 1 (numeric)

---

## Performance Examples

### Small Files (< 1,000 rows)
- Load time: < 1 second
- Compare time: < 2 seconds
- Save time: < 2 seconds

### Medium Files (1,000 - 10,000 rows)
- Load time: 1-3 seconds
- Compare time: 2-5 seconds
- Save time: 3-8 seconds

### Large Files (10,000 - 100,000 rows)
- Load time: 3-10 seconds
- Compare time: 5-15 seconds
- Save time: 10-30 seconds

### Very Large Files (> 100,000 rows)
- Consider filtering data first
- Use 64-bit Python with more RAM
- Expected processing time: 30 seconds - several minutes

---

## Automation Possibilities

While the GUI is interactive, the underlying logic can be automated:

```python
# See test_excel_lookup.py for automation example
import pandas as pd
from excel_lookup_gui import compare_and_save

# Could be extended for batch processing
```

For batch automation needs, consider modifying the code to accept command-line arguments or configuration files.
