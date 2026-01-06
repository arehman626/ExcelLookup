# GUI Layout Reference

This document describes the visual layout and organization of the Excel Lookup GUI application.

## Application Window

**Window Title**: Excel Lookup - Data Comparison Tool  
**Default Size**: 1200x700 pixels  
**Resizable**: Yes

---

## Layout Structure

```
┌─────────────────────────────────────────────────────────────────────────┐
│                   Excel Data Comparison Tool                            │
├─────────────────────────────────────────┬───────────────────────────────┤
│                                         │                               │
│  Left File (Source)                     │  Right File (Compare)        │
│  ┌───────────────────────────────────┐  │  ┌─────────────────────────┐ │
│  │ File: [No file selected]          │  │  │ File: [No file selected]│ │
│  │ [Browse...]                        │  │  │ [Browse...]             │ │
│  ├───────────────────────────────────┤  │  ├─────────────────────────┤ │
│  │ Sheet: [Dropdown▼]                │  │  │ Sheet: [Dropdown▼]      │ │
│  ├───────────────────────────────────┤  │  ├─────────────────────────┤ │
│  │ Info: Rows: X, Columns: Y         │  │  │ Info: Rows: X, Columns: Y│ │
│  └───────────────────────────────────┘  │  └─────────────────────────┘ │
│                                         │                               │
├─────────────────────────────────────────┴───────────────────────────────┤
│  Column Selection for Comparison                                        │
│  ┌─────────────────────────────────┬───────────────────────────────┐   │
│  │ Left Column:  [Dropdown▼]       │ Right Column: [Dropdown▼]     │   │
│  ├─────────────────────────────────┼───────────────────────────────┤   │
│  │ Additional Left Columns:        │ Additional Right Columns:     │   │
│  │ ┌─────────────────────────────┐ │ ┌─────────────────────────┐   │   │
│  │ │ [Column 1]                  │ │ │ [Column 1]              │   │   │
│  │ │ [Column 2]                  │ │ │ [Column 2]              │   │   │
│  │ │ [Column 3]                  │ │ │ [Column 3]              │   │   │
│  │ │ [Column 4]                  │ │ │ [Column 4]              │   │   │
│  │ └─────────────────────────────┘ │ └─────────────────────────┘   │   │
│  └─────────────────────────────────┴───────────────────────────────┘   │
├─────────────────────────────────────────────────────────────────────────┤
│              [Compare Data]  [Save Results]  [Reset]                    │
├─────────────────────────────────────────────────────────────────────────┤
│  Status: Ready to compare Excel files                                   │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Component Details

### 1. Title Bar
- **Location**: Top center
- **Font**: Arial, 16pt, Bold
- **Text**: "Excel Data Comparison Tool"

### 2. Left Panel - Source File Section
- **Label**: "Left File (Source)"
- **Components**:
  - File path label (shows selected filename or "No file selected")
  - Browse button - Opens file dialog for Excel file selection
  - Sheet dropdown - Lists all sheets in the selected Excel file
  - Info label - Shows row and column count (e.g., "Rows: 100, Columns: 5")
- **Color Coding**:
  - Gray text: No file selected
  - Black text: File selected
  - Blue text: Info label

### 3. Right Panel - Comparison File Section
- **Label**: "Right File (Compare)"
- **Components**: Same as Left Panel
- **Purpose**: Select the file to compare against the source

### 4. Column Selection Section
- **Label**: "Column Selection for Comparison"
- **Components**:
  
  #### Key Columns (Top Row)
  - **Left Column Dropdown**: Select the key/matching column from left file
  - **Right Column Dropdown**: Select the key/matching column from right file
  - These columns are used to match records between files
  
  #### Additional Columns (Bottom Section)
  - **Left Additional Listbox**: 
    - Multi-select listbox showing all columns from left file
    - Ctrl+Click to select multiple columns
    - Selected columns will be included in results with "Left_" prefix
  - **Right Additional Listbox**:
    - Multi-select listbox showing all columns from right file
    - Ctrl+Click to select multiple columns
    - Selected columns will be included in results with "Right_" prefix

### 5. Action Buttons (Center)
- **Compare Data**: 
  - Enabled when: Both files loaded, sheets selected, key columns selected
  - Action: Performs the comparison and shows summary dialog
- **Save Results**: 
  - Enabled when: Comparison has been performed
  - Action: Opens save dialog to export results to Excel with highlighting
- **Reset**: 
  - Always enabled
  - Action: Clears all selections and resets the application state

### 6. Status Bar (Bottom)
- **Location**: Bottom of window
- **Style**: Sunken, left-aligned text
- **Purpose**: Shows current operation status and messages
- **Example Messages**:
  - "Ready to compare Excel files"
  - "Left file loaded: employees.xlsx"
  - "Comparing data..."
  - "Comparison complete: 50 matches, 10 left only, 5 right only"
  - "Results saved: Comparison_file1_vs_file2_20240106_153045.xlsx"

---

## Color Scheme

### Application Colors
- **Background**: Light gray (default system color)
- **Frame borders**: Default system style
- **Labels**: Black text
- **Disabled elements**: Gray
- **Status info**: Blue
- **Error messages**: Red (in dialogs)

### Result File Colors
- **Matched records**: Light green background (90EE90)
- **Non-matched records**: Light red/pink background (FFB6C1)

---

## Dialog Boxes

### File Selection Dialog
- **Type**: Standard OS file picker
- **Filter**: Excel files (*.xlsx, *.xls)
- **Title**: 
  - "Select Left Excel File" (for left panel)
  - "Select Right Excel File" (for right panel)

### Save Results Dialog
- **Type**: Standard OS save file picker
- **Default Name**: Comparison_[left]_vs_[right]_[timestamp].xlsx
- **Example**: Comparison_employees_vs_payroll_20240106_153045.xlsx
- **Filter**: Excel files (*.xlsx)

### Comparison Summary Dialog
- **Type**: Information message box
- **Title**: "Comparison Complete"
- **Content**: Shows counts of matches and non-matches
- **Buttons**: OK

### Error Dialogs
- **Type**: Error message box
- **Examples**:
  - "Failed to load file"
  - "Please select columns to compare"
  - "Invalid columns selected"

---

## Interaction Flow

### Standard Workflow
```
1. Click "Browse..." (Left) → Select file
2. Select sheet from dropdown (Left)
3. Click "Browse..." (Right) → Select file
4. Select sheet from dropdown (Right)
5. Select key column (Left)
6. Select key column (Right)
7. [Optional] Select additional columns
8. Click "Compare Data" → View summary
9. Click "Save Results" → Choose location
10. Open saved Excel file to review results
```

### Alternative Workflow
```
1. Load both files
2. Use "Reset" to clear and try different sheets/columns
3. Repeat comparison with different parameters
```

---

## Keyboard Navigation

- **Tab**: Move between input fields
- **Ctrl+Click**: Multi-select in listboxes
- **Enter**: Activate focused button
- **Escape**: Close dialog boxes

---

## Responsiveness

- Window is resizable
- Panels expand proportionally
- Listboxes scroll when content exceeds visible area
- Long filenames are truncated with ellipsis if needed

---

## Accessibility Features

- Clear labels for all inputs
- Descriptive button text
- Status messages for all operations
- Error messages with specific guidance
- Consistent layout and navigation

---

## State Management

### Initial State
- All dropdowns and listboxes: Empty/Disabled
- Compare button: Disabled
- Save button: Disabled
- Status: "Ready to compare Excel files"

### After Loading Left File
- Left sheet dropdown: Enabled with options
- Left column selectors: Enabled after sheet selection

### After Loading Right File
- Right sheet dropdown: Enabled with options
- Right column selectors: Enabled after sheet selection

### After Both Sheets Selected
- Compare button: Enabled if key columns selected

### After Comparison
- Save button: Enabled
- Status shows comparison results

### After Reset
- Returns to Initial State

---

## Technical Implementation

### Framework
- **GUI Library**: tkinter (Python standard library)
- **Layout Manager**: grid
- **Widget Types**: 
  - Frame, LabelFrame (organization)
  - Label (static text)
  - Button (actions)
  - Combobox (dropdowns)
  - Listbox (multi-select lists)

### Data Flow
```
File Selection → Pandas → Sheet Loading → Column Population
     ↓
Column Selection → Comparison Logic → Result DataFrame
     ↓
Save Action → OpenPyXL → Excel with Formatting
```

---

## Tips for Users

1. **File Selection**: Files can be the same if comparing different sheets
2. **Column Names**: Don't need to match between files
3. **Multiple Columns**: Use Ctrl+Click in listboxes
4. **Performance**: Larger files take longer to process
5. **Results**: Open in Excel to see color highlighting
6. **Troubleshooting**: Check status bar for current operation

---

This layout provides a clear, intuitive interface for comparing Excel files with minimal learning curve.
