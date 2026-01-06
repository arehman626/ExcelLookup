"""
Excel Lookup Utility - Main Application
A comprehensive Excel comparison and reconciliation tool with GUI
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from pathlib import Path
import traceback


class ExcelLookupApp:
    """Main application class for Excel Lookup utility"""
    
    # Display limits for results
    MAX_DISPLAY_ITEMS = 100
    MAX_DISPLAY_MATCHES = 50
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Lookup & Reconciliation Tool")
        self.root.geometry("1200x800")
        
        # Data storage
        self.left_file_path = None
        self.right_file_path = None
        self.left_df = None
        self.right_df = None
        self.left_sheets = []
        self.right_sheets = []
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel Lookup & Reconciliation Tool", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Left Panel
        self.create_file_panel(main_frame, "left", 1, 0)
        
        # Right Panel
        self.create_file_panel(main_frame, "right", 1, 1)
        
        # Compare Button
        compare_frame = ttk.Frame(main_frame)
        compare_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        self.compare_btn = ttk.Button(compare_frame, text="Compare & Reconcile", 
                                      command=self.compare_data, state='disabled')
        self.compare_btn.pack()
        
        # Results Panel
        results_label = ttk.Label(main_frame, text="Reconciliation Results:", 
                                 font=('Arial', 12, 'bold'))
        results_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        self.results_text = scrolledtext.ScrolledText(main_frame, height=15, wrap=tk.WORD)
        self.results_text.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(4, weight=1)
        
        # Export Button
        export_frame = ttk.Frame(main_frame)
        export_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        self.export_btn = ttk.Button(export_frame, text="Export Results to Excel", 
                                     command=self.export_results, state='disabled')
        self.export_btn.pack()
        
    def create_file_panel(self, parent, side, row, col):
        """Create a file selection panel"""
        panel_frame = ttk.LabelFrame(parent, text=f"{side.capitalize()} File", padding="10")
        panel_frame.grid(row=row, column=col, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # File selection
        file_frame = ttk.Frame(panel_frame)
        file_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        file_label = ttk.Label(panel_frame, text="No file selected", wraplength=250)
        file_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        select_btn = ttk.Button(file_frame, text="Select Excel File", 
                               command=lambda: self.select_file(side, file_label))
        select_btn.pack()
        
        # Sheet selection
        sheet_label = ttk.Label(panel_frame, text="Select Sheet:")
        sheet_label.grid(row=2, column=0, sticky=tk.W, pady=(10, 5))
        
        sheet_combo = ttk.Combobox(panel_frame, state='disabled', width=30)
        sheet_combo.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        sheet_combo.bind('<<ComboboxSelected>>', 
                        lambda e: self.load_sheet(side, sheet_combo.get()))
        
        # Column selection
        col_label = ttk.Label(panel_frame, text="Select Compare Column:")
        col_label.grid(row=4, column=0, sticky=tk.W, pady=(10, 5))
        
        col_combo = ttk.Combobox(panel_frame, state='disabled', width=30)
        col_combo.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Preview
        preview_label = ttk.Label(panel_frame, text="Data Preview:")
        preview_label.grid(row=6, column=0, sticky=tk.W, pady=(10, 5))
        
        preview_text = scrolledtext.ScrolledText(panel_frame, height=10, width=40)
        preview_text.grid(row=7, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        panel_frame.rowconfigure(7, weight=1)
        
        # Store references
        if side == "left":
            self.left_file_label = file_label
            self.left_sheet_combo = sheet_combo
            self.left_col_combo = col_combo
            self.left_preview = preview_text
        else:
            self.right_file_label = file_label
            self.right_sheet_combo = sheet_combo
            self.right_col_combo = col_combo
            self.right_preview = preview_text
            
    def select_file(self, side, label):
        """Handle file selection"""
        file_path = filedialog.askopenfilename(
            title=f"Select {side.capitalize()} Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Read Excel file to get sheet names
                excel_file = pd.ExcelFile(file_path)
                sheets = excel_file.sheet_names
                
                if side == "left":
                    self.left_file_path = file_path
                    self.left_sheets = sheets
                    self.left_sheet_combo['values'] = sheets
                    self.left_sheet_combo['state'] = 'readonly'
                    label.config(text=Path(file_path).name)
                else:
                    self.right_file_path = file_path
                    self.right_sheets = sheets
                    self.right_sheet_combo['values'] = sheets
                    self.right_sheet_combo['state'] = 'readonly'
                    label.config(text=Path(file_path).name)
                    
                self.check_ready_to_compare()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")
                
    def load_sheet(self, side, sheet_name):
        """Load selected sheet and populate column dropdown"""
        if not sheet_name:
            return
            
        try:
            if side == "left":
                self.left_df = pd.read_excel(self.left_file_path, sheet_name=sheet_name)
                columns = list(self.left_df.columns)
                self.left_col_combo['values'] = columns
                self.left_col_combo['state'] = 'readonly'
                
                # Show preview
                preview = f"Sheet: {sheet_name}\n"
                preview += f"Rows: {len(self.left_df)}\n"
                preview += f"Columns: {len(self.left_df.columns)}\n\n"
                preview += "First 5 rows:\n"
                preview += self.left_df.head().to_string()
                
                self.left_preview.delete(1.0, tk.END)
                self.left_preview.insert(1.0, preview)
            else:
                self.right_df = pd.read_excel(self.right_file_path, sheet_name=sheet_name)
                columns = list(self.right_df.columns)
                self.right_col_combo['values'] = columns
                self.right_col_combo['state'] = 'readonly'
                
                # Show preview
                preview = f"Sheet: {sheet_name}\n"
                preview += f"Rows: {len(self.right_df)}\n"
                preview += f"Columns: {len(self.right_df.columns)}\n\n"
                preview += "First 5 rows:\n"
                preview += self.right_df.head().to_string()
                
                self.right_preview.delete(1.0, tk.END)
                self.right_preview.insert(1.0, preview)
                
            self.check_ready_to_compare()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {str(e)}")
            
    def check_ready_to_compare(self):
        """Check if all required selections are made"""
        if (self.left_df is not None and self.right_df is not None):
            self.compare_btn['state'] = 'normal'
        else:
            self.compare_btn['state'] = 'disabled'
            
    def compare_data(self):
        """Compare the two datasets and generate reconciliation report"""
        left_col = self.left_col_combo.get()
        right_col = self.right_col_combo.get()
        
        if not left_col or not right_col:
            messagebox.showwarning("Warning", "Please select columns to compare from both files")
            return
            
        try:
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, "Processing comparison...\n\n")
            self.root.update()
            
            # Perform comparison
            results = self.perform_reconciliation(left_col, right_col)
            
            # Display results
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, results)
            
            self.export_btn['state'] = 'normal'
            
        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {str(e)}\n\n{traceback.format_exc()}")
            
    def perform_reconciliation(self, left_col, right_col):
        """Perform the actual reconciliation logic"""
        results = []
        results.append("=" * 80)
        results.append("EXCEL LOOKUP & RECONCILIATION REPORT")
        results.append("=" * 80)
        results.append("")
        
        # Basic statistics
        results.append(f"Left File: {Path(self.left_file_path).name}")
        results.append(f"  - Rows: {len(self.left_df)}")
        results.append(f"  - Compare Column: '{left_col}'")
        results.append("")
        
        results.append(f"Right File: {Path(self.right_file_path).name}")
        results.append(f"  - Rows: {len(self.right_df)}")
        results.append(f"  - Compare Column: '{right_col}'")
        results.append("")
        
        # Convert columns to sets for comparison
        left_values = set(self.left_df[left_col].dropna().astype(str))
        right_values = set(self.right_df[right_col].dropna().astype(str))
        
        # Find matches and differences
        matches = left_values & right_values
        only_in_left = left_values - right_values
        only_in_right = right_values - left_values
        
        results.append("-" * 80)
        results.append("SUMMARY")
        results.append("-" * 80)
        results.append(f"Total unique values in left:  {len(left_values)}")
        results.append(f"Total unique values in right: {len(right_values)}")
        results.append(f"Matching values:              {len(matches)} ({len(matches)/max(len(left_values), 1)*100:.1f}%)")
        results.append(f"Only in left:                 {len(only_in_left)}")
        results.append(f"Only in right:                {len(only_in_right)}")
        results.append("")
        
        # Detailed results
        if only_in_left:
            results.append("-" * 80)
            results.append(f"VALUES ONLY IN LEFT FILE ({len(only_in_left)} items)")
            results.append("-" * 80)
            for i, value in enumerate(sorted(only_in_left)[:self.MAX_DISPLAY_ITEMS], 1):
                results.append(f"{i}. {value}")
            if len(only_in_left) > self.MAX_DISPLAY_ITEMS:
                results.append(f"... and {len(only_in_left) - self.MAX_DISPLAY_ITEMS} more")
            results.append("")
            
        if only_in_right:
            results.append("-" * 80)
            results.append(f"VALUES ONLY IN RIGHT FILE ({len(only_in_right)} items)")
            results.append("-" * 80)
            for i, value in enumerate(sorted(only_in_right)[:self.MAX_DISPLAY_ITEMS], 1):
                results.append(f"{i}. {value}")
            if len(only_in_right) > self.MAX_DISPLAY_ITEMS:
                results.append(f"... and {len(only_in_right) - self.MAX_DISPLAY_ITEMS} more")
            results.append("")
            
        if matches:
            results.append("-" * 80)
            results.append(f"MATCHING VALUES (showing first {self.MAX_DISPLAY_MATCHES} of {len(matches)})")
            results.append("-" * 80)
            for i, value in enumerate(sorted(matches)[:self.MAX_DISPLAY_MATCHES], 1):
                results.append(f"{i}. {value}")
            if len(matches) > self.MAX_DISPLAY_MATCHES:
                results.append(f"... and {len(matches) - self.MAX_DISPLAY_MATCHES} more")
            results.append("")
        
        results.append("=" * 80)
        results.append("END OF REPORT")
        results.append("=" * 80)
        
        # Store results for export
        self.reconciliation_results = {
            'matches': matches,
            'only_in_left': only_in_left,
            'only_in_right': only_in_right,
            'left_col': left_col,
            'right_col': right_col
        }
        
        return "\n".join(results)
        
    def export_results(self):
        """Export reconciliation results to Excel"""
        if not hasattr(self, 'reconciliation_results'):
            messagebox.showwarning("Warning", "No results to export")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="Save Reconciliation Results",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                results = self.reconciliation_results
                
                # Create DataFrames for each category
                matches_df = pd.DataFrame(sorted(results['matches']), 
                                         columns=[f"Matching Values"])
                
                only_left_df = pd.DataFrame(sorted(results['only_in_left']), 
                                           columns=[f"Only in Left ({results['left_col']})"])
                
                only_right_df = pd.DataFrame(sorted(results['only_in_right']), 
                                            columns=[f"Only in Right ({results['right_col']})"])
                
                # Write to Excel with multiple sheets
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    matches_df.to_excel(writer, sheet_name='Matches', index=False)
                    only_left_df.to_excel(writer, sheet_name='Only in Left', index=False)
                    only_right_df.to_excel(writer, sheet_name='Only in Right', index=False)
                    
                    # Create summary sheet
                    summary_data = {
                        'Metric': ['Total Matches', 'Only in Left', 'Only in Right', 
                                  'Left Column', 'Right Column'],
                        'Value': [len(results['matches']), 
                                 len(results['only_in_left']),
                                 len(results['only_in_right']),
                                 results['left_col'],
                                 results['right_col']]
                    }
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                messagebox.showinfo("Success", f"Results exported to:\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export results: {str(e)}")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ExcelLookupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
