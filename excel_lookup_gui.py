"""
Excel Lookup GUI Application
A comprehensive tool to compare data between two Excel files with visual highlighting.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
import os


class ExcelLookupGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Lookup - Data Comparison Tool")
        self.root.geometry("1200x700")
        
        # Data storage
        self.left_file_path = None
        self.right_file_path = None
        self.left_df = None
        self.right_df = None
        self.left_sheets = []
        self.right_sheets = []
        self.left_columns = []
        self.right_columns = []
        
        self.create_widgets()
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel Data Comparison Tool", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Left Panel
        self.create_left_panel(main_frame)
        
        # Right Panel
        self.create_right_panel(main_frame)
        
        # Column Selection Section
        self.create_column_selection(main_frame)
        
        # Action Buttons
        self.create_action_buttons(main_frame)
        
        # Status Bar
        self.status_var = tk.StringVar(value="Ready to compare Excel files")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                relief=tk.SUNKEN, anchor=tk.W)
        status_label.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
    
    def create_left_panel(self, parent):
        """Create left panel for source file selection"""
        left_frame = ttk.LabelFrame(parent, text="Left File (Source)", padding="10")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # File selection
        ttk.Label(left_frame, text="File:").grid(row=0, column=0, sticky=tk.W)
        self.left_file_label = ttk.Label(left_frame, text="No file selected", 
                                         foreground="gray")
        self.left_file_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        left_browse_btn = ttk.Button(left_frame, text="Browse...", 
                                     command=self.browse_left_file)
        left_browse_btn.grid(row=0, column=2, padx=5)
        
        # Sheet selection
        ttk.Label(left_frame, text="Sheet:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.left_sheet_combo = ttk.Combobox(left_frame, state="disabled", width=30)
        self.left_sheet_combo.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5)
        self.left_sheet_combo.bind("<<ComboboxSelected>>", self.on_left_sheet_selected)
        
        # File info
        self.left_info_label = ttk.Label(left_frame, text="", foreground="blue")
        self.left_info_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
    
    def create_right_panel(self, parent):
        """Create right panel for comparison file selection"""
        right_frame = ttk.LabelFrame(parent, text="Right File (Compare)", padding="10")
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # File selection
        ttk.Label(right_frame, text="File:").grid(row=0, column=0, sticky=tk.W)
        self.right_file_label = ttk.Label(right_frame, text="No file selected", 
                                          foreground="gray")
        self.right_file_label.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        right_browse_btn = ttk.Button(right_frame, text="Browse...", 
                                      command=self.browse_right_file)
        right_browse_btn.grid(row=0, column=2, padx=5)
        
        # Sheet selection
        ttk.Label(right_frame, text="Sheet:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.right_sheet_combo = ttk.Combobox(right_frame, state="disabled", width=30)
        self.right_sheet_combo.grid(row=1, column=1, columnspan=2, sticky=tk.W, padx=5)
        self.right_sheet_combo.bind("<<ComboboxSelected>>", self.on_right_sheet_selected)
        
        # File info
        self.right_info_label = ttk.Label(right_frame, text="", foreground="blue")
        self.right_info_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
    
    def create_column_selection(self, parent):
        """Create column selection section"""
        column_frame = ttk.LabelFrame(parent, text="Column Selection for Comparison", 
                                      padding="10")
        column_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                         padx=5, pady=5)
        
        # Left column selection
        ttk.Label(column_frame, text="Left Column:").grid(row=0, column=0, sticky=tk.W)
        self.left_column_combo = ttk.Combobox(column_frame, state="disabled", width=30)
        self.left_column_combo.grid(row=0, column=1, padx=5, pady=5)
        
        # Right column selection
        ttk.Label(column_frame, text="Right Column:").grid(row=0, column=2, sticky=tk.W, 
                                                           padx=(20, 0))
        self.right_column_combo = ttk.Combobox(column_frame, state="disabled", width=30)
        self.right_column_combo.grid(row=0, column=3, padx=5, pady=5)
        
        # Additional columns to include
        ttk.Label(column_frame, text="Additional Left Columns (optional):").grid(
            row=1, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        self.left_additional_listbox = tk.Listbox(column_frame, height=4, 
                                                   selectmode=tk.MULTIPLE, width=35)
        self.left_additional_listbox.grid(row=2, column=0, columnspan=2, 
                                         sticky=(tk.W, tk.E), padx=5, pady=5)
        
        ttk.Label(column_frame, text="Additional Right Columns (optional):").grid(
            row=1, column=2, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        self.right_additional_listbox = tk.Listbox(column_frame, height=4, 
                                                    selectmode=tk.MULTIPLE, width=35)
        self.right_additional_listbox.grid(row=2, column=2, columnspan=2, 
                                          sticky=(tk.W, tk.E), padx=5, pady=5)
    
    def create_action_buttons(self, parent):
        """Create action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)
        
        self.compare_btn = ttk.Button(button_frame, text="Compare Data", 
                                      command=self.compare_data, state="disabled")
        self.compare_btn.grid(row=0, column=0, padx=10)
        
        self.save_btn = ttk.Button(button_frame, text="Save Results", 
                                   command=self.save_results, state="disabled")
        self.save_btn.grid(row=0, column=1, padx=10)
        
        reset_btn = ttk.Button(button_frame, text="Reset", command=self.reset_all)
        reset_btn.grid(row=0, column=2, padx=10)
    
    def browse_left_file(self):
        """Browse and select left file"""
        file_path = filedialog.askopenfilename(
            title="Select Left Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.left_file_path = file_path
                filename = os.path.basename(file_path)
                self.left_file_label.config(text=filename, foreground="black")
                
                # Load sheets
                excel_file = pd.ExcelFile(file_path)
                self.left_sheets = excel_file.sheet_names
                
                self.left_sheet_combo.config(state="readonly")
                self.left_sheet_combo['values'] = self.left_sheets
                if self.left_sheets:
                    self.left_sheet_combo.current(0)
                    self.on_left_sheet_selected(None)
                
                self.status_var.set(f"Left file loaded: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load left file:\n{str(e)}")
                self.status_var.set("Error loading left file")
    
    def browse_right_file(self):
        """Browse and select right file"""
        file_path = filedialog.askopenfilename(
            title="Select Right Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.right_file_path = file_path
                filename = os.path.basename(file_path)
                self.right_file_label.config(text=filename, foreground="black")
                
                # Load sheets
                excel_file = pd.ExcelFile(file_path)
                self.right_sheets = excel_file.sheet_names
                
                self.right_sheet_combo.config(state="readonly")
                self.right_sheet_combo['values'] = self.right_sheets
                if self.right_sheets:
                    self.right_sheet_combo.current(0)
                    self.on_right_sheet_selected(None)
                
                self.status_var.set(f"Right file loaded: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load right file:\n{str(e)}")
                self.status_var.set("Error loading right file")
    
    def on_left_sheet_selected(self, event):
        """Handle left sheet selection"""
        if not self.left_file_path or not self.left_sheet_combo.get():
            return
        
        try:
            sheet_name = self.left_sheet_combo.get()
            self.left_df = pd.read_excel(self.left_file_path, sheet_name=sheet_name)
            self.left_columns = list(self.left_df.columns)
            
            # Update column combo
            self.left_column_combo.config(state="readonly")
            self.left_column_combo['values'] = self.left_columns
            if self.left_columns:
                self.left_column_combo.current(0)
            
            # Update additional columns listbox
            self.left_additional_listbox.delete(0, tk.END)
            for col in self.left_columns:
                self.left_additional_listbox.insert(tk.END, col)
            
            self.left_info_label.config(
                text=f"Rows: {len(self.left_df)}, Columns: {len(self.left_columns)}")
            
            self.check_ready_to_compare()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load left sheet:\n{str(e)}")
    
    def on_right_sheet_selected(self, event):
        """Handle right sheet selection"""
        if not self.right_file_path or not self.right_sheet_combo.get():
            return
        
        try:
            sheet_name = self.right_sheet_combo.get()
            self.right_df = pd.read_excel(self.right_file_path, sheet_name=sheet_name)
            self.right_columns = list(self.right_df.columns)
            
            # Update column combo
            self.right_column_combo.config(state="readonly")
            self.right_column_combo['values'] = self.right_columns
            if self.right_columns:
                self.right_column_combo.current(0)
            
            # Update additional columns listbox
            self.right_additional_listbox.delete(0, tk.END)
            for col in self.right_columns:
                self.right_additional_listbox.insert(tk.END, col)
            
            self.right_info_label.config(
                text=f"Rows: {len(self.right_df)}, Columns: {len(self.right_columns)}")
            
            self.check_ready_to_compare()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load right sheet:\n{str(e)}")
    
    def check_ready_to_compare(self):
        """Check if all requirements are met to enable compare button"""
        if (self.left_df is not None and self.right_df is not None and 
            self.left_column_combo.get() and self.right_column_combo.get()):
            self.compare_btn.config(state="normal")
        else:
            self.compare_btn.config(state="disabled")
    
    def compare_data(self):
        """Compare data between left and right files"""
        try:
            left_key_col = self.left_column_combo.get()
            right_key_col = self.right_column_combo.get()
            
            if not left_key_col or not right_key_col:
                messagebox.showwarning("Warning", "Please select columns to compare")
                return
            
            self.status_var.set("Comparing data...")
            self.root.update()
            
            # Get additional columns
            left_additional = [self.left_additional_listbox.get(i) 
                              for i in self.left_additional_listbox.curselection()]
            right_additional = [self.right_additional_listbox.get(i) 
                               for i in self.right_additional_listbox.curselection()]
            
            # Prepare columns to include
            left_cols = [left_key_col] + [col for col in left_additional 
                                          if col != left_key_col]
            right_cols = [right_key_col] + [col for col in right_additional 
                                            if col != right_key_col]
            
            # Validate columns exist
            invalid_left = [col for col in left_cols if col not in self.left_df.columns]
            invalid_right = [col for col in right_cols if col not in self.right_df.columns]
            
            if invalid_left:
                messagebox.showerror("Error", 
                    f"Invalid left columns: {', '.join(invalid_left)}")
                return
            
            if invalid_right:
                messagebox.showerror("Error", 
                    f"Invalid right columns: {', '.join(invalid_right)}")
                return
            
            # Create subsets
            left_subset = self.left_df[left_cols].copy()
            right_subset = self.right_df[right_cols].copy()
            
            # Rename key columns for merging
            left_subset = left_subset.rename(columns={left_key_col: 'key_column'})
            right_subset = right_subset.rename(columns={right_key_col: 'key_column'})
            
            # Add suffixes to other columns
            left_subset.columns = ['key_column'] + [f'Left_{col}' 
                                                     for col in left_subset.columns[1:]]
            right_subset.columns = ['key_column'] + [f'Right_{col}' 
                                                      for col in right_subset.columns[1:]]
            
            # Perform merge (outer join to capture all data)
            self.result_df = pd.merge(left_subset, right_subset, on='key_column', 
                                     how='outer', indicator=True)
            
            # Rename merge indicator
            self.result_df = self.result_df.rename(columns={'_merge': 'Match_Status'})
            
            # Map status
            status_map = {
                'both': 'Match',
                'left_only': 'Left Only',
                'right_only': 'Right Only'
            }
            self.result_df['Match_Status'] = self.result_df['Match_Status'].map(status_map)
            
            # Rename key_column back
            self.result_df = self.result_df.rename(
                columns={'key_column': f'Key_({left_key_col}={right_key_col})'})
            
            matches = len(self.result_df[self.result_df['Match_Status'] == 'Match'])
            left_only = len(self.result_df[self.result_df['Match_Status'] == 'Left Only'])
            right_only = len(self.result_df[self.result_df['Match_Status'] == 'Right Only'])
            
            self.status_var.set(
                f"Comparison complete: {matches} matches, {left_only} left only, "
                f"{right_only} right only")
            
            self.save_btn.config(state="normal")
            
            messagebox.showinfo("Comparison Complete", 
                              f"Results:\n"
                              f"Matches: {matches}\n"
                              f"Left Only: {left_only}\n"
                              f"Right Only: {right_only}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to compare data:\n{str(e)}")
            self.status_var.set("Error during comparison")
    
    def save_results(self):
        """Save comparison results with highlighting"""
        if not hasattr(self, 'result_df') or self.result_df is None:
            messagebox.showwarning("Warning", "No results to save. Please compare data first.")
            return
        
        try:
            # Generate default filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            left_name = os.path.splitext(os.path.basename(self.left_file_path))[0]
            right_name = os.path.splitext(os.path.basename(self.right_file_path))[0]
            default_name = f"Comparison_{left_name}_vs_{right_name}_{timestamp}.xlsx"
            
            # Ask user for save location
            save_path = filedialog.asksaveasfilename(
                title="Save Comparison Results",
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not save_path:
                return
            
            self.status_var.set("Saving results...")
            self.root.update()
            
            # Save to Excel with formatting
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                self.result_df.to_excel(writer, sheet_name='Comparison_Results', 
                                       index=False)
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Comparison_Results']
                
                # Define fill colors
                green_fill = PatternFill(start_color='90EE90', end_color='90EE90', 
                                        fill_type='solid')
                red_fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', 
                                      fill_type='solid')
                
                # Pre-extract match status for performance
                match_statuses = self.result_df['Match_Status'].tolist()
                
                # Apply formatting
                for row_idx, row in enumerate(worksheet.iter_rows(min_row=2, 
                                                                  max_row=len(self.result_df) + 1), 
                                             start=2):
                    status = match_statuses[row_idx - 2]
                    
                    for cell in row:
                        if status == 'Match':
                            cell.fill = green_fill
                        else:
                            cell.fill = red_fill
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except (TypeError, AttributeError):
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            self.status_var.set(f"Results saved: {os.path.basename(save_path)}")
            messagebox.showinfo("Success", f"Results saved successfully to:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save results:\n{str(e)}")
            self.status_var.set("Error saving results")
    
    def reset_all(self):
        """Reset all selections and data"""
        self.left_file_path = None
        self.right_file_path = None
        self.left_df = None
        self.right_df = None
        self.left_sheets = []
        self.right_sheets = []
        self.left_columns = []
        self.right_columns = []
        
        self.left_file_label.config(text="No file selected", foreground="gray")
        self.right_file_label.config(text="No file selected", foreground="gray")
        
        self.left_sheet_combo.set('')
        self.left_sheet_combo.config(state="disabled")
        self.right_sheet_combo.set('')
        self.right_sheet_combo.config(state="disabled")
        
        self.left_column_combo.set('')
        self.left_column_combo.config(state="disabled")
        self.right_column_combo.set('')
        self.right_column_combo.config(state="disabled")
        
        self.left_additional_listbox.delete(0, tk.END)
        self.right_additional_listbox.delete(0, tk.END)
        
        self.left_info_label.config(text="")
        self.right_info_label.config(text="")
        
        self.compare_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        
        if hasattr(self, 'result_df'):
            self.result_df = None
        
        self.status_var.set("Ready to compare Excel files")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ExcelLookupGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
