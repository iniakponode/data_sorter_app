#!/usr/bin/env python3
"""
Data Sorter Application
A Tkinter-based desktop application that parses structured text records,
groups them by CO-OP NAME, and exports to a multi-sheet Excel file.
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from collections import defaultdict


class DataSorterApp:
    """Main application class for the Data Sorter."""
    
    def __init__(self, root):
        """Initialize the application UI."""
        self.root = root
        self.root.title("Data Sorter Application")
        self.root.geometry("800x600")
        
        # Create UI elements
        self.create_widgets()
    
    def create_widgets(self):
        """Create and layout UI widgets."""
        # Instructions label
        instructions = tk.Label(
            self.root,
            text="Paste records in KEY: VALUE format (separated by blank lines):",
            font=("Arial", 10),
            pady=10
        )
        instructions.pack()
        
        # Text area for input
        self.text_area = scrolledtext.ScrolledText(
            self.root,
            width=90,
            height=30,
            font=("Courier", 9),
            wrap=tk.WORD
        )
        self.text_area.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        # Process button
        self.process_button = tk.Button(
            self.root,
            text="Process and Export to Excel",
            command=self.process_data,
            font=("Arial", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10
        )
        self.process_button.pack(pady=10)
    
    def parse_records(self, text):
        """
        Parse the input text into records.
        
        Args:
            text (str): Input text with KEY: VALUE format records
            
        Returns:
            list: List of dictionaries, each representing a record
        """
        records = []
        current_record = {}
        
        lines = text.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Blank line indicates end of record
            if not line:
                if current_record:
                    records.append(current_record)
                    current_record = {}
            else:
                # Parse KEY: VALUE format
                if ':' in line:
                    parts = line.split(':', 1)
                    key = parts[0].strip()
                    value = parts[1].strip() if len(parts) > 1 else ""
                    current_record[key] = value
        
        # Don't forget the last record if text doesn't end with blank line
        if current_record:
            records.append(current_record)
        
        return records
    
    def group_by_coop_name(self, records):
        """
        Group records by CO-OP NAME field.
        
        Args:
            records (list): List of record dictionaries
            
        Returns:
            dict: Dictionary with CO-OP NAME as keys and list of records as values
        """
        grouped = defaultdict(list)
        
        for record in records:
            coop_name = record.get('CO-OP NAME', 'Unknown')
            grouped[coop_name].append(record)
        
        return grouped
    
    def create_excel_file(self, grouped_data, filepath):
        """
        Create an Excel file with separate sheets for each CO-OP NAME.
        
        Args:
            grouped_data (dict): Grouped records by CO-OP NAME
            filepath (str): Path where Excel file should be saved
        """
        workbook = Workbook()
        # Remove default sheet
        if 'Sheet' in workbook.sheetnames:
            workbook.remove(workbook['Sheet'])
        
        for coop_name, records in grouped_data.items():
            if not records:
                continue
            
            # Create sheet name (Excel has 31 char limit)
            sheet_name = coop_name[:31] if coop_name else "Unknown"
            
            # Ensure unique sheet name
            original_name = sheet_name
            counter = 1
            while sheet_name in workbook.sheetnames:
                suffix = f"_{counter}"
                max_len = 31 - len(suffix)
                sheet_name = original_name[:max_len] + suffix
                counter += 1
            
            sheet = workbook.create_sheet(title=sheet_name)
            
            # Get all unique keys from all records for this coop
            all_keys = []
            seen_keys = set()
            for record in records:
                for key in record.keys():
                    if key not in seen_keys:
                        all_keys.append(key)
                        seen_keys.add(key)
            
            # Write header row
            for col_idx, key in enumerate(all_keys, start=1):
                sheet.cell(row=1, column=col_idx, value=key)
            
            # Write data rows
            for row_idx, record in enumerate(records, start=2):
                for col_idx, key in enumerate(all_keys, start=1):
                    value = record.get(key, '')
                    sheet.cell(row=row_idx, column=col_idx, value=value)
            
            # Auto-adjust column widths
            for col_idx, key in enumerate(all_keys, start=1):
                max_length = len(key)
                for record in records:
                    value = str(record.get(key, ''))
                    max_length = max(max_length, len(value))
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
        
        # Save the workbook
        workbook.save(filepath)
    
    def process_data(self):
        """Process the input data and export to Excel."""
        # Get text from text area
        text = self.text_area.get("1.0", tk.END)
        
        if not text.strip():
            messagebox.showwarning("No Data", "Please paste some data into the text area.")
            return
        
        # Parse records
        try:
            records = self.parse_records(text)
            
            if not records:
                messagebox.showwarning("No Records", "No valid records found. Please check the format.")
                return
            
            # Group by CO-OP NAME
            grouped_data = self.group_by_coop_name(records)
            
            if not grouped_data:
                messagebox.showwarning("No Data", "No records to process.")
                return
            
            # Ask user for save location
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Excel File As"
            )
            
            if not filepath:
                # User cancelled
                return
            
            # Create Excel file
            self.create_excel_file(grouped_data, filepath)
            
            # Show confirmation
            messagebox.showinfo(
                "Success",
                f"Data successfully exported!\n\nFile saved to:\n{filepath}\n\n"
                f"Created {len(grouped_data)} sheet(s) with {len(records)} total record(s)."
            )
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = DataSorterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
