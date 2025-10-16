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
            text="Paste your data below - the app will automatically filter out noise!\n" +
                 "• First record: KEY: VALUE format (establishes columns)\n" +
                 "• Subsequent records: KEY: VALUE or single values per line\n" +
                 "• Headers, instructions, and other text will be filtered out automatically",
            font=("Arial", 10),
            pady=10,
            justify=tk.LEFT
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
    
    def is_noise_line(self, line):
        """
        Determine if a line is noise/extraneous text that should be filtered out.
        
        Args:
            line (str): Line to check
            
        Returns:
            bool: True if the line is noise, False if it's potentially valid data
        """
        line_upper = line.upper()
        
        # First check if it's a valid key-value pair - if so, it's NOT noise
        if self.is_valid_key_value_pair(line):
            return False
        
        # Common noise patterns (only for non-key-value lines)
        noise_patterns = [
            # Headers and titles
            'PERSONAL DATA', 'SOCIETY', 'LIMITED', 'LTD',
            'FARMERS', 'UNION', 'MPCS', 'FCSL', 'MPCSL',
            # Instructions and messages
            'YOU JUST HAVE', 'TILL', 'TOMORROW', 'SEND YOUR DETAILS',
            'WHATSAPP', 'SMS', 'DON\'T SEND', 'OTHER NUMBERS',
            'PLEASE', 'HELP', 'CORRECTION', 'THANK YOU', 'CORRECT',
            'SERIAL NO', 'INSTEAD OF', 'SPELLED', 'SIR',
            # Security/company names
            'PROTECTUS SECURITY', 'INTERNATIONAL',
            # Special characters and formatting
            '👆', '👇🏻', '👇', '🏻'
        ]
        
        # Check if line contains noise patterns
        for pattern in noise_patterns:
            if pattern in line_upper:
                return True
        
        # Check for lines that are all caps and look like titles (but not key-value pairs)
        if (line.isupper() and 
            len(line) > 20 and 
            not ':' in line and
            ('COOPERATIVE' in line or 'SOCIETY' in line or 'LIMITED' in line)):
            return True
        
        # Check for lines with excessive punctuation
        punct_count = sum(1 for c in line if c in '.,!?;')
        if len(line) > 10 and punct_count / len(line) > 0.3:
            return True
        
        # Check for lines that look like serial numbers or corrections
        if (line_upper.startswith('SERIAL') or 
            'CORRECT' in line_upper or 
            'SPELLING' in line_upper):
            return True
        
        return False
    
    def is_valid_key_value_pair(self, line):
        """
        Check if a line represents a valid KEY: VALUE pair for data records.
        
        Args:
            line (str): Line to check
            
        Returns:
            bool: True if it's a valid key-value pair
        """
        if ':' not in line:
            return False
        
        parts = line.split(':', 1)
        key = parts[0].strip().upper()
        value = parts[1].strip()
        
        # Valid key patterns (common data fields)
        valid_keys = [
            'NAME', 'PHONE', 'PHONE NO', 'BANK', 'BANK NAME', 'ACCT', 'ACCT NO', 
            'ACCOUNT', 'ACCOUNT NO', 'SEX', 'EMAIL', 'ADDRESS', 'CEO', 'CEO NAME',
            'CO-OP NAME', 'COOP NAME', 'COOPERATIVE NAME', 'COOPERATIVE', 'CO-OP',
            'MEMBER ID', 'ID', 'GENDER'
        ]
        
        # Check if key matches valid patterns
        for valid_key in valid_keys:
            if valid_key in key:
                return True
        
        # Additional checks for common variations
        if (any(word in key for word in ['NAME', 'PHONE', 'BANK', 'ACCT', 'SEX']) and 
            len(key) <= 30):  # Reasonable key length
            return True
        
        return False
    
    def normalize_key_name(self, key):
        """
        Normalize a key name to standard format.
        
        Args:
            key (str): Original key name
            
        Returns:
            str: Normalized key name
        """
        key_upper = key.upper().strip()
        
        # Standardize common variations - order matters!
        if 'CO-OP' in key_upper or 'COOP' in key_upper or 'COOPERATIVE' in key_upper:
            return 'CO-OP NAME'
        elif 'PHONE' in key_upper:
            return 'PHONE NO'
        elif 'BANK' in key_upper and 'NAME' in key_upper:
            return 'BANK NAME'
        elif 'BANK' in key_upper and ('ACCT' in key_upper or 'ACCOUNT' in key_upper):
            return 'ACCT NO'
        elif ('ACCT' in key_upper or 'ACCOUNT' in key_upper) and 'NO' in key_upper:
            return 'ACCT NO'
        elif 'ACCT' in key_upper or 'ACCOUNT' in key_upper:
            return 'ACCT NO'
        elif 'CEO' in key_upper and 'NAME' in key_upper:
            return 'CEO NAME'
        elif 'CEO' in key_upper:
            return 'CEO NAME'
        elif key_upper == 'NAME':
            return 'NAME'
        elif key_upper in ['SEX', 'GENDER']:
            return 'SEX'
        elif key_upper in ['EMAIL', 'E-MAIL']:
            return 'EMAIL'
        elif key_upper in ['ADDRESS', 'LOCATION']:
            return 'ADDRESS'
        elif 'BANK' in key_upper:
            return 'BANK NAME'
        else:
            return key_upper
    
    def clean_and_normalize_data(self, records, headers):
        """
        Clean and normalize the extracted data records.
        
        Args:
            records (list): List of record rows
            headers (list): List of column headers
            
        Returns:
            tuple: (cleaned_headers, cleaned_records)
        """
        if not records or not headers:
            return headers, records
        
        # Normalize headers
        normalized_headers = []
        for header in headers:
            normalized_headers.append(self.normalize_key_name(header))
        
        # Remove duplicates while preserving order
        seen = set()
        final_headers = []
        for header in normalized_headers:
            if header not in seen:
                seen.add(header)
                final_headers.append(header)
        
        # Clean records
        cleaned_records = []
        for record in records:
            cleaned_record = []
            for i, value in enumerate(record):
                if i < len(final_headers):
                    # Clean up values
                    clean_value = str(value).strip()
                    
                    # Remove common suffixes/prefixes
                    clean_value = clean_value.replace('Account Name:', '').strip()
                    clean_value = clean_value.replace('ACC No', '').strip()
                    clean_value = clean_value.replace('Acct. N0.', '').strip()
                    clean_value = clean_value.replace('Phone no.', '').strip()
                    clean_value = clean_value.replace('CEO:', '').strip()
                    
                    cleaned_record.append(clean_value)
                else:
                    cleaned_record.append(str(value).strip())
            
            # Ensure all records have the same number of columns
            while len(cleaned_record) < len(final_headers):
                cleaned_record.append('')
            
            # Only add records that have some actual data
            if any(val.strip() for val in cleaned_record):
                cleaned_records.append(cleaned_record[:len(final_headers)])
        
        return final_headers, cleaned_records
    
    def parse_records(self, text):
        """
        Parse the input text into records with robust noise filtering.
        
        Args:
            text (str): Input text with KEY: VALUE format records
            
        Returns:
            tuple: (column_headers, records) where records is list of lists
        """
        lines = text.strip().split('\n')
        
        # Pre-filter to remove noise and group into blocks
        blocks = []
        current_block = []
        
        for line in lines:
            line = line.strip()
            
            if not line:  # Empty line - potential block separator
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            elif not self.is_noise_line(line):
                current_block.append(line)
        
        # Add final block if it exists
        if current_block:
            blocks.append(current_block)
        
        # Extract records from blocks
        records_data = []
        all_keys = set()
        
        for block in blocks:
            record = {}
            has_valid_data = False
            
            for line in block:
                if self.is_valid_key_value_pair(line):
                    parts = line.split(':', 1)
                    key = self.normalize_key_name(parts[0].strip())
                    value = parts[1].strip() if len(parts) > 1 else ""
                    
                    # Add all valid key-value pairs, even if value is empty
                    record[key] = value
                    all_keys.add(key)
                    if value.strip():  # Has meaningful data
                        has_valid_data = True
            
            # Only add records that have at least some data
            if has_valid_data and len(record) >= 1:  # At least 1 field
                records_data.append(record)
        
        if not records_data:
            return [], []
        
        # Create consistent column order
        # Prioritize common fields
        priority_fields = ['NAME', 'CEO NAME', 'CO-OP NAME', 'PHONE NO', 'BANK NAME', 'ACCT NO', 'SEX', 'EMAIL', 'ADDRESS']
        column_headers = []
        
        # Add priority fields that exist
        for field in priority_fields:
            if field in all_keys:
                column_headers.append(field)
        
        # Add any remaining fields
        for key in sorted(all_keys):
            if key not in column_headers:
                column_headers.append(key)
        
        # Convert to list format
        records = []
        for record_dict in records_data:
            record_list = []
            for header in column_headers:
                record_list.append(record_dict.get(header, ''))
            records.append(record_list)
        
        return column_headers, records
    
    def group_by_coop_name(self, column_headers, records):
        """
        Group records by CO-OP NAME field.
        
        Args:
            column_headers (list): List of column header names
            records (list): List of record rows (each row is a list of values)
            
        Returns:
            dict: Dictionary with CO-OP NAME as keys and list of records as values
        """
        grouped = defaultdict(list)
        
        # Find the index of CO-OP NAME column
        coop_name_index = None
        for i, header in enumerate(column_headers):
            if header == 'CO-OP NAME':
                coop_name_index = i
                break
        
        for record in records:
            if coop_name_index is not None and coop_name_index < len(record):
                coop_name = record[coop_name_index] if record[coop_name_index] else 'Unknown'
            else:
                coop_name = 'Unknown'
            grouped[coop_name].append(record)
        
        return grouped
    
    def create_excel_file(self, column_headers, grouped_data, filepath):
        """
        Create an Excel file with separate sheets for each CO-OP NAME.
        
        Args:
            column_headers (list): List of column header names
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
            
            # Write header row
            for col_idx, header in enumerate(column_headers, start=1):
                sheet.cell(row=1, column=col_idx, value=header)
            
            # Write data rows
            for row_idx, record in enumerate(records, start=2):
                for col_idx, value in enumerate(record, start=1):
                    if col_idx <= len(column_headers):  # Don't exceed column count
                        sheet.cell(row=row_idx, column=col_idx, value=value)
            
            # Auto-adjust column widths
            for col_idx, header in enumerate(column_headers, start=1):
                max_length = len(str(header))
                for record in records:
                    if col_idx - 1 < len(record):
                        value = str(record[col_idx - 1])
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
            column_headers, records = self.parse_records(text)
            
            if not records:
                messagebox.showwarning("No Records", "No valid records found. Please check the format.")
                return
            
            if not column_headers:
                messagebox.showwarning("No Headers", "No column headers found. First record must use KEY: VALUE format.")
                return
            
            # Group by CO-OP NAME
            grouped_data = self.group_by_coop_name(column_headers, records)
            
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
            self.create_excel_file(column_headers, grouped_data, filepath)
            
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
