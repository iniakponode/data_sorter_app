#!/usr/bin/env python3
"""
Example usage of the Data Sorter Application
Demonstrates the core functionality without GUI
"""

import sys
from app import DataSorterApp
import tkinter as tk

# Sample data
SAMPLE_DATA = """Name: John Doe
CO-OP NAME: Alpha Co-op
Member ID: 12345
Email: john@example.com

Name: Jane Smith
CO-OP NAME: Alpha Co-op
Member ID: 67890
Email: jane@example.com

Name: Bob Johnson
CO-OP NAME: Beta Co-op
Member ID: 11111
Email: bob@example.com"""

def main():
    """Demonstrate core functionality."""
    print("Data Sorter Application - Example Usage")
    print("=" * 60)
    
    # Create a simple root window (won't be shown)
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Create app instance
    app = DataSorterApp(root)
    
    # Parse sample data
    print("\n1. Parsing sample data...")
    records = app.parse_records(SAMPLE_DATA)
    print(f"   Found {len(records)} records")
    
    # Group by CO-OP NAME
    print("\n2. Grouping by CO-OP NAME...")
    grouped = app.group_by_coop_name(records)
    for coop_name, coop_records in grouped.items():
        print(f"   - {coop_name}: {len(coop_records)} record(s)")
    
    # Create Excel file
    print("\n3. Creating Excel file...")
    output_file = 'example_output.xlsx'
    app.create_excel_file(grouped, output_file)
    print(f"   Excel file created: {output_file}")
    
    print("\n" + "=" * 60)
    print("Example completed successfully!")
    print(f"Check {output_file} to see the results.")
    
    root.destroy()

if __name__ == "__main__":
    try:
        main()
    except ImportError as e:
        print("Error: tkinter module not found.")
        print("To run the full GUI application, tkinter must be installed.")
        print("\nOn Ubuntu/Debian: sudo apt-get install python3-tk")
        print("On Fedora: sudo dnf install python3-tkinter")
        print("On macOS: tkinter is included with Python")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
