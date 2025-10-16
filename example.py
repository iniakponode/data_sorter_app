#!/usr/bin/env python3
"""
Example usage of the Data Sorter Application
Demonstrates the core functionality without GUI
"""

import sys
from app import DataSorterApp
import tkinter as tk

# Sample data with noise (realistic example)
SAMPLE_DATA = """PERSONAL DATA OF COOPERATIVE OWNERS

NAME: GAD ELEMOKUMO
CO-OP NAME: Bayelsa Farmers Union
PHONE NO: 08037374009
BANK NAME: FIRST BANK
ACCT NO: 1000383461
SEX: FEMALE

YOU JUST HAVE NOW TILL 3PM TOMORROW TO SEND YOUR DETAILS TO 08030822597 ON WHATSAPP OR SMS
PLZ DON'T SEND TO OTHER NUMBERS

PROTECTUS SECURITY INTERNATIONAL
CEO NAME: EWA OBIA OMINI
CO-OP NAME: Delta Agricultural Coop
PHONE NO: 08035533020
BANK NAME: GTB
ACCT NO: 22281345092
EMAIL: eomini51@gmail.com

MAREWOMA COOPERATIVE SOCIETY IRRI. PERSONAL DATA OF COOPERATIVE OWNER.

NAME: ODOGWA QUEEN ENIFOME
CO-OP NAME: Marewoma Coop Society
PHONE NO: 08138379531
BANK NAME: ZENITH
ACCT NO: 2191621677
SEX: FEMALE

ROYAL WOMEN FARMERS Cooperative SOCIETY LIMITED 

CEO NAME: Prefa Oyinduobra Helen 
CO-OP NAME: Royal Women Farmers
Phone no: 08037806976
Bank Name: U.B.A
Acc No: 2151124183
Sex: Female

EVERGREEN AGRO MPCSL

CEO NAME: KAYODE JANET NIKE
CO-OP NAME: Evergreen Agro
Phone no: 08140301646
Bank Name: FIRST BANK
Account no: 3053891740

Please help me with the corrections sir 
Thank you

Ebieyerin(Angalabiri) M.P.C.S. LTD. 
CEO: Agononama Priestley Beer
CO-OP NAME: Ebieyerin Angalabiri
Phone no: 08064300259
Sex: Male
Bank: Sterling
ACCT No: 0023279323"""

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
    column_headers, records = app.parse_records(SAMPLE_DATA)
    print(f"   Found {len(records)} records")
    print(f"   Column headers: {column_headers}")
    
    # Group by CO-OP NAME
    print("\n2. Grouping by CO-OP NAME...")
    grouped = app.group_by_coop_name(column_headers, records)
    for coop_name, coop_records in grouped.items():
        print(f"   - {coop_name}: {len(coop_records)} record(s)")
    
    # Create Excel file
    print("\n3. Creating Excel file...")
    output_file = 'example_output.xlsx'
    app.create_excel_file(column_headers, grouped, output_file)
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
