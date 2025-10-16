# Data Sorter Application

Build a standalone Python desktop application that parses structured text records, groups them, and exports them to a multi-sheet Excel file.

## Description

This is a Python desktop application built with Tkinter that allows users to:
- Paste text records in KEY: VALUE format
- Automatically group records by CO-OP NAME
- Export grouped data to Excel with separate sheets for each CO-OP NAME

## Requirements

- Python 3.6 or higher
- openpyxl library

## Installation

1. Clone this repository:
```bash
git clone https://github.com/iniakponode/data_sorter_app.git
cd data_sorter_app
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python app.py
```

2. In the application window:
   - Paste your records into the text area (records should be in KEY: VALUE format, separated by blank lines)
   - Click the "Process and Export to Excel" button
   - Choose where to save the Excel file
   - View the confirmation message with the file location

### Input Format

Records should be formatted as follows:
```
KEY1: Value1
KEY2: Value2
CO-OP NAME: Cooperative Name
KEY3: Value3

KEY1: Value4
KEY2: Value5
CO-OP NAME: Another Cooperative
KEY3: Value6
```

Each record is separated by a blank line. The application will group all records with the same CO-OP NAME into a single Excel sheet.

## Features

- Simple and intuitive GUI
- Automatic grouping by CO-OP NAME
- Multi-sheet Excel export
- Auto-adjusted column widths
- Confirmation messages
- Error handling

## Application Interface

The application window consists of:
1. **Instructions Label** - Shows how to format input data
2. **Text Area** - Large scrollable text field for pasting records
3. **Process Button** - Green button to trigger data processing and export

```
┌─────────────────────────────────────────────────────────┐
│         Data Sorter Application                         │
├─────────────────────────────────────────────────────────┤
│ Paste records in KEY: VALUE format (separated by       │
│ blank lines):                                           │
│                                                         │
│ ┌─────────────────────────────────────────────────────┐│
│ │                                                       ││
│ │  Name: John Doe                                      ││
│ │  CO-OP NAME: Alpha Co-op                            ││
│ │  Member ID: 12345                                   ││
│ │                                                       ││
│ │  Name: Jane Smith                                   ││
│ │  CO-OP NAME: Alpha Co-op                            ││
│ │  Member ID: 67890                                   ││
│ │                                                       ││
│ │  Name: Bob Johnson                                  ││
│ │  CO-OP NAME: Beta Co-op                             ││
│ │  Member ID: 11111                                   ││
│ │                                                       ││
│ │                                                       ││
│ └─────────────────────────────────────────────────────┘│
│                                                         │
│         ┌─────────────────────────────────┐            │
│         │ Process and Export to Excel     │            │
│         └─────────────────────────────────┘            │
└─────────────────────────────────────────────────────────┘
```

## Example

Input:
```
Name: John Doe
CO-OP NAME: Alpha Co-op
Member ID: 12345

Name: Jane Smith
CO-OP NAME: Alpha Co-op
Member ID: 67890

Name: Bob Johnson
CO-OP NAME: Beta Co-op
Member ID: 11111
```

Output: An Excel file with two sheets:
- "Alpha Co-op" sheet containing John Doe and Jane Smith's records
- "Beta Co-op" sheet containing Bob Johnson's record

## License

This project is open source and available under the MIT License.
