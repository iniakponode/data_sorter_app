# Data Sorter Application

![Version](https://img.shields.io/badge/version-2.0.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.8%2B-brightgreen.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

A robust desktop application that intelligently parses messy text data, filters out noise, and exports clean records to multi-sheet Excel files. Built with Python and Tkinter, packaged as standalone executables for easy distribution.

**Latest Release**: v2.0.0 (October 16, 2025)  
**New Features**: Multi-architecture support, intelligent noise filtering, robust data parsing

## 🚀 Quick Start (No Python Required)

**For End Users**: Download and run the pre-built executable:

1. Download the appropriate package for your system:
   - **Windows 32-bit**: `DataSorterApp_32_Distribution.zip` (21.1 MB)
   - **Windows 64-bit**: `DataSorterApp_64_Distribution.zip` (22.0 MB)
   - **Universal Windows**: `DataSorterApp_universal_Distribution.zip` (22.0 MB)
   - **All Architectures**: `DataSorterApp_Complete_All_Architectures.zip` (65.0 MB)

2. Extract the ZIP file and run `DataSorterApp.exe`
3. No Python installation required!

## 🛠️ Developer Installation

For developers who want to modify the source code:

1. Clone this repository:
```bash
git clone https://github.com/iniakponode/data_sorter_app.git
cd data_sorter_app
```

2. Create and activate virtual environment:
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Requirements

- **For End Users**: Windows OS (32-bit or 64-bit)
- **For Developers**: Python 3.8+ and dependencies listed in `requirements.txt`

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

The application now intelligently filters noise and extracts data from messy text:

```
PERSONAL DATA OF COOPERATIVE OWNERS

NAME: John Doe
CO-OP NAME: Alpha Co-op
PHONE NO: 08012345678
BANK NAME: First Bank
ACCT NO: 1234567890
SEX: MALE

YOU JUST HAVE NOW TILL 3PM TOMORROW TO SEND YOUR DETAILS
PLZ DON'T SEND TO OTHER NUMBERS

CEO NAME: Jane Smith
CO-OP NAME: Beta Co-op
PHONE NO: 08087654321
BANK NAME: GTB
ACCT NO: 0987654321
SEX: FEMALE
```

**Key Features:**
- **🧠 Intelligent Noise Filtering**: Automatically removes headers, instructions, and irrelevant text
- **🔍 Smart Data Detection**: Recognizes valid KEY: VALUE pairs amidst noise
- **📋 Flexible Format Support**: Handles variations in key names and formatting
- **📊 Automatic Grouping**: Records are grouped by the `CO-OP NAME` field value
- **✨ Robust Parsing**: Works with real-world messy data formats
- **📝 Data Validation**: Ensures only valid records with sufficient data are processed

**Supported Field Variations:**
- Names: `NAME`, `CEO NAME`, `CEO`
- Phone: `PHONE NO`, `PHONE`, `Phone no`
- Bank: `BANK NAME`, `BANK`, `Bank Name`
- Account: `ACCT NO`, `ACCOUNT NO`, `ACC No`, `Acct. N0.`
- Cooperative: `CO-OP NAME`, `COOP NAME`, `COOPERATIVE NAME`

## ✨ Key Features

### 🧠 Intelligent Data Processing
- **Robust Noise Filtering**: Automatically removes headers, instructions, and irrelevant text
- **Smart Data Detection**: Recognizes valid KEY: VALUE pairs amidst messy data
- **Flexible Format Support**: Handles variations in field names and formatting
- **Field Name Normalization**: Intelligently matches similar field names (e.g., "PHONE NO", "Phone no", "PHONE")

### 📊 Advanced Parsing Capabilities
- **Automatic Column Detection**: Uses first valid record to establish column structure
- **Mixed Format Support**: Handles both KEY: VALUE format and single values per line
- **Record Boundary Detection**: Intelligently separates records using blank lines
- **Data Validation**: Ensures only valid records with sufficient data are processed

### 💼 Professional Output
- **Multi-sheet Excel Export**: Automatically groups records by CO-OP NAME
- **Clean Formatting**: Auto-adjusted column widths and professional styling
- **Error Handling**: Comprehensive error reporting and user feedback

### 🖥️ User Experience
- **Simple GUI**: Intuitive drag-and-drop or paste interface
- **Real-time Processing**: Instant feedback during data processing
- **Cross-platform Executables**: Standalone apps for Windows (32-bit, 64-bit, universal)

## 📦 Distribution Packages

The application is available as ready-to-run executables:

| Package | Size | Target System | Contents |
|---------|------|---------------|----------|
| `DataSorterApp_32_Distribution.zip` | 21.1 MB | Windows 32-bit | Executable + Documentation + Examples |
| `DataSorterApp_64_Distribution.zip` | 22.0 MB | Windows 64-bit | Executable + Documentation + Examples |
| `DataSorterApp_universal_Distribution.zip` | 22.0 MB | Universal Windows | Executable + Documentation + Examples |
| `DataSorterApp_Complete_All_Architectures.zip` | 65.0 MB | All Windows | All executables + Documentation |

### Package Contents
- ✅ Ready-to-run executable (no Python installation required)
- ✅ User documentation with installation and usage instructions
- ✅ Example data file for testing
- ✅ Launcher script for easy execution

## Application Interface

The application window consists of:
1. **Instructions Label** - Shows how to format input data
2. **Text Area** - Large scrollable text field for pasting records
3. **Process Button** - Green button to trigger data processing and export

```
┌─────────────────────────────────────────────────────────┐
│         Data Sorter Application                         │
├─────────────────────────────────────────────────────────┤
│ First record: KEY: VALUE format (establishes columns)  │
│ Subsequent records: KEY: VALUE or single values per    │
│ line (Records separated by blank lines):               │
│                                                         │
│ ┌─────────────────────────────────────────────────────┐│
│ │                                                       ││
│ │  Name: John Doe                                      ││
│ │  CO-OP NAME: Alpha Co-op                            ││
│ │  Member ID: 12345                                   ││
│ │  Email: john@example.com                            ││
│ │                                                       ││
│ │  Jane Smith                                         ││
│ │  Alpha Co-op                                        ││
│ │  67890                                              ││
│ │  jane@example.com                                   ││
│ │                                                       ││
│ │  Bob Johnson                                        ││
│ │  Beta Co-op                                         ││
│ │  11111                                              ││
│ │  bob@example.com                                    ││
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

## 🔧 Building Executables

For developers who want to build their own executables:

### Single Architecture Build
```bash
# Activate virtual environment
.venv\Scripts\activate

# Install build dependencies
pip install pyinstaller

# Build for current architecture
python build_exe.py
```

### Multi-Architecture Build
```bash
# Build for all Windows architectures
python create_packages.py
```

This will create distribution packages for:
- 32-bit Windows systems
- 64-bit Windows systems  
- Universal Windows compatibility
- Complete package with all architectures

## 🧪 Testing

The project includes a comprehensive test suite using pytest:

```bash
# Install test dependencies
pip install -r requirements.txt

# Run all tests
pytest test_app.py -v

# Run specific test categories
pytest test_app.py::test_parse_records -v
pytest test_app.py::test_group_by_coop -v
```

**Test Coverage:**
- ✅ Noise filtering and data extraction
- ✅ Record parsing from messy text input
- ✅ Field name normalization and matching
- ✅ Grouping records by CO-OP NAME
- ✅ Excel file creation with multiple sheets
- ✅ Edge cases (empty input, missing fields, malformed data)
- ✅ Integration testing of complete workflow

## 📋 Project Structure

```
data_sorter_app/
├── app.py                    # Main application with GUI
├── example.py               # Test data examples
├── test_app.py             # Comprehensive test suite
├── requirements.txt        # Python dependencies
├── build_exe.py           # Multi-architecture build script
├── create_packages.py     # Distribution package creator
├── dist/                  # Built executables
│   ├── DataSorterApp.exe           # 64-bit executable
│   ├── DataSorterApp_x64.exe       # 64-bit explicit
│   └── DataSorterApp_x86.exe       # 32-bit executable
└── *.zip                  # Distribution packages
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is open source and available under the MIT License.

## 📞 Support

For issues, questions, or contributions, please open an issue on the GitHub repository.
