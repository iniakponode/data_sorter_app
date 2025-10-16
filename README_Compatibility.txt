Data Sorter Application - Windows 7 Compatibility Package
===========================================================

WINDOWS 7 COMPATIBILITY FIXES:
- Removed pywin32 dependencies that cause DLL loading errors
- Excluded problematic Windows-specific modules
- Optimized for older Windows systems
- Uses only essential Tkinter and openpyxl libraries

FILES INCLUDED:
- DataSorterApp_Win7Compatible.exe - Main application (Windows 7+ compatible)
- DataSorterApp_Minimal.exe - Ultra-lightweight version
- README_Compatibility.txt - This file

SYSTEM REQUIREMENTS:
- Windows 7 SP1 or later (32-bit or 64-bit)
- Microsoft Visual C++ 2015-2019 Redistributable (usually pre-installed)
- 30 MB available RAM
- 20 MB free disk space

TROUBLESHOOTING:
If you still encounter DLL errors:
1. Try the minimal version first: DataSorterApp_Minimal.exe
2. Install Microsoft Visual C++ 2015-2019 Redistributable
3. Run as administrator (right-click > "Run as administrator")
4. Check Windows Updates are installed

TESTED ON:
- Windows 7 SP1 (32-bit and 64-bit)
- Windows 8/8.1
- Windows 10
- Windows 11

For support, please create an issue on the GitHub repository.
