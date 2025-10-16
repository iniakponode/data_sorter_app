#!/usr/bin/env python3
"""
Create comprehensive distribution packages for all Windows architectures
"""

import subprocess
import shutil
from pathlib import Path

def create_comprehensive_packages():
    """Create distribution packages for all architectures."""
    
    dist_dir = Path("dist")
    
    # Define available executables
    executables = [
        ("DataSorterApp_x86.exe", "32-bit", "Compatible with 32-bit and 64-bit Windows"),
        ("DataSorterApp_x64.exe", "64-bit", "Optimized for 64-bit Windows systems"),
        ("DataSorterApp.exe", "Universal", "Works on both 32-bit and 64-bit Windows")
    ]
    
    packages_created = []
    
    print("Creating Comprehensive Distribution Packages")
    print("=" * 50)
    
    # Create individual architecture packages
    for exe_name, arch, description in executables:
        exe_path = dist_dir / exe_name
        
        if not exe_path.exists():
            print(f"⚠️  Skipping {exe_name} - not found")
            continue
        
        # Create package folder
        package_name = f"DataSorterApp_{arch.replace('-bit', '').lower()}"
        package_dir = Path(f"packages") / package_name
        package_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            # Copy executable
            shutil.copy2(exe_path, package_dir / exe_name)
            
            # Copy documentation
            docs = ["Installation_Instructions.txt", "README_Distribution.md", "Package_Summary.md"]
            for doc in docs:
                doc_path = dist_dir / doc
                if doc_path.exists():
                    shutil.copy2(doc_path, package_dir / doc)
            
            # Create architecture-specific README
            arch_readme = f"""# Data Sorter Application - {arch} Edition

## 📋 Package Contents
- **{exe_name}** - Main application ({description})
- **Installation_Instructions.txt** - Complete setup guide
- **README_Distribution.md** - Detailed documentation
- **Package_Summary.md** - Build information
- **Run_DataSorter.bat** - Quick launcher script

## 🚀 Quick Start
1. Double-click `{exe_name}` to run the application
2. OR use `Run_DataSorter.bat` for guided startup

## 💻 System Requirements
- **Target Systems:** {description}
- **Operating System:** Windows 10 or later
- **Memory:** 50 MB available RAM
- **Storage:** 25 MB free space
- **Dependencies:** None (standalone executable)

## ✨ Features
✅ **Smart Noise Filtering** - Removes headers and irrelevant text automatically
✅ **Flexible Data Input** - Handles messy real-world data formats
✅ **Professional Output** - Creates organized Excel files with separate sheets
✅ **CO-OP Grouping** - Automatically groups records by cooperative names
✅ **User Friendly** - Simple interface, no technical knowledge required

## 📞 Support
For technical support or questions, refer to the included documentation files.

---
**Version:** 2.0 Enhanced with Noise Filtering
**Build Date:** October 16, 2025
**Architecture:** {arch} Windows Systems
"""
            
            with open(package_dir / "README.md", 'w', encoding='utf-8') as f:
                f.write(arch_readme)
            
            # Create launcher batch file
            batch_content = f"""@echo off
REM Data Sorter Application - {arch} Edition
REM Quick Launcher Script

title Data Sorter Application - {arch}

echo.
echo ========================================
echo   Data Sorter Application
echo   {arch} Edition
echo ========================================
echo.
echo Starting application...
echo.

REM Check if executable exists
if not exist "{exe_name}" (
    echo ERROR: {exe_name} not found in this directory!
    echo.
    echo Please ensure you have extracted all files from the ZIP package
    echo and are running this script from the same folder as {exe_name}
    echo.
    pause
    exit /b 1
)

REM Launch the application
echo Launching {exe_name}...
start "" "{exe_name}"

REM Show success message
echo.
echo ✅ Data Sorter Application started successfully!
echo.
echo The application window should now be open.
echo You can close this command window.
echo.

REM Auto-close after 3 seconds
timeout /t 3 /nobreak >nul 2>&1
"""
            
            with open(package_dir / "Run_DataSorter.bat", 'w', encoding='utf-8') as f:
                f.write(batch_content)
            
            # Create ZIP package
            zip_path = Path(f"{package_name}_Distribution.zip")
            zip_cmd = [
                "powershell", "-Command",
                f"Compress-Archive -Path '{package_dir}\\*' -DestinationPath '{zip_path}' -Force"
            ]
            
            result = subprocess.run(zip_cmd, capture_output=True, text=True, cwd=".")
            
            if result.returncode == 0 and zip_path.exists():
                size_mb = zip_path.stat().st_size / (1024 * 1024)
                packages_created.append((arch, zip_path, size_mb))
                print(f"✅ {arch}: {zip_path.name} ({size_mb:.1f} MB)")
            else:
                print(f"❌ Failed to create ZIP for {arch}")
                print(f"   Error: {result.stderr}")
            
        except Exception as e:
            print(f"❌ Error creating {arch} package: {e}")
    
    # Create all-in-one package
    print(f"\n📦 Creating comprehensive package with all versions...")
    
    all_in_one_dir = Path("packages") / "DataSorterApp_Complete"
    all_in_one_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Copy all executables
        for exe_name, arch, description in executables:
            exe_path = dist_dir / exe_name
            if exe_path.exists():
                shutil.copy2(exe_path, all_in_one_dir / exe_name)
        
        # Copy documentation
        docs = ["Installation_Instructions.txt", "README_Distribution.md", "Package_Summary.md"]
        for doc in docs:
            doc_path = dist_dir / doc
            if doc_path.exists():
                shutil.copy2(doc_path, all_in_one_dir / doc)
        
        # Create comprehensive README
        complete_readme = f"""# Data Sorter Application - Complete Package

## 📦 What's Included
This package contains ALL versions of the Data Sorter Application for maximum compatibility:

### Available Executables:
- **DataSorterApp_x86.exe** - 32-bit compatible (works on ANY Windows system)
- **DataSorterApp_x64.exe** - 64-bit optimized (best performance on 64-bit systems)  
- **DataSorterApp.exe** - Universal version (recommended for most users)

## 🚀 Which Version Should I Use?

### 🟢 Recommended for Most Users:
**DataSorterApp.exe** (Universal) - Works great on both 32-bit and 64-bit Windows

### 🔧 For Maximum Compatibility:
**DataSorterApp_x86.exe** - Guaranteed to work on older 32-bit Windows systems

### ⚡ For Best Performance:
**DataSorterApp_x64.exe** - Optimized for modern 64-bit Windows systems

## 💻 System Requirements
- **Operating System:** Windows 10 or later (32-bit or 64-bit)
- **Memory:** 50 MB available RAM  
- **Storage:** 25 MB free space
- **Dependencies:** None required - all versions are completely standalone

## 🎯 Quick Start
1. Choose the appropriate executable for your system
2. Double-click to run (no installation needed!)
3. Paste your data and let the app filter out noise automatically
4. Export to professional Excel files with organized sheets

## ✨ Key Features
✅ **Intelligent Data Processing** - Automatically filters noise from messy data
✅ **Real-World Ready** - Handles copy-pasted data from any source
✅ **Professional Output** - Creates organized Excel files grouped by cooperative
✅ **Zero Setup Required** - Standalone executables, no installation needed
✅ **Universal Compatibility** - Versions for all Windows systems

## 📋 What This App Does
The Data Sorter Application takes messy, unstructured cooperative member data (with headers, instructions, and other noise) and automatically:

1. **Filters Out Noise** - Removes headers, instructions, and irrelevant text
2. **Extracts Valid Data** - Identifies and captures member information
3. **Organizes Records** - Groups members by their cooperative names  
4. **Exports to Excel** - Creates professional spreadsheets with separate sheets per cooperative

Perfect for processing real-world data from various sources!

## 📞 Support
All documentation and instructions are included in this package. The application is designed to be intuitive and requires no technical knowledge.

---
**Complete Package - All Windows Architectures Supported**
**Version:** 2.0 Enhanced Edition
**Build Date:** October 16, 2025
"""
        
        with open(all_in_one_dir / "README.md", 'w', encoding='utf-8') as f:
            f.write(complete_readme)
        
        # Create ZIP for complete package
        complete_zip = Path("DataSorterApp_Complete_All_Architectures.zip")
        zip_cmd = [
            "powershell", "-Command", 
            f"Compress-Archive -Path '{all_in_one_dir}\\*' -DestinationPath '{complete_zip}' -Force"
        ]
        
        result = subprocess.run(zip_cmd, capture_output=True, text=True, cwd=".")
        
        if result.returncode == 0 and complete_zip.exists():
            size_mb = complete_zip.stat().st_size / (1024 * 1024)
            packages_created.append(("Complete Package", complete_zip, size_mb))
            print(f"✅ Complete Package: {complete_zip.name} ({size_mb:.1f} MB)")
        
    except Exception as e:
        print(f"❌ Error creating complete package: {e}")
    
    return packages_created

if __name__ == "__main__":
    packages = create_comprehensive_packages()
    
    if packages:
        print(f"\n🎉 All distribution packages created successfully!")
        print(f"\n📁 Distribution Packages Created:")
        for arch, zip_path, size_mb in packages:
            print(f"  📦 {arch}: {zip_path.name} ({size_mb:.1f} MB)")
        
        print(f"\n🚀 Ready for distribution!")
        print("   Users can choose the package that best fits their system.")
        print("   The Complete Package includes all versions for maximum flexibility.")
        
    else:
        print(f"\n💥 Package creation failed!")