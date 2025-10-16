#!/usr/bin/env python3
"""
Build script to create standalone executables for Data Sorter Application
Supports both 32-bit and 64-bit Windows systems
"""

import os
import subprocess
import sys
import platform
from pathlib import Path

def get_system_info():
    """Get system architecture information."""
    arch = platform.architecture()[0]
    machine = platform.machine().lower()
    
    print(f"System Architecture: {arch}")
    print(f"Machine Type: {machine}")
    
    # Determine if we're on 64-bit system
    is_64bit = arch == '64bit' or 'amd64' in machine or 'x86_64' in machine
    
    return is_64bit

def build_executable(target_arch="auto", output_suffix=""):
    """Build the executable using PyInstaller for specified architecture."""
    
    # Get the current directory
    current_dir = Path(__file__).parent
    
    # Determine output name based on architecture
    exe_name = f"DataSorterApp{output_suffix}"
    
    # Define the PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",                    # Create a single executable file
        "--windowed",                   # Don't show console window (GUI app)
        f"--name={exe_name}",          # Name of the executable
        "--icon=NONE",                  # No icon (can be added later)
        "--add-data=README.md;.",       # Include README in the bundle
        "--distpath=dist",              # Output directory
        "--workpath=build",             # Build directory
        "--specpath=.",                 # Spec file location
        "app.py"                        # Main Python file
    ]
    
    print(f"Building Data Sorter Application executable ({target_arch})...")
    print("Command:", " ".join(cmd))
    print("-" * 50)
    
    try:
        # Run PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        print("Build successful!")
        if result.stdout.strip():
            print("\nOutput:")
            print(result.stdout)
        
        # Check if the executable was created
        exe_path = current_dir / "dist" / f"{exe_name}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✅ Executable created: {exe_path}")
            print(f"📁 File size: {size_mb:.1f} MB")
            return exe_path
        else:
            print("❌ Executable not found in expected location")
            return None
            
    except subprocess.CalledProcessError as e:
        print("❌ Build failed!")
        print("Error:", e.stderr)
        return None

def build_multi_architecture():
    """Build executables for multiple architectures."""
    
    print("Data Sorter Application - Multi-Architecture Builder")
    print("=" * 60)
    
    # Get system info
    is_64bit_system = get_system_info()
    
    built_executables = []
    
    # Build for current architecture (auto-detect)
    print(f"\n🔧 Building for current system architecture...")
    
    if is_64bit_system:
        # Build 64-bit version
        exe_path = build_executable("64-bit", "_x64")
        if exe_path:
            built_executables.append(("64-bit", exe_path))
    else:
        # Build 32-bit version
        exe_path = build_executable("32-bit", "_x86")
        if exe_path:
            built_executables.append(("32-bit", exe_path))
    
    # Build universal version (works on both architectures)
    print(f"\n🔧 Building universal version...")
    exe_path = build_executable("Universal", "")
    if exe_path:
        built_executables.append(("Universal", exe_path))
    
    return built_executables

def create_installer_info(built_executables):
    """Create installation instructions for multiple architectures."""
    
    # Create executable list for instructions
    exe_list = []
    for arch, exe_path in built_executables:
        exe_name = exe_path.name
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        exe_list.append(f"- {exe_name} ({arch}) - {size_mb:.1f} MB")
    
    exe_descriptions = "\n".join(exe_list)
    
    info_content = f"""Data Sorter Application - Installation Instructions
=====================================================

AVAILABLE VERSIONS:
{exe_descriptions}

WHICH VERSION TO CHOOSE:
- DataSorterApp.exe (Universal) - Works on both 32-bit and 64-bit Windows
- DataSorterApp_x64.exe - Optimized for 64-bit Windows systems  
- DataSorterApp_x86.exe - Compatible with 32-bit Windows systems

SYSTEM REQUIREMENTS:
- Windows 10 or later (32-bit or 64-bit)
- 50 MB available RAM
- 25 MB free disk space
- No additional software required (standalone executables)

INSTALLATION:
1. Download the appropriate executable for your system
2. Save it to any folder on your computer
3. Double-click the executable to run

USAGE:
1. The application window will open
2. Paste your messy data directly - the app filters noise automatically!
   
   Example input (with noise):
   PERSONAL DATA OF COOPERATIVE OWNERS
   
   Name: John Doe
   CO-OP NAME: Alpha Co-op
   Phone No: 08012345678
   Bank Name: First Bank
   Acct No: 1234567890
   Sex: Male
   
   YOU HAVE UNTIL 3PM TO SEND DETAILS
   
   CEO Name: Jane Smith
   CO-OP NAME: Beta Co-op
   Phone No: 08087654321
   
3. Click "Process and Export to Excel"
4. Choose where to save your Excel file
5. The application automatically groups records by CO-OP NAME

KEY FEATURES:
✅ Intelligent noise filtering - removes headers and instructions
✅ Flexible format support - handles various field name variations
✅ Smart data detection - extracts valid records from messy text
✅ Professional Excel output - separate sheets per cooperative
✅ Real-world ready - works with copy-pasted data from any source

TECHNICAL SUPPORT:
- Application automatically handles noise and formatting variations
- Records are grouped by CO-OP NAME field
- Supports various field names (CEO/CEO NAME, PHONE/PHONE NO, etc.)
- Creates professional Excel files with proper formatting

For additional support, contact the developer.
"""
    
    info_path = Path("dist") / "Installation_Instructions.txt"
    info_path.parent.mkdir(exist_ok=True)
    
    with open(info_path, 'w', encoding='utf-8') as f:
        f.write(info_content)
    
    print(f"📄 Installation instructions created: {info_path}")

def create_distribution_packages(built_executables):
    """Create distribution ZIP packages for each architecture."""
    
    packages_created = []
    
    # Create individual packages for each architecture
    for arch, exe_path in built_executables:
        arch_clean = arch.replace("-", "").replace(" ", "").lower()
        package_name = f"DataSorterApp_{arch_clean}_Distribution.zip"
        package_path = Path("dist").parent / package_name
        
        # Create temporary folder structure for this architecture
        temp_folder = Path("dist") / f"temp_{arch_clean}"
        temp_folder.mkdir(exist_ok=True)
        
        try:
            # Copy executable to temp folder
            import shutil
            shutil.copy2(exe_path, temp_folder / exe_path.name)
            
            # Copy documentation
            for doc_file in ["Installation_Instructions.txt", "README_Distribution.md", "Package_Summary.md"]:
                doc_path = Path("dist") / doc_file
                if doc_path.exists():
                    shutil.copy2(doc_path, temp_folder / doc_file)
            
            # Create launcher batch file
            batch_content = f"""@echo off
REM Data Sorter Application Launcher ({arch})
echo Starting Data Sorter Application ({arch})...
echo.

if not exist "{exe_path.name}" (
    echo ERROR: {exe_path.name} not found!
    pause
    exit /b 1
)

start "" "{exe_path.name}"
echo Application started successfully!
timeout /t 2 /nobreak >nul
"""
            
            with open(temp_folder / f"Run_DataSorter_{arch_clean}.bat", 'w') as f:
                f.write(batch_content)
            
            # Create ZIP package
            import subprocess
            zip_cmd = [
                "powershell", "-Command",
                f"Compress-Archive -Path '{temp_folder}\\*' -DestinationPath '{package_path}' -Force"
            ]
            
            result = subprocess.run(zip_cmd, capture_output=True, text=True)
            if result.returncode == 0 and package_path.exists():
                size_mb = package_path.stat().st_size / (1024 * 1024)
                packages_created.append((arch, package_path, size_mb))
                print(f"📦 {arch} package: {package_path.name} ({size_mb:.1f} MB)")
            
            # Clean up temp folder
            shutil.rmtree(temp_folder, ignore_errors=True)
            
        except Exception as e:
            print(f"⚠️  Failed to create {arch} package: {e}")
            # Clean up temp folder on error
            import shutil
            shutil.rmtree(temp_folder, ignore_errors=True)
    
    return packages_created

if __name__ == "__main__":
    print("Data Sorter Application - Multi-Architecture Builder")
    print("=" * 60)
    
    # Build executables for multiple architectures
    built_executables = build_multi_architecture()
    
    if built_executables:
        print(f"\n📋 Successfully built {len(built_executables)} executable(s):")
        for arch, exe_path in built_executables:
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"  ✅ {arch}: {exe_path.name} ({size_mb:.1f} MB)")
        
        # Create installation instructions
        create_installer_info(built_executables)
        
        # Create distribution packages
        print(f"\n📦 Creating distribution packages...")
        packages = create_distribution_packages(built_executables)
        
        if packages:
            print(f"\n🎉 Build and packaging completed successfully!")
            print(f"\n📁 Files created:")
            print("   Executables in 'dist' folder:")
            for arch, exe_path in built_executables:
                size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"     - {exe_path.name} ({arch}) - {size_mb:.1f} MB")
            
            print("   Distribution packages:")
            for arch, package_path, size_mb in packages:
                print(f"     - {package_path.name} ({arch}) - {size_mb:.1f} MB")
            
            print("   Documentation:")
            print("     - Installation_Instructions.txt")
            print("     - README_Distribution.md") 
            print("     - Package_Summary.md")
            
        else:
            print("\n⚠️  Executables built but package creation failed!")
    else:
        print("\n💥 Build process failed!")
        sys.exit(1)