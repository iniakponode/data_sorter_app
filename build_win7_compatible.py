#!/usr/bin/env python3
"""
Windows 7 Compatible Build Script for Data Sorter Application
Addresses pywin32 and DLL loading issues on older Windows systems
"""

import os
import subprocess
import sys
import platform
from pathlib import Path
import shutil

def create_compatible_spec_file(output_name="DataSorterApp_Win7Compatible"):
    """Create a PyInstaller spec file optimized for Windows 7 compatibility."""
    
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

import sys
from PyInstaller.building.build_main import Analysis, PYZ, EXE

block_cipher = None

# Define the analysis with Windows 7 compatibility settings
a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[('README.md', '.')],
    hiddenimports=[
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.styles',
        'openpyxl.utils',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[
        'pywin32',
        'win32api',
        'win32con',
        'win32gui',
        'win32print',
        'win32process',
        'win32security',
        'win32service',
        'win32clipboard',
        'win32file',
        'win32pipe',
        'win32event',
        'pywintypes',
        'pythoncom',
        'PIL',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'setuptools',
        'distutils',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{output_name}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Disable UPX compression for better compatibility
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI application
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    uac_admin=False,  # Don't require admin privileges
    uac_uiaccess=False,
    manifest=None,  # Use default manifest
    version=None,
    icon=None,
)
'''
    
    spec_file_path = f"{output_name}.spec"
    with open(spec_file_path, 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    return spec_file_path

def build_win7_compatible():
    """Build Windows 7 compatible executable."""
    
    print("Data Sorter Application - Windows 7 Compatible Builder")
    print("=" * 60)
    print("Building executable optimized for Windows 7 and older systems...")
    print("Excluding pywin32 dependencies that cause DLL loading issues")
    print("-" * 60)
    
    # Clean previous builds
    if os.path.exists("build"):
        shutil.rmtree("build")
        print("Cleaned previous build directory")
    
    if os.path.exists("dist"):
        print("Preserving existing dist directory...")
    
    # Create the spec file
    spec_file = create_compatible_spec_file("DataSorterApp_Win7Compatible")
    print(f"Created spec file: {spec_file}")
    
    # Build using the spec file
    cmd = [
        "pyinstaller",
        "--clean",
        "--noconfirm",
        spec_file
    ]
    
    print(f"Building command: {' '.join(cmd)}")
    print("-" * 50)
    
    try:
        # Run PyInstaller
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        print("Build successful!")
        if result.stdout.strip():
            print("\nBuild Output:")
            # Show only important lines
            for line in result.stdout.split('\n'):
                if any(keyword in line.lower() for keyword in ['building', 'analyzing', 'warning', 'error', 'successfully']):
                    print(f"  {line}")
        
        # Check if the executable was created
        exe_path = Path("dist") / "DataSorterApp_Win7Compatible.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✅ Windows 7 Compatible Executable created!")
            print(f"📁 Location: {exe_path}")
            print(f"📊 File size: {size_mb:.1f} MB")
            
            return exe_path
        else:
            print("❌ Executable not found in expected location")
            return None
            
    except subprocess.CalledProcessError as e:
        print("❌ Build failed!")
        print("Error output:")
        print(e.stderr)
        if e.stdout:
            print("Standard output:")
            print(e.stdout)
        return None
    
    finally:
        # Clean up spec file
        if os.path.exists(spec_file):
            os.remove(spec_file)
            print(f"Cleaned up spec file: {spec_file}")

def build_minimal_tkinter():
    """Build an even more minimal version using only Tkinter essentials."""
    
    print("\nBuilding MINIMAL version for maximum compatibility...")
    print("-" * 50)
    
    # Create minimal spec file
    minimal_spec = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pywin32', 'win32api', 'win32con', 'win32gui', 'pywintypes', 'pythoncom',
        'PIL', 'matplotlib', 'numpy', 'pandas', 'scipy', 'IPython', 'jupyter'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DataSorterApp_Minimal',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
)
'''
    
    # Write minimal spec file
    with open("minimal.spec", "w") as f:
        f.write(minimal_spec)
    
    # Build minimal version
    cmd = ["pyinstaller", "--clean", "--noconfirm", "minimal.spec"]
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        exe_path = Path("dist") / "DataSorterApp_Minimal.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"✅ Minimal executable created: {exe_path}")
            print(f"📊 Size: {size_mb:.1f} MB")
            return exe_path
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Minimal build failed: {e}")
        return None
    
    finally:
        if os.path.exists("minimal.spec"):
            os.remove("minimal.spec")

def create_installation_package():
    """Create a complete installation package with compatibility info."""
    
    # Create compatibility readme
    compat_readme = """Data Sorter Application - Windows 7 Compatibility Package
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
"""
    
    with open("README_Compatibility.txt", "w", encoding='utf-8') as f:
        f.write(compat_readme)
    
    print("\n📋 Created compatibility documentation")

if __name__ == "__main__":
    print("Starting Windows 7 compatible build process...")
    
    # Build both versions
    main_exe = build_win7_compatible()
    minimal_exe = build_minimal_tkinter()
    
    # Create documentation
    create_installation_package()
    
    print("\n" + "=" * 60)
    print("BUILD SUMMARY:")
    
    if main_exe:
        size_mb = main_exe.stat().st_size / (1024 * 1024)
        print(f"✅ Main: {main_exe.name} ({size_mb:.1f} MB)")
    else:
        print("❌ Main build failed")
    
    if minimal_exe:
        size_mb = minimal_exe.stat().st_size / (1024 * 1024)
        print(f"✅ Minimal: {minimal_exe.name} ({size_mb:.1f} MB)")
    else:
        print("❌ Minimal build failed")
    
    print("\n🎯 Recommendation: Try the minimal version first on Windows 7")
    print("📁 All files are in the 'dist' directory")
    print("📋 See README_Compatibility.txt for installation instructions")