#!/usr/bin/env python3
"""
Quick start script for Excel Lookup GUI
"""

import sys
import os

def check_dependencies():
    """Check if required dependencies are installed"""
    missing = []
    
    try:
        import pandas
    except ImportError:
        missing.append('pandas')
    
    try:
        import openpyxl
    except ImportError:
        missing.append('openpyxl')
    
    try:
        import xlrd
    except ImportError:
        missing.append('xlrd')
    
    try:
        import tkinter
    except ImportError:
        print("\nWarning: tkinter is not available.")
        print("On Linux, you may need to install it:")
        print("  - Ubuntu/Debian: sudo apt-get install python3-tk")
        print("  - Fedora: sudo dnf install python3-tkinter")
        print("  - macOS: tkinter should be included with Python")
        print("  - Windows: tkinter should be included with Python")
        return False
    
    if missing:
        print("\nMissing required packages:", ", ".join(missing))
        print("\nPlease install dependencies:")
        print("  pip install -r requirements.txt")
        return False
    
    return True


def main():
    """Main entry point"""
    print("="*60)
    print("Excel Lookup - Data Comparison Tool")
    print("="*60)
    print()
    
    if not check_dependencies():
        print("\nPlease install missing dependencies and try again.")
        sys.exit(1)
    
    print("✓ All dependencies installed")
    print("\nStarting application...\n")
    
    # Import and run the GUI
    from excel_lookup_gui import main as run_gui
    run_gui()


if __name__ == "__main__":
    main()
