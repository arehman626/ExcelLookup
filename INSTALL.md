# Installation Guide

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation Steps

### 1. Clone the Repository

```bash
git clone https://github.com/arehman626/ExcelLookup.git
cd ExcelLookup
```

### 2. Install Dependencies

#### Option A: Using pip (Recommended)
```bash
pip install -r requirements.txt
```

#### Option B: Using pip with virtual environment
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Verify Installation

Run the test script to verify everything is working:
```bash
python test_excel_lookup.py
```

### 4. Run the Application

#### Option A: Using the quick start script
```bash
python run.py
```

#### Option B: Direct launch
```bash
python excel_lookup_gui.py
```

## Troubleshooting

### tkinter not found (Linux)

If you encounter a "No module named 'tkinter'" error on Linux:

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install python3-tk
```

**Fedora:**
```bash
sudo dnf install python3-tkinter
```

**Arch Linux:**
```bash
sudo pacman -S tk
```

### tkinter not found (macOS)

tkinter should be included with Python on macOS. If it's missing:

1. Reinstall Python from python.org (not Homebrew)
2. Or use Homebrew with tkinter support:
```bash
brew install python-tk@3.11
```

### tkinter not found (Windows)

tkinter should be included with Python on Windows. If it's missing:

1. Reinstall Python from python.org
2. Make sure to check "tcl/tk and IDLE" during installation

### Dependencies installation fails

If pip fails to install dependencies:

1. Upgrade pip:
```bash
pip install --upgrade pip
```

2. Try installing packages individually:
```bash
pip install pandas
pip install openpyxl
pip install xlrd
```

### Excel file format errors

- **For .xls files**: Ensure xlrd is installed (it's in requirements.txt)
- **For .xlsx files**: Ensure openpyxl is installed (it's in requirements.txt)

## System Requirements

- **RAM**: Minimum 2GB (4GB+ recommended for large files)
- **Disk Space**: 100MB for application and dependencies
- **Display**: Any resolution supporting 1200x700 pixels or higher
- **Operating System**: 
  - Windows 7 or higher
  - macOS 10.12 or higher
  - Linux (any modern distribution)

## Getting Help

If you continue to experience issues:

1. Check the [README.md](README.md) for usage instructions
2. Verify your Python version: `python --version`
3. Check installed packages: `pip list`
4. Create an issue on GitHub with:
   - Your operating system and version
   - Python version
   - Error message (if any)
   - Steps to reproduce the issue
