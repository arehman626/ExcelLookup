"""
Setup script for ExcelLookup package.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding='utf-8')

setup(
    name="excel-lookup",
    version="1.0.0",
    author="Abdul Rehman",
    author_email="arehman626@gmail.com",
    description="Comprehensive Excel Lookup and Reconciliation Utility",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/arehman626/ExcelLookup",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Utilities",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
        "xlsxwriter>=3.1.0",
        "numpy>=1.24.0",
        "matplotlib>=3.7.0",
        "seaborn>=0.12.0",
        "tabulate>=0.9.0",
        "colorama>=0.4.6",
        "tqdm>=4.65.0",
        "click>=8.1.0",
        "python-dotenv>=1.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.4.0",
            "pytest-cov>=4.1.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.5.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "excellookup=main:cli",
        ],
    },
    include_package_data=True,
    package_data={
        "excel_lookup": ["config/*.json", "templates/*.html"],
    },
)