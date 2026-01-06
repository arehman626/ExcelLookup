# ExcelLookup

A comprehensive Excel Lookup and Reconciliation Utility for comparing and reconciling data between Excel files with advanced matching algorithms and detailed reporting capabilities.

## Features

- 📊 **Excel File Comparison**: Compare data between multiple Excel files (.xlsx, .xls) and CSV files
- 🔍 **Advanced Matching**: Configurable matching algorithms with support for:
  - Case-sensitive/insensitive comparison
  - Numeric tolerance for floating-point numbers  
  - Date tolerance for datetime comparisons
  - Whitespace handling options
- 📈 **Comprehensive Reports**: Generate detailed reports in Excel, HTML, and CSV formats
- 🚀 **Performance Optimized**: Handles large datasets efficiently with chunking and memory management
- 🎯 **Reconciliation Suggestions**: Automated suggestions for data reconciliation
- 🛠️ **CLI Interface**: Easy-to-use command-line interface
- 📝 **Detailed Logging**: Comprehensive logging and error handling

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Install from Source
```bash
# Clone the repository
git clone https://github.com/arehman626/ExcelLookup.git
cd ExcelLookup

# Install in development mode
pip install -e ".[dev]"
```

### Production Installation
```bash
pip install excel-lookup
```

## Quick Start

### Command Line Usage

#### Compare Two Excel Files
```bash
# Basic comparison
excellookup compare file1.xlsx file2.xlsx --output comparison_report.xlsx

# Advanced comparison with custom settings
excellookup compare file1.xlsx file2.xlsx \
  --output report.xlsx \
  --key-columns "ID,Name" \
  --ignore-columns "UpdatedDate,Version" \
  --case-sensitive \
  --numeric-tolerance 0.01
```

#### Validate File Structure
```bash
excellookup validate data.xlsx --required-columns "ID,Name,Amount"
```

#### Get File Information
```bash
excellookup info data.xlsx
```

### Python API Usage

```python
from excel_lookup import ExcelFileHandler, DataComparator, ReportGenerator
from excel_lookup.core.comparator import ComparisonConfig

# Initialize components
file_handler = ExcelFileHandler()
comparator = DataComparator()
report_generator = ReportGenerator()

# Read Excel files
df1 = file_handler.read_file("file1.xlsx")
df2 = file_handler.read_file("file2.xlsx")

# Configure comparison
config = ComparisonConfig(
    case_sensitive=False,
    numeric_tolerance=0.001,
    key_columns=["ID", "Name"],
    ignore_columns=["LastUpdated"]
)

# Perform comparison
result = comparator.compare_datasets(df1, df2, "File1", "File2")

# Generate report
report_generator.generate_comparison_report(
    result, 
    "comparison_report.xlsx",
    "file1.xlsx",
    "file2.xlsx"
)

# Access results programmatically
print(f"Matching records: {len(result.matches)}")
print(f"Different records: {len(result.differences)}")
print(f"Only in File1: {len(result.only_in_left)}")
print(f"Only in File2: {len(result.only_in_right)}")
```

## Configuration

### Environment Variables
Create a `.env` file in your project directory:

```bash
# Default directories
INPUT_DIR=./input
OUTPUT_DIR=./output

# Comparison settings
DEFAULT_TOLERANCE=0.001
CASE_SENSITIVE=false
IGNORE_WHITESPACE=true

# Report settings
REPORT_FORMAT=xlsx
INCLUDE_CHARTS=true
MAX_DIFFERENCES=1000
```

### Configuration File
Create a `config.json` file:

```json
{
  "input_dir": "./input",
  "output_dir": "./output", 
  "default_tolerance": 0.001,
  "case_sensitive": false,
  "ignore_whitespace": true,
  "report_format": "xlsx",
  "include_charts": true,
  "max_differences": 1000,
  "log_level": "INFO"
}
```

## Project Structure

```
ExcelLookup/
├── src/excel_lookup/           # Main package source
│   ├── core/                   # Core business logic
│   │   ├── file_handler.py     # File I/O operations
│   │   ├── comparator.py       # Data comparison logic
│   │   └── reconciler.py       # Reconciliation algorithms
│   ├── utils/                  # Utility modules
│   │   ├── config.py           # Configuration management
│   │   └── exceptions.py       # Custom exceptions
│   └── reports/                # Report generation
│       └── generator.py        # Report generators
├── tests/                      # Test suite
├── examples/                   # Example files and scripts
├── docs/                       # Documentation
├── config/                     # Configuration files
├── main.py                     # CLI entry point
├── requirements.txt            # Dependencies
├── setup.py                    # Package setup
└── README.md                   # This file
```

## Development

### Setup Development Environment
```bash
# Clone repository
git clone https://github.com/arehman626/ExcelLookup.git
cd ExcelLookup

# Create virtual environment (recommended)
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Install in development mode
pip install -e ".[dev]"
```

### Running Tests
```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src --cov-report=html

# Run specific test file
pytest tests/test_file_handler.py
```

### Code Quality
```bash
# Format code
black src tests

# Run linting
flake8 src tests

# Type checking
mypy src
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

Please ensure:
- All tests pass
- Code is properly formatted with Black
- Type hints are included
- Documentation is updated

## Use Cases

### Financial Reconciliation
Compare financial datasets to identify discrepancies:
```bash
excellookup compare transactions_source.xlsx transactions_target.xlsx \
  --key-columns "TransactionID" \
  --numeric-tolerance 0.01 \
  --output financial_reconciliation.xlsx
```

### Data Migration Validation
Validate data after migration between systems:
```bash
excellookup compare old_system_export.xlsx new_system_export.xlsx \
  --ignore-columns "CreatedDate,ModifiedDate" \
  --output migration_validation.xlsx
```

### Inventory Reconciliation
Compare inventory records from different sources:
```bash
excellookup compare warehouse_inventory.xlsx system_inventory.xlsx \
  --key-columns "SKU,Location" \
  --output inventory_reconciliation.xlsx
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

- 📧 Email: arehman626@gmail.com
- 🐛 Bug Reports: [GitHub Issues](https://github.com/arehman626/ExcelLookup/issues)
- 💬 Discussions: [GitHub Discussions](https://github.com/arehman626/ExcelLookup/discussions)

## Changelog

### Version 1.0.0
- Initial release
- Excel file comparison capabilities
- CLI interface
- Comprehensive reporting
- Python API
- Configuration management
