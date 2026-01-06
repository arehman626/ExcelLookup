#!/usr/bin/env python3
"""
Example script demonstrating ExcelLookup usage.

This script shows how to use the ExcelLookup library to compare two Excel files
and generate a detailed comparison report.
"""

import sys
import logging
from pathlib import Path

# Add the src directory to the path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from excel_lookup import ExcelFileHandler, DataComparator, ReportGenerator
from excel_lookup.core.comparator import ComparisonConfig
from excel_lookup.utils.config import Config


def main():
    """Run the example comparison."""
    
    # Setup logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger(__name__)
    
    # Create example data if it doesn't exist
    create_sample_files()
    
    try:
        # Initialize components
        logger.info("Initializing ExcelLookup components...")
        file_handler = ExcelFileHandler()
        report_generator = ReportGenerator()
        
        # Read the sample files
        logger.info("Reading sample files...")
        file1_path = Path(__file__).parent / 'sample_file1.xlsx'
        file2_path = Path(__file__).parent / 'sample_file2.xlsx'
        
        df1 = file_handler.read_file(file1_path)
        df2 = file_handler.read_file(file2_path)
        
        logger.info(f"File 1: {len(df1)} rows, {len(df1.columns)} columns")
        logger.info(f"File 2: {len(df2)} rows, {len(df2.columns)} columns")
        
        # Configure comparison
        config = ComparisonConfig(
            case_sensitive=False,
            ignore_whitespace=True,
            numeric_tolerance=0.01,
            key_columns=["ID"],  # Use ID as the key column
            ignore_columns=["LastUpdated"]  # Ignore timestamp columns
        )
        
        # Perform comparison
        logger.info("Performing data comparison...")
        comparator = DataComparator(config)
        result = comparator.compare_datasets(df1, df2, "SampleFile1", "SampleFile2")
        
        # Display results
        logger.info("\n=== Comparison Results ===")
        logger.info(f"Matching records: {len(result.matches)}")
        logger.info(f"Different records: {len(result.differences)}")
        logger.info(f"Only in File 1: {len(result.only_in_left)}")
        logger.info(f"Only in File 2: {len(result.only_in_right)}")
        logger.info(f"Match percentage: {result.summary['match_percentage']:.2f}%")
        
        # Generate report
        output_path = Path(__file__).parent / 'comparison_report.xlsx'
        logger.info(f"Generating report: {output_path}")
        
        report_generator.generate_comparison_report(
            result,
            str(output_path),
            "sample_file1.xlsx",
            "sample_file2.xlsx"
        )
        
        logger.info(f"\nExample completed successfully!")
        logger.info(f"Report saved to: {output_path}")
        logger.info(f"You can open the report in Excel to view the detailed comparison results.")
        
    except Exception as e:
        logger.error(f"Error during example execution: {e}")
        sys.exit(1)


def create_sample_files():
    """Create sample Excel files for demonstration."""
    import pandas as pd
    from datetime import datetime, timedelta
    
    # Sample data for file 1
    data1 = {
        'ID': [1, 2, 3, 4, 5],
        'Name': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown', 'Charlie Wilson'],
        'Department': ['IT', 'HR', 'Finance', 'IT', 'Marketing'],
        'Salary': [75000, 65000, 80000, 72000, 68000],
        'StartDate': [
            datetime(2020, 1, 15),
            datetime(2019, 6, 1),
            datetime(2021, 3, 10),
            datetime(2020, 9, 5),
            datetime(2022, 2, 20)
        ],
        'LastUpdated': [datetime.now() - timedelta(days=i) for i in range(5)]
    }
    
    # Sample data for file 2 (with some differences)
    data2 = {
        'ID': [1, 2, 3, 4, 6],  # ID 5 removed, ID 6 added
        'Name': ['John Doe', 'Jane Smith', 'Robert Johnson', 'Alice Brown', 'David Lee'],  # Bob -> Robert, new person
        'Department': ['IT', 'HR', 'Finance', 'Engineering', 'Sales'],  # Alice's dept changed, new dept
        'Salary': [77000, 65000, 80000, 75000, 70000],  # John's salary increased, Alice's salary increased
        'StartDate': [
            datetime(2020, 1, 15),
            datetime(2019, 6, 1),
            datetime(2021, 3, 10),
            datetime(2020, 9, 5),
            datetime(2023, 1, 10)
        ],
        'LastUpdated': [datetime.now() - timedelta(days=i+10) for i in range(5)]
    }
    
    # Create DataFrames
    df1 = pd.DataFrame(data1)
    df2 = pd.DataFrame(data2)
    
    # Save to Excel files
    file1_path = Path(__file__).parent / 'sample_file1.xlsx'
    file2_path = Path(__file__).parent / 'sample_file2.xlsx'
    
    df1.to_excel(file1_path, index=False)
    df2.to_excel(file2_path, index=False)
    
    print(f"Created sample files:")
    print(f"  - {file1_path}")
    print(f"  - {file2_path}")


if __name__ == '__main__':
    main()