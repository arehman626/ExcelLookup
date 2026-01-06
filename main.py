"""
Main CLI entry point for ExcelLookup application.
"""

import click
import logging
import sys
from pathlib import Path
from typing import Optional

from excel_lookup import ExcelFileHandler, DataComparator, Config, ReportGenerator
from excel_lookup.core.comparator import ComparisonConfig
from excel_lookup.utils.exceptions import ExcelLookupError


@click.group()
@click.option('--verbose', '-v', is_flag=True, help='Enable verbose logging')
@click.option('--config-file', '-c', type=click.Path(exists=True), help='Configuration file path')
@click.pass_context
def cli(ctx, verbose: bool, config_file: Optional[str]):
    """
    ExcelLookup - Comprehensive Excel file comparison and reconciliation tool.
    
    Compare Excel files, identify differences, and generate detailed reports.
    """
    # Setup logging
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('excellookup.log')
        ]
    )
    
    # Initialize context
    ctx.ensure_object(dict)
    ctx.obj['config_file'] = config_file
    ctx.obj['verbose'] = verbose


@cli.command()
@click.argument('file1', type=click.Path(exists=True))
@click.argument('file2', type=click.Path(exists=True))
@click.option('--output', '-o', type=click.Path(), help='Output file path for comparison report')
@click.option('--sheet1', help='Sheet name in first file')
@click.option('--sheet2', help='Sheet name in second file')
@click.option('--key-columns', help='Comma-separated list of key columns for matching')
@click.option('--ignore-columns', help='Comma-separated list of columns to ignore')
@click.option('--case-sensitive', is_flag=True, help='Enable case-sensitive comparison')
@click.option('--numeric-tolerance', type=float, default=0.001, help='Numeric tolerance for comparisons')
@click.pass_context
def compare(ctx, file1: str, file2: str, output: Optional[str], 
           sheet1: Optional[str], sheet2: Optional[str],
           key_columns: Optional[str], ignore_columns: Optional[str],
           case_sensitive: bool, numeric_tolerance: float):
    """
    Compare two Excel files and generate a comparison report.
    
    FILE1 and FILE2 are the paths to the Excel files to compare.
    """
    try:
        logger = logging.getLogger(__name__)
        logger.info(f"Starting comparison between {file1} and {file2}")
        
        # Initialize file handler
        file_handler = ExcelFileHandler()
        
        # Read the files
        click.echo(f"Reading {file1}...")
        df1 = file_handler.read_file(file1, sheet_name=sheet1)
        
        click.echo(f"Reading {file2}...")
        df2 = file_handler.read_file(file2, sheet_name=sheet2)
        
        # Setup comparison configuration
        config = ComparisonConfig(
            case_sensitive=case_sensitive,
            numeric_tolerance=numeric_tolerance,
            key_columns=key_columns.split(',') if key_columns else None,
            ignore_columns=ignore_columns.split(',') if ignore_columns else None
        )
        
        # Perform comparison
        comparator = DataComparator(config)
        click.echo("Performing comparison...")
        
        result = comparator.compare_datasets(
            df1, df2, 
            left_name=Path(file1).stem,
            right_name=Path(file2).stem
        )
        
        # Display summary
        summary = result.summary
        click.echo("\\n=== Comparison Summary ===")
        click.echo(f"Total records in {Path(file1).stem}: {summary['total_left_records']}")
        click.echo(f"Total records in {Path(file2).stem}: {summary['total_right_records']}")
        click.echo(f"Matching records: {summary['matching_records']}")
        click.echo(f"Different records: {summary['different_records']}")
        click.echo(f"Only in {Path(file1).stem}: {summary['only_in_left']}")
        click.echo(f"Only in {Path(file2).stem}: {summary['only_in_right']}")
        click.echo(f"Match percentage: {summary['match_percentage']:.2f}%")
        
        # Generate report if output specified
        if output:
            click.echo(f"Generating report: {output}")
            report_generator = ReportGenerator()
            report_generator.generate_comparison_report(result, output, file1, file2)
            click.echo(f"Report saved to: {output}")
        
        click.echo("\\nComparison completed successfully!")
        
    except ExcelLookupError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Unexpected error: {e}", err=True)
        sys.exit(1)


@cli.command()
@click.argument('file_path', type=click.Path(exists=True))
@click.option('--sheet', help='Sheet name to validate')
@click.option('--required-columns', help='Comma-separated list of required columns')
def validate(file_path: str, sheet: Optional[str], required_columns: Optional[str]):
    """
    Validate the structure and content of an Excel file.
    
    FILE_PATH is the path to the Excel file to validate.
    """
    try:
        file_handler = ExcelFileHandler()
        
        click.echo(f"Validating file: {file_path}")
        
        required_cols = required_columns.split(',') if required_columns else None
        result = file_handler.validate_file_structure(file_path, required_cols)
        
        if result['valid']:
            click.echo("✓ File validation passed")
        else:
            click.echo("✗ File validation failed")
        
        # Display info
        info = result['info']
        click.echo(f"\\n=== File Information ===")
        click.echo(f"Rows: {info['rows']}")
        click.echo(f"Columns: {info['columns']}")
        click.echo(f"Column names: {', '.join(info['column_names'])}")
        click.echo(f"Has duplicates: {info['has_duplicates']}")
        click.echo(f"Empty cells: {info['empty_cells']}")
        
        # Display errors and warnings
        if result['errors']:
            click.echo("\\n=== Errors ===")
            for error in result['errors']:
                click.echo(f"✗ {error}")
        
        if result['warnings']:
            click.echo("\\n=== Warnings ===")
            for warning in result['warnings']:
                click.echo(f"⚠ {warning}")
                
    except ExcelLookupError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


@cli.command()
@click.argument('file_path', type=click.Path(exists=True))
def info(file_path: str):
    """
    Display information about an Excel file.
    
    FILE_PATH is the path to the Excel file to inspect.
    """
    try:
        file_handler = ExcelFileHandler()
        
        click.echo(f"File: {file_path}")
        
        # Get sheet names
        sheets = file_handler.get_sheet_names(file_path)
        click.echo(f"Sheets: {', '.join(sheets)}")
        
        # Read and display info for each sheet
        for sheet in sheets:
            if len(sheets) > 1:
                click.echo(f"\\n=== Sheet: {sheet} ===")
            
            df = file_handler.read_file(file_path, sheet_name=sheet if sheet != 'Sheet1' or len(sheets) > 1 else None)
            
            click.echo(f"Rows: {len(df)}")
            click.echo(f"Columns: {len(df.columns)}")
            click.echo(f"Column names: {', '.join(df.columns.tolist())}")
            
            # Data types
            click.echo("\\nData types:")
            for col, dtype in df.dtypes.items():
                click.echo(f"  {col}: {dtype}")
                
    except ExcelLookupError as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


if __name__ == '__main__':
    cli()