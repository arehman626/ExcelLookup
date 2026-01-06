"""
Report Generator Module

Generates comprehensive reports from comparison and reconciliation results
in various formats including Excel, CSV, and HTML.
"""

import pandas as pd
from pathlib import Path
import xlsxwriter
from typing import Optional, Dict, Any
import logging
from datetime import datetime
from ..core.comparator import ComparisonResult
from ..utils.exceptions import ReportGenerationError


class ReportGenerator:
    """
    Generates detailed reports from comparison and reconciliation results.
    
    Supports multiple output formats and customizable report templates.
    """
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """Initialize the report generator."""
        self.logger = logger or logging.getLogger(__name__)
    
    def generate_comparison_report(self, 
                                 result: ComparisonResult,
                                 output_path: str,
                                 file1_name: str,
                                 file2_name: str,
                                 template: str = 'standard') -> None:
        """
        Generate a comprehensive comparison report.
        
        Args:
            result: Comparison results
            output_path: Path for output file
            file1_name: Name of first file
            file2_name: Name of second file
            template: Report template to use
        """
        try:
            output_path = Path(output_path)
            
            if output_path.suffix.lower() == '.xlsx':
                self._generate_excel_report(result, output_path, file1_name, file2_name)
            elif output_path.suffix.lower() == '.html':
                self._generate_html_report(result, output_path, file1_name, file2_name)
            else:
                # Default to Excel
                output_path = output_path.with_suffix('.xlsx')
                self._generate_excel_report(result, output_path, file1_name, file2_name)
            
            self.logger.info(f"Report generated successfully: {output_path}")
            
        except Exception as e:
            raise ReportGenerationError(f"Error generating report: {str(e)}")
    
    def _sanitize_sheet_name(self, name: str, max_length: int = 31) -> str:
        """
        Sanitize sheet name to comply with Excel requirements.
        
        Args:
            name: Original sheet name
            max_length: Maximum allowed length (Excel limit is 31)
            
        Returns:
            Sanitized sheet name
        """
        # Remove invalid characters
        invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
        for char in invalid_chars:
            name = name.replace(char, '_')
        
        # Truncate if too long
        if len(name) > max_length:
            name = name[:max_length]
        
        return name
    
    def _generate_excel_report(self, 
                             result: ComparisonResult,
                             output_path: Path,
                             file1_name: str,
                             file2_name: str) -> None:
        """Generate Excel format report."""
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D7E4BC',
                'border': 1
            })
            
            summary_format = workbook.add_format({
                'bold': True,
                'bg_color': '#F2F2F2'
            })
            
            # Summary sheet
            self._create_summary_sheet(writer, result, file1_name, file2_name, 
                                     header_format, summary_format)
            
            # Matches sheet
            if not result.matches.empty:
                result.matches.to_excel(writer, sheet_name='Matches', index=False)
                worksheet = writer.sheets['Matches']
                self._format_worksheet(worksheet, header_format)
            
            # Differences sheet
            if not result.differences.empty:
                result.differences.to_excel(writer, sheet_name='Differences', index=False)
                worksheet = writer.sheets['Differences']
                self._format_worksheet(worksheet, header_format)
            
            # Only in first file
            if not result.only_in_left.empty:
                sheet_name = self._sanitize_sheet_name(f'Only_in_{Path(file1_name).stem}')
                result.only_in_left.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                self._format_worksheet(worksheet, header_format)
            
            # Only in second file
            if not result.only_in_right.empty:
                sheet_name = self._sanitize_sheet_name(f'Only_in_{Path(file2_name).stem}')
                result.only_in_right.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                self._format_worksheet(worksheet, header_format)
    
    def _create_summary_sheet(self, writer, result: ComparisonResult, 
                            file1_name: str, file2_name: str,
                            header_format, summary_format) -> None:
        """Create summary sheet with comparison statistics."""
        summary_data = [
            ['Comparison Summary', ''],
            ['Generated on', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            ['', ''],
            ['File 1', file1_name],
            ['File 2', file2_name],
            ['', ''],
            ['Results', ''],
            ['Total records in file 1', result.summary['total_left_records']],
            ['Total records in file 2', result.summary['total_right_records']],
            ['Matching records', result.summary['matching_records']],
            ['Different records', result.summary['different_records']],
            ['Only in file 1', result.summary['only_in_left']],
            ['Only in file 2', result.summary['only_in_right']],
            ['Match percentage', f"{result.summary['match_percentage']:.2f}%"],
            ['', ''],
            ['Configuration', ''],
            ['Case sensitive', result.summary['comparison_config']['case_sensitive']],
            ['Ignore whitespace', result.summary['comparison_config']['ignore_whitespace']],
            ['Numeric tolerance', result.summary['comparison_config']['numeric_tolerance']],
            ['Date tolerance (days)', result.summary['comparison_config']['date_tolerance_days']]
        ]
        
        summary_df = pd.DataFrame(summary_data, columns=['Metric', 'Value'])
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        worksheet = writer.sheets['Summary']
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 30)
        
        # Apply formatting to headers
        for row, (metric, value) in enumerate(summary_data):
            if metric in ['Comparison Summary', 'Results', 'Configuration']:
                worksheet.write(row + 1, 0, metric, header_format)
    
    def _format_worksheet(self, worksheet, header_format) -> None:
        """Apply formatting to worksheet."""
        # Auto-adjust column widths
        for col_num in range(20):  # Adjust first 20 columns
            worksheet.set_column(col_num, col_num, 15)
        
        # Format header row
        for col_num in range(20):
            worksheet.write(0, col_num, '', header_format)
    
    def _generate_html_report(self, 
                            result: ComparisonResult,
                            output_path: Path,
                            file1_name: str,
                            file2_name: str) -> None:
        """Generate HTML format report."""
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>ExcelLookup Comparison Report</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                h1, h2 {{ color: #333; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .summary {{ background-color: #f9f9f9; padding: 20px; border-radius: 5px; }}
            </style>
        </head>
        <body>
            <h1>Excel Comparison Report</h1>
            <div class="summary">
                <h2>Summary</h2>
                <p><strong>Generated on:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                <p><strong>File 1:</strong> {file1_name}</p>
                <p><strong>File 2:</strong> {file2_name}</p>
                <p><strong>Total records in file 1:</strong> {result.summary['total_left_records']}</p>
                <p><strong>Total records in file 2:</strong> {result.summary['total_right_records']}</p>
                <p><strong>Matching records:</strong> {result.summary['matching_records']}</p>
                <p><strong>Different records:</strong> {result.summary['different_records']}</p>
                <p><strong>Only in file 1:</strong> {result.summary['only_in_left']}</p>
                <p><strong>Only in file 2:</strong> {result.summary['only_in_right']}</p>
                <p><strong>Match percentage:</strong> {result.summary['match_percentage']:.2f}%</p>
            </div>
        """
        
        # Add tables for each result set
        if not result.matches.empty:
            html_content += f"<h2>Matches ({len(result.matches)} records)</h2>\n"
            html_content += result.matches.head(100).to_html(index=False, classes='data-table')
        
        if not result.differences.empty:
            html_content += f"<h2>Differences ({len(result.differences)} records)</h2>\n"
            html_content += result.differences.head(100).to_html(index=False, classes='data-table')
        
        html_content += "</body></html>"
        
        output_path.write_text(html_content, encoding='utf-8')