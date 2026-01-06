"""
Tests for file handler module.
"""

import pytest
import pandas as pd
from pathlib import Path
import tempfile
from src.excel_lookup.core.file_handler import ExcelFileHandler
from src.excel_lookup.utils.exceptions import FileHandlerError


class TestExcelFileHandler:
    """Test cases for ExcelFileHandler class."""
    
    def setup_method(self):
        """Setup for each test method."""
        self.handler = ExcelFileHandler()
        self.temp_dir = Path(tempfile.mkdtemp())
        
        # Create sample data
        self.sample_data = pd.DataFrame({
            'Name': ['John', 'Jane', 'Bob'],
            'Age': [25, 30, 35],
            'City': ['New York', 'London', 'Tokyo']
        })
    
    def test_write_and_read_excel(self):
        """Test writing and reading Excel files."""
        file_path = self.temp_dir / 'test.xlsx'
        
        # Write file
        self.handler.write_file(self.sample_data, file_path)
        assert file_path.exists()
        
        # Read file
        result = self.handler.read_file(file_path)
        pd.testing.assert_frame_equal(result, self.sample_data)
    
    def test_write_and_read_csv(self):
        """Test writing and reading CSV files."""
        file_path = self.temp_dir / 'test.csv'
        
        # Write file
        self.handler.write_file(self.sample_data, file_path)
        assert file_path.exists()
        
        # Read file
        result = self.handler.read_file(file_path)
        pd.testing.assert_frame_equal(result, self.sample_data)
    
    def test_file_not_found(self):
        """Test handling of non-existent files."""
        with pytest.raises(FileHandlerError, match="File not found"):
            self.handler.read_file('nonexistent.xlsx')
    
    def test_unsupported_format(self):
        """Test handling of unsupported file formats."""
        file_path = self.temp_dir / 'test.txt'
        file_path.write_text('test content')
        
        with pytest.raises(FileHandlerError, match="Unsupported file format"):
            self.handler.read_file(file_path)
    
    def test_validate_file_structure(self):
        """Test file structure validation."""
        file_path = self.temp_dir / 'test.xlsx'
        self.handler.write_file(self.sample_data, file_path)
        
        # Valid validation
        result = self.handler.validate_file_structure(file_path, ['Name', 'Age'])
        assert result['valid'] is True
        assert result['info']['rows'] == 3
        assert result['info']['columns'] == 3
        
        # Invalid validation (missing columns)
        result = self.handler.validate_file_structure(file_path, ['Name', 'Salary'])
        assert result['valid'] is False
        assert 'Missing required columns' in result['errors'][0]
    
    def test_get_sheet_names_excel(self):
        """Test getting sheet names from Excel file."""
        file_path = self.temp_dir / 'test.xlsx'
        
        # Create multi-sheet Excel file
        with pd.ExcelWriter(file_path) as writer:
            self.sample_data.to_excel(writer, sheet_name='Sheet1', index=False)
            self.sample_data.to_excel(writer, sheet_name='Sheet2', index=False)
        
        sheet_names = self.handler.get_sheet_names(file_path)
        assert 'Sheet1' in sheet_names
        assert 'Sheet2' in sheet_names
    
    def test_get_sheet_names_csv(self):
        """Test getting sheet names from CSV file."""
        file_path = self.temp_dir / 'test.csv'
        self.handler.write_file(self.sample_data, file_path)
        
        sheet_names = self.handler.get_sheet_names(file_path)
        assert sheet_names == ['Sheet1']