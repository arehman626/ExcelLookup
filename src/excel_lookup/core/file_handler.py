"""
Excel File Handler Module

Handles reading, writing, and basic operations on Excel files.
Supports multiple file formats and provides data validation.
"""

import pandas as pd
import openpyxl
from pathlib import Path
from typing import Dict, List, Optional, Union, Any
import logging
from ..utils.exceptions import FileHandlerError


class ExcelFileHandler:
    """
    Handles Excel file operations including reading, writing, and validation.
    
    Supports .xlsx, .xls, and .csv formats with comprehensive error handling
    and data validation capabilities.
    """
    
    SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.csv'}
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """Initialize the Excel file handler."""
        self.logger = logger or logging.getLogger(__name__)
        self._validate_dependencies()
    
    def _validate_dependencies(self) -> None:
        """Validate that required dependencies are available."""
        try:
            import pandas as pd
            import openpyxl
        except ImportError as e:
            raise FileHandlerError(f"Missing required dependency: {e}")
    
    def read_file(self, file_path: Union[str, Path], 
                  sheet_name: Optional[str] = None,
                  **kwargs) -> pd.DataFrame:
        """
        Read an Excel file and return a pandas DataFrame.
        
        Args:
            file_path: Path to the Excel file
            sheet_name: Name of the sheet to read (optional)
            **kwargs: Additional arguments passed to pandas read functions
            
        Returns:
            pandas.DataFrame: The loaded data
            
        Raises:
            FileHandlerError: If file cannot be read or processed
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileHandlerError(f"File not found: {file_path}")
        
        if file_path.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            raise FileHandlerError(
                f"Unsupported file format: {file_path.suffix}. "
                f"Supported formats: {', '.join(self.SUPPORTED_EXTENSIONS)}"
            )
        
        try:
            if file_path.suffix.lower() == '.csv':
                return pd.read_csv(file_path, **kwargs)
            else:
                if sheet_name is not None:
                    return pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
                else:
                    return pd.read_excel(file_path, **kwargs)
                
        except Exception as e:
            raise FileHandlerError(f"Error reading file {file_path}: {str(e)}")
    
    def get_sheet_names(self, file_path: Union[str, Path]) -> List[str]:
        """
        Get list of sheet names in an Excel file.
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            List[str]: List of sheet names
            
        Raises:
            FileHandlerError: If file cannot be processed
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileHandlerError(f"File not found: {file_path}")
        
        if file_path.suffix.lower() == '.csv':
            return ['Sheet1']  # CSV files have only one "sheet"
        
        try:
            with pd.ExcelFile(file_path) as xl:
                return xl.sheet_names
        except Exception as e:
            raise FileHandlerError(f"Error reading sheet names from {file_path}: {str(e)}")
    
    def write_file(self, data: pd.DataFrame, 
                   file_path: Union[str, Path],
                   sheet_name: str = 'Sheet1',
                   **kwargs) -> None:
        """
        Write a pandas DataFrame to an Excel file.
        
        Args:
            data: DataFrame to write
            file_path: Output file path
            sheet_name: Name of the sheet to write to
            **kwargs: Additional arguments passed to pandas write functions
            
        Raises:
            FileHandlerError: If file cannot be written
        """
        file_path = Path(file_path)
        
        # Create directory if it doesn't exist
        file_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            if file_path.suffix.lower() == '.csv':
                data.to_csv(file_path, index=False, **kwargs)
            else:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    data.to_excel(writer, sheet_name=sheet_name, index=False, **kwargs)
                    
        except Exception as e:
            raise FileHandlerError(f"Error writing file {file_path}: {str(e)}")
    
    def validate_file_structure(self, file_path: Union[str, Path], 
                              required_columns: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        Validate the structure of an Excel file.
        
        Args:
            file_path: Path to the Excel file
            required_columns: List of required column names
            
        Returns:
            Dict with validation results
        """
        validation_result = {
            'valid': True,
            'errors': [],
            'warnings': [],
            'info': {}
        }
        
        try:
            df = self.read_file(file_path)
            
            # Basic info
            validation_result['info'] = {
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': df.columns.tolist(),
                'has_duplicates': df.duplicated().any(),
                'empty_cells': df.isnull().sum().sum()
            }
            
            # Check required columns
            if required_columns:
                missing_columns = set(required_columns) - set(df.columns)
                if missing_columns:
                    validation_result['valid'] = False
                    validation_result['errors'].append(
                        f"Missing required columns: {', '.join(missing_columns)}"
                    )
            
            # Check for empty DataFrame
            if df.empty:
                validation_result['warnings'].append("File contains no data")
            
            self.logger.info(f"Validated file {file_path}: {validation_result['info']}")
            
        except Exception as e:
            validation_result['valid'] = False
            validation_result['errors'].append(f"Validation error: {str(e)}")
        
        return validation_result