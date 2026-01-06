"""
Data Comparator Module

Handles comparison logic between datasets with various matching algorithms
and tolerance settings for different data types.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Union, Any, Tuple, Set
from dataclasses import dataclass
import logging
from ..utils.exceptions import ComparisonError


@dataclass
class ComparisonConfig:
    """Configuration for data comparison operations."""
    case_sensitive: bool = False
    ignore_whitespace: bool = True
    numeric_tolerance: float = 0.001
    date_tolerance_days: int = 0
    ignore_columns: Optional[List[str]] = None
    key_columns: Optional[List[str]] = None
    comparison_columns: Optional[List[str]] = None


@dataclass 
class ComparisonResult:
    """Result of data comparison operation."""
    matches: pd.DataFrame
    differences: pd.DataFrame
    only_in_left: pd.DataFrame
    only_in_right: pd.DataFrame
    summary: Dict[str, Any]


class DataComparator:
    """
    Handles comparison of datasets with configurable matching logic.
    
    Supports various data types including strings, numbers, and dates
    with customizable tolerance levels and matching criteria.
    """
    
    def __init__(self, config: Optional[ComparisonConfig] = None, 
                 logger: Optional[logging.Logger] = None):
        """Initialize the data comparator."""
        self.config = config or ComparisonConfig()
        self.logger = logger or logging.getLogger(__name__)
    
    def compare_datasets(self, 
                        left_df: pd.DataFrame, 
                        right_df: pd.DataFrame,
                        left_name: str = "Left",
                        right_name: str = "Right") -> ComparisonResult:
        """
        Compare two datasets and return detailed comparison results.
        
        Args:
            left_df: First dataset to compare
            right_df: Second dataset to compare
            left_name: Name for the first dataset
            right_name: Name for the second dataset
            
        Returns:
            ComparisonResult: Detailed comparison results
            
        Raises:
            ComparisonError: If comparison cannot be performed
        """
        try:
            self.logger.info(f"Starting comparison between {left_name} and {right_name}")
            
            # Validate inputs
            self._validate_dataframes(left_df, right_df)
            
            # Prepare data for comparison
            left_prepared = self._prepare_dataframe(left_df.copy(), left_name)
            right_prepared = self._prepare_dataframe(right_df.copy(), right_name)
            
            # Determine key columns for matching
            key_cols = self._determine_key_columns(left_prepared, right_prepared)
            
            # Perform the comparison
            matches, differences, only_left, only_right = self._perform_comparison(
                left_prepared, right_prepared, key_cols, left_name, right_name
            )
            
            # Generate summary statistics
            summary = self._generate_summary(
                left_df, right_df, matches, differences, only_left, only_right
            )
            
            self.logger.info(f"Comparison completed. Found {len(matches)} matches, "
                           f"{len(differences)} differences")
            
            return ComparisonResult(
                matches=matches,
                differences=differences,
                only_in_left=only_left,
                only_in_right=only_right,
                summary=summary
            )
            
        except Exception as e:
            raise ComparisonError(f"Error during dataset comparison: {str(e)}")
    
    def _validate_dataframes(self, left_df: pd.DataFrame, right_df: pd.DataFrame) -> None:
        """Validate input dataframes."""
        if left_df.empty or right_df.empty:
            raise ComparisonError("Cannot compare empty dataframes")
        
        if self.config.key_columns:
            missing_left = set(self.config.key_columns) - set(left_df.columns)
            missing_right = set(self.config.key_columns) - set(right_df.columns)
            
            if missing_left:
                raise ComparisonError(f"Key columns missing in left dataset: {missing_left}")
            if missing_right:
                raise ComparisonError(f"Key columns missing in right dataset: {missing_right}")
    
    def _prepare_dataframe(self, df: pd.DataFrame, name: str) -> pd.DataFrame:
        """Prepare dataframe for comparison by cleaning and formatting."""
        # Add source column
        df['_source'] = name
        
        # Handle string columns
        string_cols = df.select_dtypes(include=['object']).columns
        for col in string_cols:
            if col != '_source':
                if not self.config.case_sensitive:
                    df[col] = df[col].astype(str).str.lower()
                if self.config.ignore_whitespace:
                    df[col] = df[col].astype(str).str.strip()
        
        # Remove ignored columns
        if self.config.ignore_columns:
            cols_to_drop = [col for col in self.config.ignore_columns if col in df.columns]
            df = df.drop(columns=cols_to_drop)
        
        return df
    
    def _determine_key_columns(self, left_df: pd.DataFrame, right_df: pd.DataFrame) -> List[str]:
        """Determine which columns to use as keys for matching."""
        if self.config.key_columns:
            return self.config.key_columns
        
        # Use common columns as keys
        common_cols = list(set(left_df.columns) & set(right_df.columns))
        common_cols = [col for col in common_cols if col != '_source']
        
        if not common_cols:
            raise ComparisonError("No common columns found for comparison")
        
        self.logger.warning(f"No key columns specified, using all common columns: {common_cols}")
        return common_cols
    
    def _perform_comparison(self, 
                          left_df: pd.DataFrame, 
                          right_df: pd.DataFrame,
                          key_cols: List[str],
                          left_name: str,
                          right_name: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Perform the actual comparison logic."""
        # Merge dataframes on key columns
        merged = pd.merge(
            left_df, right_df, 
            on=key_cols, 
            how='outer', 
            suffixes=(f'_{left_name}', f'_{right_name}'),
            indicator=True
        )
        
        # Separate results
        matches = merged[merged['_merge'] == 'both'].copy()
        only_left = merged[merged['_merge'] == 'left_only'].copy()
        only_right = merged[merged['_merge'] == 'right_only'].copy()
        
        # Find differences in matching records
        differences = self._find_differences_in_matches(matches, left_name, right_name)
        
        # Clean up the dataframes
        for df in [matches, differences, only_left, only_right]:
            if '_merge' in df.columns:
                df.drop(columns=['_merge'], inplace=True)
        
        return matches, differences, only_left, only_right
    
    def _find_differences_in_matches(self, 
                                   matches: pd.DataFrame,
                                   left_name: str,
                                   right_name: str) -> pd.DataFrame:
        """Find differences within matched records."""
        if matches.empty:
            return pd.DataFrame()
        
        differences_list = []
        
        # Get columns to compare (excluding key columns and source columns)
        left_cols = [col for col in matches.columns if col.endswith(f'_{left_name}')]
        right_cols = [col for col in matches.columns if col.endswith(f'_{right_name}')]
        
        # Compare corresponding columns
        base_cols = set(col.replace(f'_{left_name}', '') for col in left_cols)
        base_cols &= set(col.replace(f'_{right_name}', '') for col in right_cols)
        
        for base_col in base_cols:
            left_col = f"{base_col}_{left_name}"
            right_col = f"{base_col}_{right_name}"
            
            if left_col in matches.columns and right_col in matches.columns:
                # Compare values with appropriate tolerance
                mask = self._compare_columns(
                    matches[left_col], matches[right_col], base_col
                )
                
                if mask.any():
                    diff_rows = matches[mask].copy()
                    diff_rows['_difference_column'] = base_col
                    differences_list.append(diff_rows)
        
        if differences_list:
            return pd.concat(differences_list, ignore_index=True)
        else:
            return pd.DataFrame()
    
    def _compare_columns(self, left_series: pd.Series, right_series: pd.Series, col_name: str) -> pd.Series:
        """Compare two series with appropriate tolerance based on data type."""
        # Handle NaN values
        both_null = left_series.isnull() & right_series.isnull()
        one_null = left_series.isnull() ^ right_series.isnull()
        
        # For numeric columns, use numeric tolerance
        if pd.api.types.is_numeric_dtype(left_series) and pd.api.types.is_numeric_dtype(right_series):
            numeric_diff = np.abs(left_series - right_series) > self.config.numeric_tolerance
            return (numeric_diff & ~both_null) | one_null
        
        # For datetime columns, use date tolerance
        elif pd.api.types.is_datetime64_any_dtype(left_series) and pd.api.types.is_datetime64_any_dtype(right_series):
            date_diff = np.abs((left_series - right_series).dt.days) > self.config.date_tolerance_days
            return (date_diff & ~both_null) | one_null
        
        # For other types, use exact comparison
        else:
            return (left_series != right_series) & ~both_null | one_null
    
    def _generate_summary(self, 
                         left_df: pd.DataFrame,
                         right_df: pd.DataFrame, 
                         matches: pd.DataFrame,
                         differences: pd.DataFrame,
                         only_left: pd.DataFrame,
                         only_right: pd.DataFrame) -> Dict[str, Any]:
        """Generate summary statistics for the comparison."""
        return {
            'total_left_records': len(left_df),
            'total_right_records': len(right_df),
            'matching_records': len(matches),
            'different_records': len(differences),
            'only_in_left': len(only_left),
            'only_in_right': len(only_right),
            'match_percentage': len(matches) / max(len(left_df), len(right_df)) * 100 if max(len(left_df), len(right_df)) > 0 else 0,
            'comparison_config': {
                'case_sensitive': self.config.case_sensitive,
                'ignore_whitespace': self.config.ignore_whitespace,
                'numeric_tolerance': self.config.numeric_tolerance,
                'date_tolerance_days': self.config.date_tolerance_days
            }
        }