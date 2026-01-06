"""
Custom exceptions for the ExcelLookup package.
"""


class ExcelLookupError(Exception):
    """Base exception for ExcelLookup package."""
    pass


class FileHandlerError(ExcelLookupError):
    """Exception raised for errors in file handling operations."""
    pass


class ComparisonError(ExcelLookupError):
    """Exception raised for errors in data comparison operations."""
    pass


class ReconciliationError(ExcelLookupError):
    """Exception raised for errors in reconciliation operations."""
    pass


class ConfigurationError(ExcelLookupError):
    """Exception raised for configuration-related errors."""
    pass


class ReportGenerationError(ExcelLookupError):
    """Exception raised for errors in report generation."""
    pass