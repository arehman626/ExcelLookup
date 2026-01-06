"""
ExcelLookup - Comprehensive Excel Lookup and Reconciliation Utility

This package provides tools for comparing and reconciling data between Excel files,
offering comprehensive matching algorithms and detailed reporting capabilities.
"""

__version__ = "1.0.0"
__author__ = "Abdul Rehman"
__email__ = "arehman626@gmail.com"

from .core.file_handler import ExcelFileHandler
from .core.comparator import DataComparator
from .core.reconciler import DataReconciler
from .utils.config import Config
from .reports.generator import ReportGenerator

__all__ = [
    "ExcelFileHandler",
    "DataComparator", 
    "DataReconciler",
    "Config",
    "ReportGenerator"
]