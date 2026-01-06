"""
Configuration Management Module

Handles application configuration, settings management,
and environment variable loading.
"""

import os
from pathlib import Path
from typing import Dict, Any, Optional, List
import json
from dataclasses import dataclass, asdict
from dotenv import load_dotenv
from ..utils.exceptions import ConfigurationError


@dataclass
class Config:
    """
    Main configuration class for ExcelLookup application.
    """
    # File handling settings
    input_dir: str = './input'
    output_dir: str = './output'
    temp_dir: str = './temp'
    
    # Comparison settings
    default_tolerance: float = 0.001
    case_sensitive: bool = False
    ignore_whitespace: bool = True
    date_tolerance_days: int = 0
    
    # Report settings
    report_format: str = 'xlsx'
    include_charts: bool = True
    max_differences: int = 1000
    
    # Performance settings
    chunk_size: int = 10000
    max_memory_usage: str = '1GB'
    
    # Logging settings
    log_level: str = 'INFO'
    log_file: str = 'excellookup.log'
    
    @classmethod
    def from_env(cls, env_file: Optional[str] = None) -> 'Config':
        """
        Create configuration from environment variables.
        
        Args:
            env_file: Path to .env file
            
        Returns:
            Config instance with values from environment
        """
        if env_file:
            load_dotenv(env_file)
        else:
            load_dotenv()  # Load from default .env file
        
        return cls(
            input_dir=os.getenv('INPUT_DIR', './input'),
            output_dir=os.getenv('OUTPUT_DIR', './output'),
            temp_dir=os.getenv('TEMP_DIR', './temp'),
            default_tolerance=float(os.getenv('DEFAULT_TOLERANCE', '0.001')),
            case_sensitive=os.getenv('CASE_SENSITIVE', 'false').lower() == 'true',
            ignore_whitespace=os.getenv('IGNORE_WHITESPACE', 'true').lower() == 'true',
            date_tolerance_days=int(os.getenv('DATE_TOLERANCE_DAYS', '0')),
            report_format=os.getenv('REPORT_FORMAT', 'xlsx'),
            include_charts=os.getenv('INCLUDE_CHARTS', 'true').lower() == 'true',
            max_differences=int(os.getenv('MAX_DIFFERENCES', '1000')),
            chunk_size=int(os.getenv('CHUNK_SIZE', '10000')),
            max_memory_usage=os.getenv('MAX_MEMORY_USAGE', '1GB'),
            log_level=os.getenv('LOG_LEVEL', 'INFO'),
            log_file=os.getenv('LOG_FILE', 'excellookup.log')
        )
    
    @classmethod
    def from_file(cls, config_file: str) -> 'Config':
        """
        Create configuration from JSON file.
        
        Args:
            config_file: Path to JSON configuration file
            
        Returns:
            Config instance with values from file
        """
        config_path = Path(config_file)
        
        if not config_path.exists():
            raise ConfigurationError(f"Configuration file not found: {config_file}")
        
        try:
            with open(config_path, 'r') as f:
                config_data = json.load(f)
            
            return cls(**config_data)
            
        except json.JSONDecodeError as e:
            raise ConfigurationError(f"Invalid JSON in configuration file: {e}")
        except TypeError as e:
            raise ConfigurationError(f"Invalid configuration parameters: {e}")
    
    def to_file(self, config_file: str) -> None:
        """
        Save configuration to JSON file.
        
        Args:
            config_file: Path to output JSON file
        """
        config_path = Path(config_file)
        config_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            with open(config_path, 'w') as f:
                json.dump(asdict(self), f, indent=2)
                
        except Exception as e:
            raise ConfigurationError(f"Error saving configuration file: {e}")
    
    def create_directories(self) -> None:
        """
        Create necessary directories based on configuration.
        """
        directories = [self.input_dir, self.output_dir, self.temp_dir]
        
        for directory in directories:
            Path(directory).mkdir(parents=True, exist_ok=True)
    
    def validate(self) -> List[str]:
        """
        Validate configuration settings.
        
        Returns:
            List of validation error messages
        """
        errors = []
        
        # Validate numeric values
        if self.default_tolerance < 0:
            errors.append("Default tolerance must be non-negative")
        
        if self.date_tolerance_days < 0:
            errors.append("Date tolerance days must be non-negative")
        
        if self.max_differences <= 0:
            errors.append("Max differences must be positive")
        
        if self.chunk_size <= 0:
            errors.append("Chunk size must be positive")
        
        # Validate report format
        valid_formats = ['xlsx', 'csv', 'html']
        if self.report_format not in valid_formats:
            errors.append(f"Report format must be one of: {', '.join(valid_formats)}")
        
        # Validate log level
        valid_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
        if self.log_level.upper() not in valid_levels:
            errors.append(f"Log level must be one of: {', '.join(valid_levels)}")
        
        return errors
    
    def get_memory_limit_bytes(self) -> int:
        """
        Convert memory usage string to bytes.
        
        Returns:
            Memory limit in bytes
        """
        memory_str = self.max_memory_usage.upper()
        
        if memory_str.endswith('GB'):
            return int(float(memory_str[:-2]) * 1024 * 1024 * 1024)
        elif memory_str.endswith('MB'):
            return int(float(memory_str[:-2]) * 1024 * 1024)
        elif memory_str.endswith('KB'):
            return int(float(memory_str[:-2]) * 1024)
        else:
            # Assume bytes
            return int(memory_str)