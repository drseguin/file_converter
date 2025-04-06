"""
Base classes for file format conversion.

This module provides the abstract base classes for file format conversion
that will be implemented by specific converter classes.
"""

import logging
import os
from abc import ABC, abstractmethod
from typing import BinaryIO, List, Optional, Union

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class ConverterBase(ABC):
    """Abstract base class for all file converters."""
    
    def __init__(self):
        """Initialize the converter."""
        self.logger = logging.getLogger(f"{__name__}.{self.__class__.__name__}")
        self.logger.info(f"Initializing {self.__class__.__name__}")
        
    @abstractmethod
    def convert(self, input_file: Union[str, BinaryIO], output_format: str, 
                output_path: Optional[str] = None, **kwargs) -> str:
        """
        Convert a file from one format to another.
        
        Args:
            input_file: Path to the input file or file-like object
            output_format: The target format for conversion
            output_path: Optional path where the output file should be saved
            **kwargs: Additional conversion options
            
        Returns:
            str: Path to the converted file
        """
        pass
    
    @abstractmethod
    def get_supported_input_formats(self) -> List[str]:
        """
        Get a list of supported input formats.
        
        Returns:
            List[str]: List of supported input format extensions
        """
        pass
    
    @abstractmethod
    def get_supported_output_formats(self) -> List[str]:
        """
        Get a list of supported output formats.
        
        Returns:
            List[str]: List of supported output format extensions
        """
        pass
    
    def validate_file_exists(self, file_path: str) -> bool:
        """
        Validate that a file exists.
        
        Args:
            file_path: Path to the file to validate
            
        Returns:
            bool: True if the file exists, False otherwise
        """
        if not os.path.isfile(file_path):
            self.logger.error(f"File not found: {file_path}")
            return False
        return True
    
    def get_file_extension(self, file_path: str) -> str:
        """
        Get the extension of a file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            str: File extension without the dot
        """
        _, ext = os.path.splitext(file_path)
        return ext.lower().lstrip('.')
    
    def create_output_path(self, input_path: str, output_format: str, output_dir: Optional[str] = None) -> str:
        """
        Create an output path based on the input path and output format.
        
        Args:
            input_path: Path to the input file
            output_format: The target format extension
            output_dir: Optional directory for the output file
            
        Returns:
            str: Path for the output file
        """
        base_name = os.path.basename(input_path)
        name_without_ext = os.path.splitext(base_name)[0]
        
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            return os.path.join(output_dir, f"{name_without_ext}.{output_format}")
        
        return os.path.join(os.path.dirname(input_path), f"{name_without_ext}.{output_format}")


class BatchConverter:
    """Class for handling batch conversion of multiple files."""
    
    def __init__(self, converter: ConverterBase):
        """
        Initialize the batch converter.
        
        Args:
            converter: The converter to use for individual file conversions
        """
        self.converter = converter
        self.logger = logging.getLogger(f"{__name__}.BatchConverter")
        self.logger.info("Initializing BatchConverter")
        
    def convert_batch(self, input_files: List[str], output_format: str, 
                     output_dir: Optional[str] = None, **kwargs) -> List[str]:
        """
        Convert multiple files to the specified format.
        
        Args:
            input_files: List of paths to input files
            output_format: The target format for conversion
            output_dir: Optional directory for output files
            **kwargs: Additional conversion options
            
        Returns:
            List[str]: List of paths to converted files
        """
        self.logger.info(f"Starting batch conversion of {len(input_files)} files to {output_format}")
        
        converted_files = []
        for input_file in input_files:
            try:
                output_path = self.converter.convert(
                    input_file, 
                    output_format, 
                    output_path=self.converter.create_output_path(input_file, output_format, output_dir),
                    **kwargs
                )
                converted_files.append(output_path)
                self.logger.info(f"Successfully converted {input_file} to {output_path}")
            except Exception as e:
                self.logger.error(f"Failed to convert {input_file}: {str(e)}")
                
        return converted_files
