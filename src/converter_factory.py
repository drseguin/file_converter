"""
File converter factory module.

This module provides a factory class for creating appropriate converters
based on file types.
"""

import logging
import os
from typing import List, Optional, Union

from .converter_base import ConverterBase
from .document_converter import DocumentConverter
from .presentation_converter import PresentationConverter
from .spreadsheet_converter import SpreadsheetConverter

logger = logging.getLogger(__name__)


class ConverterFactory:
    """Factory class for creating appropriate converters based on file types."""
    
    def __init__(self):
        """Initialize the converter factory."""
        self.logger = logging.getLogger(f"{__name__}.ConverterFactory")
        self.logger.info("Initializing ConverterFactory")
        
        # Initialize converters
        self._document_converter = None
        self._presentation_converter = None
        self._spreadsheet_converter = None
        
        # Map file extensions to converter types
        self._extension_map = {
            # Document formats
            'md': 'document',
            'markdown': 'document',
            'docx': 'document',
            'doc': 'document',
            'pdf': 'document',
            'txt': 'document',
            'html': 'document',
            'odt': 'document',
            'rtf': 'document',
            
            # Presentation formats
            'ppt': 'presentation',
            'pptx': 'presentation',
            
            # Spreadsheet formats
            'csv': 'spreadsheet',
            'xlsx': 'spreadsheet',
            'xls': 'spreadsheet',
            'json': 'spreadsheet',
            'tsv': 'spreadsheet',
            'ods': 'spreadsheet'
        }
    
    def get_converter(self, file_extension: str) -> Optional[ConverterBase]:
        """
        Get the appropriate converter for a file extension.
        
        Args:
            file_extension: File extension without the dot
            
        Returns:
            ConverterBase: Appropriate converter instance or None if not supported
        """
        converter_type = self._extension_map.get(file_extension.lower())
        
        if not converter_type:
            self.logger.warning(f"No converter found for extension: {file_extension}")
            return None
        
        if converter_type == 'document':
            if not self._document_converter:
                self._document_converter = DocumentConverter()
            return self._document_converter
        
        elif converter_type == 'presentation':
            if not self._presentation_converter:
                self._presentation_converter = PresentationConverter()
            return self._presentation_converter
        
        elif converter_type == 'spreadsheet':
            if not self._spreadsheet_converter:
                self._spreadsheet_converter = SpreadsheetConverter()
            return self._spreadsheet_converter
        
        return None
    
    def get_all_supported_formats(self) -> List[str]:
        """
        Get a list of all supported file formats.
        
        Returns:
            List[str]: List of all supported file extensions
        """
        return list(self._extension_map.keys())
    
    def is_format_supported(self, file_extension: str) -> bool:
        """
        Check if a file format is supported.
        
        Args:
            file_extension: File extension without the dot
            
        Returns:
            bool: True if the format is supported, False otherwise
        """
        return file_extension.lower() in self._extension_map
