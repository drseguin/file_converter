"""
Spreadsheet format converter implementation.

This module provides classes for converting between spreadsheet formats
such as CSV, Excel, and other tabular data formats.
"""

import logging
import os
import tempfile
from typing import BinaryIO, List, Optional, Union

import pandas as pd
import numpy as np
import openpyxl
import csv
import json

from .converter_base import ConverterBase

logger = logging.getLogger(__name__)


class SpreadsheetConverter(ConverterBase):
    """Class for converting between spreadsheet formats."""
    
    # Define supported formats
    _SUPPORTED_INPUT_FORMATS = ['csv', 'xlsx', 'xls', 'json', 'tsv', 'ods']
    _SUPPORTED_OUTPUT_FORMATS = ['csv', 'xlsx', 'xls', 'json', 'tsv', 'html', 'md', 'ods']
    
    def __init__(self):
        """Initialize the spreadsheet converter."""
        super().__init__()
        self.logger.info("Initializing SpreadsheetConverter")
    
    def get_supported_input_formats(self) -> List[str]:
        """
        Get a list of supported input formats.
        
        Returns:
            List[str]: List of supported input format extensions
        """
        return self._SUPPORTED_INPUT_FORMATS
    
    def get_supported_output_formats(self) -> List[str]:
        """
        Get a list of supported output formats.
        
        Returns:
            List[str]: List of supported output format extensions
        """
        return self._SUPPORTED_OUTPUT_FORMATS
    
    def convert(self, input_file: Union[str, BinaryIO], output_format: str, 
                output_path: Optional[str] = None, **kwargs) -> str:
        """
        Convert a spreadsheet from one format to another.
        
        Args:
            input_file: Path to the input file or file-like object
            output_format: The target format for conversion
            output_path: Optional path where the output file should be saved
            **kwargs: Additional conversion options
                - sheet_name: Name or index of the sheet to convert (for Excel files)
                - delimiter: Delimiter for CSV/TSV files (default: ',' for CSV, '\t' for TSV)
                - encoding: File encoding (default: 'utf-8')
                - preserve_formatting: Whether to preserve formatting (default: True)
                - include_index: Whether to include DataFrame index in output (default: False)
                
        Returns:
            str: Path to the converted file
        """
        # Handle file-like objects
        if not isinstance(input_file, str):
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=f".{kwargs.get('input_format', 'tmp')}")
            temp_file.write(input_file.read())
            temp_file.close()
            input_file = temp_file.name
            self.logger.info(f"Created temporary file: {input_file}")
        
        # Validate input file
        if not self.validate_file_exists(input_file):
            raise FileNotFoundError(f"Input file not found: {input_file}")
        
        # Get input format from file extension if not specified
        input_format = kwargs.get('input_format', self.get_file_extension(input_file))
        
        # Create output path if not provided
        if not output_path:
            output_path = self.create_output_path(input_file, output_format)
        
        self.logger.info(f"Converting {input_file} ({input_format}) to {output_path} ({output_format})")
        
        # Read the input file into a pandas DataFrame
        df = self._read_spreadsheet(input_file, input_format, **kwargs)
        
        # Write the DataFrame to the output file
        return self._write_spreadsheet(df, output_path, output_format, **kwargs)
    
    def _read_spreadsheet(self, input_file: str, input_format: str, **kwargs) -> pd.DataFrame:
        """
        Read a spreadsheet file into a pandas DataFrame.
        
        Args:
            input_file: Path to the input file
            input_format: Input file format
            **kwargs: Additional options
                - sheet_name: Name or index of the sheet to read (for Excel files)
                - delimiter: Delimiter for CSV/TSV files
                - encoding: File encoding
                
        Returns:
            pd.DataFrame: DataFrame containing the spreadsheet data
        """
        self.logger.info(f"Reading {input_format.upper()} file: {input_file}")
        
        # Get options
        encoding = kwargs.get('encoding', 'utf-8')
        sheet_name = kwargs.get('sheet_name', 0)
        
        # Read based on format
        if input_format == 'csv':
            delimiter = kwargs.get('delimiter', ',')
            return pd.read_csv(input_file, delimiter=delimiter, encoding=encoding)
        
        elif input_format == 'tsv':
            delimiter = kwargs.get('delimiter', '\t')
            return pd.read_csv(input_file, delimiter=delimiter, encoding=encoding)
        
        elif input_format in ['xlsx', 'xls']:
            return pd.read_excel(input_file, sheet_name=sheet_name)
        
        elif input_format == 'json':
            return pd.read_json(input_file, encoding=encoding)
        
        elif input_format == 'ods':
            return pd.read_excel(input_file, engine='odf')
        
        else:
            raise ValueError(f"Unsupported input format: {input_format}")
    
    def _write_spreadsheet(self, df: pd.DataFrame, output_path: str, output_format: str, **kwargs) -> str:
        """
        Write a pandas DataFrame to a spreadsheet file.
        
        Args:
            df: DataFrame to write
            output_path: Path where the output file should be saved
            output_format: Output file format
            **kwargs: Additional options
                - sheet_name: Name of the sheet (for Excel files, default: 'Sheet1')
                - delimiter: Delimiter for CSV/TSV files
                - encoding: File encoding
                - include_index: Whether to include DataFrame index in output
                
        Returns:
            str: Path to the output file
        """
        self.logger.info(f"Writing {output_format.upper()} file: {output_path}")
        
        # Get options
        encoding = kwargs.get('encoding', 'utf-8')
        include_index = kwargs.get('include_index', False)
        sheet_name = kwargs.get('sheet_name', 'Sheet1')
        
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        # Write based on format
        if output_format == 'csv':
            delimiter = kwargs.get('delimiter', ',')
            df.to_csv(output_path, sep=delimiter, index=include_index, encoding=encoding)
        
        elif output_format == 'tsv':
            delimiter = kwargs.get('delimiter', '\t')
            df.to_csv(output_path, sep=delimiter, index=include_index, encoding=encoding)
        
        elif output_format == 'xlsx':
            df.to_excel(output_path, sheet_name=sheet_name, index=include_index)
        
        elif output_format == 'xls':
            df.to_excel(output_path, sheet_name=sheet_name, index=include_index, engine='openpyxl')
        
        elif output_format == 'json':
            df.to_json(output_path, orient='records', indent=4)
        
        elif output_format == 'html':
            df.to_html(output_path, index=include_index)
        
        elif output_format == 'md':
            # Convert DataFrame to markdown table
            with open(output_path, 'w', encoding=encoding) as f:
                # Write table header
                header = '| ' + ' | '.join(df.columns) + ' |'
                separator = '| ' + ' | '.join(['---'] * len(df.columns)) + ' |'
                f.write(header + '\n')
                f.write(separator + '\n')
                
                # Write table rows
                for _, row in df.iterrows():
                    row_str = '| ' + ' | '.join([str(val) for val in row]) + ' |'
                    f.write(row_str + '\n')
        
        elif output_format == 'ods':
            df.to_excel(output_path, sheet_name=sheet_name, index=include_index, engine='odf')
        
        else:
            raise ValueError(f"Unsupported output format: {output_format}")
        
        return output_path
