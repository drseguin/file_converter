"""
File conversion utility functions.

This module provides utility functions for file conversion operations.
"""

import logging
import os
import tempfile
import shutil
from typing import List, Optional, Tuple, Union

logger = logging.getLogger(__name__)


def create_temp_directory() -> str:
    """
    Create a temporary directory for file operations.
    
    Returns:
        str: Path to the created temporary directory
    """
    temp_dir = tempfile.mkdtemp()
    logger.info(f"Created temporary directory: {temp_dir}")
    return temp_dir


def clean_temp_directory(temp_dir: str) -> None:
    """
    Clean up a temporary directory.
    
    Args:
        temp_dir: Path to the temporary directory to clean up
    """
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
        logger.info(f"Cleaned up temporary directory: {temp_dir}")


def get_file_extension(file_path: str) -> str:
    """
    Get the extension of a file.
    
    Args:
        file_path: Path to the file
        
    Returns:
        str: File extension without the dot
    """
    _, ext = os.path.splitext(file_path)
    return ext.lower().lstrip('.')


def get_mime_type(file_format: str) -> str:
    """
    Get the MIME type for a file format.
    
    Args:
        file_format: File format extension
        
    Returns:
        str: MIME type
    """
    mime_types = {
        'md': 'text/markdown',
        'markdown': 'text/markdown',
        'txt': 'text/plain',
        'html': 'text/html',
        'pdf': 'application/pdf',
        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'doc': 'application/msword',
        'rtf': 'application/rtf',
        'odt': 'application/vnd.oasis.opendocument.text',
        'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'ppt': 'application/vnd.ms-powerpoint',
        'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'xls': 'application/vnd.ms-excel',
        'csv': 'text/csv',
        'tsv': 'text/tab-separated-values',
        'json': 'application/json',
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'ods': 'application/vnd.oasis.opendocument.spreadsheet'
    }
    
    return mime_types.get(file_format, 'application/octet-stream')


def validate_file_exists(file_path: str) -> bool:
    """
    Validate that a file exists.
    
    Args:
        file_path: Path to the file to validate
        
    Returns:
        bool: True if the file exists, False otherwise
    """
    if not os.path.isfile(file_path):
        logger.error(f"File not found: {file_path}")
        return False
    return True


def create_output_path(input_path: str, output_format: str, output_dir: Optional[str] = None) -> str:
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


def save_uploaded_file(uploaded_file, directory: str) -> str:
    """
    Save an uploaded file to a directory.
    
    Args:
        uploaded_file: Streamlit uploaded file object
        directory: Directory to save the file to
        
    Returns:
        str: Path to the saved file
    """
    os.makedirs(directory, exist_ok=True)
    file_path = os.path.join(directory, uploaded_file.name)
    
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    logger.info(f"Saved uploaded file: {file_path}")
    return file_path


def check_dependencies() -> Tuple[bool, List[str]]:
    """
    Check if all required external dependencies are installed.
    
    Returns:
        Tuple[bool, List[str]]: (True if all dependencies are available, List of missing dependencies)
    """
    import subprocess
    
    dependencies = {
        'pandoc': 'Pandoc is required for document format conversions',
        'libreoffice': 'LibreOffice is required for some document and presentation conversions',
        'pdf2htmlEX': 'pdf2htmlEX is optional for PDF to HTML conversion'
    }
    
    missing = []
    
    for dep, message in dependencies.items():
        try:
            subprocess.run(['which', dep], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            logger.info(f"Dependency check: {dep} is available")
        except subprocess.CalledProcessError:
            if dep == 'pdf2htmlEX':
                logger.warning(f"Optional dependency {dep} is not available: {message}")
            else:
                logger.warning(f"Dependency {dep} is not available: {message}")
                missing.append(f"{dep}: {message}")
    
    return len(missing) == 0, missing
