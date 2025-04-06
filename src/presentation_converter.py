"""
Presentation format converter implementation.

This module provides classes for converting between presentation formats
such as PowerPoint, PDF, and HTML.
"""

import logging
import os
import tempfile
from typing import BinaryIO, List, Optional, Union

import pypandoc
from pptx import Presentation
from pdf2image import convert_from_path

from .converter_base import ConverterBase

logger = logging.getLogger(__name__)


class PresentationConverter(ConverterBase):
    """Class for converting between presentation formats."""
    
    # Define supported formats
    _SUPPORTED_INPUT_FORMATS = ['ppt', 'pptx', 'pdf', 'html']
    _SUPPORTED_OUTPUT_FORMATS = ['ppt', 'pptx', 'pdf', 'html', 'png', 'jpg']
    
    def __init__(self):
        """Initialize the presentation converter."""
        super().__init__()
        self.logger.info("Initializing PresentationConverter")
        
        # Check if pandoc is installed
        try:
            pypandoc.get_pandoc_version()
            self.pandoc_available = True
            self.logger.info("Pandoc is available")
        except OSError:
            self.pandoc_available = False
            self.logger.warning("Pandoc is not available. Some conversions may not work.")
    
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
        Convert a presentation from one format to another.
        
        Args:
            input_file: Path to the input file or file-like object
            output_format: The target format for conversion
            output_path: Optional path where the output file should be saved
            **kwargs: Additional conversion options
                - style_template: Path to a style template for the output presentation
                - preserve_formatting: Whether to preserve formatting (default: True)
                - image_quality: Quality for image outputs (1-100, default: 90)
                
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
        
        # Handle specific conversion cases
        if input_format in ['ppt', 'pptx'] and output_format == 'pdf':
            return self._convert_pptx_to_pdf(input_file, output_path, **kwargs)
        elif input_format == 'pdf' and output_format in ['png', 'jpg']:
            return self._convert_pdf_to_images(input_file, output_format, output_path, **kwargs)
        elif input_format in ['ppt', 'pptx'] and output_format in ['png', 'jpg']:
            # Convert to PDF first, then to images
            pdf_path = self._convert_pptx_to_pdf(input_file, tempfile.mktemp(suffix='.pdf'), **kwargs)
            return self._convert_pdf_to_images(pdf_path, output_format, output_path, **kwargs)
        elif input_format in ['ppt', 'pptx', 'pdf'] and output_format == 'html':
            return self._convert_to_html(input_file, input_format, output_path, **kwargs)
        else:
            # Use pandoc for other conversions if available
            if not self.pandoc_available:
                raise RuntimeError("Pandoc is required for this conversion but is not available")
            
            return self._convert_with_pandoc(input_file, input_format, output_format, output_path, **kwargs)
    
    def _convert_pptx_to_pdf(self, input_file: str, output_path: str, **kwargs) -> str:
        """
        Convert PowerPoint to PDF.
        
        Args:
            input_file: Path to the input PowerPoint file
            output_path: Path where the output PDF file should be saved
            **kwargs: Additional conversion options
                
        Returns:
            str: Path to the converted PDF file
        """
        self.logger.info(f"Converting PowerPoint to PDF: {input_file} -> {output_path}")
        
        # For this conversion, we'll use LibreOffice in headless mode
        # This is a simplified implementation - in a real app, you'd need to handle LibreOffice properly
        import subprocess
        
        try:
            # Check if LibreOffice is available
            subprocess.run(['which', 'libreoffice'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            # Convert using LibreOffice
            cmd = [
                'libreoffice', 
                '--headless', 
                '--convert-to', 'pdf', 
                '--outdir', os.path.dirname(output_path),
                input_file
            ]
            
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            # LibreOffice creates the file with the same name but .pdf extension
            # We need to rename it if the output_path is different
            generated_pdf = os.path.join(
                os.path.dirname(output_path),
                os.path.splitext(os.path.basename(input_file))[0] + '.pdf'
            )
            
            if generated_pdf != output_path:
                os.rename(generated_pdf, output_path)
            
            self.logger.info(f"Successfully converted PowerPoint to PDF: {output_path}")
            return output_path
            
        except subprocess.CalledProcessError:
            self.logger.error("LibreOffice is not available for PowerPoint to PDF conversion")
            raise RuntimeError("LibreOffice is required for PowerPoint to PDF conversion but is not available")
    
    def _convert_pdf_to_images(self, input_file: str, output_format: str, output_path: str, **kwargs) -> str:
        """
        Convert PDF to images (PNG or JPG).
        
        Args:
            input_file: Path to the input PDF file
            output_format: Output image format ('png' or 'jpg')
            output_path: Path where the output image file should be saved
            **kwargs: Additional conversion options
                - image_quality: Quality for JPG outputs (1-100, default: 90)
                
        Returns:
            str: Path to the directory containing the converted images
        """
        self.logger.info(f"Converting PDF to {output_format.upper()} images: {input_file}")
        
        # Get image quality for JPG
        image_quality = kwargs.get('image_quality', 90)
        
        # Create output directory if it's a file path
        output_dir = os.path.dirname(output_path)
        base_name = os.path.splitext(os.path.basename(output_path))[0]
        
        # Convert PDF to images
        images = convert_from_path(
            input_file,
            dpi=300,
            fmt=output_format,
            output_folder=output_dir,
            output_file=base_name,
            thread_count=4
        )
        
        # Return the directory containing the images
        self.logger.info(f"Successfully converted PDF to {len(images)} {output_format.upper()} images in {output_dir}")
        return output_dir
    
    def _convert_to_html(self, input_file: str, input_format: str, output_path: str, **kwargs) -> str:
        """
        Convert presentation to HTML.
        
        Args:
            input_file: Path to the input file
            input_format: Input file format
            output_path: Path where the output HTML file should be saved
            **kwargs: Additional conversion options
                
        Returns:
            str: Path to the converted HTML file
        """
        self.logger.info(f"Converting {input_format.upper()} to HTML: {input_file} -> {output_path}")
        
        # For PowerPoint to HTML, we'll use LibreOffice
        if input_format in ['ppt', 'pptx']:
            import subprocess
            
            try:
                # Check if LibreOffice is available
                subprocess.run(['which', 'libreoffice'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                # Convert using LibreOffice
                cmd = [
                    'libreoffice', 
                    '--headless', 
                    '--convert-to', 'html', 
                    '--outdir', os.path.dirname(output_path),
                    input_file
                ]
                
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                # LibreOffice creates the file with the same name but .html extension
                # We need to rename it if the output_path is different
                generated_html = os.path.join(
                    os.path.dirname(output_path),
                    os.path.splitext(os.path.basename(input_file))[0] + '.html'
                )
                
                if generated_html != output_path:
                    os.rename(generated_html, output_path)
                
                self.logger.info(f"Successfully converted {input_format.upper()} to HTML: {output_path}")
                return output_path
                
            except subprocess.CalledProcessError:
                self.logger.error("LibreOffice is not available for PowerPoint to HTML conversion")
                raise RuntimeError("LibreOffice is required for PowerPoint to HTML conversion but is not available")
        
        # For PDF to HTML, we'll use pdf2htmlEX if available, otherwise fall back to pandoc
        elif input_format == 'pdf':
            import subprocess
            
            try:
                # Check if pdf2htmlEX is available
                subprocess.run(['which', 'pdf2htmlEX'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                # Convert using pdf2htmlEX
                cmd = [
                    'pdf2htmlEX',
                    '--zoom', '1.5',
                    input_file,
                    output_path
                ]
                
                subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                self.logger.info(f"Successfully converted PDF to HTML: {output_path}")
                return output_path
                
            except subprocess.CalledProcessError:
                self.logger.warning("pdf2htmlEX is not available, falling back to pandoc")
                
                if self.pandoc_available:
                    return self._convert_with_pandoc(input_file, input_format, 'html', output_path, **kwargs)
                else:
                    raise RuntimeError("Neither pdf2htmlEX nor pandoc is available for PDF to HTML conversion")
    
    def _convert_with_pandoc(self, input_file: str, input_format: str, 
                            output_format: str, output_path: str, **kwargs) -> str:
        """
        Convert a presentation using pandoc.
        
        Args:
            input_file: Path to the input file
            input_format: Input file format
            output_format: Output file format
            output_path: Path where the output file should be saved
            **kwargs: Additional conversion options
                
        Returns:
            str: Path to the converted file
        """
        self.logger.info(f"Converting with pandoc: {input_file} ({input_format}) -> {output_path} ({output_format})")
        
        # Map our format names to pandoc format names
        format_mapping = {
            'ppt': 'pptx',
            'pptx': 'pptx',
            'pdf': 'pdf',
            'html': 'html'
        }
        
        pandoc_input_format = format_mapping.get(input_format, input_format)
        pandoc_output_format = format_mapping.get(output_format, output_format)
        
        # Set pandoc options
        extra_args = []
        
        # Add options for preserving formatting if specified
        if kwargs.get('preserve_formatting', True):
            if pandoc_output_format == 'html':
                extra_args.extend(['--standalone', '--self-contained'])
        
        # Convert using pandoc
        pypandoc.convert_file(
            input_file,
            pandoc_output_format,
            format=pandoc_input_format,
            outputfile=output_path,
            extra_args=extra_args
        )
        
        self.logger.info(f"Successfully converted with pandoc: {output_path}")
        return output_path
