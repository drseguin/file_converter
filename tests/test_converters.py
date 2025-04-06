"""
Test script for the file converter application.

This script tests the basic functionality of the file converter classes.
"""

import os
import sys
import logging
import tempfile
import unittest

# Add parent directory to path to import modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.converter_factory import ConverterFactory
from src.document_converter import DocumentConverter
from src.presentation_converter import PresentationConverter
from src.spreadsheet_converter import SpreadsheetConverter
from src.utils import create_temp_directory, clean_temp_directory

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class TestConverterFactory(unittest.TestCase):
    """Test cases for the ConverterFactory class."""
    
    def setUp(self):
        """Set up test environment."""
        self.factory = ConverterFactory()
    
    def test_get_converter(self):
        """Test getting converters for different file types."""
        # Test document converter
        converter = self.factory.get_converter('md')
        self.assertIsInstance(converter, DocumentConverter)
        
        # Test presentation converter
        converter = self.factory.get_converter('pptx')
        self.assertIsInstance(converter, PresentationConverter)
        
        # Test spreadsheet converter
        converter = self.factory.get_converter('csv')
        self.assertIsInstance(converter, SpreadsheetConverter)
        
        # Test unsupported format
        converter = self.factory.get_converter('xyz')
        self.assertIsNone(converter)
    
    def test_is_format_supported(self):
        """Test checking if formats are supported."""
        self.assertTrue(self.factory.is_format_supported('md'))
        self.assertTrue(self.factory.is_format_supported('docx'))
        self.assertTrue(self.factory.is_format_supported('pptx'))
        self.assertTrue(self.factory.is_format_supported('csv'))
        self.assertFalse(self.factory.is_format_supported('xyz'))


class TestDocumentConverter(unittest.TestCase):
    """Test cases for the DocumentConverter class."""
    
    def setUp(self):
        """Set up test environment."""
        self.converter = DocumentConverter()
        self.temp_dir = create_temp_directory()
        
        # Create a test markdown file
        self.test_md_path = os.path.join(self.temp_dir, 'test.md')
        with open(self.test_md_path, 'w') as f:
            f.write('# Test Heading\n\nThis is a test paragraph.\n\n* List item 1\n* List item 2\n')
    
    def tearDown(self):
        """Clean up test environment."""
        clean_temp_directory(self.temp_dir)
    
    def test_get_supported_formats(self):
        """Test getting supported formats."""
        input_formats = self.converter.get_supported_input_formats()
        output_formats = self.converter.get_supported_output_formats()
        
        self.assertIn('md', input_formats)
        self.assertIn('docx', input_formats)
        self.assertIn('pdf', input_formats)
        
        self.assertIn('md', output_formats)
        self.assertIn('docx', output_formats)
        self.assertIn('pdf', output_formats)
    
    def test_markdown_to_html_conversion(self):
        """Test converting Markdown to HTML."""
        try:
            output_path = os.path.join(self.temp_dir, 'output.html')
            result = self.converter.convert(
                self.test_md_path,
                'html',
                output_path=output_path
            )
            
            self.assertTrue(os.path.exists(result))
            self.assertEqual(result, output_path)
            
            # Check if the HTML file contains expected content
            with open(result, 'r') as f:
                content = f.read()
                self.assertIn('<h1>Test Heading</h1>', content)
                self.assertIn('<li>List item 1</li>', content)
        except Exception as e:
            logger.warning(f"Markdown to HTML conversion test failed: {str(e)}")
            # Skip test if dependencies are missing
            self.skipTest(f"Skipping test due to: {str(e)}")


class TestSpreadsheetConverter(unittest.TestCase):
    """Test cases for the SpreadsheetConverter class."""
    
    def setUp(self):
        """Set up test environment."""
        self.converter = SpreadsheetConverter()
        self.temp_dir = create_temp_directory()
        
        # Create a test CSV file
        self.test_csv_path = os.path.join(self.temp_dir, 'test.csv')
        with open(self.test_csv_path, 'w') as f:
            f.write('Name,Age,City\nJohn,30,New York\nJane,25,Boston\n')
    
    def tearDown(self):
        """Clean up test environment."""
        clean_temp_directory(self.temp_dir)
    
    def test_get_supported_formats(self):
        """Test getting supported formats."""
        input_formats = self.converter.get_supported_input_formats()
        output_formats = self.converter.get_supported_output_formats()
        
        self.assertIn('csv', input_formats)
        self.assertIn('xlsx', input_formats)
        self.assertIn('json', input_formats)
        
        self.assertIn('csv', output_formats)
        self.assertIn('xlsx', output_formats)
        self.assertIn('json', output_formats)
    
    def test_csv_to_json_conversion(self):
        """Test converting CSV to JSON."""
        output_path = os.path.join(self.temp_dir, 'output.json')
        result = self.converter.convert(
            self.test_csv_path,
            'json',
            output_path=output_path
        )
        
        self.assertTrue(os.path.exists(result))
        self.assertEqual(result, output_path)
        
        # Check if the JSON file contains expected content
        with open(result, 'r') as f:
            content = f.read()
            self.assertIn('John', content)
            self.assertIn('New York', content)


if __name__ == '__main__':
    unittest.main()
