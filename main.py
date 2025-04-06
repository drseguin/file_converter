"""
Main module for file format conversion application.

This module integrates all components and provides the entry point for the application.
"""

import logging
import os
import sys
import tempfile
from typing import List, Dict, Any, Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("file_converter.log")
    ]
)
logger = logging.getLogger(__name__)

# Import after logging configuration
import streamlit as st
from src.converter_factory import ConverterFactory
from src.converter_base import BatchConverter
from src.utils import (
    create_temp_directory, 
    clean_temp_directory, 
    get_file_extension,
    get_mime_type,
    save_uploaded_file,
    check_dependencies
)


class FileConverterApp:
    """Main class for the Streamlit file converter application."""
    
    def __init__(self):
        """Initialize the application."""
        self.logger = logging.getLogger(f"{__name__}.FileConverterApp")
        self.logger.info("Initializing FileConverterApp")
        
        # Initialize converter factory
        self.converter_factory = ConverterFactory()
        
        # Set up the page configuration
        st.set_page_config(
            page_title="File Format Converter",
            page_icon="📄",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
        # Create temp directory for file operations
        self.temp_dir = create_temp_directory()
        
        # Check dependencies
        self._check_dependencies()
    
    def _check_dependencies(self):
        """Check if all required external dependencies are installed."""
        all_available, missing = check_dependencies()
        if not all_available:
            self.logger.warning(f"Missing dependencies: {missing}")
            # We'll show warnings in the UI when the app runs
    
    def run(self):
        """Run the Streamlit application."""
        self.logger.info("Starting FileConverterApp")
        
        # Display header
        st.title("📄 File Format Converter")
        st.write("Convert between various file formats with ease.")
        
        # Create tabs for different conversion types
        tabs = st.tabs([
            "Document Conversion", 
            "Presentation Conversion", 
            "Spreadsheet Conversion",
            "Batch Conversion"
        ])
        
        # Document Conversion Tab
        with tabs[0]:
            self._document_conversion_tab()
        
        # Presentation Conversion Tab
        with tabs[1]:
            self._presentation_conversion_tab()
        
        # Spreadsheet Conversion Tab
        with tabs[2]:
            self._spreadsheet_conversion_tab()
        
        # Batch Conversion Tab
        with tabs[3]:
            self._batch_conversion_tab()
        
        # Display footer
        st.markdown("---")
        st.markdown("### About")
        st.write(
            "This application allows you to convert files between various formats. "
            "Upload your file, select the desired output format, and download the converted file."
        )
    
    def _document_conversion_tab(self):
        """Display the document conversion tab."""
        st.header("Document Conversion")
        st.write("Convert between document formats like Markdown, Word, PDF, HTML, etc.")
        
        # Get document converter
        document_converter = self.converter_factory.get_converter('md')
        
        # Get supported formats
        input_formats = document_converter.get_supported_input_formats()
        output_formats = document_converter.get_supported_output_formats()
        
        # File upload
        uploaded_file = st.file_uploader(
            "Upload a document file",
            type=input_formats,
            key="document_uploader"
        )
        
        if uploaded_file is not None:
            # Display file info
            st.write(f"File name: {uploaded_file.name}")
            st.write(f"File size: {uploaded_file.size / 1024:.2f} KB")
            
            # Get file extension
            file_ext = get_file_extension(uploaded_file.name)
            
            # Select output format
            output_format = st.selectbox(
                "Select output format",
                options=output_formats,
                key="document_output_format"
            )
            
            # Advanced options
            with st.expander("Advanced Options"):
                preserve_formatting = st.checkbox("Preserve formatting", value=True, key="document_preserve_formatting")
                style_template = st.file_uploader(
                    "Upload style template (optional)",
                    type=["docx", "css"],
                    key="document_style_template"
                )
            
            # Convert button
            if st.button("Convert Document", key="document_convert_button"):
                try:
                    with st.spinner("Converting document..."):
                        # Save uploaded file to temp directory
                        input_path = save_uploaded_file(uploaded_file, self.temp_dir)
                        
                        # Save style template if provided
                        style_template_path = None
                        if style_template:
                            style_template_path = save_uploaded_file(style_template, self.temp_dir)
                        
                        # Create output path
                        output_filename = f"{os.path.splitext(uploaded_file.name)[0]}.{output_format}"
                        output_path = os.path.join(self.temp_dir, output_filename)
                        
                        # Convert file
                        converted_file = document_converter.convert(
                            input_path,
                            output_format,
                            output_path=output_path,
                            preserve_formatting=preserve_formatting,
                            style_template=style_template_path
                        )
                        
                        # Display success message
                        st.success(f"Conversion successful! File saved as {output_filename}")
                        
                        # Preview converted file if possible
                        self._preview_file(converted_file, output_format)
                        
                        # Download button
                        with open(converted_file, "rb") as f:
                            st.download_button(
                                label="Download Converted File",
                                data=f,
                                file_name=output_filename,
                                mime=get_mime_type(output_format),
                                key="document_download_button"
                            )
                
                except Exception as e:
                    st.error(f"Error during conversion: {str(e)}")
                    self.logger.error(f"Conversion error: {str(e)}", exc_info=True)
    
    def _presentation_conversion_tab(self):
        """Display the presentation conversion tab."""
        st.header("Presentation Conversion")
        st.write("Convert between presentation formats like PowerPoint, PDF, HTML, etc.")
        
        # Get presentation converter
        presentation_converter = self.converter_factory.get_converter('pptx')
        
        # Get supported formats
        input_formats = presentation_converter.get_supported_input_formats()
        output_formats = presentation_converter.get_supported_output_formats()
        
        # File upload
        uploaded_file = st.file_uploader(
            "Upload a presentation file",
            type=input_formats,
            key="presentation_uploader"
        )
        
        if uploaded_file is not None:
            # Display file info
            st.write(f"File name: {uploaded_file.name}")
            st.write(f"File size: {uploaded_file.size / 1024:.2f} KB")
            
            # Get file extension
            file_ext = get_file_extension(uploaded_file.name)
            
            # Select output format
            output_format = st.selectbox(
                "Select output format",
                options=output_formats,
                key="presentation_output_format"
            )
            
            # Advanced options
            with st.expander("Advanced Options"):
                preserve_formatting = st.checkbox("Preserve formatting", value=True, key="presentation_preserve_formatting")
                
                # Image quality option for image outputs
                if output_format in ['png', 'jpg']:
                    image_quality = st.slider(
                        "Image Quality",
                        min_value=1,
                        max_value=100,
                        value=90,
                        key="presentation_image_quality"
                    )
                else:
                    image_quality = 90
            
            # Convert button
            if st.button("Convert Presentation", key="presentation_convert_button"):
                try:
                    with st.spinner("Converting presentation..."):
                        # Save uploaded file to temp directory
                        input_path = save_uploaded_file(uploaded_file, self.temp_dir)
                        
                        # Create output path
                        output_filename = f"{os.path.splitext(uploaded_file.name)[0]}.{output_format}"
                        output_path = os.path.join(self.temp_dir, output_filename)
                        
                        # Convert file
                        converted_file = presentation_converter.convert(
                            input_path,
                            output_format,
                            output_path=output_path,
                            preserve_formatting=preserve_formatting,
                            image_quality=image_quality
                        )
                        
                        # Display success message
                        st.success(f"Conversion successful! File saved as {output_filename}")
                        
                        # Preview converted file if possible
                        self._preview_file(converted_file, output_format)
                        
                        # Download button
                        if os.path.isdir(converted_file):
                            # For image outputs that return a directory
                            st.write("Multiple files were generated. Please download them individually:")
                            for file in os.listdir(converted_file):
                                if file.endswith(f".{output_format}"):
                                    file_path = os.path.join(converted_file, file)
                                    with open(file_path, "rb") as f:
                                        st.download_button(
                                            label=f"Download {file}",
                                            data=f,
                                            file_name=file,
                                            mime=get_mime_type(output_format),
                                            key=f"presentation_download_button_{file}"
                                        )
                        else:
                            # For single file outputs
                            with open(converted_file, "rb") as f:
                                st.download_button(
                                    label="Download Converted File",
                                    data=f,
                                    file_name=output_filename,
                                    mime=get_mime_type(output_format),
                                    key="presentation_download_button"
                                )
                
                except Exception as e:
                    st.error(f"Error during conversion: {str(e)}")
                    self.logger.error(f"Conversion error: {str(e)}", exc_info=True)
    
    def _spreadsheet_conversion_tab(self):
        """Display the spreadsheet conversion tab."""
        st.header("Spreadsheet Conversion")
        st.write("Convert between spreadsheet formats like CSV, Excel, JSON, etc.")
        
        # Get spreadsheet converter
        spreadsheet_converter = self.converter_factory.get_converter('csv')
        
        # Get supported formats
        input_formats = spreadsheet_converter.get_supported_input_formats()
        output_formats = spreadsheet_converter.get_supported_output_formats()
        
        # File upload
        uploaded_file = st.file_uploader(
            "Upload a spreadsheet file",
            type=input_formats,
            key="spreadsheet_uploader"
        )
        
        if uploaded_file is not None:
            # Display file info
            st.write(f"File name: {uploaded_file.name}")
            st.write(f"File size: {uploaded_file.size / 1024:.2f} KB")
            
            # Get file extension
            file_ext = get_file_extension(uploaded_file.name)
            
            # Select output format
            output_format = st.selectbox(
                "Select output format",
                options=output_formats,
                key="spreadsheet_output_format"
            )
            
            # Advanced options
            with st.expander("Advanced Options"):
                # Sheet selection for Excel files
                if file_ext in ['xlsx', 'xls']:
                    sheet_name = st.text_input("Sheet name (leave blank for first sheet)", key="spreadsheet_sheet_name")
                else:
                    sheet_name = None
                
                # Delimiter for CSV/TSV files
                if file_ext in ['csv', 'tsv'] or output_format in ['csv', 'tsv']:
                    delimiter = st.text_input(
                        "Delimiter (leave blank for default)",
                        key="spreadsheet_delimiter",
                        help="Default is comma for CSV, tab for TSV"
                    )
                else:
                    delimiter = None
                
                # Encoding option
                encoding = st.selectbox(
                    "File encoding",
                    options=["utf-8", "latin-1", "ascii", "utf-16"],
                    index=0,
                    key="spreadsheet_encoding"
                )
                
                # Include index option
                include_index = st.checkbox("Include index in output", value=False, key="spreadsheet_include_index")
            
            # Convert button
            if st.button("Convert Spreadsheet", key="spreadsheet_convert_button"):
                try:
                    with st.spinner("Converting spreadsheet..."):
                        # Save uploaded file to temp directory
                        input_path = save_uploaded_file(uploaded_file, self.temp_dir)
                        
                        # Create output path
                        output_filename = f"{os.path.splitext(uploaded_file.name)[0]}.{output_format}"
                        output_path = os.path.join(self.temp_dir, output_filename)
                        
                        # Prepare conversion options
                        options = {
                            'encoding': encoding,
                            'include_index': include_index
                        }
                        
                        if sheet_name:
                            options['sheet_name'] = sheet_name
                        
                        if delimiter:
                            options['delimiter'] = delimiter
                        
                        # Convert file
                        converted_file = spreadsheet_converter.convert(
                            input_path,
                            output_format,
                            output_path=output_path,
                            **options
                        )
                        
                        # Display success message
                        st.success(f"Conversion successful! File saved as {output_filename}")
                        
                        # Preview converted file if possible
                        self._preview_file(converted_file, output_format)
                        
                        # Download button
                        with open(converted_file, "rb") as f:
                            st.download_button(
                                label="Download Converted File",
                                data=f,
                                file_name=output_filename,
                                mime=get_mime_type(output_format),
                                key="spreadsheet_download_button"
                            )
                
                except Exception as e:
                    st.error(f"Error during conversion: {str(e)}")
                    self.logger.error(f"Conversion error: {str(e)}", exc_info=True)
    
    def _batch_conversion_tab(self):
        """Display the batch conversion tab."""
        st.header("Batch Conversion")
        st.write("Convert multiple files at once.")
        
        # Select converter type
        converter_type = st.selectbox(
            "Select file type",
            options=["Document", "Presentation", "Spreadsheet"],
            key="batch_converter_type"
        )
        
        # Get appropriate converter and formats based on selection
        if converter_type == "Document":
            converter = self.converter_factory.get_converter('md')
            input_formats = converter.get_supported_input_formats()
            output_formats = converter.get_supported_output_formats()
        elif converter_type == "Presentation":
            converter = self.converter_factory.get_converter('pptx')
            input_formats = converter.get_supported_input_formats()
            output_formats = converter.get_supported_output_formats()
        else:  # Spreadsheet
            converter = self.converter_factory.get_converter('csv')
            input_formats = converter.get_supported_input_formats()
            output_formats = converter.get_supported_output_formats()
        
        # File upload (multiple)
        uploaded_files = st.file_uploader(
            "Upload files",
            type=input_formats,
            accept_multiple_files=True,
            key="batch_uploader"
        )
        
        if uploaded_files:
            # Display file info
            st.write(f"Number of files: {len(uploaded_files)}")
            
            # Select output format
            output_format = st.selectbox(
                "Select output format",
                options=output_formats,
                key="batch_output_format"
            )
            
            # Advanced options
            with st.expander("Advanced Options"):
                preserve_formatting = st.checkbox("Preserve formatting", value=True, key="batch_preserve_formatting")
                
                # Additional options based on converter type
                if converter_type == "Spreadsheet":
                    encoding = st.selectbox(
                        "File encoding",
                        options=["utf-8", "latin-1", "ascii", "utf-16"],
                        index=0,
                        key="batch_encoding"
                    )
                    include_index = st.checkbox("Include index in output", value=False, key="batch_include_index")
                elif converter_type == "Presentation" and output_format in ['png', 'jpg']:
                    image_quality = st.slider(
                        "Image Quality",
                        min_value=1,
                        max_value=100,
                        value=90,
                        key="batch_image_quality"
                    )
            
            # Convert button
            if st.button("Convert All Files", key="batch_convert_button"):
                try:
                    with st.spinner(f"Converting {len(uploaded_files)} files..."):
                        # Save uploaded files to temp directory
                        input_paths = []
                        for uploaded_file in uploaded_files:
                            input_path = save_uploaded_file(uploaded_file, self.temp_dir)
                            input_paths.append(input_path)
                        
                        # Create output directory
                        output_dir = os.path.join(self.temp_dir, "batch_output")
                        os.makedirs(output_dir, exist_ok=True)
                        
                        # Prepare conversion options
                        options = {'preserve_formatting': preserve_formatting}
                        
                        if converter_type == "Spreadsheet":
                            options['encoding'] = encoding
                            options['include_index'] = include_index
                        elif converter_type == "Presentation" and output_format in ['png', 'jpg']:
                            options['image_quality'] = image_quality
                        
                        # Convert files
                        batch_converter = BatchConverter(converter)
                        converted_files = batch_converter.convert_batch(
                            input_paths,
                            output_format,
                            output_dir=output_dir,
                            **options
                        )
                        
                        # Display success message
                        st.success(f"Conversion successful! {len(converted_files)} files converted.")
                        
                        # Create a zip file containing all converted files
                        import zipfile
                        zip_path = os.path.join(self.temp_dir, "converted_files.zip")
                        with zipfile.ZipFile(zip_path, 'w') as zipf:
                            for file in converted_files:
                                if os.path.isdir(file):
                                    # For directories (like image outputs)
                                    for root, _, files in os.walk(file):
                                        for f in files:
                                            file_path = os.path.join(root, f)
                                            zipf.write(
                                                file_path,
                                                arcname=os.path.relpath(file_path, self.temp_dir)
                                            )
                                else:
                                    # For single files
                                    zipf.write(
                                        file,
                                        arcname=os.path.relpath(file, self.temp_dir)
                                    )
                        
                        # Download button for zip file
                        with open(zip_path, "rb") as f:
                            st.download_button(
                                label="Download All Converted Files (ZIP)",
                                data=f,
                                file_name="converted_files.zip",
                                mime="application/zip",
                                key="batch_download_button"
                            )
                
                except Exception as e:
                    st.error(f"Error during batch conversion: {str(e)}")
                    self.logger.error(f"Batch conversion error: {str(e)}", exc_info=True)
    
    def _preview_file(self, file_path: str, file_format: str):
        """
        Preview the converted file if possible.
        
        Args:
            file_path: Path to the file to preview
            file_format: Format of the file
        """
        st.subheader("Preview")
        
        # Check if file is a directory (for image outputs)
        if os.path.isdir(file_path):
            st.write("Preview of first few images:")
            image_files = [f for f in os.listdir(file_path) if f.endswith(f".{file_format}")]
            for i, image_file in enumerate(image_files[:3]):  # Show first 3 images
                image_path = os.path.join(file_path, image_file)
                st.image(image_path, caption=f"Page {i+1}", use_column_width=True)
            if len(image_files) > 3:
                st.write(f"... and {len(image_files) - 3} more pages")
            return
        
        # Preview based on format
        if file_format in ['md', 'markdown']:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            st.markdown(content)
        
        elif file_format in ['txt', 'csv', 'tsv', 'json']:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            st.text_area("File Content", content, height=300)
        
        elif file_format == 'html':
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            st.components.v1.html(content, height=500, scrolling=True)
        
        elif file_format in ['png', 'jpg', 'jpeg']:
            st.image(file_path, caption="Converted Image", use_column_width=True)
        
        elif file_format in ['xlsx', 'xls', 'ods']:
            try:
                import pandas as pd
                df = pd.read_excel(file_path)
                st.dataframe(df.head(10))
                if len(df) > 10:
                    st.write(f"... and {len(df) - 10} more rows")
            except Exception as e:
                st.write("Preview not available for this Excel file.")
                self.logger.warning(f"Excel preview error: {str(e)}")
        
        elif file_format == 'pdf':
            st.write("PDF preview not available. Please download the file to view it.")
        
        elif file_format in ['docx', 'doc', 'ppt', 'pptx']:
            st.write(f"{file_format.upper()} preview not available. Please download the file to view it.")
        
        else:
            st.write("Preview not available for this file format.")
    
    def __del__(self):
        """Clean up resources when the application is closed."""
        try:
            clean_temp_directory(self.temp_dir)
        except Exception as e:
            self.logger.error(f"Error cleaning up temporary directory: {str(e)}")


def main():
    """Main entry point for the application."""
    app = FileConverterApp()
    app.run()


if __name__ == "__main__":
    main()
