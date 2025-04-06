# File Format Converter

A comprehensive Streamlit application for converting between various file formats with an intuitive user interface, batch processing capabilities, and file preview functionality.

## Features

- **Multiple Format Support**: Convert between a wide range of file formats:
  - **Documents**: Markdown, Word (DOCX/DOC), PDF, HTML, TXT, RTF, ODT
  - **Presentations**: PowerPoint (PPTX/PPT), PDF, HTML, PNG/JPG
  - **Spreadsheets**: CSV, Excel (XLSX/XLS), JSON, TSV, ODS, HTML, Markdown tables

- **User-Friendly Interface**: Tabbed layout for easy navigation between different conversion types

- **Advanced Options**: Customize your conversions with format-specific options:
  - Preserve formatting
  - Custom style templates
  - Sheet selection for Excel files
  - Custom delimiters for CSV/TSV
  - File encoding options
  - Image quality settings

- **Batch Processing**: Convert multiple files at once with the same settings

- **File Preview**: Preview converted files directly in the application before downloading

- **Comprehensive Logging**: Detailed logs for troubleshooting and monitoring

## Installation

### Prerequisites

- Python 3.8 or higher
- Pandoc (for document format conversions)
- LibreOffice (for some document and presentation conversions)
- pdf2htmlEX (optional, for enhanced PDF to HTML conversion)

### Installing External Dependencies

#### Pandoc

```bash
# Ubuntu/Debian
sudo apt-get install pandoc

# macOS
brew install pandoc

# Windows
# Download from https://pandoc.org/installing.html
```

#### LibreOffice

```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# macOS
brew install --cask libreoffice

# Windows
# Download from https://www.libreoffice.org/download/
```

#### pdf2htmlEX (Optional)

```bash
# Ubuntu/Debian
sudo apt-get install pdf2htmlex

# For other platforms, see: https://github.com/pdf2htmlEX/pdf2htmlEX
```

### Installing the Application

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/file-converter-app.git
   cd file-converter-app
   ```

2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Application Locally

```bash
streamlit run main.py
```

The application will start and open in your default web browser at `http://localhost:8501`.

### Using the Application

1. **Select Conversion Type**:
   - Choose the appropriate tab for your file type (Document, Presentation, Spreadsheet, or Batch Conversion)

2. **Upload File(s)**:
   - Click "Browse files" to upload your file(s)
   - For batch conversion, you can select multiple files

3. **Select Output Format**:
   - Choose your desired output format from the dropdown menu

4. **Configure Options** (optional):
   - Expand the "Advanced Options" section to customize your conversion

5. **Convert**:
   - Click the "Convert" button to process your file(s)

6. **Preview and Download**:
   - Preview the converted file(s) in the application
   - Click "Download" to save the converted file(s) to your computer

## Docker Deployment

### Building the Docker Image

```bash
docker build -t file-converter-app .
```

### Running the Docker Container

```bash
docker run -p 8501:8501 file-converter-app
```

Access the application at `http://localhost:8501`.

## Streamlit Cloud Deployment

1. Fork this repository to your GitHub account

2. Log in to [Streamlit Cloud](https://streamlit.io/cloud)

3. Create a new app and select your forked repository

4. Set the main file path to `main.py`

5. Deploy the application

## Project Structure

```
file_converter_app/
├── main.py                  # Main application entry point
├── requirements.txt         # Python dependencies
├── Dockerfile              # Docker configuration
├── README.md               # This documentation
├── src/
│   ├── __init__.py         # Package initialization
│   ├── converter_base.py   # Base classes for converters
│   ├── converter_factory.py # Factory for creating converters
│   ├── document_converter.py # Document format converter
│   ├── presentation_converter.py # Presentation format converter
│   ├── spreadsheet_converter.py # Spreadsheet format converter
│   └── utils.py            # Utility functions
└── tests/                  # Test files
    └── ...
```

## Limitations

- Some conversions require external dependencies (Pandoc, LibreOffice)
- PDF to editable format conversions may not preserve all formatting
- Complex document layouts might not convert perfectly between all formats
- Large files may take longer to process

## Troubleshooting

### Common Issues

1. **Missing Dependencies**:
   - Check that all required external dependencies are installed
   - The application will show warnings for missing dependencies

2. **Conversion Errors**:
   - Check the application logs for detailed error information
   - Ensure your input file is not corrupted

3. **Performance Issues**:
   - Large files may require more memory and processing time
   - Consider splitting large files into smaller chunks

### Logs

Logs are saved to `file_converter.log` in the application directory.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- [Streamlit](https://streamlit.io/) for the web application framework
- [Pandas](https://pandas.pydata.org/) for data processing
- [Pandoc](https://pandoc.org/) for document conversions
- [python-docx](https://python-docx.readthedocs.io/) for Word document processing
- [pdf2docx](https://github.com/dothinking/pdf2docx) for PDF to Word conversions
