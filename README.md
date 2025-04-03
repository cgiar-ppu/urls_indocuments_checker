# Document URL Checker

This Streamlit application allows you to check the validity of URLs found within Microsoft Word (.docx) documents. It extracts all URLs from the document, attempts to access each one, and provides a report showing which URLs are working and which are not.

## Features

- Upload .docx files through a simple web interface
- Automatically extracts URLs from both document text and tables
- Checks each URL's accessibility
- Displays a preview of the content from working URLs
- Shows detailed error messages for failed URLs
- Progress tracking for URL checking process

## Installation

1. Make sure you have Python 3.7+ installed
2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the Streamlit application:
```bash
streamlit run app.py
```

2. Open your web browser and navigate to the URL shown in the terminal (typically http://localhost:8501)
3. Upload your .docx file using the file uploader
4. Wait for the application to process and check all URLs
5. Review the results in the "Working URLs" and "Failed URLs" sections

## Notes

- The application includes a timeout of 10 seconds for each URL request
- Content previews are limited to 500 characters
- The application uses a custom User-Agent to avoid being blocked by some websites
- URLs are deduplicated before checking 