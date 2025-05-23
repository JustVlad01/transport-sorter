# Driver PDF Page Sorter

A Python desktop application for processing driver information from Excel files and sorting PDF documents using OCR technology.

## Features

### Excel Processor Tab
- Upload Excel files (.xlsx, .xls)
- Extract and store values from specific columns (C and J by default)
- Save mappings to a JSON file for later use
- View extracted data in a searchable table

### Route Splitter Tab
- Upload and process PDF files using OCR
- Extract customer references from each page
- Match customer references to routes using the stored data
- Split PDF into separate files by route
- Generate summary and logs of the processing

## Requirements

- Python 3.7+
- Dependencies listed in `requirements.txt`
- Tkinter (included in standard Python installation)
- Tesseract OCR engine installed on your system (for OCR functionality)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/transport-sorter.git
   cd transport-sorter
   ```

2. Create a virtual environment (recommended):
   ```
   python -m venv venv
   ```

3. Activate the virtual environment:
   - Windows:
     ```
     venv\Scripts\activate
     ```
   - macOS/Linux:
     ```
     source venv/bin/activate
     ```

4. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

5. Install Tesseract OCR:
   - Windows: Download and install from https://github.com/UB-Mannheim/tesseract/wiki
   - macOS: `brew install tesseract`
   - Linux: `sudo apt install tesseract-ocr`

## Usage

1. Start the application:
   ```
   python app.py
   ```

2. Excel Processor Tab:
   - Click "Select Excel File" to choose an Excel file from your computer
   - Click "Process File" to extract the data
   - View the extracted data in the table below
   - Use the search box to find specific values

3. Route Splitter Tab:
   - First, make sure you've processed an Excel file in the Excel Processor tab
   - Click "Select PDF File" to choose a PDF to process
   - Select an output directory where the split PDFs will be saved
   - Click "Process PDF" to start the OCR and splitting process
   - Monitor progress in the log area
   - When complete, check the output directory for the split PDFs

## Data Storage

- Extracted data is stored in `data/driver_data.json`
- The PDF splitting process creates individual PDF files in the selected output directory:
  - One PDF per route, containing all pages for that route
  - An "Unassigned_Pages.pdf" for pages without a recognized customer reference

## Project Structure

```
transport-sorter/
├── app.py                  # Main application file
├── requirements.txt        # Python dependencies
├── README.md               # Project documentation
├── uploads/                # Directory for uploaded files
└── data/                   # Directory for storing extracted data
    └── driver_data.json    # JSON file with extracted data
```

## How It Works

1. Excel Processor:
   - Reads Excel file and extracts values from columns C and J
   - Creates a mapping between customer references and routes
   - Stores the mapping in a JSON file

2. Route Splitter:
   - Uses OCR to extract text from each PDF page
   - Searches for customer reference patterns
   - Matches customer references against the stored data
   - Groups pages by route
   - Creates new PDFs for each route with the relevant pages

## License

[MIT License](LICENSE)

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Adjust path as needed
