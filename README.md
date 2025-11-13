# pdf_to_excel
A lightweight, free-to-use Python utility that extracts all tables from a PDF file using pdfplumber, converts them into structured pandas DataFrames, and exports them into a single Excel spreadsheet. Ideal for automating data extraction workflows from multi-page PDFs.
## Features

- Extracts **all tables** from every page in a PDF.
- Automatically **tags each row** with the page number where it was found.
- Merges all tables into **one unified DataFrame**.
- Exports the final output to **Excel (.xlsx)**.
- Works on Windows, macOS, and Linux.

## Requirements

Install dependencies using:

```
pip install pdfplumber pandas
```

Or for macOS:

```
pip3 install pdfplumber pandas
# or
python3 -m pip install pdfplumber pandas
```

## How It Works

1. Opens the PDF with pdfplumber.
2. Iterates through each page.
3. Detects and extracts any table structures.
4. Converts each table into a DataFrame.
5. Appends a column called **PAGE** indicating the PDF page number.
6. Concatenates all tables into one final DataFrame.
7. Exports the result to an Excel file.

```

## Example Usage

```
python transform_pdf_to_excel.py

```
Default example paths used in the script:

```
pdf_path = r"C:\Example\Path\Input_Document.pdf"
output_path = r"C:\Example\Path\Extracted_Tables.xlsx"
```

Replace these with any valid paths on your system.

## Output

The resulting Excel file will contain:

- All detected tables merged into one sheet.
- A `PAGE` column showing where each row came from.

## Notes & Limitations

- Extraction accuracy depends on the structure and quality of the source PDF.
- Complex layouts or scanned PDFs may require OCR or custom parsing.
- pdfplumber works best with **digitally-generated PDFs**, not scanned images.

## Contributing

Pull requests and improvements are welcome!
