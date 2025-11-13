# Required libraries: pdfplumber and pandas
# Installation:
#   Windows:  pip install pdfplumber pandas
#   macOS:    pip3 install pdfplumber pandas
#             or: python3 -m pip install pdfplumber pandas

import pdfplumber
import pandas as pd

# Example file paths — replace these with your own as needed
pdf_path = r"C:\Example\Path\Input_Document.pdf"
output_path = r"C:\Example\Path\Extracted_Tables.xlsx"

all_tables = []

with pdfplumber.open(pdf_path) as pdf:
    for page_num, page in enumerate(pdf.pages, start=1):
        print(f"Processing page {page_num}...")

        # Extract ALL tables found on the current page (returns a list of lists)
        tables = page.extract_tables()

        if not tables:
            print(f"  ✖ Page {page_num}: No tables detected")
            continue

        print(f"  ✔ Page {page_num}: {len(tables)} table(s) detected")

        for table in tables:
            df = pd.DataFrame(table)
            
            # Optional: insert a column indicating the PDF page number
            df.insert(0, "PAGE", page_num)
            
            all_tables.append(df)

# Validate that at least one table was found
if not all_tables:
    raise Exception("No tables were detected on any page of the PDF.")

# Combine all extracted tables into a single DataFrame
df_final = pd.concat(all_tables, ignore_index=True)

# Optional: if the first row represents headers, uncomment and adjust:
# header = df_final.iloc[0]
# df_final = df_final[1:]
# df_final.columns = header

# Export to Excel
df_final.to_excel(output_path, index=False)

print("\nExtraction completed successfully.")
print("Spreadsheet generated at:", output_path)
