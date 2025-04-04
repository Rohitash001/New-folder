import os
import pdfplumber
import pandas as pd

INPUT_DIR = "input_pdfs"
OUTPUT_DIR = "output_excels"

os.makedirs(OUTPUT_DIR, exist_ok=True)

def extract_tables_from_pdf(pdf_path):
    extracted_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"Processing page {page_num} of {pdf_path}...")
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables):
                if table:  # Avoid empty tables
                    df = pd.DataFrame(table)
                    df.columns = df.iloc[0]  # First row as header
                    df = df[1:].reset_index(drop=True)
                    extracted_tables.append((f"Page{page_num}_Table{table_idx+1}", df))

    return extracted_tables

def process_all_pdfs():
    for filename in os.listdir(INPUT_DIR):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(INPUT_DIR, filename)
            excel_filename = os.path.splitext(filename)[0] + ".xlsx"
            excel_path = os.path.join(OUTPUT_DIR, excel_filename)

            tables = extract_tables_from_pdf(pdf_path)

            if tables:
                with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                    for sheet_name, df in tables:
                        # Ensure sheet names are within Excel limits
                        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                print(f"Saved extracted tables to {excel_path}")
            else:
                print(f"No tables found in {pdf_path}")

if __name__ == "__main__":
    process_all_pdfs()
