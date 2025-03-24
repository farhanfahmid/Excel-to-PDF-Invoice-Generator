import pandas as pd
import glob #for loading multiple filepaths
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx") #get the filepaths in a python list called "filepaths"
print(filepaths)

#now read each filepath from the list filepaths
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    print("\n")
    #now generate pdf docs
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # extract file name into the variable filename as the file name is the invoice number that must be printed in the pdf text
    filename = Path(filepath).stem #.stem gives the file name without the .xlsx extension

    parts = filename.split("-")  # Splits 10001-2023.1.18 into ["10001", "2023.1.18"] by splitting it at the hyphen "-"
    first_part = parts[0]  # Selects "10001"

    invoice_no = first_part
    print(invoice_no)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_no}")
    pdf.output(f"PDFs/{filename}.pdf") # automatically name and save the pdf docs in a folder




