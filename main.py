import pandas as pd
import glob #for loading multiple filepaths
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx") #get the filepaths in a python list called "filepaths"
print(filepaths)

#now read each filepath from the list filepaths
for filepath in filepaths:

    #now generate pdf docs
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # extract file name into the variable filename as the file name is the invoice number that must be printed in the pdf text
    filename = Path(filepath).stem #.stem gives the file name without the .xlsx extension

    parts = filename.split("-")  # Splits 10001-2023.1.18 into ["10001", "2023.1.18"] by splitting it at the hyphen "-"
    first_part = parts[0]  # Selects "10001"
    second_part = parts[1] # Selects the date "2023.1.18"

    invoice_no = first_part
    invoice_date = second_part
    print(invoice_no)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_no}")
    pdf.ln(10)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}")
    pdf.ln(15)

    # extract data from tables in xlxs files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    print("\n")

    #set table headers for each file
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=20, h=8, txt=f"Product ID", border=1, align="C")
    pdf.cell(w=50, h=8, txt=f"Product Name", border=1, align="C")
    pdf.cell(w=35, h=8, txt=f"Amount Purchased", border=1, align="C")
    pdf.cell(w=25, h=8, txt=f"Price Per Unit", border=1, align="C")
    pdf.cell(w=20, h=8, txt=f"Total Price", border=1, align="C")
    pdf.ln()
    #find total price of all items by summing the total_price column data
    total_price = df['total_price'].sum()

    #iterate over each row
    for index, row in df.iterrows():

        #extract data from each row
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=20, h=8, txt=str(row['product_id']), border=1, align="C")
        pdf.cell(w=50, h=8, txt=row['product_name'], border=1, align="C")
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), border=1, align="C")
        pdf.cell(w=25, h=8, txt=str(row['price_per_unit']), border=1, align="C")
        pdf.cell(w=20, h=8, txt=str(row['total_price']), border=1, align="C")

        pdf.ln()

    # Add final row for total price
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=130, h=8, txt="", border=0)  # Empty cells for alignment
    # pdf.cell(w=25, h=8, txt="Total:", border=1)  # Label
    pdf.cell(w=20, h=8, txt=str(total_price), border=1, align="C")  # Total value
    pdf.ln(15)

    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"The total due amount is {total_price} Euros")
    pdf.ln()
    pdf.cell(w=50, h=8, txt="PythonHow")
    pdf.image(r"G:\PYTHON PROJECTS\Excel to PDF Invoice Generator\pythonhow.png", 33, pdf.get_y(), w=10)


    pdf.output(f"PDFs/{filename}.pdf")  # automatically name and save the pdf docs in a folder




