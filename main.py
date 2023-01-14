import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
# Creating a variable with glob package to have a string of filepath
filepaths = glob.glob("invoices/*")
# Iterate over all the objects in directory
for file in filepaths:
    # Create a variable of a read dataframe of an Excel file, you should define
    # a sheet_name and file
    df = pd.read_excel(file, sheet_name="Sheet 1")
    # Create a pdf instance with FPDF class
    pdf = FPDF("P", unit="mm", format="A4")
    # Add a page to a pdf
    pdf.add_page()
    # Create an instance of a file excluding directory and extension.
    filename = Path(file).stem
    # Extract from a filename only an invoice number
    invoice_nr, date = filename.split("-")
    # Select a font for pdf
    pdf.set_font(family="Times", size=16, style="B")
    # Create a cell for a new input, ln=1 means create a breakline
    pdf.cell(w=50, h=10, txt=f"Invoice number: {invoice_nr}", ln=1)

    # Create a pdf file in PDFs directory using filename and f-string
    pdf.set_font(family="Times", size=16, style="B")
    # Create a cell for a new input
    pdf.cell(w=50, h=10, txt=f"Date: {date}", ln=1)

    # Add a header
    columns = df.columns
    # Use list comprehension to rid out of underscore and capitalize
    columns = [item.replace("_", " ").capitalize() for item in columns]
    pdf.set_font(family="Times", style="B", size=8)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows
    # Iterate over rows in df file
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
    # We create cells, but we should convert txt=in string as row["product_id"] is int and
    # we'll get this error AttributeError: 'int' object has no attribute 'replace'
    # We add border here with border=1, and ln=1 to break the line
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Create a pdf file in PDFs directory using filename and f-string
    pdf.output(f"PDFs/{filename}.pdf")
