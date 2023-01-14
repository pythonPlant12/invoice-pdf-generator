import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path
# Creating a variable with glob package to have a string of filepath
filepaths = glob.glob("invoices/*")
# Iterate over all the objects in directory
for file in filepaths:
    # Create a variable of a read_excel function, you should define a sheet_name and file
    df = pd.read_excel(file, sheet_name="Sheet 1")
    # Create a pdf instance with FPDF class
    pdf = FPDF("P", unit="mm", format="A4")
    # Add a page to a pdf
    pdf.add_page()
    # Create an instance of a file excluding directory and extension.
    filename = Path(file).stem
    # Extract from a filename only an invoice number
    invoice_nr = filename.split("-")[0]
    # Select a font for pdf
    pdf.set_font(family="Times", size=16, style="B")
    # Create a cell for a new input
    pdf.cell(w=50, h=16, txt=f"Invoice number: {invoice_nr}")
    # Create a pdf file in PDFs directory using filename and f-string
    pdf.output(f"PDFs/{filename}.pdf")
