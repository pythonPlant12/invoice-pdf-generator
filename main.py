import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*")

for file in filepaths:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    pdf = FPDF("P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(file).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=10, txt=f"Invoice number: {invoice_nr}", ln=1)
    pdf.cell(w=50, h=10, txt=f"Date: {date}", ln=1)

    columns = df.columns
    columns = [item.replace("_", " ").capitalize() for item in columns]
    pdf.set_font(family="Times", style="B", size=8)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt=str(""))
    pdf.cell(w=70, h=8, txt=str(""))
    pdf.cell(w=30, h=8, txt=str(""))
    pdf.cell(w=30, h=8, txt=str(""))
    pdf.cell(w=30, h=8, txt=str(f"{total} EUR"), border=1, ln=1)

    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=70, h=8, txt=str(f"The total price of this invoice is: {total} EUR."), ln=1)
    pdf.cell(w=65, h=8, txt="Thanks for working with us")
    pdf.image("5968350.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
