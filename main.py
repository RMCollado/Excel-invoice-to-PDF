import pandas as pd
import glob
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice number: {invoice_nr}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    #  add header
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = [s.replace("_", " ").title() for s in df.columns]

    pdf.set_font(family="Times", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    #  add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    #  get total
    total_sum = str(df["total_price"].sum())
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=65, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)

    pdf.cell(w=30, h=8, txt=total_sum, border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=30, txt=f"The total price is {total_sum}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font(family="Times", size=10, style="b")
    pdf.cell(w=50, h=8, txt=f"Intelligent Technology Solutions")
    pdf.image("image.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
