import pandas as pd
from fpdf import FPDF
from pathlib import Path
import glob


filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(50, 8, f"Invoice nr. {invoice_nr}", ln=True)
    pdf.cell(50, 8, f"Date: {date}", ln=True)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = list(df.columns)

    columns = [item.replace("_", " ").title() for item in columns]

    # add a header
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 8, str(columns[0]), border=1)
    pdf.cell(70, 8, str(columns[1]), border=1)
    pdf.cell(30, 8, str(columns[2]), border=1)
    pdf.cell(30, 8, str(columns[3]), border=1)
    pdf.cell(30, 8, str(columns[4]), border=1, ln=True)

    # add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", style="B", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(30, 8, str(row["product_id"]), border=1)
        pdf.cell(70, 8, str(row["product_name"]), border=1)
        pdf.cell(30, 8, str(row["amount_purchased"]), border=1)
        pdf.cell(30, 8, str(row["price_per_unit"]), border=1)
        pdf.cell(30, 8, str(row["total_price"]), border=1, ln=True)

    pdf.output(f"PDFs/{filename}.pdf")
