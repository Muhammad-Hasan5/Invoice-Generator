import pandas as pd
from fpdf import FPDF
from pathlib import Path
import glob


filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(50, 8, f"Invoice nr. {invoice_nr}", ln=True)
    pdf.cell(50, 8, f"Date: {date}", ln=True)
    pdf.output(f"PDFs/{filename}.pdf")
