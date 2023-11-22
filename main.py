import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Ta funkcja tworzy listę z filepathsami
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # 1. Import invoices from the folder
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # 2. Extract the date and a filename from the filepath
    filename = Path(filepath).stem
    filename_splited = filename.split('-')
    invoice_number = filename_splited[0]
    invoice_date = filename_splited[1]

    # 2. Generate a PDF file for each invoice
    pdf = FPDF(orientation="P", unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=0, h=8, txt=f"Invoice nr {invoice_number}", ln=1)
    pdf.cell(w=0, h=8, txt=f"Date {invoice_date}", ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
