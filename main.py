import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Ta funkcja tworzy listę z filepathsami
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # 1. Extract the date and a filename from the filepath
    filename = Path(filepath).stem
    filename_splited = filename.split('-')
    invoice_number, invoice_date = filename.split('-')

    # 2. Generate a PDF file for each invoice
    pdf = FPDF(orientation="P", unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=0, h=8, txt=f"Invoice nr {invoice_number}", ln=1)
    pdf.cell(w=0, h=8, txt=f"Date {invoice_date}", ln=1)

    # 3. Import invoices from the folder
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # 4. Generate headers of the columns
    columns = list(df.columns)
    for column_title in columns:
        pdf.set_font(family='Times', size=10, style='B')

        title_to_display = 'Amount' \
            if column_title == 'amount_purchased' \
            else column_title.replace('_', ' ').title()

        if column_title == 'product_name':
            pdf.cell(w=70, h=8, txt=title_to_display, border=1)
        else:
            pdf.cell(w=30, h=8, txt=title_to_display, border=1)
    pdf.ln()

    # 5. Generate cells and calculate total price
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        # pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # 6. Generate one single cell on the right side holding a total price
    for column_title in columns:
        if column_title == 'product_name':
            pdf.cell(w=70, h=8)
        elif column_title == 'total_price':
            pdf.cell(w=30, h=8, border=1, ln=1, txt=str(df['total_price'].sum()))
        else:
            pdf.cell(w=30, h=8)

    # 7. Display a total due amount and a company logo
    pdf.ln(h=15)
    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=0, h=8, txt=f"The total due amount is {str(df['total_price'].sum())} euros.", ln=1)
    pdf.cell(w=23, h=8, txt='PythonHow')
    pdf.image('pythonhow.png', w=10)

    pdf.output(f"PDFs/{filename}.pdf")
