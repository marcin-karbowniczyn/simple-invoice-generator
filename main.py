import pandas as pd
import glob

# Ta funkcja tworzy listę z filepathsami
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    print(df)
