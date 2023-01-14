import pandas as pd
import glob


filepaths = glob.glob("invoices/*")

for file in filepaths:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    print(df)