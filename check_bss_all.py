import pandas as pd
import os

base = r'C:\Users\aryan\Desktop\WORK TANMAY SIR\po\PurchaseOrder\PO'

# Check BSS 24-25 sheet
full = os.path.join(base, r'2023-2024\BSS.xlsx')
xl = pd.ExcelFile(full)

for sh in ['24-25', 'SUMMARY 23-24', 'SUMMARY 24-25']:
    if sh in xl.sheet_names:
        df = pd.read_excel(full, sheet_name=sh, header=0)
        print(f"\nSheet '{sh}' shape: {df.shape}")
        print(f"Columns: {list(df.columns[:13])}")
        print(f"First 3 rows:\n{df.head(3).to_string()}")
        if 'AMOUNT' in df.columns:
            df['AMOUNT'] = pd.to_numeric(df['AMOUNT'], errors='coerce')
            print(f"Total: {df['AMOUNT'].sum():,.2f}")

# Also check W.Liquid index for 2023-24 coverage
wellness_files_2324 = [
    (r'2023-2024\W. LIQUID -  APPROVED ORDER FORM.xlsx', 'LIQUID 23-24'),
]
for f, label in wellness_files_2324:
    full2 = os.path.join(base, f)
    if os.path.exists(full2):
        xl2 = pd.ExcelFile(full2)
        print(f"\n=== {label} ===")
        print(f"Sheets: {xl2.sheet_names}")
