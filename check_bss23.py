import pandas as pd
import os

base = r'C:\Users\aryan\Desktop\WORK TANMAY SIR\po\PurchaseOrder\PO'
full = os.path.join(base, r'2023-2024\BSS.xlsx')
xl = pd.ExcelFile(full)

# Read the 23-24 sheet
df = pd.read_excel(full, sheet_name='23-24', header=0)
print(f"23-24 sheet shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
print(f"\nFirst 5 rows:")
print(df.head(5).to_string())
print(f"\nLast 5 rows:")
print(df.tail(5).to_string())
print(f"\nData types:")
print(df.dtypes)
print(f"\nAMOUNT column stats:")
print(df['AMOUNT'].describe() if 'AMOUNT' in df.columns else "No AMOUNT column")

# Filter Apr 2023 onwards
from datetime import date
df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
april_2023 = pd.Timestamp('2023-04-01')
df_2324 = df[df['DATE'] >= april_2023]
print(f"\nRows from Apr 2023 onwards: {len(df_2324)}")
print(f"Date range: {df_2324['DATE'].min()} to {df_2324['DATE'].max()}")

# Total amount
if 'AMOUNT' in df_2324.columns:
    df_2324_clean = df_2324.copy()
    df_2324_clean['AMOUNT'] = pd.to_numeric(df_2324_clean['AMOUNT'], errors='coerce')
    total = df_2324_clean['AMOUNT'].sum()
    print(f"Total AMOUNT 2023-24: {total:,.2f}")
    print(f"Non-null AMOUNT rows: {df_2324_clean['AMOUNT'].notna().sum()}")
