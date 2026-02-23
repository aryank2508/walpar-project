import pandas as pd
import os

base = r'C:\Users\aryan\Desktop\WORK TANMAY SIR\po\PurchaseOrder\PO'

# Check BSS 23-24 sheet
full = os.path.join(base, r'2023-2024\BSS.xlsx')
xl = pd.ExcelFile(full)
print("BSS sheets:", xl.sheet_names)

# Find the 23-24 sheet
for sh in xl.sheet_names:
    if '23' in sh:
        df = pd.read_excel(full, sheet_name=sh, header=0)
        print(f"\nSheet '{sh}' shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"Sample rows:")
        print(df.head(5).to_string())
        count_2023 = df[df.iloc[:,0].apply(lambda x: hasattr(x, 'year') and x.year >= 2023)].shape[0] if len(df) > 0 else 0
        print(f"\nRows from 2023+: {count_2023}")
        break

# Check wellness index files
print("\n=== WELLNESS INDEX FILES (2023-24) ===")
wellness_files = [
    r'2023-2024\W. LIQUID -  APPROVED ORDER FORM.xlsx',
    r'2023-2024\W. SACHET POWDER - WELLNESS - APPROVED ORDER FORM.xlsx',
    r'2023-2024\W. TAB & CAP WNL - APPROVED ORDER FORM.xlsx',
]
for f in wellness_files:
    full2 = os.path.join(base, f)
    if os.path.exists(full2):
        xl2 = pd.ExcelFile(full2)
        print(f"\n{os.path.basename(f)}: sheets={xl2.sheet_names[:5]}")
        df2 = pd.read_excel(full2, sheet_name=xl2.sheet_names[0], header=None, nrows=5)
        for i, row in df2.iterrows():
            non_null = [(j, str(v)[:60]) for j, v in enumerate(row) if str(v).strip() not in ['', 'nan']]
            if non_null:
                print(f"  Row {i}: {non_null}")
