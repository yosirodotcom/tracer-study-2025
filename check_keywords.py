import pandas as pd
import os

file_path = 'cleaned_data.xlsx'
if not os.path.exists(file_path):
    print("cleaned_data.xlsx not found, trying data.xlsx")
    file_path = 'data.xlsx'

try:
    df = pd.read_excel(file_path)
    col_to_check = 'prodi' if 'prodi' in df.columns else 'Program Studi'
    
    if col_to_check in df.columns:
        print(f"Checking column: {col_to_check}")
        keywords = ['Kapuas Hulu', 'Sanggau', 'Sukamara']
        for k in keywords:
            matches = df[df[col_to_check].astype(str).str.contains(k, case=False, na=False)]
            print(f"Keyword '{k}': {len(matches)} matches")
            if not matches.empty:
                print(matches[col_to_check].unique())
    else:
        print(f"Column '{col_to_check}' not found.")
        print(df.columns.tolist())

except Exception as e:
    print(f"Error: {e}")
