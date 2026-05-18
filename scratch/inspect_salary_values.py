import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_FILE = os.path.join(BASE_DIR, 'data', 'processed', 'cleaned_data.xlsx')

if not os.path.exists(DATA_FILE):
    DATA_FILE = os.path.join(BASE_DIR, 'data', 'raw', 'data.xlsx')

try:
    df = pd.read_excel(DATA_FILE)
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    if col_salary in df.columns:
        print("Unique values in salary column:")
        print(df[col_salary].value_counts(dropna=False))
    else:
        print(f"Column '{col_salary}' not found. Available columns:")
        for col in df.columns:
            if 'pendapatan' in col.lower() or 'gaji' in col.lower() or 'berapa' in col.lower():
                print(f"- {col}")
except Exception as e:
    print(f"Error: {e}")
