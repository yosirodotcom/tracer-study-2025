import pandas as pd
import os

file_path = 'cleaned_data.xlsx' if os.path.exists('cleaned_data.xlsx') else 'data.xlsx'
print(f"Reading {file_path}...")
try:
    df = pd.read_excel(file_path)
    print("Columns:")
    for col in df.columns:
        print(f"- {col}")
except Exception as e:
    print(f"Error: {e}")
