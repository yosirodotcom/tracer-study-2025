import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data', 'processed')
DATA_FILE = os.path.join(DATA_DIR, 'cleaned_data.xlsx')

try:
    df = pd.read_excel(DATA_FILE)
    print("Columns found:")
    for col in df.columns:
        if any(keyword in col.lower() for keyword in ['metode', 'pembelajaran', 'kuliah', 'praktikum', 'diskusi', 'lapangan', 'magang']):
            print(f"- {col}")
            
    # Also print first few rows of these columns to check values
    relevant_cols = [c for c in df.columns if any(k in c.lower() for k in ['metode', 'pembelajaran', 'kuliah', 'praktikum', 'diskusi', 'lapangan', 'magang'])]
    if relevant_cols:
        print("\nSample Data:")
        print(df[relevant_cols].head(3).to_string())
        
except Exception as e:
    print(f"Error: {e}")
