import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data', 'processed')
DATA_FILE = os.path.join(DATA_DIR, 'cleaned_data.xlsx')

try:
    df = pd.read_excel(DATA_FILE)
    print("Additional Columns search:")
    for col in df.columns:
        if any(keyword in col.lower() for keyword in ['riset', 'teliti', 'demonstrasi']):
            print(f"- {col}")
            
except Exception as e:
    print(f"Error: {e}")
