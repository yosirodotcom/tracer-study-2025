import pandas as pd

try:
    import os
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DATA_CLEANED = os.path.join(BASE_DIR, 'data', 'processed', 'cleaned_data.xlsx')
    df = pd.read_excel(DATA_CLEANED)
    print(f"Total columns: {len(df.columns)}")
    print("--- Column List ---")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")
except Exception as e:
    print(f"Error: {e}")
