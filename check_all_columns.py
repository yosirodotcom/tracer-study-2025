import pandas as pd

try:
    df = pd.read_excel('cleaned_data.xlsx')
    print(f"Total columns: {len(df.columns)}")
    print("--- Column List ---")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")
except Exception as e:
    print(f"Error: {e}")
