import pandas as pd

df = pd.read_excel('cleaned_data.xlsx')

cols_to_check = ['Jurusan', 'prodi', 'diploma']

for col in cols_to_check:
    print(f"\n--- Unique {col} ---")
    if col in df.columns:
        unique_vals = df[col].unique()
        print(f"Count: {len(unique_vals)}")
        for val in unique_vals:
            print(f"- {val}")
    else:
        print(f"Column '{col}' NOT FOUND.")
