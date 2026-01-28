import pandas as pd

df = pd.read_excel('cleaned_data.xlsx')
cols_removed = ['Timestamp', 'Email Address', 'Nama Mahasiswa', 'Nomor Handphone', 'NIK']
present_cols = [c for c in cols_removed if c in df.columns]

if not present_cols:
    print("SUCCESS: All sensitive columns have been removed.")
else:
    print(f"FAILURE: The following columns are still present: {present_cols}")
    
print(f"Remaining columns count: {len(df.columns)}")
print("First 5 columns:", df.columns.tolist()[:5])
