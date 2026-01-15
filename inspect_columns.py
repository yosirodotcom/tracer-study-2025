import pandas as pd

df = pd.read_excel('cleaned_data.xlsx')

col_status = "Jelaskan status Anda saat ini?"
col_duration = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan)"

print(f"--- Column: {col_status} ---")
if col_status in df.columns:
    print(df[col_status].value_counts().head(10))
else:
    print("Column Status NOT FOUND")

print(f"\n--- Column: {col_duration} ---")
if col_duration in df.columns:
    print(df[col_duration].unique()[:20])
else:
    print("Column Duration NOT FOUND")
