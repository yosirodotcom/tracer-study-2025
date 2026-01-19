import pandas as pd

df = pd.read_excel('data.xlsx')
col = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan)"

print("Checking raw data...")
for val in df[col]:
    try:
        # Check if it looks like 24228
        s = str(val)
        if "24228" in s:
             print(f"RAW FOUND: '{val}' type {type(val)}")
    except:
        pass
