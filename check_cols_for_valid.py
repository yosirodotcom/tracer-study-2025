import pandas as pd

try:
    df = pd.read_excel('cleaned_data.xlsx')
    print("Columns:")
    for c in df.columns:
        print(f"- {c}")
        
    target_col = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2"
    if target_col in df.columns:
        print(f"\nFound '{target_col}'")
    else:
        print(f"\n'{target_col}' NOT FOUND")
        
    check_col = "Apakah anda telah mendapatkan pekerjaan <=6 bulan / termasuk bekerja sebelum lulus?"
    if check_col in df.columns:
        print(f"Found '{check_col}'")
    else:
        print(f"'{check_col}' NOT FOUND")

except Exception as e:
    print(e)
