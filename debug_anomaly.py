import pandas as pd

try:
    df = pd.read_excel('cleaned_data.xlsx')
    col_rev2 = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2"
    
    if col_rev2 in df.columns:
        print(f"Checking column: {col_rev2}")
        # Check for 24228
        matches = df[df[col_rev2] == 24228]
        if not matches.empty:
            print(f"FOUND {len(matches)} rows with value 24228.")
            print(matches[["Nomor Induk Mahasiswa (NIM)", col_rev2]])
            # store index or NIM
            target_nim = matches["Nomor Induk Mahasiswa (NIM)"].iloc[0]
            print(f"Target NIM: {target_nim}")
            # We need to see the ORIGINAL value to understand why logic failed.
            # But wait, original column "rev2" is created FROM original. 
            # I can't see the original value from cleaned_data.xlsx unless I kept it.
            # The original column name is in cleaning.py
            
            # Let's check original data.xlsx for this 24228
            df_orig = pd.read_excel('data.xlsx')
            col_orig = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan)"
            
            # Find row with 24228 in original or something similar
            # Or better, since we have the cleaned file, verify if 24228 is physically there
            print("Verifying if 24228 is in cleaned file...")
            
        else:
            print("Value 24228 NOT FOUND in cleaned file (Exact Match).")
            
        print("\n--- Inspecting Large Values > 100 ---")
        numeric_vals = pd.to_numeric(df[col_rev2], errors='coerce')
        large = numeric_vals[numeric_vals > 100]
        if not large.empty:
            print(large)
        else:
            print("No values > 100 found.")
            
        # Check if file was updated recently?
        import os
        import time
        mtime = os.path.getmtime('cleaned_data.xlsx')
        print(f"\nFile modified: {time.ctime(mtime)}")

    else:
        print(f"Column {col_rev2} not found.")

except Exception as e:
    print(e)
