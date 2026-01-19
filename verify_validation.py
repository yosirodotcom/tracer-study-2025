import pandas as pd
import numpy as np

try:
    df = pd.read_excel('cleaned_data.xlsx')
    
    col_valid = "valid column for Dalam berapa bulan Anda mendapatkan pekerjaan"
    col_duration = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2"
    col_check = "Apakah anda telah mendapatkan pekerjaan <=6 bulan / termasuk bekerja sebelum lulus?"
    
    if col_valid in df.columns:
        print(f"Checking '{col_valid}'...")
        # Check Total Flagged
        flagged = df[df[col_valid] == 1]
        print(f"Total flagged (1): {len(flagged)}")
        
        # Check logic for flagged rows
        # Expectation: (Ya AND > 6) OR (Tidak AND <= 6)
        if not flagged.empty:
            errors = 0
            for idx, row in flagged.iterrows():
                val_6 = str(row[col_check]).strip()
                dur = row[col_duration]
                
                # Expect VALID rows here
                cond1 = (val_6 == "Ya") and (pd.notna(dur) and dur <= 6)
                cond2 = (val_6 == "Tidak") and ((pd.notna(dur) and dur > 6) or pd.isna(dur))
                
                if not (cond1 or cond2):
                     print(f"Logic Mismatch at row {idx}: Flagged but condition not met? {val_6} | {dur}")
                     errors += 1
            
            if errors == 0:
                print("SUCCESS: All flagged rows meet the logic criteria.")
            else:
                print(f"FAILURE: {errors} flagged rows did NOT meet criteria.")
        
        # Check logic for non-flagged rows (should be 0)
        # Expectation: NOT ((Ya AND > 6) OR (Tidak AND <= 6))
        # Meaning: (Ya AND <= 6) OR (Tidak AND > 6) OR (NaNs etc)
        non_flagged = df[df[col_valid] == 0]
        print(f"Total non-flagged (0): {len(non_flagged)}")
        
        print("\nSample Flagged Rows:")
        print(flagged[[col_check, col_duration, col_valid]].head())
    else:
        print(f"Column '{col_valid}' NOT FOUND.")

except Exception as e:
    print(f"Error: {e}")
