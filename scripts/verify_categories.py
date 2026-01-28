import pandas as pd

# Load the CLEANED data
try:
    df = pd.read_excel('cleaned_data.xlsx')
    col_rev = "Apa jenis Perusahaan/Instansi/Institusi tempat Anda bekerja sekarang? rev"
    
    allowed_list = [
        "Instansi Pemerintah", 
        "Organisasi non-profit/Lembaga Swadaya Masyarakat",
        "Perusahaan Swasta",
        "Wiraswasta/perusahaan sendiri",
        "BUMN/BUMD",
        "Institusi/Organisasi Multilateral",
        "lainnya"
    ]
    
    if col_rev in df.columns:
        print(f"Checking column: '{col_rev}'")
        unique_vals = df[col_rev].unique()
        print("Unique values found:")
        for val in unique_vals:
            status = "OK" if val in allowed_list else "FAIL"
            print(f"[{status}] {val}")
            
        # Assert
        invalid_vals = [v for v in unique_vals if v not in allowed_list]
        if not invalid_vals:
            print("\nSUCCESS: All values are within the allowed list.")
        else:
            print(f"\nFAILURE: Found invalid values: {invalid_vals}")
            
        print("\nDistribution:")
        print(df[col_rev].value_counts())
        
    else:
        print(f"Column '{col_rev}' NOT FOUND in cleaned_data.xlsx")

except Exception as e:
    print(f"Error: {e}")
