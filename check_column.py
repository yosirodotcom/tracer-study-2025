import pandas as pd

try:
    df = pd.read_excel('data.xlsx')
    col = "Apa jenis Perusahaan/Instansi/Institusi tempat Anda bekerja sekarang?"
    
    if col in df.columns:
        print(f"Total rows: {len(df)}")
        print(f"Unique values in '{col}': {df[col].nunique()}")
        
        allowed_list = [
            "Instansi Pemerintah", 
            "Organisasi non-profit/Lembaga Swadaya Masyarakat",
            "Perusahaan Swasta",
            "Wiraswasta/perusahaan sendiri",
            "BUMN/BUMD",
            "Institusi/Organisasi Multilateral"
        ]
        
        print("\n--- Allowed List ---")
        for x in allowed_list:
            print(f"- {x}")
            
        print("\n--- Data Distribution (Top 20) ---")
        counts = df[col].value_counts().head(20)
        print(counts)
        
        # Check coverage
        in_list = df[col].isin(allowed_list).sum()
        print(f"\nRows exactly matching list: {in_list} ({in_list/len(df)*100:.2f}%)")
        
        # Check case-insensitive coverage
        allowed_lower = [x.lower() for x in allowed_list]
        in_list_lower = df[col].astype(str).str.lower().isin(allowed_lower).sum()
        print(f"Rows matching list (case-insensitive): {in_list_lower} ({in_list_lower/len(df)*100:.2f}%)")

    else:
        print(f"Column '{col}' not found.")
        print(df.columns.tolist())

except Exception as e:
    print(f"Error: {e}")
