import pandas as pd
import table_jml_responden as tr
import os

print("--- Verifying table_jml_responden.py ---")

# Load data
file_path = 'cleaned_data.xlsx' if os.path.exists('cleaned_data.xlsx') else 'data.xlsx'
print(f"Loading data from {file_path}...")
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"Failed to load data: {e}")
    exit(1)

# Ensure necessary columns exist (if using data.xlsx, 'prodi' might not exist if created in cleaning.py)
# Check if we need to mock or if cleaning.py results are available
if 'prodi' not in df.columns:
    print("Warning: 'prodi' column not found (likely reading raw data.xlsx).")
    # Quick fix for testing if 'cleaned_data.xlsx' isn't available or fully correct
    if 'Program Studi' in df.columns:
        print("Creating 'prodi' from 'Program Studi' for testing purposes...")
        split_data = df['Program Studi'].str.split(' - ', n=1, expand=True)
        if len(split_data.columns) > 1:
            df['prodi'] = split_data[1]
        else:
             df['prodi'] = df['Program Studi'] # Fallback
else:
    print("'prodi' column found.")

if 'Provinsi rev' not in df.columns:
    print("Warning: 'Provinsi rev' column not found.")
    # Assuming user wants to test with available data, let's see if we can use 'Provinsi' if it exists or mock it
    if 'Provinsi' in df.columns:
         print("Using 'Provinsi' as 'Provinsi rev' for testing...")
         df['Provinsi rev'] = df['Provinsi']
    else:
        print("Mocking 'Provinsi rev' for testing...")
        df['Provinsi rev'] = 'Test Province'

# Test Function 1
print("\n1. Testing Location vs Tahun Lulus...")
try:
    res1 = tr.create_distribution_campus_loc_tahun(df)
    print("Result Shape:", res1.shape)
    print(res1.head())
except Exception as e:
    print(f"FAILED: {e}")

# Test Function 2
print("\n2. Testing Jurusan vs Tahun Lulus...")
try:
    res2 = tr.create_distribution_jurusan_tahun(df)
    print("Result Shape:", res2.shape)
    print(res2.head())
except Exception as e:
    print(f"FAILED: {e}")

# Test Function 3
print("\n3. Testing Prodi vs Tahun Lulus...")
try:
    res3 = tr.create_distribution_prodi_tahun(df)
    print("Result Shape:", res3.shape)
    print(res3.head())
except Exception as e:
    print(f"FAILED: {e}")

print("\n--- Verification Complete ---")
