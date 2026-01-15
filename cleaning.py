import pandas as pd
import numpy as np
import re

df = pd.read_excel('data.xlsx')
initial_rows = len(df)
print(f"Initial Row Count: {initial_rows}")


nim_col = "Nomor Induk Mahasiswa (NIM)"
name_col = "Nama Mahasiswa"
email_col = "Email Address"

# 1. Remove rows where "Nama Mahasiswa" is empty
print(f"Rows before removing empty names: {len(df)}")
df = df.dropna(subset=[name_col])
print(f"Rows after removing empty names: {len(df)}")

# 2. Handle duplicates
# Convert Timestamp to datetime for accurate sorting
# Debug: Check for invalid timestamps
df['Timestamp_parsed'] = pd.to_datetime(df['Timestamp'], errors='coerce')
invalid_timestamps = df[df['Timestamp_parsed'].isna()]
if not invalid_timestamps.empty:
    print(f"Found {len(invalid_timestamps)} invalid timestamps:")
    print(invalid_timestamps['Timestamp'].head())
    # decided to drop them or keep them? For now, let's just use the coerced column and maybe drop NaT if essential for sorting
    # If timestamp is NaT, we can't reliably sort "latest".
    # Let's assume we drop them or keep them at the end. 
    # For now, let's just proceed with coercion so the script runs.
    df['Timestamp'] = df['Timestamp_parsed']
else:
    df['Timestamp'] = pd.to_datetime(df['Timestamp'])

# Sort by Timestamp descending (latest first)
df = df.sort_values(by='Timestamp', ascending=False)

# Identify the exception NIM
exception_nim = 4202014111

# Split data
exception_rows = df[df[nim_col] == exception_nim]
other_rows = df[df[nim_col] != exception_nim]

# Deduplicate 'other_rows' keeping the first (latest)
print(f"Rows before deduplication (excluding exception): {len(other_rows)}")
other_rows_cleaned = other_rows.drop_duplicates(subset=[nim_col], keep='first')
print(f"Rows after deduplication: {len(other_rows_cleaned)}")

# Combine back
df_cleaned = pd.concat([other_rows_cleaned, exception_rows])

print(f"Final row count: {len(df_cleaned)}")
print(f"Exception NIM {exception_nim} count: {len(exception_rows)} (should be preserved)")

# Assign back to df for further processing
df = df_cleaned

# Debug: Clean column names to ensure 'Jurusan' is accessible
df.columns = df.columns.str.strip()

# Deduplicate columns (in case stripping caused duplicates)
df = df.loc[:, ~df.columns.duplicated()]

# Split "Program Studi" into "diploma" and "prodi"
# User requested adding 'diploma', and then grouping by Jurusan, prodi, diploma.
# We will generate 'diploma' and 'prodi' (cleaned name) while KEEPING 'Program Studi'.
prodi_col = "Program Studi"
print("Splitting 'Program Studi' into 'diploma' and 'prodi'...")
split_data = df[prodi_col].str.split(' - ', n=1, expand=True)
df['diploma'] = split_data[0]
df['prodi'] = split_data[1]

# Verify columns exist
print("Columns after split:", df.columns.tolist())

# Fix inconsistent data (Jurusan="Ilmu Kelautan dan Perikanan", prodi="Teknik Sipil")
# User requested to change Jurusan to "Teknik Sipil dan Perencanaan" instead of removing.
print("\n--- Fixing Inconsistent Data ---")
rows_before = len(df)

# Identify the rows
mask_inconsistent = (df['Jurusan'] == 'Ilmu Kelautan dan Perikanan') & (df['prodi'] == 'Teknik Sipil')
inconsistent_count = mask_inconsistent.sum()
print(f"Found {inconsistent_count} inconsistent rows to fix.")

if inconsistent_count > 0:
    # Update Jurusan
    df.loc[mask_inconsistent, 'Jurusan'] = 'Teknik Sipil dan Perencanaan'
    print("Inconsistent rows updated.")

rows_after = len(df)
print(f"Rows after fixing inconsistent data: {rows_after} (Should be same as before)")


with open('cleaning_report.txt', 'w') as f:
    f.write(f"Rows before removing inconsistent data: {rows_before}\n")
    f.write(f"Inconsistent rows removed: {inconsistent_count}\n")
    f.write(f"Rows after removing inconsistent data: {rows_after}\n")



# Create Group Table
print("\n--- Group Table (Jurusan, prodi, diploma) ---")
if 'Jurusan' in df.columns:
    group_table = df.groupby(['Jurusan', 'prodi', 'diploma']).size().reset_index(name='Count')
    print(group_table)
    
    # Optional: Save group table
    group_table.to_excel('group_table.xlsx', index=False)
    print("Group table saved to 'group_table.xlsx'")
else:
    print("CRITICAL: 'Jurusan' column still NOT FOUND.")
    # Debug print to help identify the issue if it persists
    print("Available columns:", df.columns.tolist())

# Remove columns as requested by user
cols_to_remove = [
    'Timestamp', 
    'Timestamp_parsed', # Also remove the parsed one if it exists
    'Email Address', 
    'Nama Mahasiswa', 
    'Nomor Handphone', 
    'NIK'
]
print(f"\n--- Removing Columns: {cols_to_remove} ---")
# Only drop columns that exist to avoid errors
cols_to_drop = [c for c in cols_to_remove if c in df.columns]
df = df.drop(columns=cols_to_drop)
print(f"Columns removed. Remaining columns: {len(df.columns)}")

# --- Status and Duration Analysis ---
print("\n--- Processing Status and Duration ---")
col_status = "Jelaskan status Anda saat ini?"
col_duration = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan)"

def clean_duration_value(val):
    if pd.isna(val) or str(val).strip() == '-':
        return np.nan
    
    val_str = str(val).lower().replace(',', '.')
    
    # Handle "seblum lulus" or "sebelum lulus" -> 0
    if "seblum lulus" in val_str or "sebelum lulus" in val_str:
        return 0
        
    # Handle "0.8 = 8 bulan" specific case because regex would catch 0.8
    if "0.8 = 8" in val_str:
        return 8
        
    multiplier = 1
    if "tahun" in val_str:
        multiplier = 12
        
    # Handle "1/2" or "setengah"
    if "1/2" in val_str or "setengah" in val_str:
        return 6 * multiplier
            
    # Regex to find numbers
    numbers = re.findall(r"[-+]?\d*\.\d+|\d+", val_str)
    if numbers:
        try:
            # Take the first number found
            num = float(numbers[0])
            return int(round(num * multiplier))
        except ValueError:
            return np.nan
            
    return np.nan

# Clean Duration Column if it exists (it might have been removed if user requested, but check list)
# Wait, did user remove it? The previous step removed "Timestamp, Email, Name, Phone, NIK".
# Duration is "Dalam berapa bulan...". Use checks.

if col_duration in df.columns:
    print(f"Cleaning column: {col_duration}")
    df[col_duration] = df[col_duration].apply(clean_duration_value)
    print("Duration column cleaned to numeric/NaN.")
else:
    print(f"WARNING: Duration column '{col_duration}' not found.")

# Create Status Table
if col_status in df.columns:
    print(f"Analyzing column: {col_status}")
    status_counts = df[col_status].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    print(status_counts)
    
    # Save Status Table
    status_counts.to_excel('status_table.xlsx', index=False)
    print("Status table saved to 'status_table.xlsx'")
else:
    print(f"WARNING: Status column '{col_status}' not found.")

# --- Mapping Kategori Perusahaan ---
print("\n--- Mapping Company Categories ---")

# --- Mapping Kategori Perusahaan ---
print("\n--- Mapping Company Categories ---")

# --- Mapping Kategori Perusahaan ---
print("\n--- Mapping Company Categories ---")

def mapping_rev_v2(teks):
    t = str(teks).strip().lower()
    
    # --- 1. KELOMPOK TIDAK BEKERJA / TRANSISI ---
    list_kosong = ['-', '.', '0', 'tidak ada', 'belum', 'blm', 'belm', 'mencari kerja', 'resign', 'irt']
    if any(x == t for x in ['-', '.', '0']) or any(x in t for x in ['belum', 'tidak ada', 'mencari kerja', 'resign', 'irt']):
        return 'Tidak Bekerja / Transisi'

    # --- 2. ORGANISASI MULTILATERAL ---
    if 'multilateral' in t:
        return 'Organisasi Multilateral'

    # --- 3. PENDIDIKAN (Termasuk Yayasan Sekolah) ---
    if any(x in t for x in ['guru', 'sekolah', 'universitas', 'pendidikan', 'bimbel', 'dosen', 'nurul ummah']):
        return 'Pendidikan'

    # --- 4. ORGANISASI NON-PROFIT / LSM ---
    if any(x in t for x in ['non-profit', 'lsm', 'yayasan pengembangan anak']):
        return 'Organisasi Non-Profit / LSM'

    # --- 5. PEMERINTAH (Instansi Murni) ---
    if any(x in t for x in ['pemerintah', 'badan gizi', 'dinas', 'kementerian']):
        return 'Pemerintah'

    # --- 6. BUMN / BUMD ---
    if any(x in t for x in ['bumn', 'bumd', 'pertamina', 'pln', 'bank bri']):
        return 'BUMN / BUMD'

    # --- 7. FREELANCE / KREATIF / GIG (Termasuk Usaha Online) ---
    if any(x in t for x in ['freelance', 'creator', 'desain', 'interior', 'shoppefood', 'online']):
        return 'Freelance / Kreatif'

    # --- 8. PERUSAHAAN SWASTA (Termasuk Bengkel, Salon, Toko, Ekspedisi) ---
    list_swasta_spesifik = ['bengkel', 'salon', 'tatto', 'toko baju', 'skincare', 'ekspedisi', 'klinik', 'kontraktor', 'telekomunikasi', 'hospitality', 'hiburan', 'pramuniaga', 'marketing', 'semua lini']
    list_swasta_umum = ['swasta', 'perusahaan', 'pt', 'cv', 'fnb', 'f&b', 'teknisi', 'ritel', 'kilang', 'smelter', 'tambang']
    
    if any(x in t for x in list_swasta_spesifik + list_swasta_umum):
        return 'Perusahaan Swasta'

    # --- 9. WIRASWASTA / MANDIRI (Sisa dari Mandiri) ---
    if any(x in t for x in ['wiraswasta', 'wirausaha', 'usaha sendiri', 'jualan', 'warung', 'founder', 'kerja sendiri', 'mandiri']):
        return 'Wiraswasta / Mandiri'

    # --- 10. INFORMAL / PERTANIAN ---
    if any(x in t for x in ['petani', 'kebun', 'sawit', 'buruh', 'babysitter']):
        return 'Informal / Pertanian'

    return 'Lainnya / Belum Terklasifikasi'


# 2. Identifikasi Nama Kolom
kolom_asal = "Apa jenis Perusahaan/Instansi/Institusi tempat Anda bekerja sekarang?"
kolom_rev = "Apa jenis Perusahaan/Instansi/Institusi tempat Anda bekerja sekarang? rev"

if kolom_asal in df.columns:
    # 3. Eksekusi Perubahan
    print("Applying V2 revised mapping function...")
    df[kolom_rev] = df[kolom_asal].apply(mapping_rev_v2)
    print(f"Created new column '{kolom_rev}' based on V2 mapping.")
    
    # Check for 'Lainnya / Belum Terklasifikasi' to see what's left
    unmapped = df[df[kolom_rev] == 'Lainnya / Belum Terklasifikasi']
    if not unmapped.empty:
        print(f"INFO: {len(unmapped)} rows classified as 'Lainnya'. Sample values:")
        print(unmapped[kolom_asal].unique()[:20])
else:
    print(f"WARNING: Column '{kolom_asal}' NOT FOUND. Skipping mapping.")

# Save to new file
output_file = 'cleaned_data.xlsx'
df.to_excel(output_file, index=False)
print(f"Cleaned data saved to {output_file}")

print("\n=== FINAL REPORT ===")
print(f"Initial Rows: {initial_rows}")
print(f"Final Rows:   {len(df)}")
print(f"Rows Removed: {initial_rows - len(df)}")

with open('cleaning_report.txt', 'w') as f:
    f.write(f"Initial Rows: {initial_rows}\n")
    f.write(f"Final Rows: {len(df)}\n")
    f.write(f"Total Rows Removed: {initial_rows - len(df)}\n")

