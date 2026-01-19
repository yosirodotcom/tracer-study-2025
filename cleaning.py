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



# Clean Duration Column if it exists (it might have been removed if user requested, but check list)
# Wait, did user remove it? The previous step removed "Timestamp, Email, Name, Phone, NIK".
# Duration is "Dalam berapa bulan...". Use checks.

# User requested to keep 'rev2' as is.
col_duration_rev2 = "Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2"
if col_duration_rev2 not in df.columns:
    print(f"WARNING: Duration column '{col_duration_rev2}' not found in source.")
    col_duration_rev2 = None
else:
    print(f"Using original data for: {col_duration_rev2}")

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
    
    # 1. Instansi Pemerintah
    if any(x in t for x in ['pemerintah', 'badan gizi', 'dinas', 'kementerian']):
        return 'Instansi Pemerintah'

    # 2. Organisasi non-profit/Lembaga Swadaya Masyarakat
    if any(x in t for x in ['non-profit', 'lsm', 'yayasan']):
        return 'Organisasi non-profit/Lembaga Swadaya Masyarakat'

    # 3. BUMN/BUMD
    if any(x in t for x in ['bumn', 'bumd', 'pertamina', 'pln', 'bank', 'bni', 'mandiri']) and not 'wiraswasta' in t and not 'usaha' in t:
         # Note: 'bank' could be private too, but commonly associated with BUMN in indonesia (BRI, Mandiri, BNI). 
         # However, user list has "Perusahaan Swasta". Let's try to be specific.
         # Actually, 'Bank BRI', 'Mandiri' are BUMN. 'BCA' is Swasta.
         # Let's keep the existing logic but change the return string.
         if any(x in t for x in ['bri', 'mandiri', 'bni', 'btn']):
             return 'BUMN/BUMD'
         if any(x in t for x in ['bumn', 'bumd', 'pertamina', 'pln']):
             return 'BUMN/BUMD'

    # 4. Institusi/Organisasi Multilateral
    if 'multilateral' in t:
        return 'Institusi/Organisasi Multilateral'

    # 5. Wiraswasta/perusahaan sendiri
    if any(x in t for x in ['wiraswasta', 'wirausaha', 'usaha sendiri', 'jualan', 'warung', 'founder', 'kerja sendiri', 'mandiri', 'freelance', 'creator', 'online']):
        # Merged 'Freelance' here or 'Perusahaan Swasta'? 
        # Usually Freelance is closer to self-employed (Wiraswasta).
        # User said: "else ... lainnya". 
        # But wait, "Wiraswasta/perusahaan sendiri" is in the list.
        # Let's map Freelance to Wiraswasta if acceptable, or Lainnya?
        # The user instruction was: `if data is not in this list ... else the it is what it is` (wait)
        # "I want the data is "lainnya" if data is not in this list ... else the it is what it is"
        # This implies: If the logic determines it belongs to one of the lists, use the list name.
        # If the updated logic (which I am writing now) can't verify it belongs to a list, it becomes 'lainnya'.
        return 'Wiraswasta/perusahaan sendiri'

    # 6. Perusahaan Swasta
    # Includes many things from before
    list_swasta = [
        'swasta', 'perusahaan', 'pt', 'cv', 'fnb', 'f&b', 'teknisi', 'ritel', 'kilang', 'smelter', 'tambang',
        'bengkel', 'salon', 'tatto', 'toko', 'skincare', 'ekspedisi', 'klinik', 'kontraktor', 'telekomunikasi', 
        'hospitality', 'hiburan', 'pramuniaga', 'marketing', 'konsultan', 'lawyer', 'notaris', 'apotek'
    ]
    if any(x in t for x in list_swasta):
        return 'Perusahaan Swasta'

    # Additional Cleanup for specific known entities that might fall through
    if 'bank' in t: # General bank usually private if not BUMN caught above
        return 'Perusahaan Swasta'

    if any(x in t for x in ['sekolah', 'universitas', 'guru', 'dosen', 'pendidikan']):
        # User defined list DOES NOT have Education/Pendidikan.
        # So it must be 'lainnya'?
        # "Instansi Pemerintah" might cover public schools?
        # "Perusahaan Swasta" might cover private schools?
        # Safe bet according to prompt: "lainnya" if not in list.
        # But maybe mapped to 'Instansi Pemerintah' if 'negeri'?
        if 'negeri' in t:
             return 'Instansi Pemerintah'
        if 'swasta' in t:
             return 'Perusahaan Swasta'
        return 'lainnya'

    return 'lainnya'


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



# --- Validation Column ---
# "valid column for Dalam berapa bulan Anda mendapatkan pekerjaan"
# 1 if (<=6 bulan == "Ya" AND duration > 6) OR (<=6 bulan == "Tidak" AND duration <= 6) else 0

col_valid_name = "valid column for Dalam berapa bulan Anda mendapatkan pekerjaan"
col_check_6bulan = "Apakah anda telah mendapatkan pekerjaan <=6 bulan / termasuk bekerja sebelum lulus?"

print("\n--- Adding Validation Column ---")
if col_duration_rev2 and col_check_6bulan in df.columns:
    def validation_logic(row):
        val_6bulan = str(row[col_check_6bulan])
        duration = row[col_duration_rev2] # This is numeric or NaN
        
        # Condition 1: Claimed <= 6 months (Ya) AND duration <= 6 (VALID)
        cond1 = (val_6bulan == "Ya") and (pd.notna(duration) and duration <= 6)
        
        # Condition 2: Claimed > 6 months (Tidak) AND (duration > 6 OR Empty) (VALID)
        cond2 = (val_6bulan == "Tidak") and ((pd.notna(duration) and duration > 6) or pd.isna(duration))
        
        if cond1 or cond2:
            return 1
        else:
            return 0

    df[col_valid_name] = df.apply(validation_logic, axis=1)
    print(f"Created '{col_valid_name}'. Found {df[col_valid_name].sum()} flagged rows.")
    
    # Debug sample
    flagged = df[df[col_valid_name] == 1]
    if not flagged.empty:
        print("Sample flagged rows (Check 6bulan vs Duration):")
        print(flagged[[col_check_6bulan, col_duration_rev2]].head())

else:
    print(f"WARNING: Could not create validation column. Missing columns: {col_check_6bulan if col_check_6bulan not in df.columns else ''} {col_duration_rev2 if not col_duration_rev2 else ''}")

# --- Clean Workplace Level Column ---
col_tingkat = "Apa tingkat tempat kerja Anda?"
col_tingkat_rev = "Apa tingkat tempat kerja Anda? rev"

print("\n--- Cleaning Workplace Level Column ---")
if col_tingkat in df.columns:
    # Remove leading numbers and space (e.g., "1 Lokal..." -> "Lokal...")
    df[col_tingkat_rev] = df[col_tingkat].str.replace(r'^\d+\s+', '', regex=True)
    print(f"Created '{col_tingkat_rev}'.")
else:
    print(f"WARNING: Column '{col_tingkat}' not found.")

# --- Transform Competency Columns ---
# User requested 1-5 to String mapping
competency_mapping = {
    1: "Tidak Menguasai",
    2: "Kurang Menguasai",
    3: "Menguasai",
    4: "Cukup Menguasai",
    5: "Sangat Menguasai"
}

def map_competency(val):
    try:
        if pd.isna(val): return val
        ival = int(float(val))
        return competency_mapping.get(ival, val)
    except:
        return val

# Identify columns using robust whitespace stripping
comp_cols_1 = ['Etika', 'Keahlian berdasarkan bidang ilmu', 'Bahasa Inggris', 'Penggunaan Teknologi Informasi', 'Komunikasi', 'Kerjasama Tim', 'Pengembangan']

actual_cols_to_map = []

# Map Set 1 (Suffix 1)
for target in comp_cols_1:
    found = False
    for col in df.columns:
        # Check patterns: "Name 1", "Name  1"
        clean_col = col.strip()
        if clean_col == target + " 1" or clean_col == target + "  1":
            actual_cols_to_map.append(col)
            found = True
            break
    if not found:
        print(f"DEBUG: Set 1 target '{target}' (Suffix 1) NOT found matches.")

# Map Set 2 (Suffix 2 or .1)
for target in comp_cols_1:
    found = False
    for col in df.columns:
        # Check patterns: "Name 2", "Name  2", "Name.1"
        clean_col = col.strip()
        if clean_col == target + " 2" or clean_col == target + ".1" or clean_col == target + "  2":
             actual_cols_to_map.append(col)
             found = True
    if not found:
        print(f"DEBUG: Set 2 target '{target}' NOT found matches.")

print(f"\n--- Transforming Competency Columns ({len(actual_cols_to_map)}) ---")
print(actual_cols_to_map)

for col in actual_cols_to_map:
    print(f"Mapping {col}...")
    # Debug first value
    first_val = df[col].iloc[0]
    print(f"  First val before: {first_val} type {type(first_val)}")
    df[col] = df[col].apply(map_competency)
    print(f"  First val after: {df[col].iloc[0]}")

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

