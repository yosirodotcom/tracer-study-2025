import pandas as pd
from viz_utils import sort_crosstab_by_total

def create_distribution_masa_tunggu_status(df):
    """
    Creates a distribution table of Status Pekerjaan vs Kategori Masa Tunggu.
    """
    # 1. Seleksi dan Rename Kolom
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    
    # Check columns
    if col_status not in df.columns or col_masa_tunggu not in df.columns:
        # Try to be flexible if headers are slightly different or stripped
        print("Warning: Specific columns for Masa Tunggu not found exactly. Trying fuzzy match or skipping.")
        # Attempt to find closest match if needed, but for now rely on exact match as per request
        if col_status not in df.columns: return pd.DataFrame()
        if col_masa_tunggu not in df.columns: return pd.DataFrame()

    df_analysis = df[[col_status, col_masa_tunggu]].copy()
    df_analysis.columns = ['Status Pekerjaan', 'Masa_Tunggu_Bulan']
    
    # Save total respondents for denominator before filtering
    total_respondents = len(df)

    # 2. Filter Data Valid (Hanya yang bekerja/wiraswasta)
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    df_analysis = df_analysis[df_analysis['Status Pekerjaan'].isin(working_status)]
    
    df_analysis = df_analysis.dropna(subset=['Masa_Tunggu_Bulan'])
    df_analysis['Masa_Tunggu_Bulan'] = pd.to_numeric(df_analysis['Masa_Tunggu_Bulan'], errors='coerce')

    # 3. Fungsi Kategorisasi
    def kategorisasi_masa_tunggu(bulan):
        if pd.isna(bulan): return 'Unknown'
        if bulan < 3:
            return 'Kurang dari 3 Bulan'
        elif bulan <= 6:
            return '3 - 6 Bulan'
        elif bulan <= 12:
            return '6 - 12 Bulan'
        else:
            return 'Lebih dari 12 Bulan'

    df_analysis['Kategori_Masa_Tunggu'] = df_analysis['Masa_Tunggu_Bulan'].apply(kategorisasi_masa_tunggu)

    # 4. Membuat Pivot Table
    tabel_distribusi = pd.crosstab(
        df_analysis['Status Pekerjaan'],
        df_analysis['Kategori_Masa_Tunggu'],
        margins=True,
        margins_name='Total'
    )

    # 5. Mengurutkan kolom agar logis
    urutan_kolom = ['Kurang dari 3 Bulan', '3 - 6 Bulan', '6 - 12 Bulan', 'Lebih dari 12 Bulan', 'Total']
    col_ada = [c for c in urutan_kolom if c in tabel_distribusi.columns]
    
    tabel_final = tabel_distribusi[col_ada]
    
    # Apply sorting by Total desc (Rows) and add Percentage
    # Using total_respondents as denominator so percentages are relative to ALL alumni
    return sort_crosstab_by_total(tabel_final, denominator=total_respondents)


def create_distribution_waktu_tunggu_jurusan(df):
    """
    Creates a distribution table for Average Respondents Accepted Working within 6 months.
    """
    # 1. Preprocessing Data Masa Tunggu
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    
    # Check columns
    if col_masa_tunggu not in df.columns:
        print(f"Warning: Column '{col_masa_tunggu}' not found.")
        return pd.DataFrame()

    # 2. Filter: Hanya ambil responden yang mengisi masa tunggu DAN statusnya Bekerja/Wiraswasta
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    df_filtered = df.dropna(subset=[col_masa_tunggu]).copy()
    
    if col_status in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    else:
        print(f"Warning: Column '{col_status}' not found. Cannot filter by status.")
        return pd.DataFrame()

    # Pastikan data numerik
    df_filtered['Masa_Tunggu_Bulan'] = pd.to_numeric(df_filtered[col_masa_tunggu], errors='coerce')

    # 3. Logika Perhitungan (< 6 Bulan)
    # Buat kolom helper: 1 jika <= 6 bulan, 0 jika > 6 bulan
    df_filtered['Is_Less_6_Months'] = df_filtered['Masa_Tunggu_Bulan'].apply(lambda x: 1 if x <= 6 else 0)

    # 4. Membuat Tabel Agregat (Group by Jurusan)
    col_group = 'Jurusan'
    if col_group not in df_filtered.columns:
        return pd.DataFrame()
        
    analisis_masa_tunggu = df_filtered.groupby(col_group).agg(
        Jumlah_Responden_Bekerja=('Masa_Tunggu_Bulan', 'count'),
        Jumlah_Kurang_6_Bulan=('Is_Less_6_Months', 'sum'),
        Rata_rata_Waktu_Tunggu=('Masa_Tunggu_Bulan', 'mean')
    ).reset_index()

    # Calculate Total Alumni per Jurusan from unfiltered df
    total_per_jurusan = df.groupby(col_group).size().reset_index(name='Total_Responden_Alumni')
    analisis_masa_tunggu = analisis_masa_tunggu.merge(total_per_jurusan, on=col_group, how='left')

    # 5. Menghitung Persentase terhadap TOTAL Alumni
    analisis_masa_tunggu['Persentase_Kurang_6_Bulan'] = (
        analisis_masa_tunggu['Jumlah_Kurang_6_Bulan'] / analisis_masa_tunggu['Total_Responden_Alumni'] * 100
    )

    # Rounding rata-rata waktu tunggu
    analisis_masa_tunggu['Rata_rata_Waktu_Tunggu'] = analisis_masa_tunggu['Rata_rata_Waktu_Tunggu'].fillna(0).round(1)

    # Sort before renaming
    analisis_masa_tunggu = analisis_masa_tunggu.sort_values(by='Persentase_Kurang_6_Bulan', ascending=False)

    # 6. Formatting Tabel Akhir
    final_table = analisis_masa_tunggu[[
        'Jurusan', 
        'Total_Responden_Alumni', 
        'Jumlah_Responden_Bekerja',
        'Jumlah_Kurang_6_Bulan', 
        'Persentase_Kurang_6_Bulan', 
        'Rata_rata_Waktu_Tunggu'
    ]].copy()

    # Rename kolom untuk laporan
    final_table.columns = [
        'Jurusan', 
        'Total Responden (Alumni)',
        'Responden Bekerja', 
        'Jumlah Lulusan (<= 6 Bulan)', 
        'Persentase (<= 6 Bulan) (%)', 
        'Rata-rata Masa Tunggu (Bulan)'
    ]

    # Tambahkan Baris Total/Rata-rata Institusi
    total_alumni = len(df)
    responden_bekerja = final_table['Responden Bekerja'].sum()
    jumlah_kurang_6 = final_table['Jumlah Lulusan (<= 6 Bulan)'].sum()
    
    avg_masa_tunggu = df_filtered['Masa_Tunggu_Bulan'].mean() if not df_filtered.empty else 0

    total_row = pd.DataFrame({
        'Jurusan': ['TOTAL / RATA-RATA INSTITUSI'],
        'Total Responden (Alumni)': [total_alumni],
        'Responden Bekerja': [responden_bekerja],
        'Jumlah Lulusan (<= 6 Bulan)': [jumlah_kurang_6],
        'Persentase (<= 6 Bulan) (%)': [
            (jumlah_kurang_6 / total_alumni * 100) if total_alumni > 0 else 0
        ],
        'Rata-rata Masa Tunggu (Bulan)': [
            round(avg_masa_tunggu, 1)
        ]
    })
    
    final_table = pd.concat([final_table, total_row], ignore_index=True)
    
    # Format Percentage Column string
    final_table['Persentase (<= 6 Bulan) (%)'] = final_table['Persentase (<= 6 Bulan) (%)'].apply(lambda x: f"{x:.2f}%")
    
    return final_table


def create_waktu_tunggu_prodi_per_jurusan(df):
    """
    Creates a dictionary of tables, one per Jurusan.
    Each table shows [Prodi] vs Average Waiting Time analysis.
    Returns: dict { "Jurusan Name": pd.DataFrame }
    """
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
        
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_jurusan not in df.columns or col_prodi not in df.columns or col_masa_tunggu not in df.columns or col_status not in df.columns:
        return {}

    # 1. Filter Data
    df_filtered = df.dropna(subset=[col_masa_tunggu]).copy()
    df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    
    if df_filtered.empty:
        return {}

    # 2. Prepare numeric data
    df_filtered['Masa_Tunggu_Bulan'] = pd.to_numeric(df_filtered[col_masa_tunggu], errors='coerce')
    df_filtered['Is_Less_6_Months'] = df_filtered['Masa_Tunggu_Bulan'].apply(lambda x: 1 if x <= 6 else 0)

    # 3. Aggregate by [Jurusan, Prodi]
    analisis = df_filtered.groupby([col_jurusan, col_prodi]).agg(
        Jumlah_Responden_Bekerja=('Masa_Tunggu_Bulan', 'count'),
        Jumlah_Kurang_6_Bulan=('Is_Less_6_Months', 'sum'),
        Rata_rata_Waktu_Tunggu=('Masa_Tunggu_Bulan', 'mean')
    ).reset_index()

    # Calculate Total Alumni per Prodi from unfiltered df
    total_per_prodi = df.groupby([col_jurusan, col_prodi]).size().reset_index(name='Total_Responden_Alumni')
    analisis = analisis.merge(total_per_prodi, on=[col_jurusan, col_prodi], how='right')
    
    # Fill NaN for prodis that might have 0 working respondents
    analisis['Jumlah_Responden_Bekerja'] = analisis['Jumlah_Responden_Bekerja'].fillna(0)
    analisis['Jumlah_Kurang_6_Bulan'] = analisis['Jumlah_Kurang_6_Bulan'].fillna(0)
    analisis['Rata_rata_Waktu_Tunggu'] = analisis['Rata_rata_Waktu_Tunggu'].fillna(0)

    # 4. Split by Jurusan
    unique_jurusans = analisis[col_jurusan].unique()
    results = {}

    for jurusan in unique_jurusans:
        sub_df = analisis[analisis[col_jurusan] == jurusan].copy()
        
        # Calculate Percentage relative to TOTAL Alumni
        sub_df['Persentase (<= 6 Bulan) (%)'] = (sub_df['Jumlah_Kurang_6_Bulan'] / sub_df['Total_Responden_Alumni'] * 100)
        
        # Rounding
        sub_df['Rata_rata_Waktu_Tunggu'] = sub_df['Rata_rata_Waktu_Tunggu'].round(1)
        
        # Sort by Percentage Desc
        sub_df = sub_df.sort_values(by='Persentase (<= 6 Bulan) (%)', ascending=False)

        # Formatting percentage string
        sub_df['Persentase (<= 6 Bulan) (%)'] = sub_df['Persentase (<= 6 Bulan) (%)'].apply(lambda x: f"{x:.2f}%")

        # Select and Rename columns
        final_table = sub_df[[
            col_prodi, 
            'Total_Responden_Alumni',
            'Jumlah_Responden_Bekerja',
            'Jumlah_Kurang_6_Bulan', 
            'Persentase (<= 6 Bulan) (%)', 
            'Rata_rata_Waktu_Tunggu'
        ]].copy()
        
        final_table.columns = [
            'Program Studi', 
            'Total Responden (Alumni)',
            'Responden Bekerja',
            'Jumlah Lulusan (<= 6 Bulan)', 
            'Persentase (<= 6 Bulan) (%)', 
            'Rata-rata Masa Tunggu (Bulan)'
        ]

        # Add Total row for this Jurusan
        total_alumni_jur = final_table['Total Responden (Alumni)'].sum()
        total_bekerja_jur = final_table['Responden Bekerja'].sum()
        total_k6_jur = final_table['Jumlah Lulusan (<= 6 Bulan)'].sum()
        
        # Mean for this specific Jurusan
        jur_avg = df_filtered[df_filtered[col_jurusan] == jurusan]['Masa_Tunggu_Bulan'].mean()

        total_row = pd.DataFrame({
            'Program Studi': [f'TOTAL {jurusan}'],
            'Total Responden (Alumni)': [total_alumni_jur],
            'Responden Bekerja': [total_bekerja_jur],
            'Jumlah Lulusan (<= 6 Bulan)': [total_k6_jur],
            'Persentase (<= 6 Bulan) (%)': [f"{(total_k6_jur / total_alumni_jur * 100):.2f}%" if total_alumni_jur > 0 else "0.00%"],
            'Rata-rata Masa Tunggu (Bulan)': [round(jur_avg, 1)]
        })
        
        final_table = pd.concat([final_table, total_row], ignore_index=True)
        
        results[f"Rata-rata Waktu Tunggu - {jurusan}"] = final_table
        
    return results


def create_masa_tunggu_prodi_per_jurusan(df):
    """
    Creates a dictionary of tables, one per Jurusan.
    Each table shows [Prodi] vs Kategori Masa Tunggu.
    Returns: dict { "Jurusan Name": pd.DataFrame }
    """
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
        
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    
    if col_jurusan not in df.columns or col_prodi not in df.columns or col_masa_tunggu not in df.columns:
        return {}

    # 1. Filter and Prepare
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_status not in df.columns:
        return {}
        
    # Filter only working/wiraswasta
    df_filtered = df[df[col_status].isin(working_status)].copy()
    
    # Filter only those who filled in the Masa Tunggu column
    df_filtered = df_filtered.dropna(subset=[col_masa_tunggu])
    
    df_analysis = df_filtered[[col_jurusan, col_prodi, col_masa_tunggu]].copy()
    df_analysis['Masa_Tunggu_Bulan'] = pd.to_numeric(df_analysis[col_masa_tunggu], errors='coerce')
    
    def kategorisasi_masa_tunggu(bulan):
        if pd.isna(bulan): return 'Unknown'
        if bulan < 3: return 'Kurang dari 3 Bulan'
        elif bulan <= 6: return '3 - 6 Bulan'
        elif bulan <= 12: return '6 - 12 Bulan'
        else: return 'Lebih dari 12 Bulan'

    df_analysis['Kategori_Masa_Tunggu'] = df_analysis['Masa_Tunggu_Bulan'].apply(kategorisasi_masa_tunggu)

    # 2. Create Crosstab
    ct = pd.crosstab(
        [df_analysis[col_jurusan], df_analysis[col_prodi]],
        df_analysis['Kategori_Masa_Tunggu']
    )
    
    # Sort columns logically
    urutan_kolom = ['Kurang dari 3 Bulan', '3 - 6 Bulan', '6 - 12 Bulan', 'Lebih dari 12 Bulan']
    sorted_cols = [c for c in urutan_kolom if c in ct.columns]
    
    # Get unique Jurusans
    unique_jurusans = ct.index.get_level_values(0).unique()
    
    results = {}
    for jurusan in unique_jurusans:
        try:
            # sub_df index is Prodi
            sub_df = ct.loc[jurusan].copy()
        except KeyError:
            continue
            
        # Ensure all columns exist for consistency
        for col in sorted_cols:
            if col not in sub_df.columns:
                sub_df[col] = 0
        
        # Reorder columns
        sub_df = sub_df[sorted_cols]
        
        # Calculate Total per row
        sub_df['Total'] = sub_df.sum(axis=1)
        
        # Calculate Grand Total for this Jurusan
        jurusan_total_row = sub_df.sum(axis=0)
        jurusan_grand_total = jurusan_total_row['Total']
        
        # Prepare Rows
        final_rows = []
        
        # Add Prodi Rows
        for prodi, row_data in sub_df.iterrows():
            row_dict = {'Program Studi': prodi}
            for col in sorted_cols + ['Total']:
                row_dict[col] = row_data[col]
            
            # Percentage based on Jurusan Total
            if jurusan_grand_total > 0:
                pct = (row_data['Total'] / jurusan_grand_total) * 100
                row_dict['Persentase'] = f"{pct:.2f}%"
            else:
                row_dict['Persentase'] = "0.00%"
                
            final_rows.append(row_dict)
            
        # Add Total Row
        total_dict = {'Program Studi': f'Total {jurusan}'}
        for col in sorted_cols + ['Total']:
            total_dict[col] = jurusan_total_row[col]
        
        total_dict['Persentase'] = "100.00%"
        final_rows.append(total_dict)
        
        # Create DataFrame
        df_jurusan = pd.DataFrame(final_rows)
        # Reorder columns
        final_cols_order = ['Program Studi'] + sorted_cols + ['Total', 'Persentase']
        df_jurusan = df_jurusan[final_cols_order]
        
        results[f"Masa Tunggu - {jurusan}"] = df_jurusan
        
    return results


