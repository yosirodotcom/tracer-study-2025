import pandas as pd
from viz_utils import sort_crosstab_by_total

def create_serapan_jurusan(df):
    """
    Creates a crosstab of Jurusan vs Status Pekerjaan.
    """
    col_jurusan = 'Jurusan'
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_jurusan not in df.columns or col_status not in df.columns:
        return pd.DataFrame()
        
    tabel_jurusan = pd.crosstab(
        df[col_jurusan], 
        df[col_status], 
        margins=True, 
        margins_name='Total'
    )
    # Reuse our sorter if possible, it sorts by 'Total' column desc
    # This matches the user's general preference for sorting
    return sort_crosstab_by_total(tabel_jurusan)


def create_serapan_prodi_per_jurusan(df):
    """
    Creates a dictionary of tables, one per Jurusan.
    Each table shows [Prodi] vs Status Pekerjaan.
    Returns: dict { "Jurusan Name": pd.DataFrame }
    """
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
        
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_jurusan not in df.columns or col_prodi not in df.columns or col_status not in df.columns:
        return {}

    # Clean/Rename Schema for Display
    df_clean = df.copy()
    status_map = {
        "Bekerja (Full time/Part time)": "Bekerja",
        "Wiraswasta": "Wiraswasta", 
        "Tidak kerja tetapi sedang mencari kerja": "Sedang Mencari Kerja",
        "Melanjutkan Pendidikan": "Studi Lanjut",
        "Belum memungkinkan bekerja": "Belum Memungkinkan Bekerja",
        "Tidak kerja tetapi tidak mencari kerja": "Tidak Mencari Kerja" 
    }
    # Apply mapping
    df_clean[col_status] = df_clean[col_status].replace(status_map)

    # 1. Create Crosstab
    ct = pd.crosstab(
        [df_clean[col_jurusan], df_clean[col_prodi]],
        df_clean[col_status]
    )
    
    # Prepare Columns Order
    preferred_order = [
         "Bekerja", "Wiraswasta", "Sedang Mencari Kerja", "Studi Lanjut", "Belum Memungkinkan Bekerja"
    ]
    existing_cols = ct.columns.tolist()
    sorted_cols = []
    
    for col in preferred_order:
        if col in existing_cols:
            sorted_cols.append(col)
    
    for col in existing_cols:
        if col not in sorted_cols:
            sorted_cols.append(col)
            
    # Add Total column to list (calculated later)
    columns = sorted_cols + ['Total']
    
    # Get unique Jurusans
    unique_jurusans = ct.index.get_level_values(0).unique()
    
    results = {}
    
    for jurusan in unique_jurusans:
        try:
            # sub_df index is Prodi
            sub_df = ct.loc[jurusan].copy()
        except KeyError:
            continue
            
        # Ensure all columns exist
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
            for col in columns:
                val = row_data[col]
                row_dict[col] = val
            
            # Percentage based on Jurusan Total
            if jurusan_grand_total > 0:
                pct = (row_data['Total'] / jurusan_grand_total) * 100
                row_dict['Persentase'] = f"{pct:.2f}%"
            else:
                row_dict['Persentase'] = "0.00%"
                
            final_rows.append(row_dict)
            
        # Add Total Row
        total_dict = {'Program Studi': f'Total {jurusan}'}
        for col in columns:
            val = jurusan_total_row[col]
            total_dict[col] = val
        
        # Percentage for Total Row is 100%
        total_dict['Persentase'] = "100.00%"
        
        final_rows.append(total_dict)
        
        # Create DataFrame
        df_jurusan = pd.DataFrame(final_rows)
        # Reorder columns
        final_cols_order = ['Program Studi'] + columns + ['Persentase']
        df_jurusan = df_jurusan[final_cols_order]
        
        results[jurusan] = df_jurusan
        
    return results


