import pandas as pd
import numpy as np
from viz_utils import sort_crosstab_by_total, get_pie_chart_base64, get_divergence_chart_base64, get_stacked_bar_chart_base64

def get_serapan_prodi_facet_pie_chart_base64(df):
    """
    Generates a stacked bar chart for graduate absorption across all Program Studi.
    Uses Empirical Bayes Shrinkage (EBS) for fair ranking, as recommended by the Council of Experts.
    This method derives its smoothing parameters directly from the university's data variance.
    """
    ct = create_serapan_prodi_ranked_table(df, apply_fair_sort=True)
    if ct.empty: return None
    
    # Robustly remove Total rows if they exist for charting
    if 'Total' in ct.index: ct = ct.drop('Total')
    # Filter for original status columns only (remove calculated metrics)
    plot_cols = [c for c in ct.columns if c not in ['Total', 'Persentase', 'Fair Score', 'Rank']]
    
    return get_stacked_bar_chart_base64(
        ct[plot_cols], 
        "Peringkat Serapan Lulusan antar Program Studi (Council Consensus: EBS)",
        "serapan_prodi_facet_pie",
        is_percentage=True,
        orientation='horizontal'
    )


def create_serapan_prodi_ranked_table(df, apply_fair_sort=True):
    """
    Creates a ranked table of Program Studi vs Status Pekerjaan.
    Includes fair ranking metrics (EBS) if requested.
    """
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_prodi not in df.columns or col_status not in df.columns:
        return pd.DataFrame()

    # Clean/Rename Schema
    df_clean = df.copy()
    status_map = {
        "Bekerja (Full time/Part time)": "Bekerja",
        "Wiraswasta": "Wiraswasta", 
        "Tidak kerja tetapi sedang mencari kerja": "Sedang Mencari Kerja",
        "Melanjutkan Pendidikan": "Studi Lanjut",
        "Belum memungkinkan bekerja": "Belum Memungkinkan Bekerja",
        "Tidak kerja tetapi tidak mencari kerja": "Tidak Mencari Kerja" 
    }
    df_clean[col_status] = df_clean[col_status].replace(status_map)
    
    # Create crosstab
    ct = pd.crosstab(df_clean[col_prodi], df_clean[col_status])
    
    if apply_fair_sort:
        # --- Council Recommendation: Empirical Bayes Shrinkage (EBS) ---
        working_cols = [c for c in ["Bekerja", "Wiraswasta"] if c in ct.columns]
        N = ct.sum(axis=1)
        X = ct[working_cols].sum(axis=1) if working_cols else pd.Series(0, index=ct.index)
        
        raw_rates = X / N
        mu = raw_rates.mean()
        sigma2 = raw_rates.var()
        
        if sigma2 > 0 and mu > 0 and mu < 1:
            m = max((mu * (1 - mu) / sigma2) - 1, 0)
        else:
            m = N.median()
            
        alpha = mu * m
        beta = (1 - mu) * m
        
        fair_scores = (X + alpha) / (N + alpha + beta)
        ct['Fair Score'] = fair_scores
        
        # Sort by Fair Score descending
        ct = ct.sort_values(by='Fair Score', ascending=False)
        ct['Rank'] = range(1, len(ct) + 1)
        
    # Reorder status columns
    preferred_order = ["Bekerja", "Wiraswasta", "Sedang Mencari Kerja", "Studi Lanjut", "Belum Memungkinkan Bekerja", "Tidak Mencari Kerja"]
    status_cols = [c for c in preferred_order if c in ct.columns]
    
    # Calculate Totals
    ct['Total'] = ct[status_cols].sum(axis=1)
    
    # Final column order
    # Rank and Fair Score are used for sorting but "hidden" (excluded) from final display columns
    final_cols = status_cols + ['Total']
    
    ct = ct[final_cols]
    
    # Apply row-relative percentage logic compatible with sort_crosstab_by_total
    # But since we have a custom sort, we'll manually add Persentase if needed 
    # or let generate_html_report handle it via apply_row_percentages_for_display
    
    return ct

def get_serapan_global_pie_chart_base64(df):
    """
    Generates a global pie chart for graduate absorption status.
    """
    col_status = 'Jelaskan status Anda saat ini?'
    if col_status not in df.columns:
        return None
        
    df_clean = df.copy()
    status_map = {
        "Bekerja (Full time/Part time)": "Bekerja",
        "Wiraswasta": "Wiraswasta", 
        "Tidak kerja tetapi sedang mencari kerja": "Sedang Mencari Kerja",
        "Melanjutkan Pendidikan": "Studi Lanjut",
        "Belum memungkinkan bekerja": "Belum Memungkinkan Bekerja",
        "Tidak kerja tetapi tidak mencari kerja": "Tidak Mencari Kerja" 
    }
    df_clean[col_status] = df_clean[col_status].replace(status_map)
    
    counts = df_clean[col_status].value_counts().to_frame()
    counts.columns = ['Total']
    
    return get_pie_chart_base64(counts, "Persentase Serapan Lulusan (Global)", "serapan_global_pie")


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


def get_serapan_divergence_chart_base64(df):
    """
    Generates a divergence chart for Program Studi: 
    (Bekerja + Wiraswasta) vs (Sedang Mencari Kerja).
    """
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_prodi not in df.columns or col_status not in df.columns:
        return None

    # Filter and Map (reset_index to avoid duplicate label issues in crosstab)
    df_plot = df[[col_prodi, col_status]].copy().reset_index(drop=True)
    
    # Calculate counts per Prodi and Status
    ct = pd.crosstab(df_plot[col_prodi], df_plot[col_status])
    
    # Group categories
    absorbed_cols = ["Bekerja (Full time/Part time)", "Wiraswasta"]
    searching_cols = ["Tidak kerja tetapi sedang mencari kerja"]
    
    # Ensure columns exist
    for col in absorbed_cols + searching_cols:
        if col not in ct.columns:
            ct[col] = 0
            
    # Calculate Counts
    ct['Total'] = ct.sum(axis=1)
    ct['Bekerja'] = ct["Bekerja (Full time/Part time)"]
    ct['Wiraswasta'] = ct["Wiraswasta"]
    ct['Absorbed'] = ct['Bekerja'] + ct['Wiraswasta']
    ct['Searching'] = ct[searching_cols].sum(axis=1)
    
    # User requested: Highlight 100% absorption (Absorbed > 2 and Searching == 0)
    ct['Is_100Pct'] = (ct['Absorbed'] > 2) & (ct['Searching'] == 0)
    
    # Sort by Absorbed count descending (largest at top)
    ct = ct.sort_values(by='Absorbed', ascending=True)
    
    labels = ct.index.tolist()
    # Add a marker or suffix to labels that are 100% for easier identification in the viz function
    final_labels = [f"(100%) {label}" if ct.loc[label, 'Is_100Pct'] else label for label in labels]
    
    # Pass Working side as a list of lists for stacking: [Bekerja, Wiraswasta]
    values_left = [ct['Bekerja'].tolist(), ct['Wiraswasta'].tolist()]
    values_right = ct['Searching'].tolist()
    
    return get_divergence_chart_base64(
        final_labels, values_left, values_right, 
        "Perbandingan Jumlah Serapan Lulusan per Program Studi",
        ["Bekerja", "Wiraswasta"], "Sedang Mencari Kerja",
        is_percentage=False
    )


def create_serapan_divergence_charts_per_jurusan(df):
    """
    Generates a dictionary of divergence charts, one for each Jurusan.
    Returns: dict { "Jurusan Name": base64_chart_string }
    """
    col_jurusan = 'Jurusan'
    if col_jurusan not in df.columns:
        return {}
    
    # Clean/Get unique Jurusans
    unique_jurusans = sorted([j for j in df[col_jurusan].unique() if pd.notna(j)])
    
    results = {}
    for jurusan in unique_jurusans:
        df_jur = df[df[col_jurusan] == jurusan]
        chart = get_serapan_divergence_chart_base64(df_jur)
        if chart:
            results[jurusan] = chart
            
    return results


