import pandas as pd
from viz_utils import (
    sort_crosstab_by_total, get_salary_bell_curve_base64,
    get_salary_jurusan_lollipop_chart_base64,
    get_salary_stacked_bar_chart_base64
)

def get_salary_distribution_bell_curve(df):
    """
    Wrapper for bell curve visualization of salary distribution.
    """
    return get_salary_bell_curve_base64(df, "Kurva Distribusi Normal Pendapatan Lulusan", "gaji_bell_curve")

def get_salary_jurusan_lollipop_chart(df):
    """
    Wrapper for lollipop chart visualization of salary by jurusan.
    """
    return get_salary_jurusan_lollipop_chart_base64(df, "Rata-rata Gaji Lulusan per Jurusan", "gaji_jurusan_lollipop")

def get_salary_distribution_by_prodi_chart(df):
    """
    Wrapper for horizontal stacked bar chart of salary distribution by Prodi.
    """
    ct = create_salary_distribution_by_prodi(df)
    if ct.empty:
        return None
    return get_salary_stacked_bar_chart_base64(ct, "Distribusi Pendapatan Lulusan per Program Studi")

def create_salary_distribution_by_prodi(df):
    """
    Creates a crosstab of Prodi vs Salary Range for working respondents.
    """
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_salary not in df.columns or col_prodi not in df.columns:
        return pd.DataFrame()
        
    df_filtered = df.copy()
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    
    df_filtered = df_filtered.dropna(subset=[col_salary, col_prodi])
    
    if df_filtered.empty:
        return pd.DataFrame()

    # Custom sort order for salary categories
    order = [
        '< Rp. 1.000.000',
        'Rp. 1.000.001 - Rp. 2.000.000',
        'Rp. 2.000.001 - Rp. 3.000.000',
        'Rp. 3.000.001 - Rp. 4.000.000',
        'Rp. 4.000.001 - Rp. 5.000.000',
        'Rp. 5.000.001 - Rp. 6.000.000',
        'Rp. 6.000.001 - Rp. 7.000.000',
        'Rp. 7.000.001 - Rp. 8.000.000',
        '> Rp. 8.000.001'
    ]
    
    # Create Crosstab
    ct = pd.crosstab(df_filtered[col_prodi], df_filtered[col_salary])
    
    # Reorder columns according to salary scale
    available_cols = [o for o in order if o in ct.columns]
    ct = ct[available_cols]
    
    # User Defined Mappings for Mean Calculation
    salary_map = {
        '< Rp. 1.000.000': 500000,
        'Rp. 1.000.001 - Rp. 2.000.000': 1500000,
        'Rp. 2.000.001 - Rp. 3.000.000': 2500000,
        'Rp. 3.000.001 - Rp. 4.000.000': 3500000,
        'Rp. 4.000.001 - Rp. 5.000.000': 4500000,
        'Rp. 5.000.001 - Rp. 6.000.000': 5500000,
        'Rp. 6.000.001 - Rp. 7.000.000': 6500000,
        'Rp. 7.000.001 - Rp. 8.000.000': 7500000,
        '> Rp. 8.000.001': 8500000
    }

    # Calculate weighted mean for sorting
    # Get values for the columns we actually have in the crosstab
    weights = [salary_map.get(col, 0) for col in ct.columns]
    
    # Calculate sum(count * value) / sum(count) for each row
    ct['Total'] = ct.sum(axis=1)
    weighted_sum = (ct[ct.columns[:-1]] * weights).sum(axis=1)
    ct['Mean_Salary'] = weighted_sum / ct['Total']
    
    # Sort by Mean_Salary descending
    ct = ct.sort_values('Mean_Salary', ascending=False)
    
    # Format Mean Salary as string
    ct['Mean_Text'] = ct['Mean_Salary'].apply(lambda x: f"Rp{x/1000000:.1f} Juta" if pd.notna(x) else "")
    
    # Remove the numeric Mean_Salary column as it's not needed for the chart segments
    ct = ct.drop(columns=['Mean_Salary'])
    
    return ct

def create_salary_distribution(df):
    """
    Creates a distribution table for Salary/Income of working respondents.
    """
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_salary not in df.columns:
        return pd.DataFrame()
        
    df_filtered = df.copy()
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
        
    salary_counts = df_filtered[col_salary].value_counts().reset_index()
    salary_counts.columns = ['Rata-rata Pendapatan', 'Jumlah Responden']
    
    # Custom sort order for salary categories
    order = [
        '< Rp. 1.000.000',
        'Rp. 1.000.001 - Rp. 2.000.000',
        'Rp. 2.000.001 - Rp. 3.000.000',
        'Rp. 3.000.001 - Rp. 4.000.000',
        'Rp. 4.000.001 - Rp. 5.000.000',
        'Rp. 5.000.001 - Rp. 6.000.000',
        'Rp. 6.000.001 - Rp. 7.000.000',
        'Rp. 7.000.001 - Rp. 8.000.000',
        '> Rp. 8.000.001'
    ]
    
    salary_counts['Rata-rata Pendapatan'] = pd.Categorical(salary_counts['Rata-rata Pendapatan'], categories=order, ordered=True)
    salary_counts = salary_counts.sort_values('Rata-rata Pendapatan').reset_index(drop=True)
    
    # Calculate Percentage
    total = salary_counts['Jumlah Responden'].sum()
    if total > 0:
        salary_counts['Persentase (%)'] = (salary_counts['Jumlah Responden'] / total * 100).map('{:.2f}%'.format)
    else:
        salary_counts['Persentase (%)'] = '0.00%'
        
    # Add Total Row
    total_row = pd.DataFrame({
        'Rata-rata Pendapatan': ['Total'], 
        'Jumlah Responden': [total], 
        'Persentase (%)': ['100.00%']
    })
    
    final_table = pd.concat([salary_counts, total_row], ignore_index=True)
    
    return final_table


def create_salary_by_jurusan(df):
    """
    Calculates the Average salary per Jurusan using custom range conversions.
    """
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    col_jurusan = 'Jurusan'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_salary not in df.columns or col_jurusan not in df.columns:
        return pd.DataFrame(), pd.DataFrame()
        
    df_filtered = df.copy()
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    
    # Drop rows where salary is missing
    df_filtered = df_filtered.dropna(subset=[col_salary])
    
    if df_filtered.empty:
        return pd.DataFrame(), pd.DataFrame()

    # User Defined Mappings for Mean Calculation
    salary_map = {
        '< Rp. 1.000.000': 500000,
        'Rp. 1.000.001 - Rp. 2.000.000': 1500000,
        'Rp. 2.000.001 - Rp. 3.000.000': 2500000,
        'Rp. 3.000.001 - Rp. 4.000.000': 3500000,
        'Rp. 4.000.001 - Rp. 5.000.000': 4500000,
        'Rp. 5.000.001 - Rp. 6.000.000': 5500000,
        'Rp. 6.000.001 - Rp. 7.000.000': 6500000,
        'Rp. 7.000.001 - Rp. 8.000.000': 7500000,
        '> Rp. 8.000.001': 8500000
    }

    # Map categories to numeric
    df_filtered['salary_num'] = df_filtered[col_salary].map(salary_map)
    
    # Calculate Mean per Jurusan
    salary_by_jurusan = df_filtered.groupby(col_jurusan)['salary_num'].mean().reset_index()
    salary_by_jurusan.columns = ['Jurusan', 'Rata-rata Gaji (Estimasi)']
    
    # Sort by numeric mean descending
    salary_by_jurusan = salary_by_jurusan.sort_values(by='Rata-rata Gaji (Estimasi)', ascending=False)
    
    # Store numeric version for chart
    df_chart = salary_by_jurusan.copy()
    df_chart = df_chart.rename(columns={'Rata-rata Gaji (Estimasi)': 'Total'})
    
    # Format for display table: Compact format (e.g., Rp3.8 juta)
    salary_by_jurusan['Rata-rata Gaji (Estimasi)'] = salary_by_jurusan['Rata-rata Gaji (Estimasi)'].apply(
        lambda x: f"Rp{x/1000000:.1f} juta"
    )
    
    return salary_by_jurusan, df_chart


def create_salary_prodi_per_jurusan(df):
    """
    Creates a dictionary of tables, one per Jurusan.
    Each table shows [Prodi] vs Average Salary estimation.
    Returns: dict { "Jurusan Name": pd.DataFrame }
    """
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_salary not in df.columns or col_jurusan not in df.columns or col_prodi not in df.columns:
        return {}
        
    df_filtered = df.copy()
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    
    df_filtered = df_filtered.dropna(subset=[col_salary])
    
    if df_filtered.empty:
        return {}

    salary_map = {
        '< Rp. 1.000.000': 500000,
        'Rp. 1.000.001 - Rp. 2.000.000': 1500000,
        'Rp. 2.000.001 - Rp. 3.000.000': 2500000,
        'Rp. 3.000.001 - Rp. 4.000.000': 3500000,
        'Rp. 4.000.001 - Rp. 5.000.000': 4500000,
        'Rp. 5.000.001 - Rp. 6.000.000': 5500000,
        'Rp. 6.000.001 - Rp. 7.000.000': 6500000,
        'Rp. 7.000.001 - Rp. 8.000.000': 7500000,
        '> Rp. 8.000.001': 8500000
    }

    df_filtered['salary_num'] = df_filtered[col_salary].map(salary_map)
    
    # Calculate Mean per [Jurusan, Prodi]
    stats = df_filtered.groupby([col_jurusan, col_prodi])['salary_num'].agg(['mean', 'count']).reset_index()
    
    unique_jurusans = stats[col_jurusan].unique()
    results = {}
    
    for jurusan in unique_jurusans:
        sub_df = stats[stats[col_jurusan] == jurusan].copy()
        sub_df = sub_df.sort_values(by='mean', ascending=False)
        
        # Format for display
        sub_df['Rata-rata Gaji (Estimasi)'] = sub_df['mean'].apply(
            lambda x: f"Rp{x/1000000:.1f} juta" if pd.notnull(x) else "Rp0.0 juta"
        )
        
        final_table = sub_df[[col_prodi, 'count', 'Rata-rata Gaji (Estimasi)']].copy()
        final_table.columns = ['Program Studi', 'Jumlah Responden', 'Rata-rata Gaji (Estimasi)']
        
        # Add Total Row for Jurusan
        jur_total_data = df_filtered[df_filtered[col_jurusan] == jurusan]
        jur_mean = jur_total_data['salary_num'].mean()
        jur_count = jur_total_data['salary_num'].count()
        
        total_row = pd.DataFrame({
            'Program Studi': [f'TOTAL {jurusan}'],
            'Jumlah Responden': [jur_count],
            'Rata-rata Gaji (Estimasi)': [f"Rp{jur_mean/1000000:.1f} juta" if pd.notnull(jur_mean) else "Rp0.0 juta"]
        })
        
        final_table = pd.concat([final_table, total_row], ignore_index=True)
        results[f"Rata-rata Gaji - {jurusan}"] = final_table
        
    return results


