import pandas as pd
from viz_utils import sort_crosstab_by_total

def create_distribution_campus_loc_tahun(df):
    """
    Creates a distribution table of respondents based on Lokasi Kampus (derived from prodi) and Tahun Lulus.
    
    Logic:
    - If 'prodi' contains "Kapuas Hulu" -> "Kapuas Hulu"
    - If 'prodi' contains "Sanggau" -> "Sanggau"
    - If 'prodi' contains "Sukamara" -> "Sukamara"
    - Else -> "Kampus Polnep"
    
    Args:
        df (pd.DataFrame): The input dataframe containing 'prodi' and 'Tahun Lulus' columns.
        
    Returns:
        pd.DataFrame: A cross-tabulation of Location vs. Tahun Lulus.
    """
    column_to_check = 'prodi'
    if column_to_check not in df.columns:
        # Fallback to 'Program Studi' if 'prodi' missing
        if 'Program Studi' in df.columns:
            column_to_check = 'Program Studi'
        else:
             raise ValueError(f"Missing columns: 'prodi' or 'Program Studi' not found.")

    year_col = 'Tahun Lulus'
    if year_col not in df.columns:
         raise ValueError(f"Missing column: {year_col}")

    def get_location(val):
        s_val = str(val)
        if 'Kapuas Hulu' in s_val:
            return 'PDD Kapuas Hulu'
        if 'Sanggau' in s_val:
             return 'PSDKU Sanggau'
        if 'Sukamara' in s_val:
            return 'PSDKU Sukamara'
        return 'Kampus Polnep'

    # Create a temporary column for location to group by
    temp_loc_col = 'Lokasi Kampus'
    df = df.copy() # Avoid SettingWithCopyWarning on original df if passed directly
    df[temp_loc_col] = df[column_to_check].apply(get_location)
    
    ct = pd.crosstab(df[temp_loc_col], df[year_col], margins=True, margins_name='Total')
    return sort_crosstab_by_total(ct)


