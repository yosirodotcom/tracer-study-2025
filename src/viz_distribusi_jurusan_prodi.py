import pandas as pd
from viz_utils import sort_crosstab_by_total

def create_distribution_jurusan_tahun(df):
    """
    Creates a distribution table of respondents based on Jurusan and Tahun Lulus.
    
    Args:
        df (pd.DataFrame): The input dataframe containing 'Jurusan' and 'Tahun Lulus' columns.
        
    Returns:
        pd.DataFrame: A cross-tabulation of Jurusan vs. Tahun Lulus.
    """
    jurusan_col = 'Jurusan'
    year_col = 'Tahun Lulus'
    
    if jurusan_col not in df.columns or year_col not in df.columns:
        raise ValueError(f"Missing columns: {jurusan_col} or {year_col} not found in DataFrame.")
    
    ct = pd.crosstab(df[jurusan_col], df[year_col], margins=True, margins_name='Total')
    return sort_crosstab_by_total(ct)


def create_distribution_prodi_tahun(df):
    """
    Creates a distribution table of respondents based on Program Studi (prodi) and Tahun Lulus.
    
    Args:
        df (pd.DataFrame): The input dataframe containing 'prodi' and 'Tahun Lulus' columns.
        
    Returns:
        pd.DataFrame: A cross-tabulation of Prodi vs. Tahun Lulus.
    """
    prodi_col = 'prodi'
    if prodi_col not in df.columns:
         if 'Program Studi' in df.columns:
             prodi_col = 'Program Studi'
         else:
            raise ValueError(f"Missing columns: {prodi_col} not found in DataFrame.")
            
    year_col = 'Tahun Lulus'
    
    if year_col not in df.columns:
        raise ValueError(f"Missing columns: {year_col} not found in DataFrame.")
    
    ct = pd.crosstab(df[prodi_col], df[year_col], margins=True, margins_name='Total')
    return sort_crosstab_by_total(ct)


