import pandas as pd
import numpy as np

df = pd.DataFrame({
    '2022': [56, 0, 0, 0, 56],
    '2023': [60, 0, 1, 0, 61],
    '2024': [635, 1, 0, 1, 637],
    'Total': [751, 1, 1, 1, 754],
    'Persentase': ['99.60%', '0.13%', '0.13%', '0.13%', '100.00%']
}, index=['Kampus Polnep', 'PDD Kapuas Hulu', 'PSDKU Sanggau', 'PSDKU Sukamara', 'Total'])

print("DF Index:", df.index)
print("DF Columns:", df.columns)
print("Numeric Columns:", df.select_dtypes(include=[np.number]).columns)

rows_to_drop = [i for i in df.index if 'total' in str(i).lower()]
print("Rows to drop:", rows_to_drop)

df_plot = df.drop(index=rows_to_drop)
print("DF Plot Index after drop:", df_plot.index)

val_col = 'Total' if 'Total' in df_plot.columns else df_plot.select_dtypes(include=[np.number]).columns[-1]
print("Val Col:", val_col)
