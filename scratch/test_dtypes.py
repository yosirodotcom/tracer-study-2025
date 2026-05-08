import pandas as pd
import numpy as np

# Simulate create_distribution_campus_loc_tahun output
df = pd.DataFrame({
    '2022': [56, 0, 0, 0, 56],
    '2023': [60, 0, 1, 0, 61],
    '2024': [635, 1, 0, 1, 637],
    'Total': [751, 1, 1, 1, 754],
    'Persentase': ['99.60%', '0.13%', '0.13%', '0.13%', '100.00%']
}, index=['Kampus Polnep', 'PDD Kapuas Hulu', 'PSDKU Sanggau', 'PSDKU Sukamara', 'Total'])

print("Dtypes:\n", df.dtypes)
print("Numeric columns identified by select_dtypes:\n", df.select_dtypes(include=[np.number]).columns)
