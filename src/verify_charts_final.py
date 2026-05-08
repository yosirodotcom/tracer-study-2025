import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
import base64
import os
import sys

# Add src to path to import table_jml_responden
sys.path.append(os.path.abspath('src'))
from table_jml_responden import get_bar_chart_base64, get_pie_chart_base64, create_distribution_prodi_tahun

# Load actual data
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_CLEANED = os.path.join(BASE_DIR, 'data', 'processed', 'cleaned_data.xlsx')
df_load = pd.read_excel(DATA_CLEANED)

# Generate prodi distribution
df_prodi = create_distribution_prodi_tahun(df_load)

# Generate charts
bar_b64 = get_bar_chart_base64(df_prodi, "Prodi Bar Chart", "prodi_bar")
pie_b64 = get_pie_chart_base64(df_prodi, "Prodi Pie Chart", "prodi_pie")

# Save to files
os.makedirs(os.path.join(BASE_DIR, 'assets', 'gambar'), exist_ok=True)

with open(os.path.join(BASE_DIR, 'assets', 'gambar', 'verify_prodi_bar_final.png'), 'wb') as f:
    f.write(base64.b64decode(bar_b64))

with open(os.path.join(BASE_DIR, 'assets', 'gambar', 'verify_prodi_pie_final.png'), 'wb') as f:
    f.write(base64.b64decode(pie_b64))

print(f"Final verification charts saved to {os.path.join(BASE_DIR, 'assets', 'gambar')}")
