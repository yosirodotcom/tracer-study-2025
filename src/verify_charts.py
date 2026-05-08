import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
import base64
import os
import sys

# Add src to path to import table_jml_responden
sys.path.append(os.path.abspath('src'))
from table_jml_responden import get_bar_chart_base64, get_pie_chart_base64

# Mock data
data = {
    'Category': ['A', 'B', 'C', 'D'],
    'Total': [10, 50, 30, 20]
}
df = pd.DataFrame(data).set_index('Category')

# Generate charts
bar_b64 = get_bar_chart_base64(df, "Test Bar", "test_bar")
pie_b64 = get_pie_chart_base64(df, "Test Pie", "test_pie")

# Save to files
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.makedirs(os.path.join(BASE_DIR, 'assets', 'gambar'), exist_ok=True)

with open(os.path.join(BASE_DIR, 'assets', 'gambar', 'test_bar_verify.png'), 'wb') as f:
    f.write(base64.b64decode(bar_b64))

with open(os.path.join(BASE_DIR, 'assets', 'gambar', 'test_pie_verify.png'), 'wb') as f:
    f.write(base64.b64decode(pie_b64))

print(f"Charts saved to {os.path.join(BASE_DIR, 'assets', 'gambar')}")
