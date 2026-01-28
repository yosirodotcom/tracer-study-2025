import pandas as pd

df = pd.read_excel('cleaned_data.xlsx')
print("Verification of 'diploma' and 'prodi' columns:")
if 'diploma' in df.columns and 'prodi' in df.columns:
    print("Columns exist.")
    print(df[['Program Studi', 'diploma', 'prodi']].sample(5))
else:
    print("Columns MISSING.")
