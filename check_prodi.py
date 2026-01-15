import pandas as pd

df = pd.read_excel('cleaned_data.xlsx')
print("Unique Program Studi values:")
for val in df['Program Studi'].unique():
    print(f"- {val}")
