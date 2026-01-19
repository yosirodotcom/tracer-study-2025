import pandas as pd
df = pd.read_excel('data.xlsx')
print("Columns in data.xlsx:")
for c in df.columns:
    print(f"- {c}")
