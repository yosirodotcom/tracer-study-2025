import pandas as pd
df = pd.read_excel('cleaned_data.xlsx')

targets = ['Etika', 'Bahasa Inggris', 'Kerjasama Tim', 'Penggunaan Teknologi Informasi', 'Komunikasi', 'Pengembangan', 'Keahlian berdasarkan bidang ilmu']

print("Targets vs Columns:")
for t in targets:
    print(f"Target: '{t}'")
    for c in df.columns:
        if t in c:
            print(f"  Match candidate: '{c}' (Len: {len(c)})")
            if c.strip() == t:
                print("    STRIP MATCH!")
            else:
                print(f"    No strip match. '{c.strip()}' vs '{t}'")
