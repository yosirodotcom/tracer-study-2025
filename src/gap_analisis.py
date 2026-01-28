import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from math import pi
import io
import base64
import os
import sys
import webbrowser

# Setup Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data', 'processed')
REPORTS_DIR = os.path.join(BASE_DIR, 'reports')
ASSETS_DIR = os.path.join(BASE_DIR, 'assets', 'gambar')
DATA_FILE = os.path.join(DATA_DIR, 'cleaned_data.xlsx')

# Ensure directories exist
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)

# Competency Mappings
COMPETENCY_MAP = {
    'Etika Profesional': ['Etika'],
    'Keahlian Bidang Ilmu': ['Keahlian', 'bidang ilmu'],
    'Bahasa Inggris': ['Inggris', 'English'],
    'Teknologi Informasi': ['Teknologi', 'Informasi'],
    'Komunikasi Efektif': ['Komunikasi'],
    'Kerja Sama Tim': ['Kerja sama', 'Kerjasama', 'Tim'],
    'Pengembangan Diri': ['Pengembangan']
}

def load_data():
    """Loads the cleaned data."""
    if not os.path.exists(DATA_FILE):
        # Fallback to source/raw if needed, but per project structure use cleaned
        raw_path = os.path.join(BASE_DIR, 'data', 'raw', 'data.xlsx')
        if os.path.exists(raw_path):
            print(f"Loading raw data from {raw_path}")
            return pd.read_excel(raw_path)
        else:
            raise FileNotFoundError(f"Data file not found at {DATA_FILE} or {raw_path}")
    
    try:
        df = pd.read_excel(DATA_FILE)
    except PermissionError:
        print(f"Warning: Access denied to {DATA_FILE}. It might be open. Trying raw data...")
        raw_path = os.path.join(BASE_DIR, 'data', 'raw', 'data.xlsx')
        try:
             df = pd.read_excel(raw_path)
             # Raw data might need column stripping
             df.columns = df.columns.str.strip()
             # Deduplicate columns if any became identical
             df = df.loc[:, ~df.columns.duplicated()]
        except Exception as e:
             raise PermissionError(f"Could not load cleaned or raw data. Please close the Excel file. Error: {e}")
    except Exception as e:
        print(f"Error loading cleaned data: {e}. Trying raw data...")
        raw_path = os.path.join(BASE_DIR, 'data', 'raw', 'data.xlsx')
        df = pd.read_excel(raw_path)
        df.columns = df.columns.str.strip()
        df = df.loc[:, ~df.columns.duplicated()]
        
    return df

def get_column_pair(df, keywords):
    """Finds Acquired vs Required columns based on keywords."""
    cols = [c for c in df.columns if any(k in c.lower() for k in keywords)]
    
    # Logic: Acquired usually has "lulus" or "peran" (Context: peran saat lulus?), 
    # Actually standard tracer study: 
    # A = Kompetensi yang dikuasai saat lulus (Acquired)
    # B = Kompetensi yang dibutuhkan dalam pekerjaan (Required)
    # Let's look for "lulus" vs "kerja"/"manfaat"/"butuh"
    
    col_acq = next((c for c in cols if 'lulus' in c.lower() or 'saat ini' in c.lower()), None) 
    # Note: 'saat ini' might be ambiguous, usually it is "Manfaat ... saat ini" vs "Penguasaan ... saat lulus"
    
    # Let's try more specific Tracer Study logic if possible, otherwise fuzzy
    # If standard params:
    # f17xx -> Acquired/Diperoleh (Pada saat lulus)
    # f18xx -> Required/Dibutuhkan (Pada pekerjaan saat ini)
    
    # Debug: Print columns being searched
    # print(f"Searching for keywords: {keywords} in {len(df.columns)} columns")
    
    # Filter columns matching keywords
    cols = [c for c in df.columns if any(k.lower() in c.lower() for k in keywords)]
    
    if not cols:
        print(f"DEBUG: No columns found matching keywords: {keywords}")
        return None, None
        
    # print(f"DEBUG: Candidates for {keywords}: {cols}")

    col_acq = next((c for c in cols if 'lulus' in c.lower()), None)
    col_req = next((c for c in cols if 'kerja' in c.lower() or 'butuh' in c.lower() or 'terpakai' in c.lower()), None)
    
    if not col_acq or not col_req:
         # Try simpler fallback if raw data has "1" and "2" suffixes based on cleaning.py analysis
         col_acq = next((c for c in cols if c.strip().endswith('1') or c.strip().endswith('(1)')), col_acq)
         col_req = next((c for c in cols if c.strip().endswith('2') or c.strip().endswith('(2)')), col_req)
         
    if not col_acq or not col_req:
         # Last resort listing candidates to help debug
         print(f"DEBUG: Candidates for {keywords}: {cols} -> Acq: {col_acq}, Req: {col_req}")

    return col_acq, col_req

def print_styled_table(df, title=None):
    """
    Prints a pandas DataFrame in a styled format (Console).
    """
    if title:
        print(f"\n[{title}]")
    
    if df.empty:
        print("No data available.")
        return
        
    try:
        from tabulate import tabulate
        try:
             print(tabulate(df, headers='keys', tablefmt='psql', showindex=False))
        except UnicodeEncodeError:
             # Fallback for Windows Console encoding issues
             df_safe = df.copy()
             for col in df_safe.columns:
                 df_safe[col] = df_safe[col].apply(lambda x: str(x).encode('ascii', 'ignore').decode('ascii'))
             print(tabulate(df_safe, headers='keys', tablefmt='psql', showindex=False))
    except ImportError:
        print(df)

def calculate_gap(df, filter_val, filter_col='Jurusan'):
    """
    Calculates Gap Analysis for a specific filter (Jurusan or Prodi).
    Returns DataFrame results.
    """
    # Filter Data: Working Status only
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_status not in df.columns:
        print(f"DEBUG: Status column '{col_status}' not found. Available: {df.columns.tolist()[:5]}...")
        return pd.DataFrame() # handling mismatch names
        
    df_filtered = df[
        (df[filter_col] == filter_val) & 
        (df[col_status].isin(working_status))
    ].copy()
    
    if len(df_filtered) == 0:
        print(f"DEBUG: No working respondents for {filter_val}. Total rows for {filter_val}: {len(df[df[filter_col] == filter_val])}")
        return pd.DataFrame()
    
    results = []
    
    for comp_name, keywords in COMPETENCY_MAP.items():
        col_acq, col_req = get_column_pair(df_filtered, keywords)
        
        if not col_acq or not col_req:
             print(f"DEBUG: Could not find pair for {comp_name}. Keywords: {keywords}")
             continue
        
        if col_acq and col_req:
            # Debug: Check values
            # print(f"DEBUG: Sample values for {comp_name} ({col_acq}): {df_filtered[col_acq].dropna().unique()[:3]}")

            # Safe Convert to numeric using custom mapper if needed
            def safe_convert(series):
                # Try numeric first
                s_num = pd.to_numeric(series, errors='coerce')
                if s_num.notna().sum() > 0:
                     return s_num
                
                # If mostly NaN, try mapping from known strings
                # Inverse of cleaning.py map + common variations
                map_score = {
                    "Sangat Menguasai": 5,
                    "Cukup Menguasai": 4, # Follow cleaning.py oddity or fix? Let's treat Cukup as 4 if cleaning says so.
                    "Menguasai": 3,
                    "Kurang Menguasai": 2,
                    "Tidak Menguasai": 1,
                    # Common Likert
                    "Sangat Tinggi": 5, "Tinggi": 4, "Cukup": 3, "Rendah": 2, "Sangat Rendah": 1,
                    "Sangat Besar": 5, "Besar": 4, "Sedang": 3, "Kecil": 2, "Sangat Kecil": 1,
                    "Sangat Baik": 5, "Baik": 4, 
                }
                
                # Clean strings: remove digits "5 - ...", lowercase matching?
                # Helper to parse "5 - Sangat ..." -> 5
                def parse_val(x):
                    if pd.isna(x): return np.nan
                    s = str(x).strip()
                    # Check if starts with digit
                    if s and s[0].isdigit():
                        try:
                            return float(s[0]) # naive 1-digit check
                        except:
                            pass
                    # Check text map
                    for k, v in map_score.items():
                        if k.lower() in s.lower():
                            return v
                    return np.nan

                return series.apply(parse_val)

            acq_vals = safe_convert(df_filtered[col_acq])
            req_vals = safe_convert(df_filtered[col_req])
            
            # print(f"DEBUG: Converted Means: {acq_vals.mean()} | {req_vals.mean()}")

            acq_mean = acq_vals.mean()
            req_mean = req_vals.mean()
            
            if pd.notna(acq_mean) and pd.notna(req_mean):
                gap = req_mean - acq_mean
                
                results.append({
                    'Kompetensi': comp_name,
                    'Acquired (Diperoleh)': round(acq_mean, 2),
                    'Required (Dibutuhkan)': round(req_mean, 2),
                    'Gap': round(gap, 2)
                })
                
    return pd.DataFrame(results)

def create_radar_chart(df_gap, title):
    """Generates Radar Chart as Base64 String."""
    if df_gap.empty:
        return None
        
    categories = df_gap['Kompetensi'].tolist()
    N = len(categories)

    # Repeat first value to close the loop
    acq_values = df_gap['Acquired (Diperoleh)'].tolist()
    acq_values += acq_values[:1]

    req_values = df_gap['Required (Dibutuhkan)'].tolist()
    req_values += req_values[:1]

    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))

    # Draw Required
    ax.plot(angles, req_values, linewidth=2, linestyle='solid', color='red', label='Dibutuhkan (Required)')
    ax.fill(angles, req_values, 'red', alpha=0.1)

    # Draw Acquired
    ax.plot(angles, acq_values, linewidth=2, linestyle='solid', color='blue', label='Diperoleh (Acquired)')
    ax.fill(angles, acq_values, 'blue', alpha=0.1)

    # Labels
    plt.xticks(angles[:-1], categories, size=10)
    ax.set_rlabel_position(0)
    plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="grey", size=7)
    plt.ylim(0, 5.5)

    plt.title(f'Peta Radar Kompetensi Lulusan\n{title}', size=14, y=1.05, weight='bold')
    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))

    # Save to buffer
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight', dpi=300)
    plt.close(fig)
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')

def create_ipa_chart(df_gap, title):
    """Generates IPA Chart as Base64 String."""
    if df_gap.empty:
        return None
        
    plt.figure(figsize=(10, 8))

    x = df_gap['Acquired (Diperoleh)'] # Performance
    y = df_gap['Required (Dibutuhkan)'] # Importance
    labels = df_gap['Kompetensi']

    plt.scatter(x, y, color='darkblue', s=100, zorder=3)

    mean_x = x.mean()
    mean_y = y.mean()

    # Quadrant Lines
    plt.axvline(mean_x, color='gray', linestyle='--', linewidth=1.5, zorder=1)
    plt.axhline(mean_y, color='gray', linestyle='--', linewidth=1.5, zorder=1)

    # Annotations
    for i, txt in enumerate(labels):
        plt.annotate(txt, (x[i], y[i]), xytext=(5, 5), textcoords='offset points', fontsize=9)

    # Quadrant Labels (Dynamic positioning based on limits can be tricky, keeping fixed relative to mean)
    # A: High Imp, Low Perf
    plt.text(min(x.min(), mean_x) + 0.1, max(y.max(), mean_y) - 0.1, 'KUADRAN A\n(Prioritas Perbaikan)', 
             fontsize=10, color='red', weight='bold', ha='left', va='top')
             
    # B: High Imp, High Perf
    plt.text(max(x.max(), mean_x) - 0.1, max(y.max(), mean_y) - 0.1, 'KUADRAN B\n(Pertahankan)', 
             fontsize=10, color='green', weight='bold', ha='right', va='top')
             
    # C: Low Imp, Low Perf
    plt.text(min(x.min(), mean_x) + 0.1, min(y.min(), mean_y) + 0.1, 'KUADRAN C\n(Low Priority)', 
             fontsize=9, color='gray', ha='left', va='bottom')
             
    # D: Low Imp, High Perf
    plt.text(max(x.max(), mean_x) - 0.1, min(y.min(), mean_y) + 0.1, 'KUADRAN D\n(Excessive)', 
             fontsize=9, color='gray', ha='right', va='bottom')

    plt.title(f'Analisis Kuadran (IPA) Kompetensi\n{title}', size=14, weight='bold')
    plt.xlabel('Tingkat Kompetensi yang Diperoleh (Performance)', size=11)
    plt.ylabel('Tingkat Kompetensi yang Dibutuhkan (Importance)', size=11)
    plt.grid(True, linestyle=':', alpha=0.6)
    
    # Add buffer to limits
    plt.xlim(min(1, x.min()-0.5), 5.5)
    plt.ylim(min(1, y.min()-0.5), 5.5)

    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight', dpi=300)
    plt.close()
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')

def generate_full_report(jurusan_list=None):
    df = load_data()
    
    if jurusan_list is None:
        # Default: Analyze all Jurusan found in data
        jurusan_list = df['Jurusan'].unique().tolist()
        
    print(f"Generating Gap Analysis Report for: {jurusan_list}")
    
    # Ensure 'prodi' column exists if likely needed
    if 'prodi' not in df.columns and 'Program Studi' in df.columns:
        print("Deriving 'prodi' from 'Program Studi'...")
        try:
            split_data = df['Program Studi'].astype(str).str.split(' - ', n=1, expand=True)
            if split_data.shape[1] > 1:
                df['prodi'] = split_data[1]
            else:
                df['prodi'] = df['Program Studi'] # Fallback
        except Exception as e:
            print(f"Warning: Could not split 'Program Studi': {e}")
            df['prodi'] = df['Program Studi']
    
    html_sections = ""
    
    for i, jurusan in enumerate(jurusan_list):
        print(f"Processing Jurusan: {jurusan}")
        
        # 1. Jurusan Level Gap Analysis
        df_gap_jur = calculate_gap(df, jurusan, filter_col='Jurusan')
        
        if df_gap_jur.empty:
            print(f"Skipping {jurusan} (Not enough data)")
            continue
            
        print_styled_table(df_gap_jur, f"Gap Analysis: {jurusan}")
        
        # Charts
        ipa_chart = create_ipa_chart(df_gap_jur, f"Jurusan {jurusan}")
        jurusan_slug = f"jurusan_{i}"
        
        # HTML Block for Jurusan
        html_sections += f"""
        <div class="section">
            <h2 style="background-color: #2c3e50; color: white; padding: 10px; border-radius: 5px;">Jurusan: {jurusan}</h2>
            <div style="display: flex; flex-wrap: wrap; gap: 20px; align-items: flex-start;">
                <div style="flex: 1; min-width: 400px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                        <h3>Tabel Gap Kompetensi</h3>
                        <button onclick="saveTable('{jurusan_slug}_table', 'Tabel_Gap_{jurusan}')" style="background: #27ae60; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer;">Simpan Tabel</button>
                    </div>
                    {df_gap_jur.to_html(index=False, classes='table', border=0, table_id=f'{jurusan_slug}_table', float_format=lambda x: f'{x:.2f}')}
                </div>
                <div style="flex: 1; min-width: 400px; text-align: center;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                         <h3>Importance-Performance Analysis (IPA)</h3>
                         {f'<a href="data:image/png;base64,{ipa_chart}" download="IPA_Chart_{jurusan}.png" style="background: #2980b9; color: white; text-decoration: none; padding: 5px 10px; border-radius: 4px; font-size: 0.9em;">Simpan Grafik HD</a>' if ipa_chart else ''}
                    </div>
                    {f'<img src="data:image/png;base64,{ipa_chart}" style="max-width: 100%; border: 1px solid #ddd; border-radius: 8px;">' if ipa_chart else 'No Chart'}
                </div>
            </div>
            <hr style="margin: 40px 0; border-top: 2px dashed #ccc;">
            
            <h3>Detail per Program Studi</h3>
        """
        
        # 2. Prodi Level (Turunan)
        prodis = df[df['Jurusan'] == jurusan]['prodi'].copy().unique() # Copy to avoid SettingWithCopy warning on unique? No unique returns array.
        # Handle prodi logic again just to be safe if 'prodi' missing from earlier block? 
        # Actually it's guaranteed to exist or fallback in the header of this function.
        if 'prodi' not in df.columns:
             # Just in case fallback didn't run properly
             if 'Program Studi' in df.columns:
                  prodis = df[df['Jurusan'] == jurusan]['Program Studi'].unique()
             else:
                  prodis = []

        for j, prodi in enumerate(prodis):
            # print(f"  > Processing Prodi: {prodi}")
            df_gap_prodi = calculate_gap(df, prodi, filter_col='prodi' if 'prodi' in df.columns else 'Program Studi')
            
            if df_gap_prodi.empty:
                continue
            
            print_styled_table(df_gap_prodi, f"Gap Analysis Prodi: {prodi}")
                
            # Radar Chart for Prodi
            radar_chart = create_radar_chart(df_gap_prodi, f"Prodi {prodi}")
            prodi_slug = f"{jurusan_slug}_prodi_{j}"
            
            html_sections += f"""
            <div style="margin-left: 20px; margin-bottom: 40px; background: #f9f9f9; padding: 20px; border-radius: 8px; border-left: 5px solid #3498db;">
                <h4 style="margin-top:0; color: #2980b9;">{prodi}</h4>
                <div style="display: flex; flex-wrap: wrap; gap: 20px;">
                    <div style="flex: 1;">
                        <div style="display: flex; justify-content: space-between; margin-bottom: 5px;">
                            <span><b>Tabel Gap</b></span>
                            <button onclick="saveTable('{prodi_slug}_table', 'Tabel_Gap_{prodi}')" style="background: #27ae60; color: white; border: none; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; cursor: pointer;">Simpan Tabel</button>
                        </div>
                         {df_gap_prodi.to_html(index=False, classes='table table-sm', border=0, table_id=f'{prodi_slug}_table', float_format=lambda x: f'{x:.2f}')}
                    </div>
                    <div style="flex: 1; text-align: center;">
                        <div style="text-align: right; margin-bottom: 5px;">
                             {f'<a href="data:image/png;base64,{radar_chart}" download="Radar_Chart_{prodi}.png" style="background: #2980b9; color: white; text-decoration: none; padding: 3px 8px; border-radius: 4px; font-size: 0.8em;">Simpan Grafik HD</a>' if radar_chart else ''}
                        </div>
                         {f'<img src="data:image/png;base64,{radar_chart}" style="max-width: 100%;">' if radar_chart else 'No Chart'}
                    </div>
                </div>
            </div>
            """
            
        html_sections += "</div>" # End Jurusan Section

    # Wrap in Full HTML
    full_html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Laporan Analisis Gap Kompetensi</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
        <!-- Load html2canvas -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <style>
            body {{ font-family: 'Inter', sans-serif; background: #f4f6f9; color: #333; padding: 40px; }}
            .container {{ max-width: 1200px; margin: 0 auto; background: #fff; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.05); }}
            h1 {{ text-align: center; color: #2c3e50; margin-bottom: 40px; }}
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; background: #fff; }}
            th, td {{ padding: 10px 15px; text-align: left; border-bottom: 1px solid #eee; }}
            th {{ background-color: #3498db; color: white; }}
            tr:nth-child(even) {{ background-color: #f8fafc; }}
            .table-sm th, .table-sm td {{ padding: 8px 10px; font-size: 0.9rem; }}
        </style>
        <script>
            function saveTable(tableId, filename) {{
                const table = document.getElementById(tableId);
                html2canvas(table).then(canvas => {{
                    // Create an explicit white background since canvas transparent by default
                    const ctx = canvas.getContext('2d');
                    ctx.globalCompositeOperation = 'destination-over';
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    
                    const link = document.createElement('a');
                    link.download = filename + '.png';
                    link.href = canvas.toDataURL();
                    link.click();
                }});
            }}
        </script>
    </head>
    <body>
        <div class="container">
            <h1>Analisis Gap Kompetensi Lulusan</h1>
            <p style="text-align: center; color: #7f8c8d; margin-bottom: 40px;">
                Perbandingan antara Kompetensi yang Diperoleh (Acquired) vs Dibutuhkan (Required)
            </p>
            {html_sections}
        </div>
    </body>
    </html>
    """
    
    output_path = os.path.join(REPORTS_DIR, 'gap_analysis_report.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(full_html)
        
    print(f"Report generated: {output_path}")
    
    # Auto-open in browser
    try:
        webbrowser.open('file://' + os.path.abspath(output_path))
        print("Opening report in browser...")
    except Exception as e:
        print(f"Could not open browser: {e}")

if __name__ == "__main__":
    # Example Usage: Analyze Specific list or all
    # User said: "Tapi saya bisa mengubah-ngubah jurusan mana saja yang mau di analisis"
    # So we can define a list here easily.
    
    TARGET_JURUSAN = [
        # "Semua", or list specific ones
        "Administrasi Bisnis", 
        "Teknologi Pertanian",
        "Akuntansi",
        "Teknik Sipil dan Perencanaan",
        "Teknik Elektro",
        "Teknik Mesin",
        "Ilmu Kelautan dan Perikanan",
        "Jurusan Lainnya..."
    ]
    
    # Or just discover from data:
    try:
        df_load = load_data()
        all_jurusan = df_load['Jurusan'].unique().tolist()
        # Filter raw names if needed or use cleaned
        valid_jurusan = [j for j in all_jurusan if isinstance(j, str)]
        generate_full_report(valid_jurusan)
    except Exception as e:
        print(f"Error: {e}")
