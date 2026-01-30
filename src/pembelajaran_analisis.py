import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from math import pi
import io
import base64
import os
import webbrowser

# Setup Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data', 'processed')
REPORTS_DIR = os.path.join(BASE_DIR, 'reports')
DATA_FILE = os.path.join(DATA_DIR, 'cleaned_data.xlsx')

# Ensure directories exist
os.makedirs(REPORTS_DIR, exist_ok=True)

# Define Learning Methods Columns & Labels
LEARNING_METHODS = {
    'Perkuliahan': 'Perkuliahan',
    'Demonstrasi': 'Demonstrasi',
    'Partisipasi dalam proyek riset': 'Riset / Proyek',
    'Magang': 'Magang',
    'Praktikum': 'Praktikum',
    'Kerja Lapangan': 'Kerja Lapangan',
    'Diskusi': 'Diskusi'
}

# Likert Scale Mapping
LIKERT_MAP = {
    "Sangat Besar": 5,
    "Besar": 4,
    "Cukup Besar": 3,
    "Kurang Besar": 2,
    "Tidak Sama Sekali": 1
}

def load_data():
    """Loads the cleaned data."""
    if not os.path.exists(DATA_FILE):
        raise FileNotFoundError(f"Data file not found at {DATA_FILE}")
    
    try:
        df = pd.read_excel(DATA_FILE)
        return df
    except Exception as e:
        raise Exception(f"Error loading data: {e}")

def convert_likert(val):
    """Converts Likert string to number."""
    if pd.isna(val):
        return np.nan
    val_str = str(val).strip()
    
    # Direct Map
    if val_str in LIKERT_MAP:
        return LIKERT_MAP[val_str]
    
    # Partial match (case insensitive)
    for k, v in LIKERT_MAP.items():
        if k.lower() in val_str.lower():
            return v
            
    # Try parsing "5 - Sangat ..."
    if val_str and val_str[0].isdigit():
        try:
            return float(val_str[0])
        except:
            pass
            
    return np.nan

def calculate_means(df):
    """Calculates mean scores for each learning method."""
    stats = {}
    
    for col, label in LEARNING_METHODS.items():
        if col in df.columns:
            # Convert series
            numeric_series = df[col].apply(convert_likert)
            mean_val = numeric_series.mean()
            stats[label] = mean_val
        else:
            print(f"Warning: Column '{col}' not found in data.")
            stats[label] = 0 # Prepare safe default or skip?
            
    # Convert to DataFrame for easier plotting
    df_stats = pd.DataFrame(list(stats.items()), columns=['Metode', 'Mean Score'])
    return df_stats.sort_values(by='Mean Score', ascending=True) # Sort for Bar Chart

def create_bar_chart(df_stats):
    """Creates a horizontal bar chart ranking learning methods."""
    plt.figure(figsize=(10, 6))
    
    # Colors: Highlight top 3 vs others, or just gradient
    # Let's use a nice blue palette
    colors = plt.cm.Blues(np.linspace(0.4, 0.9, len(df_stats)))
    
    bars = plt.barh(df_stats['Metode'], df_stats['Mean Score'], color=colors)
    
    # Add values on bars
    for bar in bars:
        width = bar.get_width()
        plt.text(width + 0.05, bar.get_y() + bar.get_height()/2, 
                 f'{width:.2f}', ha='left', va='center', fontsize=10, weight='bold')

    # plt.title('Peringkat Penekanan Metode Pembelajaran\nPoliteknik Negeri Pontianak', size=14, weight='bold') # Removed per user request
    plt.xlabel('Skor Rata-rata (1-5)', size=11)
    plt.xlim(0, 5.5)
    plt.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()
    
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight', dpi=300)
    plt.close()
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')

def create_radar_chart(df_stats):
    """Creates a radar chart handling unsorted data (we need fixed order likely)."""
    # Sort specifically for Radar to make it look consistent? 
    # Or just use the input order? Let's use input order but we sorted it for Bar.
    # Re-order to arbitrary or alphabetical might be better, or keep strictly sorted?
    # Usually consistent order (like clock) is better for comparison.
    # Let's re-sort by a fixed list if we want consistent axes, 
    # but since this is a single chart, sorted by magnitude is also fine to show shape.
    # Let's stick to the dataframe order (which is sorted by score ascending currently).
    # Actually for radar, maybe "Standard Order" is better to compare if we had multiple.
    # Let's re-sort by Label name or keep as is.
    
    # Let's use a copy and re-sort by name for consistent shape if we run multiple times?
    # Actually, let's keep it simple: Sorted by Mean is fine, shows the "shape of emphasis".
    
    categories = df_stats['Metode'].tolist()
    N = len(categories)
    
    if N == 0: return None

    values = df_stats['Mean Score'].tolist()
    values += values[:1] # Close loop
    
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]
    
    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    
    ax.plot(angles, values, linewidth=2, linestyle='solid', color='#d35400')
    ax.fill(angles, values, '#d35400', alpha=0.25)
    
    # Labels
    plt.xticks(angles[:-1], categories, size=11)
    ax.set_rlabel_position(0)
    plt.yticks([1, 2, 3, 4, 5], ["1", "2", "3", "4", "5"], color="grey", size=7)
    plt.ylim(0, 5.5)
    
    # plt.title('Profil Radar Penekanan Pembelajaran', size=15, weight='bold', y=1.05) # Removed per user request
    
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight', dpi=300)
    plt.close()
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')

def calculate_jurusan_means(df):
    """Calculates mean scores for each learning method grouped by Jurusan."""
    jurusan_stats = {}
    
    # Check if Jurusan column exists
    if 'Jurusan' not in df.columns:
        print("Warning: 'Jurusan' column not found.")
        return pd.DataFrame()

    unique_jurusan = df['Jurusan'].dropna().unique()
    
    # Clean Jurusan names (strip whitespace)
    unique_jurusan = [j for j in unique_jurusan if isinstance(j, str)]
    
    for jur in unique_jurusan:
        # Filter by Jurusan
        df_jur = df[df['Jurusan'] == jur]
        
        scores = {}
        for col, label in LEARNING_METHODS.items():
            if col in df_jur.columns:
                val = df_jur[col].apply(convert_likert).mean()
                scores[label] = val
            else:
                scores[label] = np.nan
        
        jurusan_stats[jur] = scores
        
    df_heatmap = pd.DataFrame.from_dict(jurusan_stats, orient='index')
    return df_heatmap

def create_heatmap(df_heatmap):
    """Creates a heatmap visualizing method emphasis by Jurusan."""
    if df_heatmap.empty:
        return None
        
    # Visual Clustering: Sort Rows and Columns to group similar colors
    # 1. Sort Columns by overall mean (High emphasis methods on left)
    col_means = df_heatmap.mean(axis=0)
    df_heatmap = df_heatmap[col_means.sort_values(ascending=False).index]
    
    # 2. Sort Rows by overall mean (High emphasis majors on top)
    # OR Sort by the most dominant method to group by "Pattern"
    # Let's simple sort by mean first, effective for "intensity" clustering
    row_means = df_heatmap.mean(axis=1)
    df_heatmap = df_heatmap.loc[row_means.sort_values(ascending=False).index]
    
    # Create Heatmap manually using imshow since we avoid seaborn dependency if possible
    data = df_heatmap.values
    x_labels = df_heatmap.columns
    y_labels = df_heatmap.index
    
    # Dynamic height based on number of majors
    fig, ax = plt.subplots(figsize=(12, len(y_labels)*0.5 + 4))
    
    # Calculate Data Range for Dynamic Contrast
    vmin = np.nanmin(data)
    vmax = np.nanmax(data)
    
    # Plot with Diverging Colormap (Red-Yellow-Blue reversed) for better contrast
    # High (Blue) = Good, Low (Red) = Less Emphasis
    im = ax.imshow(data, cmap='RdYlBu_r', vmin=vmin, vmax=vmax)
    
    # Axes
    ax.set_xticks(np.arange(len(x_labels)))
    ax.set_yticks(np.arange(len(y_labels)))
    ax.set_xticklabels(x_labels, rotation=45, ha="right", size=10)
    ax.set_yticklabels(y_labels, size=10)
    
    # Annotate
    for i in range(len(y_labels)):
        for j in range(len(x_labels)):
            val = data[i, j]
            if pd.notna(val):
                text_color = "white" if val > 3.5 else "black"
                text = ax.text(j, i, f"{val:.1f}",
                               ha="center", va="center", color=text_color, size=9, weight='bold')

    # Colorbar
    cbar = ax.figure.colorbar(im, ax=ax, shrink=0.5)
    cbar.ax.set_ylabel("Skor Rata-rata (1-5)", rotation=-90, va="bottom")
    
    # Layout
    plt.tight_layout()
    
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', bbox_inches='tight', dpi=300)
    plt.close()
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode('utf-8')

def generate_report():
    print("Loading data...")
    df = load_data()
    
    print("Calculating scores...")
    df_stats = calculate_means(df)
    
    print("\nMean Scores:")
    print(df_stats)
    
    print("Generating charts...")
    bar_chart = create_bar_chart(df_stats)
    radar_chart = create_radar_chart(df_stats)
    
    print("Generating Heatmap...")
    df_heatmap = calculate_jurusan_means(df)
    heatmap_chart = create_heatmap(df_heatmap)

    # Print Styled Table
    try:
        from tabulate import tabulate
        # print("\n" + "="*50)
        # print("   HASIL ANALISIS PENEKANAN METODE PEMBELAJARAN") # Removed per user request
        # print("="*50)
        
        # Prepare data for display
        df_display = df_stats.sort_values(by='Mean Score', ascending=False).copy()
        df_display['Mean Score'] = df_display['Mean Score'].map('{:.2f}'.format)
        df_display['Kategori'] = df_display['Mean Score'].astype(float).apply(get_category)
        
        try:
            print(tabulate(df_display, headers='keys', tablefmt='psql', showindex=False))
        except UnicodeEncodeError:
            # Fallback for Windows Console
            print(tabulate(df_display, headers='keys', tablefmt='grid', showindex=False))
        print("\n")
    except ImportError:
        print("\n[Tips] Install 'tabulate' untuk tampilan tabel yang lebih rapi: pip install tabulate")
        print(df_stats)
    print("Calculating dimensions...")
    df_dim = calculate_dimensions(df_stats)
    
    # Print Styled Table for Dimensions
    try:
        from tabulate import tabulate
        
        print("\n" + "="*50)
        print("   ANALISIS DIMENSI PEMBELAJARAN")
        print("="*50)
        
        df_display_dim = df_dim.copy()
        df_display_dim['Skor Rata-rata Gabungan'] = df_display_dim['Skor Rata-rata Gabungan'].map('{:.2f}'.format)
        
        try:
             print(tabulate(df_display_dim, headers='keys', tablefmt='psql', showindex=False))
        except UnicodeEncodeError:
             print(tabulate(df_display_dim, headers='keys', tablefmt='grid', showindex=False))
        print("\n")
        
    except ImportError:
        pass
        
    # HTML Content
    def get_status_color(status):
        if status == "Dominan Utama": return "#27ae60" # Green
        if status == "Pendukung Kuat": return "#2980b9" # Blue
        return "#e67e22" # Orange

    html_content = f"""
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Analisis Metode Pembelajaran</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap" rel="stylesheet">
        <!-- Load html2canvas -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <script>
            function saveTable(tableId, filename) {{
                const table = document.getElementById(tableId);
                html2canvas(table).then(canvas => {{
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
        <style>
            body {{ font-family: 'Inter', sans-serif; background: #fdfdfd; color: #333; padding: 40px; }}
            .container {{ max-width: 1000px; margin: 0 auto; background: #fff; padding: 50px; border-radius: 16px; box-shadow: 0 10px 30px rgba(0,0,0,0.08); }}
            h1 {{ text-align: center; color: #2c3e50; margin-bottom: 10px; font-weight: 800; }}
            .subtitle {{ text-align: center; color: #7f8c8d; margin-bottom: 50px; font-size: 1.1em; }}
            
            .chart-section {{ margin-bottom: 60px; }}
            .chart-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px; }}
            h2 {{ margin: 0; color: #34495e; }}
            
            .btn-download {{
                background-color: #2980b9; color: white; text-decoration: none;
                padding: 8px 16px; border-radius: 6px; font-size: 0.9em; transition: background 0.2s;
            }}
            .btn-download:hover {{ background-color: #1abc9c; }}
            
            img {{ max-width: 100%; height: auto; border: 1px solid #eee; border-radius: 8px; }}
            
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ padding: 12px 15px; text-align: left; border-bottom: 1px solid #eee; }}
            th {{ background-color: #f8f9fa; font-weight: 600; color: #2c3e50; }}
            tr:hover {{ background-color: #fcfcfc; }}
            .score {{ font-weight: bold; color: #2980b9; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Analisis Penekanan Metode Pembelajaran</h1>
            <div class="subtitle">Evaluasi Persepsi Lulusan Terhadap Pendekatan Didaktik di Politeknik Negeri Pontianak</div>
            
            <!-- Ringkasan Data -->
            <div style="background: #eef2f7; padding: 20px; border-radius: 8px; margin-bottom: 40px; text-align: center;">
                <strong>Jumlah Responden:</strong> {len(df)} orang
            </div>
            
            <div class="chart-section">
                <div class="chart-header">
                    <h2>Peringkat Dominasi Metode Pembelajaran</h2>
                    <a href="data:image/png;base64,{bar_chart}" download="Peringkat_Metode_Pembelajaran.png" class="btn-download">Simpan Grafik HD</a>
                </div>
                <img src="data:image/png;base64,{bar_chart}" alt="Bar Chart">
                
                <div style="margin-top: 30px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                        <h3>Tabel Skor Rata-rata</h3>
                        <button onclick="saveTable('metodeTable', 'Tabel_Metode_Pembelajaran')" class="btn-download" style="border: none; cursor: pointer;">Simpan Tabel</button>
                    </div>
                    <table id="metodeTable" style="background: white; padding: 10px;">
                        <thead>
                            <tr>
                                <th>Metode Pembelajaran</th>
                                <th>Skor Rata-rata (1-5)</th>
                                <th>Kategori</th>
                            </tr>
                        </thead>
                        <tbody>
                            {''.join([f"<tr><td>{row['Metode']}</td><td class='score'>{row['Mean Score']:.2f}</td><td>{get_category(row['Mean Score'])}</td></tr>" for _, row in df_stats.sort_values(by='Mean Score', ascending=False).iterrows()])}
                        </tbody>
                    </table>
                </div>
                </div>
            </div>

                </div>
            </div>
            
            <div class="chart-section" style="margin-top: 50px;">
                <div class="chart-header">
                    <h2>Peta Panas (Heatmap) Intensitas Pembelajaran per Jurusan</h2>
                    <a href="data:image/png;base64,{heatmap_chart}" download="Heatmap_Jurusan.png" class="btn-download">Simpan Grafik HD</a>
                </div>
                <p style="margin-bottom: 20px; color: #666;">
                    Visualisasi ini membandingkan penekanan metode pembelajaran lintas jurusan. 
                    Warna yang lebih gelap menunjukkan intensitas yang lebih tinggi.
                </p>
                <div style="text-align: center;">
                    <img src="data:image/png;base64,{heatmap_chart}" alt="Heatmap Chart" style="max-width: 100%; border: 1px solid #ccc;">
                </div>
            </div>

            <!-- Dimension Table Section -->
            <div style="margin-top: 50px;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                     <h2>Analisis Dimensi Pembelajaran</h2>
                     <button onclick="saveTable('dimensiTable', 'Tabel_Dimensi_Pembelajaran')" class="btn-download" style="border: none; cursor: pointer;">Simpan Tabel</button>
                </div>
                <table id="dimensiTable" style="background: white; padding: 10px;">
                    <thead>
                        <tr>
                            <th>Dimensi Pembelajaran</th>
                            <th>Komponen Metode</th>
                            <th>Skor Rata-rata Gabungan</th>
                            <th>Status Dominasi</th>
                        </tr>
                    </thead>
                    <tbody>
                        {''.join([f"<tr><td>{row['Dimensi Pembelajaran']}</td><td>{row['Komponen Metode']}</td><td class='score'>{row['Skor Rata-rata Gabungan']:.2f}</td><td><span style='padding: 4px 8px; border-radius: 4px; background: {get_status_color(row['Status Dominasi'])}; color: white; font-size: 0.9em;'>{row['Status Dominasi']}</span></td></tr>" for _, row in df_dim.iterrows()])}
                    </tbody>
                </table>
            </div>
            
            <div class="chart-section" style="margin-top: 50px;">
                <div class="chart-header">
                    <h2>Profil Radar Penekanan (DNA Vokasi)</h2>
                    <a href="data:image/png;base64,{radar_chart}" download="Radar_Metode_Pembelajaran.png" class="btn-download">Simpan Grafik HD</a>
                </div>
                <div style="text-align: center;">
                    <img src="data:image/png;base64,{radar_chart}" alt="Radar Chart" style="max-height: 600px; width: auto;">
                </div>
                <p style="text-align: center; color: #666; margin-top: 15px; font-style: italic;">
                    Grafik Radar menunjukkan keseimbangan antara berbagai metode pembelajaran. <br>
                    Dominasi pada "Praktikum", "Magang", dan "Kerja Lapangan" mengindikasikan kuatnya karakteristik pendidikan vokasi.
                </p>
            </div>
            
        </div>
    </body>
    </html>
    """
    
    output_path = os.path.join(REPORTS_DIR, 'pembelajaran_analisis_report.html')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"Report generated at: {output_path}")
    
    try:
        webbrowser.open('file://' + os.path.abspath(output_path))
        print("Opening report in browser...")
    except Exception as e:
        print(f"Could not open browser: {e}")

def calculate_dimensions(df_stats):
    """Calculates grouped scores for learning dimensions."""
    # Define Groups
    groups = {
        'Dimensi Praktik (Vocational Core)': ['Kerja Lapangan', 'Magang', 'Praktikum'],
        'Dimensi Teori & Akademik': ['Perkuliahan', 'Diskusi'],
        'Dimensi Riset': ['Riset / Proyek']
    }
    
    results = []
    
    for dim_name, methods in groups.items():
        # Filter stats for these methods
        mask = df_stats['Metode'].isin(methods)
        if mask.any():
            sub_df = df_stats[mask]
            mean_score = sub_df['Mean Score'].mean()
            component_str = ", ".join(sub_df['Metode'].tolist())
            
            # Determine Status
            if mean_score >= 4.0:
                status = "Dominan Utama"
            elif mean_score >= 3.7:
                status = "Pendukung Kuat"
            else:
                status = "Area Pengembangan"
                
            results.append({
                'Dimensi Pembelajaran': dim_name,
                'Komponen Metode': component_str,
                'Skor Rata-rata Gabungan': mean_score,
                'Status Dominasi': status
            })
            
    return pd.DataFrame(results)

def get_category(score):
    if score >= 4.5: return "Sangat Besar"
    if score >= 3.5: return "Besar"
    if score >= 2.5: return "Cukup"
    if score >= 1.5: return "Kurang"
    return "Sangat Kurang"

if __name__ == "__main__":
    generate_report()
