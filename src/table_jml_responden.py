import pandas as pd
import numpy as np
import io
import time
import os

try:
    import folium
    from folium.plugins import MarkerCluster
except ImportError as e:
    print(f"Import Error: {e}")
    folium = None

# Optional Geopandas for Static Maps
try:
    import geopandas as gpd
    import matplotlib.pyplot as plt
    import matplotlib.patheffects
    from shapely.geometry import Point
    # from mpl_toolkits.axes_grid1 import make_axes_locatable # Removed
except ImportError:
    gpd = None
    plt = None

# Database Koordinat Provinsi Indonesia (Latitude, Longitude)
INDO_COORDS = {
    'Aceh': (4.6951, 96.7494),
    'Sumatera Utara': (2.1154, 99.5451),
    'Sumatera Barat': (-0.7399, 100.8000),
    'Riau': (0.2933, 101.7068),
    'Jambi': (-1.4852, 102.4381),
    'Sumatera Selatan': (-3.3194, 104.9144),
    'Bengkulu': (-3.5778, 102.3464),
    'Lampung': (-4.5586, 105.4068),
    'Kepulauan Bangka Belitung': (-2.7411, 106.4406),
    'Kepulauan Riau': (3.9456, 108.1428),
    'DKI Jakarta': (-6.2088, 106.8456),
    'Jawa Barat': (-6.9175, 107.6191),
    'Jawa Tengah': (-7.1510, 110.1403),
    'Daerah Istimewa Yogyakarta': (-7.7956, 110.3695),
    'Jawa Timur': (-7.5360, 112.2384),
    'Banten': (-6.4058, 106.0640),
    'Bali': (-8.4095, 115.1889),
    'Nusa Tenggara Barat': (-8.6529, 117.3616),
    'Nusa Tenggara Timur': (-8.6574, 121.0794),
    'Kalimantan Barat': (-0.2787, 111.4753),
    'Kalimantan Tengah': (-1.6815, 113.3824),
    'Kalimantan Selatan': (-3.0926, 115.2838),
    'Kalimantan Timur': (0.5387, 116.4194),
    'Kalimantan Utara': (3.0731, 116.0414),
    'Sulawesi Utara': (0.6247, 123.9750),
    'Sulawesi Tengah': (-1.4300, 121.4456),
    'Sulawesi Selatan': (-3.6687, 119.9740),
    'Sulawesi Tenggara': (-4.1449, 122.1746),
    'Gorontalo': (0.6999, 122.4467),
    'Sulawesi Barat': (-2.8441, 119.2321),
    'Maluku': (-3.2385, 129.4936),
    'Maluku Utara': (0.2120, 127.9791),
    'Papua Barat': (-1.3361, 133.1747),
    'Papua': (-4.2699, 138.0804),
    'Sumatera Tengah': (-0.947, 100.417)
}

# Database Koordinat Kota/Kabupaten di Kalbar
KALBAR_COORDS = {
    'Pontianak': (-0.026330, 109.342504),
    'Kubu Raya': (-0.468725, 109.378906),
    'Ketapang': (-1.595914, 110.490723), # Approx center
    'Mempawah': (0.334000, 109.116000),
    'Sanggau': (0.120800, 110.586600),
    'Sambas': (1.338700, 109.317500),
    'Landak': (0.435700, 109.957500), # Ngabang
    'Singkawang': (0.910300, 108.985000),
    'Kayong Utara': (-1.144800, 109.957900), # Sukadana
    'Sintang': (0.071100, 111.495200),
    'Kapuas Hulu': (0.814300, 112.930400), # Putussibau
    'Sekadau': (0.035700, 110.938800),
    'Melawi': (-0.686500, 111.688100), # Nanga Pinoh
    'Bengkayang': (0.931700, 109.529900)
}


def generate_static_map_geopandas(df_counts, output_path, region_name='Indonesia', total_reference=None):
    """
    Generates a static map.
    - Indonesia: Choropleth (Polygons) if GeoJSON available.
    - Kalbar: Bubble Map (Points) as fallback for missing Regency shapefiles.
    - total_reference: Optional total to use for percentage calculation (to match table).
    """
    if not gpd or not plt:
        print("Geopandas or Matplotlib not available.")
        return

    # 1. Indonesia Map -> CHOROPLETH
    if region_name == 'Indonesia':
        try:
            # Load Indonesia Province GeoJSON
            url = "https://raw.githubusercontent.com/superpikar/indonesia-geojson/master/indonesia-province-simple.json"
            gdf_indo = gpd.read_file(url)
            
            # Normalize Names for Merge
            # GeoJSON 'Propinsi' is usually UPPERCASE (e.g., 'JAWA BARAT')
            # Our DF 'Provinsi' might be Mixed (e.g., 'Jawa Barat')
            
            # Create copy to avoid mutating original
            df_plot = df_counts.copy()
            if 'Provinsi' in df_plot.columns:
                 # Normalize: Upper case
                 df_plot['Provinsi_Upper'] = df_plot['Provinsi'].astype(str).str.upper()
                 
                 # Manual fixes for common mismatches if known
                 # e.g. 'DI YOGYAKARTA' vs 'JAKARTA RAYA' - check content
                 name_map = {
                     'DI YOGYAKARTA': 'DAERAH ISTIMEWA YOGYAKARTA',
                     'DKI JAKARTA': 'DKI JAKARTA'
                 }
                 df_plot['Provinsi_Upper'] = df_plot['Provinsi_Upper'].replace(name_map)
                 
                 # GeoJSON column: 'Propinsi'
                 gdf_indo['Propinsi_Upper'] = gdf_indo['Propinsi'].astype(str).str.upper()
                 
                 # Merge
                 gdf_merged = gdf_indo.merge(df_plot, left_on='Propinsi_Upper', right_on='Provinsi_Upper', how='left')
                 
                 # Fill NaN with 0 for plotting
                 gdf_merged['Jumlah'] = gdf_merged['Jumlah'].fillna(0)
                 
                 # Plot Choropleth
                 fig, ax = plt.subplots(figsize=(15, 6))
                 
                 gdf_merged.plot(column='Jumlah',
                                 ax=ax,
                                 legend=True,
                                 legend_kwds={'label': "Jumlah Alumni", 'orientation': "horizontal"},
                                 cmap='Blues',
                                 edgecolor='black',
                                 linewidth=0.5,
                                 missing_kwds={'color': 'lightgrey'})
                                 
                 ax.set_title(f"Sebaran Alumni - {region_name} (Choropleth)", fontsize=16)
                 ax.axis('off')
                 
                 plt.tight_layout()
                 plt.savefig(output_path, dpi=150, bbox_inches='tight')
                 plt.close()
                 print(f"Choropleth Map saved: {output_path}")
                 return

        except Exception as e:
            print(f"Failed to generate Choropleth for {region_name}: {e}")
            print("Falling back to Point map...")


    # 2. Kalbar Map -> CHOROPLETH
    if region_name == 'Kalimantan Barat':
        try:
             # Load Kalbar GeoJSON
             url = "https://raw.githubusercontent.com/ghapsara/indonesia-atlas/master/kabupaten-kota/Kalimantan%20Barat/kalimantan-barat-simplified-topo.json"
             gdf_kalbar = gpd.read_file(url)
             
             # User reported duplicates (Mempawah appearing twice). Drop duplicates by name.
             gdf_kalbar = gdf_kalbar.drop_duplicates(subset='kabkot', keep='first')
             
             # Prepare Data for Merge
             # DF column usually 'Kota/Kabupaten' which has values like 'Kab. Sambas', 'Kota Pontianak'
             # GeoJSON 'kabkot' has 'Sambas', 'Pontianak' (confirmed via script)
             
             # We need to normalize DF names to match GeoJSON
             df_plot = df_counts.copy()
             
             # DEBUG: Print GeoJSON contents
             print("--- Kalbar GeoJSON Regions ---")
             print(gdf_kalbar['kabkot'].unique())
             print(f"Total Regions: {len(gdf_kalbar)}")
             
             def normalize_kalbar_name(name):
                 # Remove 'Kab. ', 'Kota ', 'Kabupaten ' prefix case-insensitive
                 # simple string replace might miss case or variations
                 import re
                 # Remove "Kabupaten", "Kab.", "Kota" with optional trailing space/dot
                 clean = re.sub(r'^(Kabupaten|Kab\.?|Kota)\s*', '', name, flags=re.IGNORECASE)
                 return clean.strip()
             
             label_col = 'Kota/Kabupaten'
             if label_col in df_plot.columns:
                 df_plot['kabkot_clean'] = df_plot[label_col].apply(normalize_kalbar_name)
                 
                 # Check if 'Pontianak' is in GeoJSON
                 if 'Pontianak' not in gdf_kalbar['kabkot'].values:
                     # Create synthetic geometry for Pontianak (approximate circle/buffer)
                     # Coordinate: (-0.026330, 109.342504)
                     # 0.05 degrees approx 5km radius
                     pt = Point(109.342504, -0.026330)
                     poly = pt.buffer(0.08) 
                     
                     # Add to GeoDataFrame
                     new_row = pd.DataFrame([{'kabkot': 'Pontianak', 'geometry': poly}])
                     gdf_kalbar = pd.concat([gdf_kalbar, new_row], ignore_index=True)
                 
                 # Merge
                 gdf_merged = gdf_kalbar.merge(df_plot, left_on='kabkot', right_on='kabkot_clean', how='left')
                 
                 # Ensure Pontianak is LAST (to be drawn on top)
                 # Create a sort key: 1 for Pontianak, 0 for others
                 gdf_merged['sort_order'] = gdf_merged['kabkot'].apply(lambda x: 1 if 'Pontianak' in str(x) else 0)
                 gdf_merged = gdf_merged.sort_values('sort_order')
                 
                 # Fill NaN
                 gdf_merged['Jumlah Responden'] = gdf_merged['Jumlah Responden'].fillna(0)
                 
                 # Calculate Percentage for Labels
                 # Use total_reference if provided to match table exactly
                 if total_reference:
                     total_responden = total_reference
                 else:
                     # Fallback to sum of mapped data (risk of mismatch if merge drops rows)
                     # Or sum of original DF?
                     # Better to use sum of input DF if columns exist
                     if 'Jumlah Responden' in df_counts.columns:
                         total_responden = df_counts['Jumlah Responden'].sum()
                     else:
                         total_responden = gdf_merged['Jumlah Responden'].sum()
                         
                 if total_responden > 0:
                     gdf_merged['pct'] = (gdf_merged['Jumlah Responden'] / total_responden) * 100
                 else:
                     gdf_merged['pct'] = 0
                 
                 # Plot Choropleth
                 fig, ax = plt.subplots(figsize=(12, 10))
                 
             import matplotlib.colors as mcolors
             
             # Split Data
             gdf_pontianak = gdf_merged[gdf_merged['kabkot'] == 'Pontianak']
             gdf_others = gdf_merged[gdf_merged['kabkot'] != 'Pontianak']
             
             # Create Custom Colormap for Others (LightGrey -> DarkBlue)
             # darkblue ke lightgrey (User request interpreted as range)
             cmap_custom = mcolors.LinearSegmentedColormap.from_list("grey_blue", ["lightgrey", "darkblue"])
             
             # Plot 'Others' first
             gdf_others.plot(column='Jumlah Responden',
                             ax=ax,
                             legend=True,
                             legend_kwds={'label': "Jumlah Alumni (Non-Pontianak)", 'orientation': "vertical"},
                             cmap=cmap_custom,
                             edgecolor='black',
                             linewidth=0.5,
                             missing_kwds={'color': 'lightgrey'})
                             
             # Plot Pontianak (Red)
             if not gdf_pontianak.empty:
                 gdf_pontianak.plot(ax=ax,
                                    color='red',
                                    edgecolor='black',
                                    linewidth=0.5)
                 # Note: Pontianak won't be in the colorbar, which is fine as it's an outlier
             
             # Add Labels (Annotate) - Iterate over full merged DF to label all
             for idx, row in gdf_merged.iterrows():
                 # Skip if geometry is missing or empty
                 if row.geometry is None: continue
                 
                 # Get Centroid
                 centroid = row.geometry.centroid
                 x, y = centroid.x, centroid.y
                 
                 # Get Name (use original kabkot from GeoJSON or cleaned)
                 name = row.get('kabkot', '')
                 val_pct = row.get('pct', 0)
                 
                 # Format Label: "Name\n(X.X%)"
                 label = f"{name}\n({val_pct:.1f}%)"
                 
                 # Add Text
                 # Use white halo for readability
                 ax.annotate(text=label, xy=(x, y), xytext=(0, 0), textcoords="offset points", # Center
                             ha='center', va='center',
                             fontsize=8, color='black', weight='bold',
                             path_effects=[matplotlib.patheffects.withStroke(linewidth=2, foreground="white")])

             ax.set_title("Sebaran Alumni - Kalimantan Barat", fontsize=16)
             ax.axis('off')
             
             plt.tight_layout()
             plt.savefig(output_path, dpi=150, bbox_inches='tight')
             plt.close()
             print(f"Kalbar Choropleth Map saved: {output_path}")
             return
                 
        except Exception as e:
             print(f"Failed to generate Kalbar Choropleth: {e}")
             print("Falling back to Bubble map logic (if any)...")

    # 3. Fallback / Generic Bubble Map (if Choropleth failed or another region)
    
    # Load Base Map (World or Indonesia Boundary)
    world = None
    try:
        url = "https://raw.githubusercontent.com/datasets/geo-countries/master/data/countries.geojson"
        world = gpd.read_file(url)
    except Exception as e:
         print(f"Could not load remote base map: {e}")

    # Prepare Points
    points = []
    values = []
    
    # Use global INDO_COORDS / KALBAR_COORDS
    if 'Provinsi' in df_counts.columns:
        coords_db = INDO_COORDS
        label_col = 'Provinsi'
        val_col = 'Jumlah'
    else:
        coords_db = KALBAR_COORDS
        label_col = 'Kota/Kabupaten'
        val_col = 'Jumlah Responden'
        
    for index, row in df_counts.iterrows():
        name = row.get(label_col)
        val = row.get(val_col)
        
        # Check direct match or partial
        matched_coords = None
        if name in coords_db:
             matched_coords = coords_db[name]
        else:
             # Try simple normalization for Kalbar keys
             # e.g. "Kab. Kubu Raya" -> "Kubu Raya"
             for k, v in coords_db.items():
                 if k.lower() in name.lower() or name.lower() in k.lower():
                     matched_coords = v
                     break
        
        if matched_coords:
            lat, lon = matched_coords
            points.append(Point(lon, lat))
            values.append(val)
            
    if not points:
        print(f"No matching coordinates found for points map ({region_name}).")
        return

    gdf_points = gpd.GeoDataFrame({'value': values}, geometry=points, crs="EPSG:4326")

    # Plotting
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # Plot Base Map if available
    if world is not None:
        # Filter for Indonesia
        base = pd.DataFrame() # Initialize as empty
        potential_cols = ['ADMIN', 'name', 'common', 'NAME', 'sovereignt']
        
        for col in potential_cols:
            if col in world.columns:
                filtered = world[world[col] == 'Indonesia']
                if not filtered.empty:
                    base = filtered
                    break
        
        if not base.empty:
            base.plot(ax=ax, color='#f0f0f0', edgecolor='#888888')
        else:
            world.plot(ax=ax, color='#f0f0f0', edgecolor='#888888')

    # Zoom Limits
    if region_name == 'Kalimantan Barat':
        ax.set_xlim(108.0, 114.5)
        ax.set_ylim(-3.5, 2.5)
        ax.set_title(f"Sebaran Alumni - {region_name}", fontsize=16)
    else:
        ax.set_xlim(95, 141)
        ax.set_ylim(-11, 6)
        ax.set_title(f"Sebaran Alumni - {region_name}", fontsize=16)
    
    # Plot Points (Bubble)
    # Size based on value
    min_size = 50
    max_size = 1000
    
    if max(values) == min(values):
        sizes = [300] * len(values)
    else:
        sizes = [(v - min(values)) / (max(values) - min(values)) * (max_size - min_size) + min_size for v in values]
    
    # Color based on value (Blue Gradient)
    gdf_points.plot(ax=ax, 
                    column='value', # Use column for color mapping
                    cmap='Blues', 
                    markersize=sizes, 
                    alpha=0.7,
                    edgecolor='k',
                    linewidth=0.5,
                    legend=True)

    ax.axis('off')
    
    # Save PNG
    plt.tight_layout()
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"Static Geopandas Map saved: {output_path}")

def map_to_png(m, output_path):
    """
    Legacy/Deprecated or wrapper. 
    If we use this name to maintain compatibility with existing calls,
    we need to redirect or handle it.
    But better to call generate_static_map_geopandas explicitly.
    Emptying this for now to avoid selenium usage.
    """
    pass

def sort_crosstab_by_total(df_crosstab):
    """
    Sorts the crosstab DataFrame by the 'Total' column in descending order,
    keeping the 'Total' row (margin) at the bottom.
    Also adds a 'Persentase' column.
    """
    if 'Total' not in df_crosstab.columns:
        return df_crosstab

    # Separate 'Total' row if it exists based on index name
    if 'Total' in df_crosstab.index:
        total_row = df_crosstab.loc[['Total']]
        df_body = df_crosstab.drop('Total')
    else:
        # Fallback if checking index value directly
        total_row = pd.DataFrame()
        df_body = df_crosstab

    # Sort body by 'Total' column descending
    df_sorted = df_body.sort_values(by='Total', ascending=False)

    # Append 'Total' row back
    if not total_row.empty:
        df_final = pd.concat([df_sorted, total_row])
    else:
         df_final = df_sorted
    
    # Add Percentage Calculation
    if 'Total' in df_final.columns:
        # Determine grand total (denominator)
        if 'Total' in df_final.index:
             grand_total = df_final.loc['Total', 'Total']
        else:
             grand_total = df_final['Total'].sum()
        
        if grand_total > 0:
            # Calculate percentage
            pct = (df_final['Total'] / grand_total) * 100
            # Format
            df_final['Persentase'] = pct.map('{:.2f}%'.format)
        else:
            df_final['Persentase'] = "0.00%"
            
    return df_final

def create_distribution_campus_loc_tahun(df):
    """
    Creates a distribution table of respondents based on Lokasi Kampus (derived from prodi) and Tahun Lulus.
    
    Logic:
    - If 'prodi' contains "Kapuas Hulu" -> "Kapuas Hulu"
    - If 'prodi' contains "Sanggau" -> "Sanggau"
    - If 'prodi' contains "Sukamara" -> "Sukamara"
    - Else -> "Kampus Polnep"
    
    Args:
        df (pd.DataFrame): The input dataframe containing 'prodi' and 'Tahun Lulus' columns.
        
    Returns:
        pd.DataFrame: A cross-tabulation of Location vs. Tahun Lulus.
    """
    column_to_check = 'prodi'
    if column_to_check not in df.columns:
        # Fallback to 'Program Studi' if 'prodi' missing
        if 'Program Studi' in df.columns:
            column_to_check = 'Program Studi'
        else:
             raise ValueError(f"Missing columns: 'prodi' or 'Program Studi' not found.")

    year_col = 'Tahun Lulus'
    if year_col not in df.columns:
         raise ValueError(f"Missing column: {year_col}")

    def get_location(val):
        s_val = str(val)
        if 'Kapuas Hulu' in s_val:
            return 'PDD Kapuas Hulu'
        if 'Sanggau' in s_val:
             return 'PSDKU Sanggau'
        if 'Sukamara' in s_val:
            return 'PSDKU Sukamara'
        return 'Kampus Polnep'

    # Create a temporary column for location to group by
    temp_loc_col = 'Lokasi Kampus'
    df = df.copy() # Avoid SettingWithCopyWarning on original df if passed directly
    df[temp_loc_col] = df[column_to_check].apply(get_location)
    
    ct = pd.crosstab(df[temp_loc_col], df[year_col], margins=True, margins_name='Total')
    return sort_crosstab_by_total(ct)

def create_distribution_jurusan_tahun(df):
    """
    Creates a distribution table of respondents based on Jurusan and Tahun Lulus.
    
    Args:
        df (pd.DataFrame): The input dataframe containing 'Jurusan' and 'Tahun Lulus' columns.
        
    Returns:
        pd.DataFrame: A cross-tabulation of Jurusan vs. Tahun Lulus.
    """
    jurusan_col = 'Jurusan'
    year_col = 'Tahun Lulus'
    
    if jurusan_col not in df.columns or year_col not in df.columns:
        raise ValueError(f"Missing columns: {jurusan_col} or {year_col} not found in DataFrame.")
    
    ct = pd.crosstab(df[jurusan_col], df[year_col], margins=True, margins_name='Total')
    return sort_crosstab_by_total(ct)

def create_distribution_prodi_tahun(df):
    """
    Creates a distribution table of respondents based on Program Studi (prodi) and Tahun Lulus.
    
    Args:
        df (pd.DataFrame): The input dataframe containing 'prodi' and 'Tahun Lulus' columns.
        
    Returns:
        pd.DataFrame: A cross-tabulation of Prodi vs. Tahun Lulus.
    """
    prodi_col = 'prodi'
    if prodi_col not in df.columns:
         if 'Program Studi' in df.columns:
             prodi_col = 'Program Studi'
         else:
            raise ValueError(f"Missing columns: {prodi_col} not found in DataFrame.")
            
    year_col = 'Tahun Lulus'
    
    if year_col not in df.columns:
        raise ValueError(f"Missing columns: {year_col} not found in DataFrame.")
    
    ct = pd.crosstab(df[prodi_col], df[year_col], margins=True, margins_name='Total')
    return sort_crosstab_by_total(ct)

def create_distribution_masa_tunggu_status(df):
    """
    Creates a distribution table of Status Pekerjaan vs Kategori Masa Tunggu.
    """
    # 1. Seleksi dan Rename Kolom
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    
    # Check columns
    if col_status not in df.columns or col_masa_tunggu not in df.columns:
        # Try to be flexible if headers are slightly different or stripped
        print("Warning: Specific columns for Masa Tunggu not found exactly. Trying fuzzy match or skipping.")
        # Attempt to find closest match if needed, but for now rely on exact match as per request
        if col_status not in df.columns: return pd.DataFrame()
        if col_masa_tunggu not in df.columns: return pd.DataFrame()

    df_analysis = df[[col_status, col_masa_tunggu]].copy()
    df_analysis.columns = ['Status Pekerjaan', 'Masa_Tunggu_Bulan']

    # 2. Filter Data Valid
    df_analysis = df_analysis.dropna(subset=['Masa_Tunggu_Bulan'])
    df_analysis['Masa_Tunggu_Bulan'] = pd.to_numeric(df_analysis['Masa_Tunggu_Bulan'], errors='coerce')

    # 3. Fungsi Kategorisasi
    def kategorisasi_masa_tunggu(bulan):
        if pd.isna(bulan): return 'Unknown'
        if bulan < 3:
            return 'Kurang dari 3 Bulan'
        elif bulan <= 6:
            return '3 - 6 Bulan'
        elif bulan <= 12:
            return '6 - 12 Bulan'
        else:
            return 'Lebih dari 12 Bulan'

    df_analysis['Kategori_Masa_Tunggu'] = df_analysis['Masa_Tunggu_Bulan'].apply(kategorisasi_masa_tunggu)

    # 4. Membuat Pivot Table
    tabel_distribusi = pd.crosstab(
        df_analysis['Status Pekerjaan'],
        df_analysis['Kategori_Masa_Tunggu'],
        margins=True,
        margins_name='Total'
    )

    # 5. Mengurutkan kolom agar logis
    urutan_kolom = ['Kurang dari 3 Bulan', '3 - 6 Bulan', '6 - 12 Bulan', 'Lebih dari 12 Bulan', 'Total']
    col_ada = [c for c in urutan_kolom if c in tabel_distribusi.columns]
    
    tabel_final = tabel_distribusi[col_ada]
    
    # Apply sorting by Total desc (Rows) and add Percentage
    return sort_crosstab_by_total(tabel_final)

def create_distribution_waktu_tunggu_jurusan(df):
    """
    Creates a distribution table for Average Respondents Accepted Working within 6 months.
    """
    # 1. Preprocessing Data Masa Tunggu
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    
    # Check columns
    if col_masa_tunggu not in df.columns:
        print(f"Warning: Column '{col_masa_tunggu}' not found.")
        return pd.DataFrame()

    # 2. Filter: Hanya ambil responden yang mengisi masa tunggu DAN statusnya Bekerja
    col_status = 'Jelaskan status Anda saat ini?'
    target_status = 'Bekerja (Full time/Part time)'
    
    df_filtered = df.dropna(subset=[col_masa_tunggu]).copy()
    
    if col_status in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_status] == target_status]
    else:
        print(f"Warning: Column '{col_status}' not found. Cannot filter by status.")
        return pd.DataFrame()

    # Pastikan data numerik
    df_filtered['Masa_Tunggu_Bulan'] = pd.to_numeric(df_filtered[col_masa_tunggu], errors='coerce')

    # 3. Logika Perhitungan (< 6 Bulan)
    # Buat kolom helper: 1 jika <= 6 bulan, 0 jika > 6 bulan
    df_filtered['Is_Less_6_Months'] = df_filtered['Masa_Tunggu_Bulan'].apply(lambda x: 1 if x <= 6 else 0)

    # 4. Membuat Tabel Agregat (Group by Jurusan)
    col_group = 'Jurusan'
    if col_group not in df_filtered.columns:
        return pd.DataFrame()
        
    analisis_masa_tunggu = df_filtered.groupby(col_group).agg(
        Jumlah_Responden=('Masa_Tunggu_Bulan', 'count'),
        Jumlah_Kurang_6_Bulan=('Is_Less_6_Months', 'sum'),
        Rata_rata_Waktu_Tunggu=('Masa_Tunggu_Bulan', 'mean')
    ).reset_index()

    # 5. Menghitung Persentase
    analisis_masa_tunggu['Persentase_Kurang_6_Bulan'] = (
        analisis_masa_tunggu['Jumlah_Kurang_6_Bulan'] / analisis_masa_tunggu['Jumlah_Responden'] * 100
    ) #.round(2) -> will map later

    # Rounding rata-rata waktu tunggu
    analisis_masa_tunggu['Rata_rata_Waktu_Tunggu'] = analisis_masa_tunggu['Rata_rata_Waktu_Tunggu'].round(1)

    # Sort before renaming
    analisis_masa_tunggu = analisis_masa_tunggu.sort_values(by='Persentase_Kurang_6_Bulan', ascending=False)

    # 6. Formatting Tabel Akhir
    final_table = analisis_masa_tunggu[[
        'Jurusan', 
        'Jumlah_Responden', 
        'Jumlah_Kurang_6_Bulan', 
        'Persentase_Kurang_6_Bulan', 
        'Rata_rata_Waktu_Tunggu'
    ]].copy()

    # Rename kolom untuk laporan
    final_table.columns = [
        'Jurusan', 
        'Total Responden (Bekerja)', 
        'Jumlah Lulusan (<= 6 Bulan)', 
        'Persentase (<= 6 Bulan) (%)', 
        'Rata-rata Masa Tunggu (Bulan)'
    ]

    # Tambahkan Baris Total/Rata-rata Institusi
    total_responden = final_table['Total Responden (Bekerja)'].sum()
    jumlah_kurang_6 = final_table['Jumlah Lulusan (<= 6 Bulan)'].sum()
    
    avg_masa_tunggu = df_filtered['Masa_Tunggu_Bulan'].mean() if not df_filtered.empty else 0

    total_row = pd.DataFrame({
        'Jurusan': ['TOTAL / RATA-RATA INSTITUSI'],
        'Total Responden (Bekerja)': [total_responden],
        'Jumlah Lulusan (<= 6 Bulan)': [jumlah_kurang_6],
        'Persentase (<= 6 Bulan) (%)': [
            (jumlah_kurang_6 / total_responden * 100) if total_responden > 0 else 0
        ],
        'Rata-rata Masa Tunggu (Bulan)': [
            round(avg_masa_tunggu, 1)
        ]
    })
    
    final_table = pd.concat([final_table, total_row], ignore_index=True)
    
    # Format Percentage Column string
    final_table['Persentase (<= 6 Bulan) (%)'] = final_table['Persentase (<= 6 Bulan) (%)'].apply(lambda x: f"{x:.2f}%")
    
    return final_table

def create_serapan_jurusan(df):
    """
    Creates a crosstab of Jurusan vs Status Pekerjaan.
    """
    col_jurusan = 'Jurusan'
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_jurusan not in df.columns or col_status not in df.columns:
        return pd.DataFrame()
        
    tabel_jurusan = pd.crosstab(
        df[col_jurusan], 
        df[col_status], 
        margins=True, 
        margins_name='Total'
    )
    # Reuse our sorter if possible, it sorts by 'Total' column desc
    # This matches the user's general preference for sorting
    return sort_crosstab_by_total(tabel_jurusan)

def create_serapan_prodi_per_jurusan(df):
    """
    Creates a dictionary of tables, one per Jurusan.
    Each table shows [Prodi] vs Status Pekerjaan.
    Returns: dict { "Jurusan Name": pd.DataFrame }
    """
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi'
    if col_prodi not in df.columns and 'Program Studi' in df.columns:
        col_prodi = 'Program Studi'
        
    col_status = 'Jelaskan status Anda saat ini?'
    
    if col_jurusan not in df.columns or col_prodi not in df.columns or col_status not in df.columns:
        return {}

    # Clean/Rename Schema for Display
    df_clean = df.copy()
    status_map = {
        "Bekerja (Full time/Part time)": "Bekerja",
        "Wiraswasta": "Wiraswasta", 
        "Tidak kerja tetapi sedang mencari kerja": "Sedang Mencari Kerja",
        "Melanjutkan Pendidikan": "Studi Lanjut",
        "Belum memungkinkan bekerja": "Belum Memungkinkan Bekerja",
        "Tidak kerja tetapi tidak mencari kerja": "Tidak Mencari Kerja" 
    }
    # Apply mapping
    df_clean[col_status] = df_clean[col_status].replace(status_map)

    # 1. Create Crosstab
    ct = pd.crosstab(
        [df_clean[col_jurusan], df_clean[col_prodi]],
        df_clean[col_status]
    )
    
    # Prepare Columns Order
    preferred_order = [
         "Bekerja", "Wiraswasta", "Sedang Mencari Kerja", "Studi Lanjut", "Belum Memungkinkan Bekerja"
    ]
    existing_cols = ct.columns.tolist()
    sorted_cols = []
    
    for col in preferred_order:
        if col in existing_cols:
            sorted_cols.append(col)
    
    for col in existing_cols:
        if col not in sorted_cols:
            sorted_cols.append(col)
            
    # Add Total column to list (calculated later)
    columns = sorted_cols + ['Total']
    
    # Get unique Jurusans
    unique_jurusans = ct.index.get_level_values(0).unique()
    
    results = {}
    
    for jurusan in unique_jurusans:
        try:
            # sub_df index is Prodi
            sub_df = ct.loc[jurusan].copy()
        except KeyError:
            continue
            
        # Ensure all columns exist
        for col in sorted_cols:
            if col not in sub_df.columns:
                sub_df[col] = 0
                
        # Reorder columns
        sub_df = sub_df[sorted_cols]
        
        # Calculate Total per row
        sub_df['Total'] = sub_df.sum(axis=1)
        
        # Calculate Grand Total for this Jurusan
        jurusan_total_row = sub_df.sum(axis=0)
        jurusan_grand_total = jurusan_total_row['Total']
        
        # Prepare Rows
        final_rows = []
        
        # Add Prodi Rows
        for prodi, row_data in sub_df.iterrows():
            row_dict = {'Program Studi': prodi}
            for col in columns:
                val = row_data[col]
                row_dict[col] = val
            
            # Percentage based on Jurusan Total
            if jurusan_grand_total > 0:
                pct = (row_data['Total'] / jurusan_grand_total) * 100
                row_dict['Persentase'] = f"{pct:.2f}%"
            else:
                row_dict['Persentase'] = "0.00%"
                
            final_rows.append(row_dict)
            
        # Add Total Row
        total_dict = {'Program Studi': f'Total {jurusan}'}
        for col in columns:
            val = jurusan_total_row[col]
            total_dict[col] = val
        
        # Percentage for Total Row is 100%
        total_dict['Persentase'] = "100.00%"
        
        final_rows.append(total_dict)
        
        # Create DataFrame
        df_jurusan = pd.DataFrame(final_rows)
        # Reorder columns
        final_cols_order = ['Program Studi'] + columns + ['Persentase']
        df_jurusan = df_jurusan[final_cols_order]
        
        results[jurusan] = df_jurusan
        
    return results

# Database Koordinat removed from here


def create_distribution_provinsi(df):
    """
    Creates a distribution table of working respondents by Province.
    """
    col_status = 'Jelaskan status Anda saat ini?'
    col_prov = 'Provinsi rev'
    
    if col_prov not in df.columns:
        if 'Provinsi' in df.columns:
            col_prov = 'Provinsi'
        else:
             return pd.DataFrame()
             
    # Filter Responden yang Bekerja/Wiraswasta
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    # Check if status column exists
    if col_status in df.columns:
        df_working = df[df[col_status].isin(working_status)].copy()
    else:
        df_working = df.copy() # Fallback if status not found? Or return empty
    
    if df_working.empty:
        return pd.DataFrame()
        
    # Hitung Jumlah per Provinsi
    prov_counts = df_working[col_prov].value_counts().reset_index()
    prov_counts.columns = ['Provinsi', 'Jumlah']
    
    # Add coordinates for reference if needed, but for table display, just simple is fine
    # Maybe add Percentage?
    total = prov_counts['Jumlah'].sum()
    if total > 0:
        prov_counts['Persentase'] = (prov_counts['Jumlah'] / total * 100).map('{:.2f}%'.format)
        
    # Add Total Row
    total_row = pd.DataFrame({'Provinsi': ['TOTAL'], 'Jumlah': [total], 'Persentase': ['100.00%']})
    prov_counts = pd.concat([prov_counts, total_row], ignore_index=True)
    
    return prov_counts


def get_density_color(count, min_val, max_val):
    """
    Returns a blue gradient hex color based on the count.
    Light Blue -> Dark Blue
    """
    # Normalize count between 0 and 1
    if max_val == min_val:
        norm = 1.0
    else:
        norm = (count - min_val) / (max_val - min_val)
    
    # Simple Bucket approach for clearer distinction or Matplotlib colormap?
    # User asked for "darkblue" for max.
    # Let's use matplotlib colors if imported, or custom hex interpolation.
    import matplotlib.colors as mcolors
    
    # Create colormap from LightBlue to DarkBlue
    cmap = mcolors.LinearSegmentedColormap.from_list("blue_density", ["#87CEEB", "#00008B"])
    
    # Get hex code
    hex_color = mcolors.to_hex(cmap(norm))
    return hex_color

def generate_alumni_map(prov_counts_df, output_file):
    """
    Generates a Folium map based on province counts.
    """
    if folium is None:
        print("Folium not installed, skipping map generation.")
        return

    # Filter out Total row for mapping
    df_map = prov_counts_df[prov_counts_df['Provinsi'] != 'TOTAL'].copy()
    
    # Calculate Min/Max for Color Scaling
    max_count = df_map['Jumlah'].max()
    min_count = df_map['Jumlah'].min()
    
    # 5. Membuat Peta Dasar (Fokus di Indonesia)
    m = folium.Map(location=[-2.5, 118], zoom_start=5, tiles='CartoDB positron')

    # 6. Menambahkan Marker ke Peta
    for index, row in df_map.iterrows():
        prov_name = row['Provinsi']
        count = row['Jumlah']
        
        # Cek apakah provinsi ada di database koordinat
        if prov_name in INDO_COORDS:
            lat, lon = INDO_COORDS[prov_name]
            
            # Menentukan ukuran lingkaran berdasarkan jumlah alumni
            # Radius dasar 5, ditambah faktor skala
            radius = 5 + (count / 5) 
            
            # Get Color based on density
            color = get_density_color(count, min_count, max_count)

            folium.CircleMarker(
                location=[lat, lon],
                radius=radius,
                popup=f"<b>{prov_name}</b><br>Jumlah Alumni: {count}",
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.7
            ).add_to(m)
            
            # Opsional: Tambahkan Label Angka Permanen untuk provinsi padat
            if count > 10: 
                folium.Marker(
                    location=[lat, lon],
                    icon=folium.DivIcon(html=f"""<div style="font-family: courier new; color: black; font-weight: bold">{count}</div>""")
                ).add_to(m)
                
    m.save(output_file)
    print(f"Map generated: {output_file}")
    
    # Save as PNG using Geopandas
    try:
        png_output = output_file.replace('reports', 'assets/gambar').replace('.html', '.png')
        if gpd:
             generate_static_map_geopandas(df_map, png_output, region_name='Indonesia')
    except Exception as e:
        print(f"Error saving PNG map: {e}")


def create_distribution_kabkota_kalbar(df):
    """
    Creates a distribution table for Kota/Kabupaten in Kalimantan Barat.
    """
    col_prov = 'Provinsi rev'
    col_city = 'Kota/Kabupate rev'
    
    if col_prov not in df.columns or col_city not in df.columns:
        return pd.DataFrame()
        
    # Filter for Kalimantan Barat
    # The user implied "Sebaran khusus", usually implies "Working" respondents too if consistent with others?
    # The example total 504 matches the user's request. 
    # Let's filter by working first to check if it matches 504.
    # Actually, previous table "Table 8" total was 556 working respondents. 
    # If Kalbar is 504, then likely it is filtered by working.
    
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    df_filtered = df.copy()
    
    # User Request: "seharusnya sebaran ini adalah yang sudah bekerja saja"
    # Ensure strict filtering
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    else:
        # If status column missing, return empty or warn? For now assume it exists
        pass
        
    df_kalbar = df_filtered[df_filtered[col_prov] == 'Kalimantan Barat'].copy()
    
    if df_kalbar.empty:
        return pd.DataFrame()

    # Count by City
    city_counts = df_kalbar[col_city].value_counts().reset_index()
    city_counts.columns = ['Kota/Kabupaten', 'Jumlah Responden']
    
    # Calculate Percentage
    total = city_counts['Jumlah Responden'].sum()
    if total > 0:
        city_counts['Persentase (%)'] = (city_counts['Jumlah Responden'] / total * 100).map('{:.2f}%'.format)
    else:
        city_counts['Persentase (%)'] = '0.00%'

    # Sort Descending (already done by value_counts, but good to ensure)
    city_counts = city_counts.sort_values(by='Jumlah Responden', ascending=False)
    
    # Add Total Row
    total_row = pd.DataFrame({
        'Kota/Kabupaten': ['Total Kalbar'], 
        'Jumlah Responden': [total], 
        'Persentase (%)': ['100.00%']
    })
    
    final_table = pd.concat([city_counts, total_row], ignore_index=True)
    
    return final_table

def create_salary_distribution(df):
    """
    Creates a distribution table for Salary/Income of working respondents.
    """
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_salary not in df.columns:
        return pd.DataFrame()
        
    df_filtered = df.copy()
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
        
    salary_counts = df_filtered[col_salary].value_counts().reset_index()
    salary_counts.columns = ['Rata-rata Pendapatan', 'Jumlah Responden']
    
    # Custom sort order for salary categories
    order = [
        '< Rp. 1.000.000',
        'Rp. 1.000.001 - Rp. 2.000.000',
        'Rp. 2.000.001 - Rp. 3.000.000',
        'Rp. 3.000.001 - Rp. 4.000.000',
        'Rp. 4.000.001 - Rp. 5.000.000',
        'Rp. 5.000.001 - Rp. 6.000.000',
        'Rp. 6.000.001 - Rp. 7.000.000',
        'Rp. 7.000.001 - Rp. 8.000.000',
        '> Rp. 8.000.001'
    ]
    
    salary_counts['Rata-rata Pendapatan'] = pd.Categorical(salary_counts['Rata-rata Pendapatan'], categories=order, ordered=True)
    salary_counts = salary_counts.sort_values('Rata-rata Pendapatan').reset_index(drop=True)
    
    # Calculate Percentage
    total = salary_counts['Jumlah Responden'].sum()
    if total > 0:
        salary_counts['Persentase (%)'] = (salary_counts['Jumlah Responden'] / total * 100).map('{:.2f}%'.format)
    else:
        salary_counts['Persentase (%)'] = '0.00%'
        
    # Add Total Row
    total_row = pd.DataFrame({
        'Rata-rata Pendapatan': ['Total'], 
        'Jumlah Responden': [total], 
        'Persentase (%)': ['100.00%']
    })
    
    final_table = pd.concat([salary_counts, total_row], ignore_index=True)
    
    return final_table

def create_salary_by_jurusan(df):
    """
    Calculates the Average salary per Jurusan using custom range conversions.
    """
    col_salary = 'Berapa rata-rata pendapatan Anda per bulan?'
    col_jurusan = 'Jurusan'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_salary not in df.columns or col_jurusan not in df.columns:
        return pd.DataFrame()
        
    df_filtered = df.copy()
    if col_status in df.columns:
        df_filtered = df_filtered[df_filtered[col_status].isin(working_status)]
    
    # Drop rows where salary is missing
    df_filtered = df_filtered.dropna(subset=[col_salary])
    
    if df_filtered.empty:
        return pd.DataFrame()

    # User Defined Mappings for Mean Calculation
    salary_map = {
        '< Rp. 1.000.000': 1000000,
        'Rp. 1.000.001 - Rp. 2.000.000': 1500000,
        'Rp. 2.000.001 - Rp. 3.000.000': 2500000,
        'Rp. 3.000.001 - Rp. 4.000.000': 3500000,
        'Rp. 4.000.001 - Rp. 5.000.000': 4500000,
        'Rp. 5.000.001 - Rp. 6.000.000': 5500000,
        'Rp. 6.000.001 - Rp. 7.000.000': 6500000,
        'Rp. 7.000.001 - Rp. 8.000.000': 7500000,
        '> Rp. 8.000.001': 8000000
    }

    # Map categories to numeric
    df_filtered['salary_num'] = df_filtered[col_salary].map(salary_map)
    
    # Calculate Mean per Jurusan
    salary_by_jurusan = df_filtered.groupby(col_jurusan)['salary_num'].mean().reset_index()
    salary_by_jurusan.columns = ['Jurusan', 'Rata-rata Gaji (Estimasi)']
    
    # Sort by numeric mean descending
    salary_by_jurusan = salary_by_jurusan.sort_values(by='Rata-rata Gaji (Estimasi)', ascending=False)
    
    # Store numeric version for chart
    df_chart = salary_by_jurusan.copy()
    df_chart = df_chart.rename(columns={'Rata-rata Gaji (Estimasi)': 'Total'})
    
    # Format for display table: Compact format (e.g., Rp3.8 juta)
    salary_by_jurusan['Rata-rata Gaji (Estimasi)'] = salary_by_jurusan['Rata-rata Gaji (Estimasi)'].apply(
        lambda x: f"Rp{x/1000000:.1f} juta"
    )
    
    return salary_by_jurusan, df_chart

def create_jurusan_ranking():
    """
    Creates a dataframe for Jurusan Ranking based on provided qualitative/quantitative analysis.
    """
    data = [
        {
            "Peringkat": 1,
            "Jurusan": "Teknik Arsitektur",
            "Skor Kekuatan": "⭐⭐⭐⭐",
            "Predikat & Analisis": """<b>"The High Quality Performer"</b><br>Meskipun jumlah respondennya paling sedikit (52), namun jurusan ini JUARA 1 di dua kategori sekaligus: Kecepatan Serapan (77,5%) dan Gaji Tertinggi (Rp 3,8 Juta). Kualitas lulusannya sangat premium di mata pasar."""
        },
        {
            "Peringkat": 2,
            "Jurusan": "Akuntansi",
            "Skor Kekuatan": "⭐⭐⭐⭐",
            "Predikat & Analisis": """<b>"The Major Contributor"</b><br>Naik signifikan ke peringkat 2 berkat Volume Responden Tertinggi (226 orang) yang mendominasi 30% data survei. Selain itu, kecepatan serapannya sangat baik (Juara 2). Nilai minus hanya pada gaji awal yang masih entry level."""
        },
        {
            "Peringkat": 3,
            "Jurusan": "Teknik Sipil & Perencanaan",
            "Skor Kekuatan": "⭐⭐⭐",
            "Predikat & Analisis": """<b>"The Fastest Hired"</b><br>Unggul mutlak di Masa Tunggu Tercepat (3,4 bulan). Sangat efisien dalam mengantarkan lulusan ke dunia kerja, meskipun volume responden dan gaji berada di level menengah."""
        },
        {
            "Peringkat": 4,
            "Jurusan": "Teknik Mesin",
            "Skor Kekuatan": "⭐⭐⭐",
            "Predikat & Analisis": """<b>"The High Valued Entrepreneur"</b><br>Unggul di Gaji Tertinggi (Rp 3,8 Juta) dan jumlah responden yang besar (110 orang). Peringkatnya tertahan karena masa tunggu rata-rata yang cukup lama (6,4 bulan), namun ini terkompensasi oleh tingginya angka wirausaha."""
        },
        {
            "Peringkat": 5,
            "Jurusan": "Ilmu Kelautan & Perikanan",
            "Skor Kekuatan": "⭐⭐",
            "Predikat & Analisis": """<b>"The Balanced Niche"</b><br>Memiliki performa yang seimbang di semua lini. Tidak terlalu menonjol di satu sisi, tapi cukup stabil dalam serapan (71,8%) dan masa tunggu (5,2 bulan)."""
        },
        {
            "Peringkat": 6,
            "Jurusan": "Teknik Elektro",
            "Skor Kekuatan": "⭐⭐",
            "Predikat & Analisis": """<b>"The Steady Player"</b><br>Konsisten di papan tengah. Memiliki gaji yang cukup baik (Rp 3,5 Juta) di atas rata-rata institusi, namun butuh peningkatan dalam kecepatan serapan (67,5%)."""
        },
        {
            "Peringkat": 7,
            "Jurusan": "Administrasi Bisnis",
            "Skor Kekuatan": "⭐⭐",
            "Predikat & Analisis": """<b>"The Salary Surprise"</b><br>Meskipun secara ranking umum ada di bawah, jurusan ini punya keunggulan Gaji Tinggi (Rp 3,7 Juta - Peringkat 3). Tantangannya ada pada masa tunggu yang paling lama (6,4 bulan) dan volume responden yang moderat."""
        },
        {
            "Peringkat": 8,
            "Jurusan": "Teknologi Pertanian",
            "Skor Kekuatan": "⭐",
            "Predikat & Analisis": """<b>"The Job Creator"</b><br>Secara statistik "pekerja", jurusan ini ada di bawah (gaji & kecepatan rendah). NAMUN, perlu dicatat: Jurusan ini adalah Raja Wirausaha (21 orang). Indikator ranking ini bias ke "karyawan", sehingga potensi wirausaha Pertanian tidak terpotret penuh di sini."""
        }
    ]
    
    df = pd.DataFrame(data)
    return df


def generate_kalbar_map(city_counts_df, output_file):
    """
    Generates a Folium map for West Kalimantan (Kalbar) distribution.
    Uses 'CartoDB positron' tiles to match the theme.
    """
    if folium is None:
        print("Folium not installed, skipping Kalbar map generation.")
        return

    # Filter out Total row
    df_map = city_counts_df[city_counts_df['Kota/Kabupaten'] != 'Total Kalbar'].copy()
    
    # Calculate Min/Max for Color Scaling
    max_count = df_map['Jumlah Responden'].max()
    min_count = df_map['Jumlah Responden'].min()

    # Center map on West Kalimantan (approx)
    m = folium.Map(location=[0.0, 111.0], zoom_start=7, tiles='CartoDB positron')

    for index, row in df_map.iterrows():
        city_name = row['Kota/Kabupaten']
        count = row['Jumlah Responden']
        
        if city_name in KALBAR_COORDS:
            lat, lon = KALBAR_COORDS[city_name]
            
            # Radius calculation
            radius = 5 + (count / 3) # Slightly larger scale for cities
            
            # Get Color based on density
            color = get_density_color(count, min_count, max_count)

            folium.CircleMarker(
                location=[lat, lon],
                radius=radius,
                popup=f"<b>{city_name}</b><br>Jumlah Alumni: {count}",
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.8
            ).add_to(m)
            
            # Label
            if count > 5:
                folium.Marker(
                    location=[lat, lon],
                    icon=folium.DivIcon(html=f"""<div style="font-family: courier new; color: black; font-weight: bold; font-size: 10pt">{count}</div>""")
                ).add_to(m)
                
    m.save(output_file)
    print(f"Kalbar Map generated: {output_file}")

    # Save as PNG using Geopandas
    try:
        png_output = output_file.replace('reports', 'assets/gambar').replace('.html', '.png')
        if gpd:
             # Extract total from original df (before filtering out Total row) or recalculate
             # city_counts_df has 'Total Kalbar' row
             total_row = city_counts_df[city_counts_df['Kota/Kabupaten'] == 'Total Kalbar']
             total_val = None
             if not total_row.empty:
                 total_val = total_row['Jumlah Responden'].values[0]
                 
             generate_static_map_geopandas(df_map, png_output, region_name='Kalimantan Barat', total_reference=total_val)
    except Exception as e:
        print(f"Error saving Kalbar PNG map: {e}")


def print_styled_table(df, title=None):
    """
    Prints a pandas DataFrame in a styled format.
    Uses 'tabulate' if available, otherwise falls back to a custom ASCII implementation.
    """
    if title:
        print(f"\n[{title}]")

    try:
        from tabulate import tabulate
        # 'psql' format looks like MySQL/PostgreSQL output, very readable
        try:
            print(tabulate(df, headers='keys', tablefmt='psql', showindex=True))
        except UnicodeEncodeError:
            # Fallback: Robustly strip ALL non-ascii chars for console print
            print("(Console output contains special characters, stripping for display...)")
            df_safe = df.copy()
            for col in df_safe.columns:
                 # Force string, encode to ascii with replace, decode back to string
                 df_safe[col] = df_safe[col].apply(lambda x: str(x).encode('ascii', 'ignore').decode('ascii'))
            print(tabulate(df_safe, headers='keys', tablefmt='psql', showindex=True))
    except ImportError:
        # Fallback implementation if tabulate is not installed
        # Calculate column widths
        # Reset index to include the index in the table columns for printing
        df_print = df.reset_index()
        
        # Convert all data to string
        df_print = df_print.astype(str)
        
        columns = [str(c) for c in df_print.columns.tolist()]
        data = df_print.values.tolist()
        
        # Calculate max width for each column
        col_widths = []
        for i, col in enumerate(columns):
            max_len = len(col)

            for row in data:
                if len(row[i]) > max_len:
                    max_len = len(row[i])
            col_widths.append(max_len + 2) # +2 for padding
            
        # Function to create a separator line
        def create_separator(chars="-", junction="+"):
            line = junction
            for w in col_widths:
                line += chars * w + junction
            return line
            
        # Print the table
        border = create_separator("-", "+")
        header_row = "|"
        for i, col in enumerate(columns):
             header_row += f" {col:<{col_widths[i]-1}}|"
             
        print(border)
        print(header_row)
        print(border)
        
        for row in data:
            row_str = "|"
            for i, val in enumerate(row):
                row_str += f" {val:<{col_widths[i]-1}}|"
            print(row_str)
            
        print(border)
        print("(Note: Install 'tabulate' for even prettier tables: pip install tabulate)")

import matplotlib.pyplot as plt
import io
import base64
import numpy as np
import matplotlib.colors as mcolors

def get_horizontal_bar_chart_base64(df, title):
    """
    Generates a horizontal bar chart from the dataframe and returns it as a base64 string.
    Theme: Gradient from DarkBlue to DarkGrey.
    """
    # Create a copy to avoid modifying original
    df_plot = df.copy()
    
    # 1. Prepare Data
    # Remove 'Total' row if it exists
    if 'Total' in df_plot.index:
        df_plot = df_plot.drop('Total')
        
    # Sort by 'Total' column descending so largest bars are at top
    if 'Total' in df_plot.columns:
        df_plot = df_plot.sort_values(by='Total', ascending=True) # Ascending for barh (bottom to top)
        values = df_plot['Total']
    else:
        # Fallback if no Total column (shouldn't happen with our data)
        values = df_plot.iloc[:, -1]
        
    labels = df_plot.index.astype(str)
    
    # 2. Create Gradient Colors
    # Define gradient: DarkBlue (#00008B) to DarkGrey (#A9A9A9)
    n_bars = len(values)
    colors = []
    if n_bars > 0:
        cmap = mcolors.LinearSegmentedColormap.from_list("my_gradient", ["#A9A9A9", "#00008B"])
        # Generate colors based on value magnitude (normalized)
        # Or just simple gradient from top to bottom?
        # Let's do magnitude based gradient (darker = larger value)
        norm = plt.Normalize(values.min(), values.max())
        colors = [cmap(norm(v)) for v in values]
    
    # 3. Plotting
    fig, ax = plt.subplots(figsize=(10, max(6, n_bars * 0.4))) # Dynamic height
    bars = ax.barh(labels, values, color=colors, edgecolor='none')
    
    # 4. Styling
    ax.set_title(f"Grafik: {title}", fontsize=14, fontweight='bold', pad=20, color='#2c3e50')
    ax.set_xlabel("Jumlah Responden", fontsize=10, color='#555')
    
    # Remove top and right spines
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#ccc')
    ax.spines['bottom'].set_color('#ccc')
    
    # Add grid on x-axis
    ax.xaxis.grid(True, linestyle='--', alpha=0.6, color='#ddd')
    ax.set_axisbelow(True)
    
    # Add value labels at the end of bars
    for i, bar in enumerate(bars):
        width = bar.get_width()
        ax.text(width + (max(values)*0.01), bar.get_y() + bar.get_height()/2, 
                f'{int(width)}', 
                va='center', fontsize=9, color='#333', fontweight='bold')
                
    plt.tight_layout()
    
    # 5. Save to Base64 (High Resolution)
    buffer = io.BytesIO()
    # Increase DPI for high resolution (e.g. 300)
    plt.savefig(buffer, format='png', bbox_inches='tight', dpi=300)
    plt.close(fig)
    buffer.seek(0)
    image_base64 = base64.b64encode(buffer.read()).decode('utf-8')
    return image_base64

def generate_html_report(data_dict, output_file='report_tables.html'):
    """
    Generates a beautiful HTML report from a dictionary.
    data_dict structure: { "Title": [dataframe, chart_base64_string] }
    """
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Laporan Tracer Study</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
        <!-- Load html2canvas for taking screenshots of tables -->
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
        <style>
            body {
                font-family: 'Inter', sans-serif;
                background-color: #f4f6f9;
                color: #333;
                margin: 0;
                padding: 40px;
            }
            .container {
                max_width: 1200px;
                margin: 0 auto;
                background: #fff;
                padding: 40px;
                border-radius: 12px;
                box-shadow: 0 4px 20px rgba(0,0,0,0.05);
            }
            .section {
                margin-bottom: 60px;
                page-break-inside: avoid;
                position: relative;
            }
            h1 {
                text-align: center;
                color: #2c3e50;
                margin-bottom: 40px;
                font-weight: 600;
            }
            h2 {
                color: #34495e;
                margin-top: 0;
                margin-bottom: 20px;
                border-bottom: 2px solid #ecf0f1;
                padding-bottom: 10px;
                display: flex;
                justify-content: space-between;
                align-items: center;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 10px; /* Reduced to make room for button */
                background: #fff;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            }
            th, td {
                padding: 12px 15px;
                text-align: left;
                border-bottom: 1px solid #edf2f7;
            }
            th {
                background-color: #3498db;
                color: #fff;
                font-weight: 600;
                text-transform: uppercase;
                font-size: 0.85rem;
                letter-spacing: 0.5px;
            }
            tr:last-child td {
                border-bottom: none;
            }
            tr:hover {
                background-color: #f8fafc;
            }
            /* Zebra striping */
            tr:nth-child(even) {
                background-color: #f9fbfd;
            }
            
            /* Total Row Highlight */
            tr.total-row td {
                font-weight: bold;
                background-color: #e2e8f0;
                border-top: 2px solid #cbd5e0;
            }
            
            /* Columns Styling (Not bold by default) */
            th:last-child {
                border-left: 1px solid #e2e8f0;
                background-color: #3498db; 
                color: #fff;
            }
            td:last-child {
                border-left: 1px solid #f1f5f9;
                background-color: #f8fafc;
            }
            
             /* Intersection of Total Row and Column (Optional) */
            tr.total-row td:last-child {
                background-color: #cbd5e0;
                color: #1a202c;
            }

            .chart-container {
                margin-top: 30px;
                text-align: center;
                border: 1px solid #edf2f7;
                padding: 20px;
                border-radius: 8px;
                background-color: #fff;
            }
            .chart-img {
                max-width: 100%;
                height: auto;
                border-radius: 4px;
            }

            .btn-group {
                text-align: right;
                margin-bottom: 10px;
            }
            
            .btn {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 0.85rem;
                font-family: inherit;
                font-weight: 600;
                transition: background-color 0.2s;
                text-decoration: none;
                display: inline-block;
                margin-left: 10px;
            }
            
            .btn:hover {
                background-color: #2980b9;
            }
            
            .btn-secondary {
                background-color: #95a5a6;
            }
             .btn-secondary:hover {
                background-color: #7f8c8d;
            }

            .footer {
                text-align: center;
                margin-top: 50px;
                color: #7f8c8d;
                font-size: 0.9rem;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Laporan Responden Tracer Study</h1>
    """
    
    section_id = 0
    for title, content_tuple in data_dict.items():
        # Unpack based on length
        if len(content_tuple) == 3:
            df, chart_base64, map_path = content_tuple
        else:
            df, chart_base64 = content_tuple
            map_path = None
            
        section_id += 1
        safe_title = title.replace(" ", "_").lower()
        table_id = f"table_{section_id}"
        
        html_content += f'<div class="section">'
        html_content += f"<h2>{title}</h2>"
        
        # --- Table Section ---
        html_content += f'<div class="btn-group">'
        html_content += f'<button class="btn" onclick="saveTable(\'{table_id}\', \'{title}_table\')">Simpan Tabel</button>'
        html_content += '</div>'
        
        # reset_index to ensure the index part (like Lokasi Kampus) is a proper column
        if df.index.name:
             df_to_html = df.reset_index()
        else:
             df_to_html = df.copy()
        
        # Fix: Clear the columns name
        df_to_html.columns.name = None
        
        # Convert to HTML without default border attribute
        # Add ID for html2canvas
        table_html = df_to_html.to_html(index=False, border=0, classes='table', table_id=table_id, escape=False)
        # Pandas to_html doesn't support table_id directly in older versions, so let's inject it via string replacement if needed
        # Actually it does support table_id in newer versions, but let's be safe.
        if f'id="{table_id}"' not in table_html:
             table_html = table_html.replace('<table', f'<table id="{table_id}"')
             
        html_content += table_html
        
        # --- Chart Section ---
        if chart_base64:
            html_content += f"""
            <div class="chart-container">
                <div class="btn-group" style="text-align: right;">
                    <a href="data:image/png;base64,{chart_base64}" download="{title}_chart.png" class="btn btn-secondary">Simpan Grafik (High Res)</a>
                </div>
                <img src="data:image/png;base64,{chart_base64}" alt="Chart for {title}" class="chart-img">
            </div>
            """
            
            """
        
        # --- Map Section --- 
        # User requested to REMOVE maps from HTML report (Step 908)
        # if map_path:
        #     html_content += f"""
        #     <div class="chart-container" style="text-align: left;">
        #         <h3>Peta Interaktif</h3>
        #         <iframe src="{map_path}" width="100%" height="500" style="border:none;"></iframe>
        #         <div style="margin-top: 10px; text-align: right;">
        #             <a href="{map_path}" target="_blank" class="btn btn-secondary">Buka Peta Fullscreen</a>
        #         </div>
        #     </div>
        #     """
        
        html_content += '</div>'

    html_content += """
            <div class="footer">
                <p>Generated by Tracer Study Analysis Tool</p>
            </div>
        </div>

        <script>
            function saveTable(tableId, filename) {
                const table = document.getElementById(tableId);
                
                // Add some padding/background for the screenshot
                const originalBg = table.style.backgroundColor;
                table.style.backgroundColor = "white";
                
                html2canvas(table, {
                    scale: 3, // High resolution screenshot
                    backgroundColor: "#ffffff",
                    logging: false
                }).then(canvas => {
                    // Restore original style
                    table.style.backgroundColor = originalBg;
                    
                    // Create download link
                    const link = document.createElement('a');
                    link.download = filename + '.png';
                    link.href = canvas.toDataURL("image/png");
                    link.click();
                 }).catch(err => {
                    alert("Error saving table: " + err);
                });
            }

            // Automatically bold rows that contain the word "Total"
            document.addEventListener("DOMContentLoaded", function() {
                const rows = document.querySelectorAll("tr");
                rows.forEach(row => {
                    // Check first cell or any cell for "Total"
                    if (row.innerText.toLowerCase().includes("total")) {
                        row.classList.add("total-row");
                    }
                });
            });
        </script>
    </body>
    </html>
    """

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    # Try to open the file automatically (optional, works on Windows)
    import webbrowser
    try:
        webbrowser.open('file://' + os.path.abspath(output_file))
    except:
        pass
        
    print(f"Report generated successfully: {output_file}")


if __name__ == "__main__":
    import os
    print("--- Running table_jml_responden.py ---")
    
    # Define paths
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DATA_CLEANED = os.path.join(BASE_DIR, 'data', 'processed', 'cleaned_data.xlsx')
    REPORTS_DIR = os.path.join(BASE_DIR, 'reports')
    REPORT_OUTPUT = os.path.join(REPORTS_DIR, 'report_tables.html')

    file_path = DATA_CLEANED
    if not os.path.exists(file_path):
        print(f"{file_path} not found, trying data.xlsx")
        DATA_RAW = os.path.join(BASE_DIR, 'data', 'raw', 'data.xlsx')
        file_path = DATA_RAW
        
    try:
        print(f"Loading data from {file_path}...")
        df_load = pd.read_excel(file_path)
        
        # Calculate dataframes
        df_campus = create_distribution_campus_loc_tahun(df_load)
        df_jurusan = create_distribution_jurusan_tahun(df_load)
        df_prodi = create_distribution_prodi_tahun(df_load)
        
        # Print to console (using our styled printer)
        print_styled_table(df_campus, "Table 1: Lokasi Kampus vs Tahun Lulus")
        print_styled_table(df_jurusan, "Table 2: Jurusan vs Tahun Lulus")
        print_styled_table(df_prodi, "Table 3: Program Studi vs Tahun Lulus")
        
        # Generate Charts
        print("\nGenerating Charts...")
        chart_campus = get_horizontal_bar_chart_base64(df_campus, "Lokasi Kampus")
        chart_jurusan = get_horizontal_bar_chart_base64(df_jurusan, "Jurusan")
        chart_prodi = get_horizontal_bar_chart_base64(df_prodi, "Program Studi")

        # Generate HTML Report
        print("\nGenerating HTML report...")
        # Passing tuple (dataframe, chart_image)
        dfs_to_report = {
            "Distribusi Berdasarkan Lokasi Kampus": (df_campus, chart_campus),
            "Distribusi Berdasarkan Jurusan": (df_jurusan, chart_jurusan),
            "Distribusi Berdasarkan Program Studi": (df_prodi, chart_prodi)
        }
        
        # New Table: Masa Tunggu
        df_masa_tunggu = create_distribution_masa_tunggu_status(df_load)
        if not df_masa_tunggu.empty:
            print_styled_table(df_masa_tunggu, "Table 4: Status Pekerjaan vs Masa Tunggu")
            chart_masa_tunggu = get_horizontal_bar_chart_base64(df_masa_tunggu, "Status Pekerjaan vs Masa Tunggu")
            dfs_to_report["Distribusi Masa Tunggu Responden"] = (df_masa_tunggu, chart_masa_tunggu)
            
        # New Table: Rata-rata Waktu Tunggu per Jurusan
        df_waktu_tunggu = create_distribution_waktu_tunggu_jurusan(df_load)
        if not df_waktu_tunggu.empty:
            print_styled_table(df_waktu_tunggu, "Table 5: Rata-rata Masa Tunggu Lulusan per Jurusan")
            # No chart requested for this yet, pass None
            dfs_to_report["Rata-rata Masa Tunggu Lulusan per Jurusan"] = (df_waktu_tunggu, None)

        # New Table: Serapan per Jurusan
        df_serapan_jurusan = create_serapan_jurusan(df_load)
        if not df_serapan_jurusan.empty:
            print_styled_table(df_serapan_jurusan, "Table 6: Serapan Lulusan per Jurusan")
            chart_serapan_jurusan = get_horizontal_bar_chart_base64(df_serapan_jurusan, "Serapan Lulusan per Jurusan")
            dfs_to_report["Serapan Lulusan per Jurusan"] = (df_serapan_jurusan, chart_serapan_jurusan)
            
        # Table 7: Serapan Prodi per Jurusan (Split Tables)
        # Now returns a dictionary of dataframes
        dict_serapan_prodi = create_serapan_prodi_per_jurusan(df_load)
        if dict_serapan_prodi:
            # We iterate and add them one by one
            # Use 7.1, 7.2 etc numbering logic or just title
            for idx, (jurusan_name, df_jur) in enumerate(dict_serapan_prodi.items(), 1):
                table_title = f"{jurusan_name}"
                # Ensure spacing
                print_styled_table(df_jur, table_title)
                dfs_to_report[table_title] = (df_jur, None)
        
        # Table 8: Sebaran Provinsi
        df_provinsi = create_distribution_provinsi(df_load)
        if not df_provinsi.empty:
            print_styled_table(df_provinsi, "Table 8: Sebaran Alumni per Provinsi")
            # Generate Map
            MAP_OUTPUT = os.path.join(REPORTS_DIR, 'Peta_Sebaran_Alumni.html')
            generate_alumni_map(df_provinsi, MAP_OUTPUT)
            # Add table to report
            # Pass map filename reference for IFrame
            dfs_to_report["Sebaran Alumni per Provinsi"] = (df_provinsi, None, 'Peta_Sebaran_Alumni.html')
            
        # New Table: Sebaran Kota/Kabupaten Kalbar
        df_kalbar = create_distribution_kabkota_kalbar(df_load)
        if not df_kalbar.empty:
             print_styled_table(df_kalbar, "Table 9: Distribusi Serapan Alumni di DUDI")
             # Prepare df for chart: Rename 'Jumlah Responden' to 'Total' and drop percentage string
             df_chart = df_kalbar.set_index('Kota/Kabupaten')[['Jumlah Responden']].rename(columns={'Jumlah Responden': 'Total'})
             chart_kalbar = get_horizontal_bar_chart_base64(
                 df_chart,
                 "Sebaran Alumni Kalbar per Kota/Kabupaten"
             )

             # Generate Kalbar Map
             MAP_KALBAR_OUTPUT = os.path.join(REPORTS_DIR, 'Peta_Sebaran_Kalbar.html')
             generate_kalbar_map(df_kalbar, MAP_KALBAR_OUTPUT)
             
             dfs_to_report["Distribusi Serapan Alumni di DUDI"] = (df_kalbar, chart_kalbar, 'Peta_Sebaran_Kalbar.html')

        # New Table: Distribusi Pendapatan
        df_salary = create_salary_distribution(df_load)
        if not df_salary.empty:
            print_styled_table(df_salary, "Table 10: Distribusi Pendapatan Responden per Bulan")
            # Prepare df for chart
            df_salary_chart = df_salary[df_salary['Rata-rata Pendapatan'] != 'Total'].copy()
            df_salary_chart = df_salary_chart.set_index('Rata-rata Pendapatan')[['Jumlah Responden']].rename(columns={'Jumlah Responden': 'Total'})
            
            chart_salary = get_horizontal_bar_chart_base64(df_salary_chart, "Distribusi Pendapatan")
            dfs_to_report["Distribusi Rata-rata Pendapatan Lulusan per Bulan"] = (df_salary, chart_salary)

        # New Table: Rata-rata Gaji per Jurusan
        result_salary = create_salary_by_jurusan(df_load)
        if result_salary and not result_salary[0].empty:
            df_salary_display, df_salary_ranked = result_salary
            print_styled_table(df_salary_display, "Table 11: Rata-rata Gaji Lulusan per Jurusan (Estimasi)")
            
            # Prepare for chart
            df_chart_sj = df_salary_ranked.set_index('Jurusan')[['Total']]
            
            chart_salary_jurusan = get_horizontal_bar_chart_base64(df_chart_sj, "Ranking Jurusan berdasarkan Rata-rata Gaji")
            dfs_to_report["Rata-rata Gaji Lulusan per Jurusan"] = (df_salary_display, chart_salary_jurusan)

        # New Table: Ranking Jurusan (Static from Analysis)
        df_ranking = create_jurusan_ranking()
        print_styled_table(df_ranking, "Table 12: Peringkat Performa Jurusan")
        dfs_to_report["Peringkat Performa Jurusan - Tracer Study 2025"] = (df_ranking, None)

        generate_html_report(dfs_to_report, output_file=REPORT_OUTPUT)
        
    except Exception as e:
        print(f"Error executing main: {e}")
        import traceback
        traceback.print_exc()
