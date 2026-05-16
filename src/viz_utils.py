import subprocess, sys, importlib
import pandas as pd
import numpy as np
import io
import base64
import numpy as np

try:
    import folium
    from folium.plugins import MarkerCluster
except ImportError as e:
    print(f"Import Error: {e}")
    folium = None

try:
    import geopandas as gpd
    import matplotlib.pyplot as plt
    import matplotlib.patheffects
    from shapely.geometry import Point
except ImportError:
    gpd = None
    plt = None


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


def _ensure_pkg(pkg):
    try:
        importlib.import_module(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])


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
                                 legend_kwds={'label': "Jumlah Alumni", 'orientation': "horizontal", 'location': "top", 'shrink': 0.6},
                                 cmap='Blues',
                                 edgecolor='black',
                                 linewidth=0.5,
                                 missing_kwds={'color': 'lightgrey'})
                                 
                 # ax.set_title(f"Sebaran Alumni - {region_name} (Choropleth)", fontsize=16)
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
                             legend_kwds={'label': "Jumlah Alumni (Non-Pontianak)", 'orientation': "horizontal", 'location': "top", 'shrink': 0.6},
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
                 val_count = row.get('Jumlah Responden', 0)
                 val_pct = row.get('pct', 0)
                 
                 # Format Label: "Name\nCount (X.X%)"
                 label = f"{name}\n{int(val_count)} ({val_pct:.1f}%)"
                 
                 # Add Text
                 # Use white halo for readability
                 ax.annotate(text=label, xy=(x, y), xytext=(0, 0), textcoords="offset points", # Center
                             ha='center', va='center',
                             fontsize=8, color='black', weight='bold',
                             path_effects=[matplotlib.patheffects.withStroke(linewidth=2, foreground="white")])

             # ax.set_title("Sebaran Alumni - Kalimantan Barat", fontsize=16)
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
        # ax.set_title(f"Sebaran Alumni - {region_name}", fontsize=16)
    else:
        ax.set_xlim(95, 141)
        ax.set_ylim(-11, 6)
        # ax.set_title(f"Sebaran Alumni - {region_name}", fontsize=16)
    
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
                    legend=True,
                    legend_kwds={'orientation': "horizontal", 'location': "top", 'shrink': 0.6})

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


def sort_crosstab_by_total(df_crosstab, denominator=None):
    """
    Sorts the crosstab DataFrame by the 'Total' column in descending order,
    keeping the 'Total' row (margin) at the bottom.
    Also adds a 'Persentase' column.
    If 'denominator' is provided, it uses it for percentage calculation 
    instead of the crosstab's own total.
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
        if denominator is not None:
             grand_total = denominator
        elif 'Total' in df_final.index:
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


def get_bar_chart_base64(df, title, chart_id, orientation='horizontal'):
    """
    Generates a bar chart (horizontal or vertical) and returns it as a base64 string.
    """
    df_plot = df.copy()
    
    # Robustly remove Total rows/columns
    rows_to_drop = [i for i in df_plot.index if 'total' in str(i).lower()]
    if rows_to_drop: df_plot = df_plot.drop(index=rows_to_drop)
    
    cols_to_drop = [c for c in df_plot.columns if 'total' in str(c).lower() and c != 'Total']
    if cols_to_drop: df_plot = df_plot.drop(columns=cols_to_drop)
    
    # Identify value column: 'Total' or the last numeric column
    numeric_cols = df_plot.select_dtypes(include=[np.number]).columns
    if 'Total' in df_plot.columns:
        val_col = 'Total'
    elif not numeric_cols.empty:
        val_col = numeric_cols[-1]
    else:
        return None
        
    df_plot = df_plot.sort_values(by=val_col, ascending=(orientation == 'horizontal'))
    values = df_plot[val_col]
    labels = df_plot.index.astype(str)

    n_bars = len(values)
    if n_bars == 0: return None

    # Style
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(10, max(5, n_bars * 0.4) if orientation == 'horizontal' else 6))
    
    max_val = values.max() if not values.empty else 0
    # Highlight max: DarkBlue, others: LightGrey
    colors = ['#00008B' if v == max_val and v > 0 else '#D3D3D3' for v in values]
    # Highlight size: Thicker for max
    sizes = [0.8 if v == max_val and v > 0 else 0.5 for v in values]

    if orientation == 'horizontal':
        bars = ax.barh(labels, values, color=colors, height=sizes)
        for bar in bars:
            width = bar.get_width()
            val_text = f'{int(width)}' if np.isfinite(width) else '0'
            ax.text(width + (max(values)*0.01 if not values.empty else 1), bar.get_y() + bar.get_height()/2, 
                    val_text, va='center', fontsize=9, fontweight='bold')
    else:
        bars = ax.bar(labels, values, color=colors, width=sizes)
        plt.xticks(rotation=45, ha='right')
        for bar in bars:
            height = bar.get_height()
            val_text = f'{int(height)}' if np.isfinite(height) else '0'
            ax.text(bar.get_x() + bar.get_width()/2, height + (max(values)*0.01 if not values.empty else 1), 
                    val_text, ha='center', va='bottom', fontsize=9, fontweight='bold')

    ax.set_title(f"{title}", fontsize=14, fontweight='bold', pad=20)
    
    # Remove grid and all spines
    ax.grid(False)
    for spine in ax.spines.values():
        spine.set_visible(False)
    
    # Hide tick marks but keep labels
    ax.tick_params(axis='both', which='both', length=0)
    
    if orientation == 'vertical':
        ax.set_yticks([]) # Hide Y values if vertical, labels are on X
    else:
        ax.set_xticks([]) # Hide X values if horizontal, labels are on Y
    
    plt.tight_layout()
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_pie_chart_base64(df, title, chart_id):
    """
    Generates a pie chart and returns it as a base64 string.
    """
    df_plot = df.copy()
    rows_to_drop = [i for i in df_plot.index if 'total' in str(i).lower()]
    if rows_to_drop: df_plot = df_plot.drop(index=rows_to_drop)
    
    numeric_cols = df_plot.select_dtypes(include=[np.number]).columns
    val_col = 'Total' if 'Total' in df_plot.columns else (numeric_cols[-1] if not numeric_cols.empty else None)
    if not val_col: return None
    
    # Sort by value
    df_plot = df_plot.sort_values(by=val_col, ascending=False)
    
    # Grouping into 'Lainnya' removed per user request to show all categories

    values = df_plot[val_col]
    labels = df_plot.index.astype(str)
    
    if len(values) == 0: return None

    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(8, 8))
    
    max_val = values.max() if not values.empty else 0
    # Explode only the highest value - reduced from 0.1 to 0.05 for better symmetry
    explode = [0.05 if v == max_val and v > 0 else 0 for v in values]
    # Highlight colors: DarkBlue for max, LightGrey for others
    colors = ['#00008B' if v == max_val and v > 0 else '#D3D3D3' for v in values]

    # Prepare display labels (Top 10 only)
    display_labels = [label if i < 10 else '' for i, label in enumerate(labels)]

    # Use width in wedgeprops for a consistent donut ring thickness
    wedges, texts, autotexts = ax.pie(values, labels=display_labels, 
                                    autopct='%1.1f%%', 
                                    startangle=140, colors=colors, pctdistance=0.85,
                                    explode=explode,
                                    wedgeprops={'edgecolor': 'white', 'linewidth': 1, 'width': 0.3})
    
    # Post-process labels and percentages to hide those beyond Top 10
    # and adjust colors for visibility
    for i, (t, at) in enumerate(zip(texts, autotexts)):
        if i >= 10:
            t.set_text('')
            at.set_text('')
        else:
            # For Top 10, ensure good contrast
            if values.iloc[i] == max_val:
                at.set_color('white') # High contrast on DarkBlue
            else:
                at.set_color('black') # High contrast on LightGrey
    
    ax.axis('equal') # Ensure the pie is drawn as a circle
    
    plt.setp(autotexts, size=9, weight="bold")
    plt.setp(texts, size=10, weight="bold")
    
    ax.set_title(f"{title}", fontsize=14, fontweight='bold', pad=20)
    plt.tight_layout()
    
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_line_chart_base64(df, title, chart_id):
    """
    Generates a line chart and returns it as a base64 string.
    Suitable for data with numeric/temporal columns.
    """
    df_plot = df.copy()
    rows_to_drop = [i for i in df_plot.index if 'total' in str(i).lower()]
    if rows_to_drop: df_plot = df_plot.drop(index=rows_to_drop)
    
    year_cols = [c for c in df_plot.columns if str(c).isdigit()]
    if not year_cols:
        return None
    
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(10, 6))
    
    for label, row in df_plot.iterrows():
        ax.plot(year_cols, row[year_cols], marker='o', label=label, linewidth=2)
    
    ax.set_title(f"{title}", fontsize=14, fontweight='bold', pad=20)
    ax.set_xlabel("Tahun")
    ax.set_ylabel("Jumlah")
    if len(df_plot) > 1:
        ax.legend(title="Kategori", bbox_to_anchor=(1.05, 1), loc='upper left')
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_horizontal_bar_chart_base64(df, title):
    """Legacy wrapper for backward compatibility if needed, but we'll use get_bar_chart_base64."""
    return get_bar_chart_base64(df, title, "legacy_bar")


def get_smooth_trend_chart_base64(df, title, chart_id):
    """
    Generates a smooth line chart showing the trend of hiring by month.
    """
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    
    if col_status not in df.columns or col_masa_tunggu not in df.columns:
        return None
        
    df_analysis = df[[col_status, col_masa_tunggu]].copy()
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    df_analysis = df_analysis[df_analysis[col_status].isin(working_status)]
    
    df_analysis[col_masa_tunggu] = pd.to_numeric(df_analysis[col_masa_tunggu], errors='coerce')
    df_analysis = df_analysis.dropna(subset=[col_masa_tunggu])
    
    # Filter for realistic months (0 to 36)
    df_analysis = df_analysis[(df_analysis[col_masa_tunggu] >= 0) & (df_analysis[col_masa_tunggu] <= 36)]
    
    monthly_counts = df_analysis.groupby(col_masa_tunggu).size().reset_index(name='Jumlah Lulusan')
    monthly_counts.columns = ['Bulan', 'Jumlah Lulusan']
    
    if monthly_counts.empty:
        return None
        
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(10, 6))
    
    x = monthly_counts['Bulan'].values
    y = monthly_counts['Jumlah Lulusan'].values
    
    sort_idx = np.argsort(x)
    x = x[sort_idx]
    y = y[sort_idx]
    
    from scipy.interpolate import make_interp_spline
    try:
        if len(x) >= 4:
            x_smooth = np.linspace(x.min(), x.max(), 300)
            spl = make_interp_spline(x, y, k=3)
            y_smooth = spl(x_smooth)
            y_smooth = np.clip(y_smooth, 0, None)
            ax.plot(x_smooth, y_smooth, color='#00008B', linewidth=2.5, label='Trend Diterima')
            ax.fill_between(x_smooth, 0, y_smooth, color='#00008B', alpha=0.1)
        else:
            ax.plot(x, y, color='#00008B', linewidth=2.5, marker='o', label='Trend Diterima')
            ax.fill_between(x, 0, y, color='#00008B', alpha=0.1)
    except Exception:
        ax.plot(x, y, color='#00008B', linewidth=2.5, marker='o', label='Trend Diterima')
        ax.fill_between(x, 0, y, color='#00008B', alpha=0.1)
        
    ax.scatter(x, y, color='#00008B', s=40, zorder=5, alpha=0.8)
    
    for i in range(len(x)):
        if y[i] > 0:
            ax.annotate(str(int(y[i])), 
                        (x[i], y[i]), 
                        textcoords="offset points", 
                        xytext=(0, 10), 
                        ha='center', 
                        fontsize=9,
                        fontweight='bold',
                        color='#00008B')

    if len(y) > 0:
        max_idx = np.argmax(y)
        max_x = x[max_idx]
        max_y = y[max_idx]
        
        ax.scatter([max_x], [max_y], color='#FF4500', s=120, zorder=6, edgecolors='white', linewidth=2)
        
        ax.annotate(f'Puncak: Bulan ke-{int(max_x)}\n({max_y} Alumni)',
                    xy=(max_x, max_y), xytext=(0, 25),
                    textcoords='offset points', ha='center', va='bottom',
                    fontsize=11, fontweight='bold', color='#FF4500',
                    bbox=dict(boxstyle='round,pad=0.5', fc='white', alpha=0.9, ec='#FF4500'))

    ax.set_title(f"{title}", fontsize=14, fontweight='bold', pad=20)
    ax.set_xlabel("Bulan (Masa Tunggu)", fontsize=12)
    ax.set_ylabel("Jumlah Lulusan Diterima Bekerja", fontsize=12)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
    
    plt.tight_layout()
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def apply_row_percentages_for_display(df):
    """
    Menambahkan persentase dalam kurung untuk setiap sel numerik 
    berdasarkan total pada baris tersebut.
    Targeting crosstab-like tables with a 'Total' column.
    """
    df_formatted = df.copy()
    
    # Identifikasi kolom yang merupakan total baris
    total_candidates = ['Total', 'Total Responden (Alumni)']
    total_col = next((c for c in total_candidates if c in df_formatted.columns), None)
    
    if total_col is None:
        return df_formatted
    
    # 1. Cari kolom Persentase yang ada (biasanya hasil sort_crosstab_by_total atau sejenisnya)
    # Cari kolom yang namanya mengandung "persentase" tapi BUKAN metriks khusus 
    # seperti 'Persentase (<= 6 Bulan) (%)'
    specific_metrics = ['Persentase (<= 6 Bulan) (%)']
    pct_col = next((c for c in df_formatted.columns if 'persentase' in str(c).lower() and c not in specific_metrics), None)
        
    # 2. Identifikasi kolom yang harus dilewati untuk row-percentages (identitas, metriks khusus, dll)
    skip_keywords = ['persentase', 'rata-rata', 'skor', 'peringkat', 'tahun', 'id', 'gaji', 'pendapatan', 'predikat']
    
    # Kolom numerik untuk row-percentages
    numeric_cols = []
    for col in df_formatted.columns:
        if col == total_col or col == pct_col:
            continue
        
        col_lower = str(col).lower()
        if any(kw in col_lower for kw in skip_keywords):
            continue
            
        if pd.api.types.is_numeric_dtype(df_formatted[col]):
            numeric_cols.append(col)
            
    # Cek apakah kolom Total sudah di-format sebagai string (mengandung '(')
    if not df_formatted[total_col].empty and isinstance(df_formatted[total_col].iloc[0], str) and '(' in df_formatted[total_col].iloc[0]:
        return df_formatted
        
    # Pastikan total_col numerik untuk kalkulasi
    df_formatted[total_col] = pd.to_numeric(df_formatted[total_col], errors='coerce').fillna(0)

    # Check for institutional total in attributes (for global tables)
    inst_total = df.attrs.get('institutional_total') if hasattr(df, 'attrs') else None

    # 3. Proses setiap kolom numerik (Row-Relative or Global-Relative Percentages)
    for col in numeric_cols:
        def row_formatter(row):
            val = row[col]
            # Use institutional total if provided, otherwise use row total
            total = inst_total if inst_total is not None else row[total_col]
            
            if pd.isna(val):
                return ""
            
            # Jika sel memang angka (termasuk 0)
            if isinstance(val, (int, float, np.number)):
                if total > 0:
                    pct = (val / total) * 100
                    # Format: "Value (XX.X%)"
                    # Gunakan integer jika val bulat
                    val_str = f"{int(val)}" if val == int(val) else f"{val:.1f}"
                    return f"{val_str} ({pct:.1f}%)"
                else:
                    # Menghindari division by zero atau total 0
                    return f"{val} (0.0%)"
            return str(val)
                
        df_formatted[col] = df_formatted.apply(row_formatter, axis=1)

    # 4. Gabungkan kolom Persentase ke dalam kolom Total jika ada
    if pct_col:
        def total_formatter(row):
            val = row[total_col]
            pct = row[pct_col]
            if pd.isna(val): return ""
            val_str = f"{int(val)}" if val == int(val) else f"{val:.1f}"
            # Kadang pct sudah ada tanda % atau format string
            return f"{val_str} ({pct})"
            
        df_formatted[total_col] = df_formatted.apply(total_formatter, axis=1)
        df_formatted.drop(columns=[pct_col], inplace=True)
        
    return df_formatted


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
                padding: 10px 12px;
                text-align: left;
                border-bottom: 1px solid #edf2f7;
            }
            th {
                background-color: #3498db;
                color: #fff;
                font-weight: 600;
                text-transform: uppercase;
                font-size: 0.75rem;
                letter-spacing: 0.5px;
                white-space: normal; /* Allow header text to wrap */
                min-width: 60px;
                max-width: 100px; /* Tight width forces multi-word headers to stack */
                word-wrap: break-word;
                hyphens: auto;
                vertical-align: middle;
                text-align: center;
            }
            td {
                white-space: nowrap; /* Prevent data values from wrapping */
                font-size: 0.85rem;
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
                width: 1%; /* Force to fit content exactly */
                white-space: nowrap;
                text-align: center;
            }
            td:last-child {
                border-left: 1px solid #f1f5f9;
                background-color: #f8fafc;
                width: 1%; /* Force to fit content exactly */
                white-space: nowrap;
                text-align: center;
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
    for title, content in data_dict.items():
        # Handle different input formats for backward compatibility
        if isinstance(content, dict):
            df = content.get('df')
            charts = content.get('charts', [])
            map_path = content.get('map')
        elif isinstance(content, tuple):
            df = content[0]
            # If it was a single chart string
            chart_val = content[1]
            if isinstance(chart_val, str):
                charts = [{"id": f"chart_{section_id}", "name": "Grafik", "base64": chart_val}]
            elif isinstance(chart_val, list):
                charts = chart_val
            else:
                charts = []
            map_path = content[2] if len(content) > 2 else None
        else:
            continue
            
        section_id += 1
        table_id = f"table_{section_id}"
        
        html_content += f'<div class="section" id="section_{section_id}">'
        html_content += f"<h2>{title}</h2>"
        
        # --- Table Section ---
        html_content += f'<div class="btn-group">'
        html_content += f'<button class="btn" onclick="saveTable(\'{table_id}\', \'{title}_table\')">Simpan Tabel</button>'
        html_content += '</div>'
        
        if df is not None:
            # Apply Row-Relative Percentages
            df_display = apply_row_percentages_for_display(df)
            
            df_to_html = df_display.reset_index() if df_display.index.name else df_display.copy()
            df_to_html.columns.name = None
            table_html = df_to_html.to_html(index=False, border=0, classes='table', table_id=table_id, escape=False)
            if f'id="{table_id}"' not in table_html:
                table_html = table_html.replace('<table', f'<table id="{table_id}"')
            html_content += table_html
        
        # --- Charts Section ---
        if charts:
            html_content += '<div class="charts-grid" style="display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-top: 20px;">'
            for chart in charts:
                c_id = chart.get('id', 'unnamed')
                c_name = chart.get('name', 'Grafik')
                c_b64 = chart.get('base64')
                if not c_b64: continue
                
                html_content += f"""
                <div class="chart-container" style="flex: 1; min-width: 45%; max-width: 100%;">
                    <div class="btn-group" style="text-align: right;">
                        <span style="float: left; font-size: 0.8rem; color: #7f8c8d;">ID: {c_id}</span>
                        <a href="data:image/png;base64,{c_b64}" download="{c_id}.png" class="btn btn-secondary">Simpan {c_name}</a>
                    </div>
                    <img src="data:image/png;base64,{c_b64}" alt="{c_name}" class="chart-img" id="{c_id}">
                </div>
                """
            html_content += '</div>'
        
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


def get_waktu_tunggu_global_smooth_line_chart_base64(df, title, chart_id):
    """
    Generates a global smooth line chart for waiting time distribution
    """
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_status = 'Jelaskan status Anda saat ini?'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_masa_tunggu not in df.columns or col_status not in df.columns:
        return None
        
    df_filtered = df[df[col_status].isin(working_status)].copy()
    df_filtered['Masa_Tunggu_Bulan'] = pd.to_numeric(df_filtered[col_masa_tunggu], errors='coerce')
    df_filtered = df_filtered.dropna(subset=['Masa_Tunggu_Bulan'])
    
    counts = df_filtered['Masa_Tunggu_Bulan'].value_counts().sort_index()
    if counts.empty: return None
    
    x = counts.index.to_numpy()
    y = counts.values
    
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(10, 6))
    
    from scipy.interpolate import make_interp_spline
    
    if len(x) > 3:
        x_new = np.linspace(x.min(), x.max(), 300)
        spline = make_interp_spline(x, y, k=3)
        y_smooth = spline(x_new)
        y_smooth = np.maximum(y_smooth, 0)
    else:
        x_new = x
        y_smooth = y
        
    ax.plot(x_new, y_smooth, color='#00008B', linewidth=2)
    ax.fill_between(x_new, y_smooth, alpha=0.3, color='#87CEEB')
    
    # Add vertical line for the mean
    mean_val = df_filtered['Masa_Tunggu_Bulan'].mean()
    ax.axvline(mean_val, color='red', linestyle='--', linewidth=2, label=f'Rata-rata: {mean_val:.1f} Bulan')
    
    ax.set_title(title, fontsize=14, fontweight='bold', pad=20)
    ax.set_xlabel('Masa Tunggu (Bulan)', fontsize=12)
    ax.set_ylabel('Jumlah Responden (Bekerja & Wiraswasta)', fontsize=12)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.legend(loc='upper right', fontsize=11)
    
    plt.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_waktu_tunggu_facet_smooth_line_chart_base64(df, title, chart_id):
    """
    Generates a facet grid smooth line chart for waiting time distribution per Jurusan
    """
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_status = 'Jelaskan status Anda saat ini?'
    col_jurusan = 'Jurusan'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_masa_tunggu not in df.columns or col_status not in df.columns or col_jurusan not in df.columns:
        return None
        
    df_filtered = df[df[col_status].isin(working_status)].copy()
    df_filtered['Masa_Tunggu_Bulan'] = pd.to_numeric(df_filtered[col_masa_tunggu], errors='coerce')
    df_filtered = df_filtered.dropna(subset=['Masa_Tunggu_Bulan'])
    
    if df_filtered.empty: return None
        
    jurusans = sorted(df_filtered[col_jurusan].unique())
    n_cats = len(jurusans)
    if n_cats == 3:
        n_cols = 3
    elif n_cats == 4:
        n_cols = 2
    elif n_cats in [5, 6]:
        n_cols = 3
    elif n_cats >= 7:
        n_cols = 4
    else:
        n_cols = max(1, n_cats)
    n_rows = (n_cats + n_cols - 1) // n_cols
    
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(5*n_cols, 4*n_rows), sharex=False, sharey=False)
    axes = axes.flatten() if n_cats > 1 else [axes]
    
    from scipy.interpolate import make_interp_spline
    import matplotlib as mpl
    colors = mpl.colormaps.get_cmap('tab20')
    
    for i, jurusan in enumerate(jurusans):
        ax = axes[i]
        df_jur = df_filtered[df_filtered[col_jurusan] == jurusan]
        counts = df_jur['Masa_Tunggu_Bulan'].value_counts().sort_index()
        
        color = colors(i % 20)
        
        if counts.empty:
            ax.set_title(jurusan, fontsize=11, fontweight='bold')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            continue
            
        x = counts.index.to_numpy()
        y = counts.values
        
        # Smooth line
        if len(x) > 3:
            x_new = np.linspace(x.min(), x.max(), 300)
            spline = make_interp_spline(x, y, k=3)
            y_smooth = spline(x_new)
            y_smooth = np.maximum(y_smooth, 0)
        else:
            x_new = x
            y_smooth = y
            
        # Plot smooth line without scatter dots
        ax.plot(x_new, y_smooth, linewidth=2, color=color)
        ax.fill_between(x_new, y_smooth, alpha=0.3, color=color)
        
        # Add average line
        mean_val = df_jur['Masa_Tunggu_Bulan'].mean()
        ax.axvline(mean_val, color='red', linestyle='--', linewidth=1.5)
        # Add text annotation for average
        y_max = y.max() if len(y) > 0 else 1
        ax.text(mean_val + 0.1, y_max * 0.85, f'Rata: {mean_val:.1f}', color='red', fontsize=10, fontweight='bold', va='center')
        
        ax.set_title(jurusan, fontsize=11, fontweight='bold')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, linestyle='--', alpha=0.5)
    
    # Hide empty subplots
    for j in range(len(jurusans), len(axes)):
        axes[j].set_visible(False)
        
    fig.suptitle(title, fontsize=16, fontweight='bold', y=1.02)
    # Add common labels
    fig.supxlabel('Masa Tunggu (Bulan)', fontsize=12, y=-0.02)
    fig.supylabel('Jumlah Responden', fontsize=12, x=-0.02)
    
    plt.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_waktu_tunggu_prodi_facet_smooth_line_chart_base64(df, title, chart_id, jurusan=None):
    """
    Generates a facet grid smooth line chart for waiting time distribution per Prodi
    """
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_status = 'Jelaskan status Anda saat ini?'
    col_prodi = 'prodi'
    col_jurusan = 'Jurusan'
    if col_prodi not in df.columns:
        if 'Program Studi' in df.columns:
            col_prodi = 'Program Studi'
        else:
            return None
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    
    if col_masa_tunggu not in df.columns or col_status not in df.columns:
        return None
        
    df_filtered = df[df[col_status].isin(working_status)].copy()
    if jurusan and col_jurusan in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_jurusan] == jurusan]
        
    df_filtered['Masa_Tunggu_Bulan'] = pd.to_numeric(df_filtered[col_masa_tunggu], errors='coerce')
    df_filtered = df_filtered.dropna(subset=['Masa_Tunggu_Bulan'])
    
    if df_filtered.empty: return None
        
    prodis = sorted(df_filtered[col_prodi].unique())
    n_cats = len(prodis)
    if n_cats == 3:
        n_cols = 3
    elif n_cats == 4:
        n_cols = 2
    elif n_cats in [5, 6]:
        n_cols = 3
    elif n_cats >= 7:
        n_cols = 4
    else:
        n_cols = max(1, n_cats)
    n_rows = (n_cats + n_cols - 1) // n_cols
    
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(10, 4*n_rows), sharex=False, sharey=False)
    axes = axes.flatten() if n_cats > 1 else [axes]
    
    from scipy.interpolate import make_interp_spline
    import matplotlib as mpl
    colors = mpl.colormaps.get_cmap('tab20')
    
    for i, prodi in enumerate(prodis):
        ax = axes[i]
        df_prod = df_filtered[df_filtered[col_prodi] == prodi]
        counts = df_prod['Masa_Tunggu_Bulan'].value_counts().sort_index()
        
        color = colors(i % 20)
        
        if counts.empty:
            ax.set_title(prodi, fontsize=11, fontweight='bold')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            continue
            
        x = counts.index.to_numpy()
        y = counts.values
        
        # Smooth line
        if len(x) > 3:
            x_new = np.linspace(x.min(), x.max(), 300)
            spline = make_interp_spline(x, y, k=3)
            y_smooth = spline(x_new)
            y_smooth = np.maximum(y_smooth, 0)
        else:
            x_new = x
            y_smooth = y
            
        # Plot smooth line without scatter dots
        ax.plot(x_new, y_smooth, linewidth=2, color=color)
        ax.fill_between(x_new, y_smooth, alpha=0.3, color=color)
        
        # Add average line
        mean_val = df_prod['Masa_Tunggu_Bulan'].mean()
        ax.axvline(mean_val, color='red', linestyle='--', linewidth=1.5)
        # Add text annotation for average
        y_max = y.max() if len(y) > 0 else 1
        ax.text(mean_val + 0.1, y_max * 0.85, f'Rata: {mean_val:.1f}', color='red', fontsize=10, fontweight='bold', va='center')
        
        ax.set_title(prodi, fontsize=11, fontweight='bold')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, linestyle='--', alpha=0.5)
    
    # Hide empty subplots
    for j in range(len(prodis), len(axes)):
        axes[j].set_visible(False)
        
    fig.suptitle(title, fontsize=16, fontweight='bold', y=1.02)
    # Add common labels
    fig.supxlabel('Masa Tunggu (Bulan)', fontsize=12, y=-0.02)
    fig.supylabel('Jumlah Responden', fontsize=12, x=-0.02)
    
    plt.tight_layout()
    import io, base64
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_divergence_chart_base64(labels, values_left, values_right, title, label_left, label_right, is_percentage=True):
    """
    Generates a divergence (bidirectional) bar chart.
    Supports stacked bars on the left side if values_left is a list of lists.
    Values should be positive; left side will be plotted as negative for divergence.
    """
    import matplotlib.pyplot as plt
    import numpy as np
    
    n_labels = len(labels)
    if n_labels == 0: return None
    
    # Style
    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(12, max(6, n_labels * 0.5)))
    
    y_pos = np.arange(n_labels)
    
    # Colors per user request:
    # Left stacked: Bekerja (DarkBlue), Wiraswasta (DarkGrey)
    # Right: Sedang Mencari Kerja (LightYellow)
    color_bekerja = '#00008B'
    color_wiraswasta = '#696969' # Dark Grey
    color_searching = '#FFFFE0' # Light Yellow
    
    # Plotting Left Side (Stacked)
    if isinstance(values_left, list) and len(values_left) > 0 and isinstance(values_left[0], list):
        # Stacked logic
        vals_bekerja = np.array(values_left[0])
        vals_wiraswasta = np.array(values_left[1])
        
        # Plot Wiraswasta first (base), then Bekerja on top (further left)
        # Note: negative values for divergence
        bars_wiraswasta = ax.barh(y_pos, -vals_wiraswasta, color=color_wiraswasta, edgecolor='white', linewidth=0.5)
        bars_bekerja = ax.barh(y_pos, -vals_bekerja, left=-vals_wiraswasta, color=color_bekerja, edgecolor='white', linewidth=0.5)
        
        # Add labels inside left bars
        for i in range(n_labels):
            # Label for Wiraswasta
            if vals_wiraswasta[i] > 0:
                ax.text(-vals_wiraswasta[i]/2, i, f'{int(vals_wiraswasta[i])}', 
                        va='center', ha='center', fontsize=8, fontweight='bold', color='white')
            # Label for Bekerja
            if vals_bekerja[i] > 0:
                ax.text(-(vals_wiraswasta[i] + vals_bekerja[i]/2), i, f'{int(vals_bekerja[i])}', 
                        va='center', ha='center', fontsize=8, fontweight='bold', color='white')
        
        max_left = (vals_bekerja + vals_wiraswasta).max()
        label_left_final = f"{label_left[0]} (Biru) & {label_left[1]} (Abu)"
    else:
        # Simple left bar (fallback)
        bars_left = ax.barh(y_pos, [-v for v in values_left], color=color_bekerja, edgecolor='grey', linewidth=0.5)
        for i, bar in enumerate(bars_left):
            val = values_left[i]
            if val > 0:
                ax.text(-val/2, i, f'{int(val)}', va='center', ha='center', fontsize=8, fontweight='bold', color='white')
        max_left = max(values_left) if any(values_left) else 0
        label_left_final = label_left

    # Plotting Right Side
    bars_right = ax.barh(y_pos, values_right, color=color_searching, edgecolor='grey', linewidth=0.5)
    for i, bar in enumerate(bars_right):
        val = values_right[i]
        if val > 0:
            val_str = f'{val:.1f}%' if is_percentage else f'{int(val)}'
            # Place searching label inside if it's large enough, or just outside
            ax.text(val + 0.5, i, val_str, va='center', ha='left', fontsize=9, fontweight='bold', color='black')

    ax.set_yticks(y_pos)
    # Highlight labels that have the (100%) marker (100% absorption)
    ax.set_yticklabels(labels, fontweight='bold')
    
    # Calculate limits first so we can use it for arrow placement
    max_right = max(values_right) if any(values_right) else 0
    limit = max(max_left, max_right) + (max(max_left, max_right) * 0.15 if max(max_left, max_right) > 0 else 10)
    ax.set_xlim(-limit, limit)

    # Apply specific colors to tick labels and add arrows
    for i, label in enumerate(labels):
        if "(100%)" in label:
            ax.get_yticklabels()[i].set_color('#0000CD') # MediumBlue
            ax.get_yticklabels()[i].set_fontsize(10)
            
            # Add red arrow pointing from label to the bar
            # Get the tip of the left bar
            if isinstance(values_left, list) and len(values_left) > 0 and isinstance(values_left[0], list):
                val_total = values_left[0][i] + values_left[1][i]
            else:
                val_total = values_left[i]
            
            # x_start is near the left edge (label area), x_end is the bar tip
            # Arrow target: tip of the bar (negative x)
            # Arrow source: at the -limit boundary where labels are
            if val_total > 0:
                ax.annotate('', 
                            xy=(-val_total, i), 
                            xytext=(-limit, i),
                            arrowprops=dict(arrowstyle='-|>', color='red', lw=2, mutation_scale=15))
    
    from matplotlib.ticker import FuncFormatter
    if is_percentage:
        ax.xaxis.set_major_formatter(FuncFormatter(lambda x, pos: f'{abs(x):.0f}%'))
    else:
        ax.xaxis.set_major_formatter(FuncFormatter(lambda x, pos: f'{int(abs(x))}'))
    
    ax.set_title(title, fontsize=16, fontweight='bold', pad=45)
    
    # Place labels berkesesuaian dengan posisi chart gradient
    ax.text(0.25, 1.02, label_left_final, transform=ax.transAxes, ha='center', va='bottom', 
            fontsize=12, fontweight='bold', color=color_bekerja) 
    
    ax.text(0.75, 1.02, label_right, transform=ax.transAxes, ha='center', va='bottom', 
            fontsize=12, fontweight='bold', color='#B8860B') 
    
    # Remove grid and spines
    ax.grid(False) 
    for spine in ax.spines.values():
        spine.set_visible(False)
    
    # Vertical line at center
    ax.axvline(0, color='black', linewidth=1, alpha=0.5)
    
    plt.tight_layout()
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_stacked_bar_chart_base64(df, title, chart_id, is_percentage=True, orientation='vertical'):
    """
    Generates a stacked bar chart and returns it as a base64 string.
    df: index is the category (e.g. Prodi), columns are the statuses.
    """
    import matplotlib.pyplot as plt
    import numpy as np
    import io
    import base64
    import textwrap

    df_plot = df.copy()
    
    # Robustly remove Total rows/columns if they exist
    rows_to_drop = [i for i in df_plot.index if 'total' in str(i).lower()]
    if rows_to_drop: df_plot = df_plot.drop(index=rows_to_drop)
    
    cols_to_drop = [c for c in df_plot.columns if 'total' in str(c).lower()]
    if cols_to_drop: df_plot = df_plot.drop(columns=cols_to_drop)

    # Wrap Program Studi Names
    df_plot.index = [textwrap.fill(str(l), width=25) for l in df_plot.index]

    # Capture totals before normalizing to 100%
    totals = df_plot.sum(axis=1)

    if is_percentage:
        # Normalize to 100%
        # Avoid division by zero
        df_plot = df_plot.div(totals.replace(0, 1), axis=0) * 100

    # Define consistent colors
    color_map = {
        "Bekerja": "#00008B",              # DarkBlue
        "Wiraswasta": "#696969",            # DarkGrey
        "Sedang Mencari Kerja": "#B8860B",  # Darkish Yellow
        "Studi Lanjut": "#4682B4",          # SteelBlue
        "Belum Memungkinkan Bekerja": "#D3D3D3", # LightGrey
        "Tidak Mencari Kerja": "#A9A9A9"    # DarkGray
    }
    
    # Get colors for columns present in df
    colors = [color_map.get(col, '#D3D3D3') for col in df_plot.columns]

    plt.style.use('seaborn-v0_8-whitegrid')
    
    if orientation == 'horizontal':
        # Flip to match top-to-bottom reading
        df_plot = df_plot.iloc[::-1] 
        totals_sorted = totals.iloc[::-1]
        
        # multiplier 0.5 increases the gap slightly compared to 0.4
        fig, ax = plt.subplots(figsize=(14, max(6, len(df_plot) * 0.5)))
        # width=0.85 increases thickness (increased by 0.3 from 0.55)
        bars_plot = df_plot.plot(kind='barh', stacked=True, ax=ax, color=colors, edgecolor='white', linewidth=0.5, width=0.85)
        
        if is_percentage:
            ax.set_xlabel("Persentase (%)", fontsize=12)
            ax.set_xlim(0, 115) # Add space for labels
        else:
            ax.set_xlabel("Jumlah", fontsize=12)
            
        # Add labels next to bars
        for i, (idx, total) in enumerate(totals_sorted.items()):
            x_pos = 101 if is_percentage else total + 0.5
            ax.text(x_pos, i, f'(N={int(total)})', va='center', ha='left', fontsize=9, fontweight='bold', color='#333')

        # Add values inside bars
        for c in ax.containers:
            # Only show if value is > 5% to avoid clutter
            labels = [f'{v:.0f}%' if v > 5 else '' for v in c.datavalues] if is_percentage else [f'{int(v)}' if v > 0 else '' for v in c.datavalues]
            ax.bar_label(c, labels=labels, label_type='center', fontsize=9, color='white', fontweight='bold')

    else:
        # Dynamic width based on number of items
        width_fig = max(12, len(df_plot) * 0.4)
        fig, ax = plt.subplots(figsize=(width_fig, 9))
        # width=0.85 increases thickness
        bars_plot = df_plot.plot(kind='bar', stacked=True, ax=ax, color=colors, edgecolor='white', linewidth=0.5, width=0.85)
        
        if is_percentage:
            ax.set_ylabel("Persentase (%)", fontsize=12)
            ax.set_ylim(0, 115) # Add space for labels
        else:
            ax.set_ylabel("Jumlah", fontsize=12)
        plt.xticks(rotation=45, ha='right', fontsize=10)
        
        # Add labels on top of bars
        for i, (idx, total) in enumerate(totals.items()):
            y_pos = 101 if is_percentage else total + 0.5
            ax.text(i, y_pos, f'N={int(total)}', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#333', rotation=90)

        # Add values inside bars
        for c in ax.containers:
            labels = [f'{v:.0f}%' if v > 5 else '' for v in c.datavalues] if is_percentage else [f'{int(v)}' if v > 0 else '' for v in c.datavalues]
            ax.bar_label(c, labels=labels, label_type='center', fontsize=9, color='white', fontweight='bold')
    
    # Remove grid
    ax.grid(False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    ax.legend(title="Status Pekerjaan", loc='lower center', bbox_to_anchor=(0.5, 1.02), 
              ncol=min(3, len(df_plot.columns)), fontsize=10, frameon=True)
    
    plt.tight_layout()
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_facet_pie_chart_base64(data_dict, title, n_cols=6):
    """
    Generates a facet grid of pie charts with a global legend.
    data_dict: { "Facet Title": pd.Series/pd.DataFrame (counts) }
    """
    import matplotlib.pyplot as plt
    import numpy as np
    import math

    n_items = len(data_dict)
    if n_items == 0: return None
    
    n_rows = math.ceil(n_items / n_cols)
    
    # Define consistent colors for common categories
    color_map = {
        "Bekerja": "#00008B",              # DarkBlue
        "Wiraswasta": "#696969",            # DarkGrey
        "Sedang Mencari Kerja": "#B8860B",  # Darkish Yellow
        "Studi Lanjut": "#4682B4",          # SteelBlue
        "Belum Memungkinkan Bekerja": "#D3D3D3", # LightGrey
        "Tidak Mencari Kerja": "#A9A9A9"    # DarkGray
    }
    
    import matplotlib as mpl
    fallback_colors = mpl.colormaps.get_cmap('Pastel1').colors

    plt.style.use('seaborn-v0_8-whitegrid')
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(3 * n_cols, 3 * n_rows + 1))
    
    if n_rows == 1 and n_cols == 1:
        axes_list = [axes]
    else:
        axes_list = axes.flatten()
    
    # Collect all unique labels
    all_labels = set()
    for df in data_dict.values():
        if isinstance(df, pd.DataFrame):
            # Assume first column is count if multi-column
            all_labels.update(df.index.astype(str))
        else:
            all_labels.update(df.index.astype(str))
    
    sorted_labels = sorted(list(all_labels))
    label_colors = {}
    for i, label in enumerate(sorted_labels):
        if label in color_map:
            label_colors[label] = color_map[label]
        else:
            label_colors[label] = fallback_colors[i % len(fallback_colors)]

    for i, (facet_title, df) in enumerate(data_dict.items()):
        ax = axes_list[i]
        
        if isinstance(df, pd.DataFrame):
            # Identify count column
            num_cols = df.select_dtypes(include=[np.number]).columns
            if not num_cols.empty:
                values = df[num_cols[0]].values
                labels = df.index.astype(str)
            else:
                ax.axis('off')
                continue
        else:
            values = df.values
            labels = df.index.astype(str)
        
        if len(values) == 0 or sum(values) == 0:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center', fontsize=8)
            ax.axis('off')
            continue
            
        colors = [label_colors[l] for l in labels]
        
        # Donut Style
        wedges, texts, autotexts = ax.pie(values, 
                                        autopct='%1.0f%%', 
                                        startangle=140, 
                                        colors=colors,
                                        pctdistance=0.75,
                                        wedgeprops={'edgecolor': 'white', 'linewidth': 0.5, 'width': 0.4})
        
        plt.setp(autotexts, size=7, weight="bold", color="white")
        # Adjust text color for light backgrounds
        for j, val in enumerate(values):
            label = labels[j]
            if label_colors[label] in ["#FFFFE0", "#D3D3D3"]:
                autotexts[j].set_color("black")
            
            # Hide if small
            if (val / sum(values)) < 0.05:
                autotexts[j].set_text("")

        ax.set_title(facet_title, fontsize=9, fontweight='bold', pad=5)
        ax.axis('equal')

    # Hide unused
    for j in range(i + 1, len(axes_list)):
        axes_list[j].axis('off')

    # Global Legend - Enlarged per user request
    legend_elements = [plt.Line2D([0], [0], marker='o', color='w', 
                                  markerfacecolor=label_colors[l], 
                                  markersize=12, label=l) for l in sorted_labels]
    
    fig.legend(handles=legend_elements, loc='lower center', ncol=min(len(sorted_labels), 3), 
               bbox_to_anchor=(0.5, 0.02), fontsize=11, frameon=True, facecolor='white', edgecolor='#D3D3D3')

    fig.suptitle(title, fontsize=14, fontweight='bold', y=0.98)
    plt.tight_layout(rect=[0, 0.12, 1, 0.95])
    
    buffer = io.BytesIO()
    plt.savefig(buffer, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buffer.getvalue()).decode('utf-8')


def get_all_charts(df, title, prefix):
    """Helper to generate bar, pie, and line charts for a dataframe."""
    charts = []
    # Prepare DF (Ensure index is set if it's a category)
    df_chart = df.copy()
    
    # Bar Chart
    bar = get_bar_chart_base64(df_chart, title, f"{prefix}_bar")
    if bar: charts.append({"id": f"{prefix}_bar", "name": "Bar Chart", "base64": bar})
    
    # Pie Chart
    pie = get_pie_chart_base64(df_chart, title, f"{prefix}_pie")
    if pie: charts.append({"id": f"{prefix}_pie", "name": "Pie Chart", "base64": pie})
    
    # Line Chart
    line = get_line_chart_base64(df_chart, title, f"{prefix}_line")
    if line: charts.append({"id": f"{prefix}_line", "name": "Line Chart", "base64": line})
    
    return charts


