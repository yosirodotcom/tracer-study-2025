import pandas as pd
from viz_utils import sort_crosstab_by_total
from viz_utils import folium, gpd, get_density_color, generate_static_map_geopandas, INDO_COORDS, KALBAR_COORDS

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
    prov_counts = df_working[col_prov].fillna('Belum Mengisi').value_counts().reset_index()
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


