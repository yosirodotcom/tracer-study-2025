import pandas as pd
from viz_utils import sort_crosstab_by_total
from viz_utils import folium, gpd, get_density_color, generate_static_map_geopandas, INDO_COORDS, KALBAR_COORDS

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


