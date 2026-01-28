import pandas as pd

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
    
    return pd.crosstab(df[temp_loc_col], df[year_col], margins=True, margins_name='Total')

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
    
    return pd.crosstab(df[jurusan_col], df[year_col], margins=True, margins_name='Total')

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
    
    return pd.crosstab(df[prodi_col], df[year_col], margins=True, margins_name='Total')

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
        print(tabulate(df, headers='keys', tablefmt='psql', showindex=True))
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
            
            /* Bold Total Row (last row) */
            tr:last-child td {
                font-weight: bold;
                background-color: #e2e8f0;
                border-top: 2px solid #cbd5e0;
            }
            
            /* Bold Total Column (last column) */
            td:last-child, th:last-child {
                font-weight: bold;
                border-left: 1px solid #e2e8f0;
            }
            
            /* Apply background color ONLY to data cells in the last column */
            td:last-child {
                background-color: #f1f5f9;
            }
            
            /* Ensure the last header cell keeps the blue background */
            th:last-child {
                background-color: #3498db; 
                color: #fff;
                border-left: 1px solid #2980b9;
            }
            
            /* Intersection of Total Row and Column */
            tr:last-child td:last-child {
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
    for title, (df, chart_base64) in data_dict.items():
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
        table_html = df_to_html.to_html(index=False, border=0, classes='table', table_id=table_id)
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
        generate_html_report(dfs_to_report, output_file=REPORT_OUTPUT)
        
    except Exception as e:
        print(f"Error executing main: {e}")
        import traceback
        traceback.print_exc()
