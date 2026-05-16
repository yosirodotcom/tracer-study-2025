import pandas as pd
import matplotlib.pyplot as plt
import io
import base64
import numpy as np
import os
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from viz_utils import sort_crosstab_by_total

URUTAN_KOLOM = ['< 3 Bulan', '3 - 6 Bulan', '6 - 12 Bulan', '>12 bulan']

def _kategorisasi(bulan):
    if bulan < 3: return '< 3 Bulan'
    elif bulan <= 6: return '3 - 6 Bulan'
    elif bulan <= 12: return '6 - 12 Bulan'
    else: return '>12 bulan'

def _fmt(n, denom):
    pct = (n / denom * 100) if denom > 0 else 0
    return f"{int(n)} ({pct:.1f}%)"

def _prepare_working_df(df):
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    df_w = df[df[col_status].isin(working_status)].copy()
    df_w['mt_numeric'] = pd.to_numeric(df_w[col_masa_tunggu], errors='coerce')
    mean_val = df_w['mt_numeric'].mean()
    df_w['mt_numeric'] = df_w['mt_numeric'].fillna(mean_val)
    df_w['Kategori'] = df_w['mt_numeric'].apply(_kategorisasi)
    return df_w, col_status, col_masa_tunggu


def create_distribution_masa_tunggu_status(df):
    """Table: Status Pekerjaan vs 4 buckets + Total + Mean."""
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    if col_status not in df.columns or col_masa_tunggu not in df.columns:
        return pd.DataFrame()

    total_respondents = len(df)
    df_w, _, _ = _prepare_working_df(df)

    # Build numeric pivot
    ct = pd.crosstab(df_w[col_status], df_w['Kategori'])
    for col in URUTAN_KOLOM:
        if col not in ct.columns: ct[col] = 0
    ct = ct[URUTAN_KOLOM]
    ct['Total'] = ct.sum(axis=1)
    means = df_w.groupby(col_status)['mt_numeric'].mean().round(1)
    ct['Mean'] = means

    # TOTAL row (numeric)
    total_counts = ct[URUTAN_KOLOM].sum()
    total_counts['Total'] = ct['Total'].sum()
    total_counts['Mean'] = df_w['mt_numeric'].mean().round(1)
    ct.loc['TOTAL'] = total_counts

    # Format into display strings
    rows = []
    for idx, row in ct.iterrows():
        r = {'Status Pekerjaan': idx}
        for col in URUTAN_KOLOM:
            r[col] = _fmt(row[col], total_respondents)
        r['Total'] = _fmt(row['Total'], total_respondents)
        r['Rata-rata masa tunggu bulan'] = row['Mean']
        rows.append(r)

    final_cols = ['Status Pekerjaan'] + URUTAN_KOLOM + ['Total', 'Rata-rata masa tunggu bulan']
    return pd.DataFrame(rows, columns=final_cols)


def create_distribution_waktu_tunggu_jurusan(df):
    """Table: Jurusan vs 4 buckets + Total + Mean."""
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_jurusan = 'Jurusan'
    if col_jurusan not in df.columns or col_status not in df.columns or col_masa_tunggu not in df.columns:
        return pd.DataFrame()

    total_respondents = len(df)
    df_w, _, _ = _prepare_working_df(df)

    # Build numeric pivot
    ct = pd.crosstab(df_w[col_jurusan], df_w['Kategori'])
    for col in URUTAN_KOLOM:
        if col not in ct.columns: ct[col] = 0
    ct = ct[URUTAN_KOLOM]
    ct['Total'] = ct.sum(axis=1)
    means = df_w.groupby(col_jurusan)['mt_numeric'].mean().round(1)
    ct['Mean'] = means

    # TOTAL row (numeric)
    total_counts = ct[URUTAN_KOLOM].sum()
    total_counts['Total'] = ct['Total'].sum()
    total_counts['Mean'] = df_w['mt_numeric'].mean().round(1)
    ct.loc['TOTAL / RATA-RATA INSTITUSI'] = total_counts

    # Format into display strings
    rows = []
    for idx, row in ct.iterrows():
        r = {'Jurusan': idx}
        for col in URUTAN_KOLOM:
            r[col] = _fmt(row[col], total_respondents)
        r['Total'] = _fmt(row['Total'], total_respondents)
        r['Rata-rata masa tunggu bulan'] = row['Mean']
        rows.append(r)

    final_cols = ['Jurusan'] + URUTAN_KOLOM + ['Total', 'Rata-rata masa tunggu bulan']
    return pd.DataFrame(rows, columns=final_cols)


def create_waktu_tunggu_prodi_per_jurusan(df):
    """Dict of Tables: Prodi vs 4 buckets + Total + Mean, one per Jurusan."""
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi' if 'prodi' in df.columns else 'Program Studi'
    if col_jurusan not in df.columns or col_prodi not in df.columns or col_status not in df.columns or col_masa_tunggu not in df.columns:
        return {}

    total_respondents = len(df)
    df_w, _, _ = _prepare_working_df(df)

    final_cols = ['Prodi'] + URUTAN_KOLOM + ['Total', 'Rata-rata masa tunggu bulan']
    results = {}

    for jurusan in sorted(df[col_jurusan].unique()):
        df_jur = df_w[df_w[col_jurusan] == jurusan]
        if df_jur.empty:
            continue

        ct = pd.crosstab(df_jur[col_prodi], df_jur['Kategori'])
        for col in URUTAN_KOLOM:
            if col not in ct.columns: ct[col] = 0
        ct = ct[URUTAN_KOLOM]
        ct['Total'] = ct.sum(axis=1)
        means = df_jur.groupby(col_prodi)['mt_numeric'].mean().round(1)
        ct['Mean'] = means

        # TOTAL row (numeric)
        total_counts = ct[URUTAN_KOLOM].sum()
        total_counts['Total'] = ct['Total'].sum()
        total_counts['Mean'] = df_jur['mt_numeric'].mean().round(1)
        ct.loc[f'TOTAL {jurusan}'] = total_counts

        # Format into display strings
        rows = []
        for idx, row in ct.iterrows():
            r = {'Prodi': idx}
            for col in URUTAN_KOLOM:
                r[col] = _fmt(row[col], total_respondents)
            r['Total'] = _fmt(row['Total'], total_respondents)
            r['Rata-rata masa tunggu bulan'] = row['Mean']
            rows.append(r)

        results[f"Rata-rata Waktu Tunggu - {jurusan}"] = pd.DataFrame(rows, columns=final_cols)

    return results


def create_masa_tunggu_prodi_per_jurusan(df):
    """Alias for backward compatibility."""
    return create_waktu_tunggu_prodi_per_jurusan(df)


def get_masa_tunggu_jurusan_line_chart_base64(df):
    """Line chart: average waiting time per Jurusan."""
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_status = 'Jelaskan status Anda saat ini?'
    col_jurusan = 'Jurusan'
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']

    if col_masa_tunggu not in df.columns or col_jurusan not in df.columns or col_status not in df.columns:
        return None

    df_w = df[df[col_status].isin(working_status)].copy()
    df_w['mt_numeric'] = pd.to_numeric(df_w[col_masa_tunggu], errors='coerce')
    df_w['mt_numeric'] = df_w['mt_numeric'].fillna(df_w['mt_numeric'].mean())
    if df_w.empty:
        return None

    stats = df_w.groupby(col_jurusan)['mt_numeric'].mean().reset_index()
    stats.columns = ['Jurusan', 'Rata-rata']
    stats = stats.sort_values('Rata-rata', ascending=True)

    plt.style.use('seaborn-v0_8-whitegrid')
    fig, ax = plt.subplots(figsize=(12, 6))
    x = stats['Jurusan'].astype(str).tolist()
    y = stats['Rata-rata'].tolist()
    
    # Plot the line and all dots
    ax.plot(x, y, marker='o', color='#00008B', linewidth=2.5, markersize=6)
    
    # Highlight the smallest average (first element because sorted ascending)
    if len(x) > 0:
        ax.plot(x[0], y[0], marker='o', color='green', markersize=12, zorder=5)

    ax.fill_between(x, 0, y, color='#00008B', alpha=0.1)
    ax.set_title("Rata-rata Masa Tunggu per Jurusan (Bulan)", fontsize=14, fontweight='bold', pad=20)
    ax.set_ylabel("Bulan", fontsize=11)
    ax.set_xlabel("Jurusan", fontsize=11)
    
    ax.set_xticks(range(len(x)))
    xtick_labels = ax.set_xticklabels(x, rotation=45, ha='right')
    
    if len(xtick_labels) > 0:
        xtick_labels[0].set_fontsize(14)
        xtick_labels[0].set_fontweight('bold')
        xtick_labels[0].set_color('darkblue')

    for i, val in enumerate(y):
        if i == 0:
            # First element: larger font and larger offset to avoid overlap with large green dot
            ax.text(i, val + (max(y) * 0.06), f'{val:.1f}', ha='center', va='bottom', 
                    fontsize=13, fontweight='bold')
        else:
            ax.text(i, val + (max(y) * 0.02), f'{val:.1f}', ha='center', va='bottom', 
                    fontsize=10, fontweight='bold')
    ax.set_ylim(0, max(y) * 1.15 if len(y) > 0 else 1)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode('utf-8')

def create_table_masa_tunggu_lt6_jurusan(df):
    """Table for waiting time <= 6 months per Jurusan."""
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_jurusan = 'Jurusan'
    if col_jurusan not in df.columns or col_status not in df.columns or col_masa_tunggu not in df.columns:
        return pd.DataFrame()

    # Pre-calculate mt_numeric for filtering
    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    df_w = df.copy()
    df_w['mt_numeric'] = pd.to_numeric(df_w[col_masa_tunggu], errors='coerce')
    mean_val = df_w[df_w[col_status].isin(working_status)]['mt_numeric'].mean()
    df_w['mt_numeric'] = df_w['mt_numeric'].fillna(mean_val)

    # Base counts (denominator)
    base_counts = df.groupby(col_jurusan).size()
    
    # Filter for <= 6 months
    df_lt6 = df_w[df_w['mt_numeric'] <= 6]
    
    # Bekerja count
    bekerja_lt6 = df_lt6[df_lt6[col_status] == 'Bekerja (Full time/Part time)'].groupby(col_jurusan).size()
    # Wiraswasta count
    wiraswasta_lt6 = df_lt6[df_lt6[col_status] == 'Wiraswasta'].groupby(col_jurusan).size()
    # Total count
    total_lt6 = df_lt6[df_lt6[col_status].isin(working_status)].groupby(col_jurusan).size()
    
    rows = []
    for jur in sorted(df[col_jurusan].unique()):
        n_total = base_counts.get(jur, 0)
        n_bek = bekerja_lt6.get(jur, 0)
        n_wir = wiraswasta_lt6.get(jur, 0)
        n_sum = total_lt6.get(jur, 0)
        
        rows.append({
            'Jurusan': jur,
            'Jumlah responden': n_total,
            'Bekerja': _fmt(n_bek, n_total),
            'Wiraswasta': _fmt(n_wir, n_total),
            'Total (Bekerja + Wiraswasta)': _fmt(n_sum, n_total),
            '_pct_numeric': (n_sum / n_total * 100) if n_total > 0 else 0
        })
        
    # TOTAL row
    n_total_all = len(df)
    n_bek_all = (df_lt6[col_status] == 'Bekerja (Full time/Part time)').sum()
    n_wir_all = (df_lt6[col_status] == 'Wiraswasta').sum()
    n_sum_all = df_lt6[col_status].isin(working_status).sum()
    
    rows.append({
        'Jurusan': 'TOTAL INSTITUSI',
        'Jumlah responden': n_total_all,
        'Bekerja': _fmt(n_bek_all, n_total_all),
        'Wiraswasta': _fmt(n_wir_all, n_total_all),
        'Total (Bekerja + Wiraswasta)': _fmt(n_sum_all, n_total_all),
        '_pct_numeric': (n_sum_all / n_total_all * 100) if n_total_all > 0 else 0
    })
    
    return pd.DataFrame(rows)

def create_table_masa_tunggu_lt6_prodi(df):
    """Breakdown table for waiting time <= 6 months per Prodi."""
    col_status = 'Jelaskan status Anda saat ini?'
    col_masa_tunggu = 'Dalam berapa bulan Anda mendapatkan pekerjaan? Tulis dengan angka (Contoh: 1, 1Tahun = 12 bulan) rev2'
    col_jurusan = 'Jurusan'
    col_prodi = 'prodi' if 'prodi' in df.columns else 'Program Studi'
    if col_jurusan not in df.columns or col_prodi not in df.columns:
        return {}

    working_status = ['Bekerja (Full time/Part time)', 'Wiraswasta']
    df_w = df.copy()
    df_w['mt_numeric'] = pd.to_numeric(df_w[col_masa_tunggu], errors='coerce')
    mean_val = df_w[df_w[col_status].isin(working_status)]['mt_numeric'].mean()
    df_w['mt_numeric'] = df_w['mt_numeric'].fillna(mean_val)

    results = {}
    for jurusan in sorted(df[col_jurusan].unique()):
        df_j = df[df[col_jurusan] == jurusan]
        df_j_w = df_w[df_w[col_jurusan] == jurusan]
        
        base_counts = df_j.groupby(col_prodi).size()
        df_lt6 = df_j_w[df_j_w['mt_numeric'] <= 6]
        
        bekerja_lt6 = df_lt6[df_lt6[col_status] == 'Bekerja (Full time/Part time)'].groupby(col_prodi).size()
        wiraswasta_lt6 = df_lt6[df_lt6[col_status] == 'Wiraswasta'].groupby(col_prodi).size()
        total_lt6 = df_lt6[df_lt6[col_status].isin(working_status)].groupby(col_prodi).size()
        
        rows = []
        for prodi in sorted(df_j[col_prodi].unique()):
            n_total = base_counts.get(prodi, 0)
            n_bek = bekerja_lt6.get(prodi, 0)
            n_wir = wiraswasta_lt6.get(prodi, 0)
            n_sum = total_lt6.get(prodi, 0)
            
            rows.append({
                'Prodi': prodi,
                'Jumlah responden': n_total,
                'Bekerja': _fmt(n_bek, n_total),
                'Wiraswasta': _fmt(n_wir, n_total),
                'Total (Bekerja + Wiraswasta)': _fmt(n_sum, n_total),
                '_pct_numeric': (n_sum / n_total * 100) if n_total > 0 else 0
            })
            
        # TOTAL row for this jurusan
        n_total_j = len(df_j)
        n_bek_j = (df_lt6[col_status] == 'Bekerja (Full time/Part time)').sum()
        n_wir_j = (df_lt6[col_status] == 'Wiraswasta').sum()
        n_sum_j = df_lt6[col_status].isin(working_status).sum()
        
        rows.append({
            'Prodi': f'TOTAL {jurusan}',
            'Jumlah responden': n_total_j,
            'Bekerja': _fmt(n_bek_j, n_total_j),
            'Wiraswasta': _fmt(n_wir_j, n_total_j),
            'Total (Bekerja + Wiraswasta)': _fmt(n_sum_j, n_total_j),
            '_pct_numeric': (n_sum_j / n_total_j * 100) if n_total_j > 0 else 0
        })
        
        results[f"Masa Tunggu <= 6 Bulan - {jurusan}"] = pd.DataFrame(rows)
        
    return results

def get_shaded_line_chart_base64(df_table, category_col, title, chart_id):
    """Horizontal bar chart with worker clipart at the end of each bar."""
    # Exclude TOTAL row
    df_plot = df_table[~df_table[category_col].str.contains('TOTAL', case=False, na=False)].copy()
    if df_plot.empty: return None
    
    # Sort by percentage
    df_plot = df_plot.sort_values('_pct_numeric', ascending=True)
    
    plt.style.use('default')
    # Increased height multiplier for more vertical space
    fig, ax = plt.subplots(figsize=(14, len(df_plot) * 1.8 + 2))
    
    categories = df_plot[category_col].astype(str).tolist()
    values = df_plot['_pct_numeric'].tolist()
    
    # Spacing out the y-coordinates
    y_coords = np.arange(len(categories)) * 2.0
    
    # Colors: Dark blue for max, White for others (to contrast with grey bg)
    max_val = max(values) if values else 0
    colors = ['#1f4e79' if v == max_val else '#ffffff' for v in values]
    
    # Add standard rectangular background patch
    from matplotlib.patches import Rectangle
    # Increase height for top margin
    bg_height = y_coords[-1] + 2.5
    p_rect = Rectangle((-12, -1.0), 132, bg_height,
                       ec="none", fc="#e0e0e0", zorder=-2)
    ax.add_patch(p_rect)
    
    # Draw horizontal lines (lollipop stalks)
    ax.hlines(y=y_coords, xmin=0, xmax=values, color=colors, linewidth=12, alpha=0.9)
    
    # Remove grid and spines
    ax.grid(False)
    for spine in ax.spines.values():
        spine.set_visible(False)
    
    # Remove ticks and default labels
    ax.set_yticks([])
    ax.tick_params(axis='both', which='both', length=0)
    
    # Add clipart, circles, and labels
    try:
        # Prioritize PNG for transparency
        img_path = os.path.join('assets', 'gambar', 'running_worker.png')
        if not os.path.exists(img_path):
             img_path = os.path.join('d:\\repos\\tracer-study-2025', 'assets', 'gambar', 'running_worker.png')
        
        if not os.path.exists(img_path):
             img_path = os.path.join('assets', 'gambar', 'running_worker.jpg')
             
        worker_img = plt.imread(img_path)
    except:
        worker_img = None
        
    for i, val in enumerate(values):
        y = y_coords[i]
        # Add Prodi name ABOVE the line
        ax.text(0, y + 0.45, categories[i], va='bottom', ha='left', fontsize=16, fontweight='semibold', color='#333333')
        
        # Draw circle at the end of the line
        ax.scatter(val, y, color=colors[i], s=2500, zorder=3)
        
        # Determine text color based on circle color
        txt_color = 'white' if val == max_val else '#333333'
        
        # Add text label inside the circle
        ax.text(val, y, f'{val:.0f}%', va='center', ha='center', fontsize=14, fontweight='bold', color=txt_color)
        
        # Add clipart
        if worker_img is not None:
            imagebox = OffsetImage(worker_img, zoom=0.06) 
            ab = AnnotationBbox(imagebox, (val, y), frameon=False, xybox=(60, 0), xycoords='data', boxcoords="offset points", pad=0)
            ax.add_artist(ab)
            
    ax.set_title(title, fontsize=28, fontweight='bold', pad=60, loc='left')
    ax.set_xlim(-12, 120) 
    ax.set_ylim(-1.5, y_coords[-1] + 2.0)
    
    # Remove x-axis labels as they might be redundant with data labels
    ax.set_xticks([])
    
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', transparent=True)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode('utf-8')

def get_masa_tunggu_lt6_facet_grid_base64(dict_df_prodi, title):
    """Facet grid (4x2) of horizontal bar charts for prodi breakdown."""
    jurusans = sorted(dict_df_prodi.keys())
    if not jurusans: return None
    
    n_cols = 2
    n_rows = 4
    
    # Increased figure height for more breathing room
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(20, 6.5 * n_rows))
    axes_flat = axes.flatten()
    
    plt.style.use('default')
    
    # Load clipart
    try:
        # Prioritize PNG for transparency
        img_path = os.path.join('assets', 'gambar', 'running_worker.png')
        if not os.path.exists(img_path):
             img_path = os.path.join('d:\\repos\\tracer-study-2025', 'assets', 'gambar', 'running_worker.png')
        
        if not os.path.exists(img_path):
             img_path = os.path.join('assets', 'gambar', 'running_worker.jpg')
             
        worker_img = plt.imread(img_path)
    except:
        worker_img = None

    idx = 0
    for table_title in jurusans:
        if idx >= len(axes_flat): break
        ax = axes_flat[idx]
        df_table = dict_df_prodi[table_title]
        jurusan_name = table_title.split(' - ')[-1]
        
        # Exclude TOTAL row
        df_plot = df_table[~df_table['Prodi'].str.contains('TOTAL', case=False, na=False)].copy()
        if df_plot.empty:
            ax.set_visible(False)
            continue
            
        df_plot = df_plot.sort_values('_pct_numeric', ascending=True)
        categories = df_plot['Prodi'].astype(str).tolist()
        values = df_plot['_pct_numeric'].tolist()
        
        # Spacing out y-coordinates
        y_coords = np.arange(len(categories)) * 2.5
        
        max_val = max(values) if values else 0
        colors = ['#1f4e79' if v == max_val else '#ffffff' for v in values]
        
        # Add standard rectangular background patch
        from matplotlib.patches import Rectangle
        # Increase height for top margin
        bg_height = y_coords[-1] + 3.0
        p_rect = Rectangle((-12, -1.2), 132, bg_height,
                           ec="none", fc="#e0e0e0", zorder=-2)
        ax.add_patch(p_rect)
        
        # Draw horizontal lines (lollipop stalks)
        ax.hlines(y=y_coords, xmin=0, xmax=values, color=colors, linewidth=10, alpha=0.8)
        
        ax.grid(False)
        for spine in ax.spines.values():
            spine.set_visible(False)
        
        ax.set_yticks([])
        ax.tick_params(axis='both', which='both', length=0)
        ax.set_xticks([])
        
        # Data labels, circles, and clipart
        for i, val in enumerate(values):
            y = y_coords[i]
            # Add Prodi name ABOVE the line - 1.5x original
            ax.text(0, y + 0.6, categories[i], va='bottom', ha='left', fontsize=13, fontweight='semibold', color='#333333')
            
            # Draw circle at the end
            ax.scatter(val, y, color=colors[i], s=1400, zorder=3)
            
            txt_color = 'white' if val == max_val else '#333333'
            # Add text label inside the circle
            ax.text(val, y, f'{val:.0f}%', va='center', ha='center', fontsize=12, fontweight='bold', color=txt_color)
            
            if worker_img is not None:
                imagebox = OffsetImage(worker_img, zoom=0.05) 
                ab = AnnotationBbox(imagebox, (val, y), frameon=False, xybox=(50, 0), xycoords='data', boxcoords="offset points", pad=0)
                ax.add_artist(ab)
        
        ax.set_title(jurusan_name, fontsize=24, fontweight='bold', pad=40)
        ax.set_xlim(-12, 120) 
        ax.set_ylim(-2.0, y_coords[-1] + 2.5)
        idx += 1

    # Hide unused axes
    for i in range(idx, len(axes_flat)):
        axes_flat[i].set_visible(False)
        
    plt.suptitle(title, fontsize=36, fontweight='bold', y=1.02)
    plt.tight_layout(rect=[0, 0, 1, 1])
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', transparent=True)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode('utf-8')
