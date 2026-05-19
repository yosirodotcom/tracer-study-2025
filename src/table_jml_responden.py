import pandas as pd
import os

from viz_utils import (
    get_all_charts, get_smooth_trend_chart_base64,
    get_waktu_tunggu_global_smooth_line_chart_base64,
    get_waktu_tunggu_facet_smooth_line_chart_base64,
    get_waktu_tunggu_prodi_facet_smooth_line_chart_base64,
    generate_html_report

)

from viz_lokasi_kampus import create_distribution_campus_loc_tahun
from viz_distribusi_jurusan_prodi import create_distribution_jurusan_tahun, create_distribution_prodi_tahun
from viz_distribusi_masa_tunggu import (
    create_distribution_masa_tunggu_status, create_distribution_waktu_tunggu_jurusan,
    create_waktu_tunggu_prodi_per_jurusan, create_masa_tunggu_prodi_per_jurusan,
    get_masa_tunggu_jurusan_line_chart_base64,
    create_table_masa_tunggu_lt6_jurusan, create_table_masa_tunggu_lt6_prodi,
    get_shaded_line_chart_base64, get_masa_tunggu_lt6_facet_grid_base64
)
from viz_serapan_lulusan import (
    create_serapan_jurusan, create_serapan_prodi_per_jurusan,
    get_serapan_global_pie_chart_base64, get_serapan_divergence_chart_base64,
    get_serapan_prodi_facet_pie_chart_base64, create_serapan_prodi_ranked_table
)
from viz_serapan_alumni_wilayah import create_distribution_provinsi, generate_alumni_map
from viz_serapan_alumni_dudi import create_distribution_kabkota_kalbar, generate_kalbar_map
from viz_pendapatan_gaji import (
    create_salary_distribution, create_salary_by_jurusan, 
    create_salary_prodi_per_jurusan, get_salary_distribution_bell_curve,
    get_salary_jurusan_lollipop_chart, get_salary_distribution_by_prodi_chart
)
from viz_peringkat_peforma import create_jurusan_ranking

if __name__ == "__main__":
    import os
    print("--- Running table_jml_responden.py ---")
    
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    DATA_CLEANED = os.path.join(BASE_DIR, 'data', 'processed', 'cleaned_data.xlsx')
    REPORTS_DIR = os.path.join(BASE_DIR, 'reports')
    REPORT_OUTPUT = os.path.join(REPORTS_DIR, 'report_tables.html')

    file_path = DATA_CLEANED
    if not os.path.exists(file_path):
        DATA_RAW = os.path.join(BASE_DIR, 'data', 'raw', 'data.xlsx')
        file_path = DATA_RAW
        
    try:
        print(f"Loading data from {file_path}...")
        df_load = pd.read_excel(file_path)
        
        while True:
            print("\n" + "="*50)
            print("MENU PILIHAN LAPORAN TRACER STUDY")
            print("="*50)
            print("1. Semua Laporan (Full Report)")
            print("2. Distribusi Lokasi Kampus")
            print("3. Distribusi Jurusan & Prodi")
            print("4. Distribusi Masa Tunggu")
            print("5. Serapan Lulusan")
            print("6. Serapan Alumni (Provinsi)")
            print("7. Serapan Alumni (DUDI / Kalbar)")
            print("8. Pendapatan / Gaji")
            print("9. Peringkat Performa Jurusan")
            print("0. Keluar")
            
            pilihan = input("\nMasukkan pilihan Anda (0-9): ")
            
            if pilihan == '0':
                print("Keluar dari program.")
                break
                
            dfs_to_report = {}
            
            if pilihan in ['1', '2']:
                df_campus = create_distribution_campus_loc_tahun(df_load)
                dfs_to_report["Distribusi Berdasarkan Lokasi Kampus"] = {"df": df_campus, "charts": get_all_charts(df_campus, "Lokasi Kampus", "campus")}
                
            if pilihan in ['1', '3']:
                df_jurusan = create_distribution_jurusan_tahun(df_load)
                dfs_to_report["Distribusi Berdasarkan Jurusan"] = {"df": df_jurusan, "charts": get_all_charts(df_jurusan, "Jurusan", "jurusan")}
                df_prodi = create_distribution_prodi_tahun(df_load)
                dfs_to_report["Distribusi Berdasarkan Program Studi"] = {"df": df_prodi, "charts": get_all_charts(df_prodi, "Program Studi", "prodi")}
                
            if pilihan in ['1', '4']:
                df_masa_tunggu = create_distribution_masa_tunggu_status(df_load)
                if not df_masa_tunggu.empty:
                    charts_list = []
                    trend_chart_b64 = get_smooth_trend_chart_base64(df_load, "Trend Waktu Diterima Bekerja (Bulan)", "masa_tunggu_trend")
                    if trend_chart_b64:
                        charts_list.append({"id": "masa_tunggu_trend", "name": "Trend Chart", "base64": trend_chart_b64})
                    dfs_to_report["Distribusi Masa Tunggu Responden"] = {"df": df_masa_tunggu, "charts": charts_list}
                    
                df_waktu_tunggu = create_distribution_waktu_tunggu_jurusan(df_load)
                if not df_waktu_tunggu.empty:
                    charts_wt = []
                    global_line = get_waktu_tunggu_global_smooth_line_chart_base64(df_load, "Distribusi Waktu Tunggu (Global)", "wt_jurusan_global_line")
                    if global_line:
                        charts_wt.append({"id": "wt_jurusan_global_line", "name": "Line Chart Global", "base64": global_line})
                    facet_line = get_waktu_tunggu_facet_smooth_line_chart_base64(df_load, "Distribusi Waktu Tunggu per Jurusan", "wt_jurusan_facet_line")
                    if facet_line:
                        charts_wt.append({"id": "wt_jurusan_facet_line", "name": "Facet Line Chart", "base64": facet_line})
                    avg_wt_line = get_masa_tunggu_jurusan_line_chart_base64(df_load)
                    if avg_wt_line:
                        charts_wt.append({"id": "wt_jurusan_avg_line", "name": "Line Chart Rata-rata Masa Tunggu", "base64": avg_wt_line})
                    dfs_to_report["Rata-rata Masa Tunggu Lulusan per Jurusan"] = {"df": df_waktu_tunggu, "charts": charts_wt}
                    
                    dict_waktu_tunggu_prodi = create_waktu_tunggu_prodi_per_jurusan(df_load)
                    for table_title, df_jur in dict_waktu_tunggu_prodi.items():
                        dfs_to_report[table_title] = {"df": df_jur, "charts": []}

                dict_masa_tunggu_prodi = create_masa_tunggu_prodi_per_jurusan(df_load)
                for table_title, df_jur in dict_masa_tunggu_prodi.items():
                    pfx = table_title.replace(" ", "_").lower()
                    jurusan = table_title.split(" - ")[-1]
                    charts_list = []
                    facet_chart = get_waktu_tunggu_prodi_facet_smooth_line_chart_base64(df_load, f"Distribusi Waktu Tunggu per Prodi - {jurusan}", f"{pfx}_facet", jurusan=jurusan)
                    if facet_chart:
                        charts_list.append({"id": f"{pfx}_facet", "name": "Facet Line Chart", "base64": facet_chart})
                    dfs_to_report[table_title] = {"df": df_jur, "charts": charts_list}
                    
                # --- Masa Tunggu <= 6 Bulan Section ---
                df_lt6_jur = create_table_masa_tunggu_lt6_jurusan(df_load)
                if not df_lt6_jur.empty:
                    chart_lt6 = get_shaded_line_chart_base64(df_lt6_jur, 'Jurusan', "Persentase Lulusan Masa Tunggu <= 6 Bulan per Jurusan", "lt6_jurusan_shaded")
                    charts = []
                    if chart_lt6:
                        charts.append({"id": "lt6_jurusan_shaded", "name": "Shaded Line Chart", "base64": chart_lt6})
                    dfs_to_report["Lulusan dengan Masa Tunggu <= 6 Bulan per Jurusan"] = {"df": df_lt6_jur.drop(columns=['_pct_numeric']), "charts": charts}
                    
                dict_lt6_prodi = create_table_masa_tunggu_lt6_prodi(df_load)
                
                # Create Facet Grid for all Prodi breakdowns
                facet_grid_base64 = get_masa_tunggu_lt6_facet_grid_base64(dict_lt6_prodi, "Breakdown Masa Tunggu <= 6 Bulan per Program Studi")
                
                for table_title, df_prodi_lt6 in dict_lt6_prodi.items():
                    pfx = table_title.replace(" ", "_").lower()
                    charts = []
                    
                    dfs_to_report[table_title] = {"df": df_prodi_lt6.drop(columns=['_pct_numeric']), "charts": charts}
                
                if facet_grid_base64:
                    dfs_to_report["Grafik Breakdown Masa Tunggu <= 6 Bulan (All Departments)"] = {
                        "df": pd.DataFrame(), 
                        "charts": [{"id": "lt6_prodi_facet", "name": "Facet Grid Chart (4x2)", "base64": facet_grid_base64}]
                    }
                    
            if pilihan in ['1', '5']:
                df_serapan_jurusan = create_serapan_jurusan(df_load)
                if not df_serapan_jurusan.empty:
                    charts_serapan = get_all_charts(df_serapan_jurusan, "Serapan Jurusan", "serapan_jurusan")
                    global_pie = get_serapan_global_pie_chart_base64(df_load)
                    if global_pie:
                        charts_serapan.append({"id": "serapan_global_pie", "name": "Global Pie Chart", "base64": global_pie})
                    
                    dfs_to_report["Serapan Lulusan per Jurusan"] = {"df": df_serapan_jurusan, "charts": charts_serapan}
                
                # Divergence Chart for all Program Studi
                div_chart = get_serapan_divergence_chart_base64(df_load)
                if div_chart:
                    dfs_to_report["Perbandingan Serapan Lulusan per Program Studi (%)"] = {
                        "df": pd.DataFrame(), 
                        "charts": [{"id": "serapan_divergence_all", "name": "Divergence Chart", "base64": div_chart}]
                    }
                
                # Ranked Table for all Program Studi
                df_ranked_prodi = create_serapan_prodi_ranked_table(df_load, apply_fair_sort=True)
                
                # Facet Pie Chart (Stacked Bar) for all Program Studi
                facet_pie_chart = get_serapan_prodi_facet_pie_chart_base64(df_load)
                if facet_pie_chart:
                    dfs_to_report["Distribusi Persentase Serapan Lulusan per Program Studi"] = {
                        "df": df_ranked_prodi, 
                        "charts": [{"id": "serapan_prodi_facet_pie", "name": "Stacked Bar Chart", "base64": facet_pie_chart}]
                    }
                    
                dict_serapan_prodi = create_serapan_prodi_per_jurusan(df_load)
                for jurusan_name, df_jur in dict_serapan_prodi.items():
                    pfx = f"serapan_{jurusan_name.replace(' ', '_').lower()}"
                    df_c = df_jur.set_index('Program Studi')
                    dfs_to_report[jurusan_name] = {"df": df_jur, "charts": get_all_charts(df_c, jurusan_name, pfx)}
                    
            if pilihan in ['1', '6']:
                df_provinsi = create_distribution_provinsi(df_load)
                if not df_provinsi.empty:
                    df_p_chart = df_provinsi.set_index('Provinsi').copy()
                    dfs_to_report["Sebaran Alumni per Provinsi"] = {"df": df_provinsi, "charts": get_all_charts(df_p_chart, "Sebaran Provinsi", "provinsi"), "map": 'Peta_Sebaran_Alumni.html'}
                    generate_alumni_map(df_provinsi, os.path.join(REPORTS_DIR, 'Peta_Sebaran_Alumni.html'))
                    
            if pilihan in ['1', '7']:
                df_kalbar = create_distribution_kabkota_kalbar(df_load)
                if not df_kalbar.empty:
                     df_k_chart = df_kalbar.set_index('Kota/Kabupaten').copy()
                     dfs_to_report["Distribusi Serapan Alumni di DUDI"] = {"df": df_kalbar, "charts": get_all_charts(df_k_chart, "Sebaran Kalbar", "kalbar"), "map": 'Peta_Sebaran_Kalbar.html'}
                     generate_kalbar_map(df_kalbar, os.path.join(REPORTS_DIR, 'Peta_Sebaran_Kalbar.html'))
                     
            if pilihan in ['1', '8']:
                df_salary = create_salary_distribution(df_load)
                if not df_salary.empty:
                    df_s_chart = df_salary.set_index('Rata-rata Pendapatan').copy()
                    charts_salary = get_all_charts(df_s_chart, "Distribusi Gaji", "gaji")
                    
                    # Add Bell Curve Chart
                    bell_curve = get_salary_distribution_bell_curve(df_load)
                    if bell_curve:
                        charts_salary.insert(0, {"id": "gaji_bell_curve", "name": "Bell Curve", "base64": bell_curve})
                        
                    dfs_to_report["Distribusi Rata-rata Pendapatan Lulusan per Bulan"] = {"df": df_salary, "charts": charts_salary}

                result_salary = create_salary_by_jurusan(df_load)
                if result_salary and not result_salary[0].empty:
                    df_salary_display, df_salary_ranked = result_salary
                    df_sj_chart = df_salary_ranked.set_index('Jurusan').copy()
                    charts_sj = [
                        {"id": "gaji_jurusan_lollipop", "name": "Lollipop Chart", "base64": get_salary_jurusan_lollipop_chart(df_sj_chart)}
                    ]
                    # Also keep the other charts if desired, but Lollipop is the main one now
                    charts_sj.extend(get_all_charts(df_sj_chart, "Gaji per Jurusan", "gaji_jurusan"))
                    
                    dfs_to_report["Rata-rata Gaji Lulusan per Jurusan"] = {"df": df_salary_display, "charts": charts_sj}
                    
                    dict_salary_prodi = create_salary_prodi_per_jurusan(df_load)
                    for table_title, df_jur in dict_salary_prodi.items():
                        pfx = table_title.replace(" ", "_").lower()
                        df_c = df_jur.set_index('Program Studi')
                        dfs_to_report[table_title] = {"df": df_jur, "charts": get_all_charts(df_c, table_title, pfx)}

                    # Add Stacked Bar Chart for Salary by Prodi
                    salary_prodi_stacked = get_salary_distribution_by_prodi_chart(df_load)
                    if salary_prodi_stacked:
                        dfs_to_report["Visualisasi Distribusi Pendapatan per Program Studi"] = {
                            "df": pd.DataFrame(), 
                            "charts": [{"id": "salary_prodi_stacked", "name": "Stacked Bar Chart", "base64": salary_prodi_stacked}]
                        }

            if pilihan in ['1', '9']:
                df_ranking = create_jurusan_ranking()
                if not df_ranking.empty:
                    df_r_chart = df_ranking.set_index('Jurusan').copy()
                    dfs_to_report["Peringkat Performa Jurusan - Tracer Study 2025"] = {"df": df_ranking, "charts": get_all_charts(df_r_chart, "Ranking Jurusan", "ranking")}

            if not dfs_to_report:
                if pilihan not in ['1', '2', '3', '4', '5', '6', '7', '8', '9']:
                    print("Pilihan tidak valid, silakan coba lagi.")
                continue

            print("\nGenerating HTML report...")
            generate_html_report(dfs_to_report, output_file=REPORT_OUTPUT)
            print("Laporan selesai dibuat!\n")
        
    except Exception as e:
        
        print(f"Error executing main: {e}")
        import traceback
        traceback.print_exc()

