import pandas as pd
from viz_utils import sort_crosstab_by_total

def create_jurusan_ranking():
    """
    Creates a dataframe for Jurusan Ranking based on provided qualitative/quantitative analysis.
    """
    data = [
        {
            "Peringkat": 1,
            "Jurusan": "Teknik Arsitektur",
            "Skor Kekuatan": "⭐⭐⭐⭐",
            "Total": 4,
            "Predikat & Analisis": """<b>"The High Quality Performer"</b><br>Meskipun jumlah respondennya paling sedikit (52), namun jurusan ini JUARA 1 di dua kategori sekaligus: Kecepatan Serapan (77,5%) dan Gaji Tertinggi (Rp 3,8 Juta). Kualitas lulusannya sangat premium di mata pasar."""
        },
        {
            "Peringkat": 2,
            "Jurusan": "Akuntansi",
            "Skor Kekuatan": "⭐⭐⭐⭐",
            "Total": 4,
            "Predikat & Analisis": """<b>"The Major Contributor"</b><br>Naik signifikan ke peringkat 2 berkat Volume Responden Tertinggi (226 orang) yang mendominasi 30% data survei. Selain itu, kecepatan serapannya sangat baik (Juara 2). Nilai minus hanya pada gaji awal yang masih entry level."""
        },
        {
            "Peringkat": 3,
            "Jurusan": "Teknik Sipil & Perencanaan",
            "Skor Kekuatan": "⭐⭐⭐",
            "Total": 3,
            "Predikat & Analisis": """<b>"The Fastest Hired"</b><br>Unggul mutlak di Masa Tunggu Tercepat (3,4 bulan). Sangat efisien dalam mengantarkan lulusan ke dunia kerja, meskipun volume responden dan gaji berada di level menengah."""
        },
        {
            "Peringkat": 4,
            "Jurusan": "Teknik Mesin",
            "Skor Kekuatan": "⭐⭐⭐",
            "Total": 3,
            "Predikat & Analisis": """<b>"The High Valued Entrepreneur"</b><br>Unggul di Gaji Tertinggi (Rp 3,8 Juta) dan jumlah responden yang besar (110 orang). Peringkatnya tertahan karena masa tunggu rata-rata yang cukup lama (6,4 bulan), namun ini terkompensasi oleh tingginya angka wirausaha."""
        },
        {
            "Peringkat": 5,
            "Jurusan": "Ilmu Kelautan & Perikanan",
            "Skor Kekuatan": "⭐⭐",
            "Total": 2,
            "Predikat & Analisis": """<b>"The Balanced Niche"</b><br>Memiliki performa yang seimbang di semua lini. Tidak terlalu menonjol di satu sisi, tapi cukup stabil dalam serapan (71,8%) dan masa tunggu (5,2 bulan)."""
        },
        {
            "Peringkat": 6,
            "Jurusan": "Teknik Elektro",
            "Skor Kekuatan": "⭐⭐",
            "Total": 2,
            "Predikat & Analisis": """<b>"The Steady Player"</b><br>Konsisten di papan tengah. Memiliki gaji yang cukup baik (Rp 3,5 Juta) di atas rata-rata institusi, namun butuh peningkatan dalam kecepatan serapan (67,5%)."""
        },
        {
            "Peringkat": 7,
            "Jurusan": "Administrasi Bisnis",
            "Skor Kekuatan": "⭐⭐",
            "Total": 2,
            "Predikat & Analisis": """<b>"The Salary Surprise"</b><br>Meskipun secara ranking umum ada di bawah, jurusan ini punya keunggulan Gaji Tinggi (Rp 3,7 Juta - Peringkat 3). Tantangannya ada pada masa tunggu yang paling lama (6,4 bulan) dan volume responden yang moderat."""
        },
        {
            "Peringkat": 8,
            "Jurusan": "Teknologi Pertanian",
            "Skor Kekuatan": "⭐",
            "Total": 1,
            "Predikat & Analisis": """<b>"The Job Creator"</b><br>Secara statistik "pekerja", jurusan ini ada di bawah (gaji & kecepatan rendah). NAMUN, perlu dicatat: Jurusan ini adalah Raja Wirausaha (21 orang). Indikator ranking ini bias ke "karyawan", sehingga potensi wirausaha Pertanian tidak terpotret penuh di sini."""
        }
    ]
    
    df = pd.DataFrame(data)
    return df


