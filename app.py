import streamlit as st
import pandas as pd
import numpy as np
import datetime
import zipfile
import xml.etree.ElementTree as ET
import re
import os

# ====================================================
# 1. KONFIGURASI TAMPILAN WEB (PREMIUM EMERALD & GOLD)
# ====================================================
st.set_page_config(page_title="AgriWarning - Siap Tanam", page_icon="🌾", layout="centered")

st.markdown("""
    <style>
    /* 1. BACKGROUND UTAMA (Emerald Green Bersih Tanpa Pola) */
    .stApp { 
        background-color: #024B30 !important; /* Hijau Emerald Gelap / Premium */
    }

    /* 2. TIPOGRAFI GLOBAL (Emas Premium) */
    html, body, p, span, h1, h2, h3, h4, h5, h6 { 
        color: #D4AF37 !important; /* Warna Emas Classic */
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }

    /* 3. HEADER & SUBHEADER */
    .judul-besar { 
        font-size: 42px; 
        font-weight: 900; 
        /* Gradasi Emas Metallic */
        background: linear-gradient(90deg, #BF953F, #FCF6BA, #B38728, #FBF5B7, #AA771C);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center; 
        margin-bottom: 0px; 
        letter-spacing: -1px;
    }
    .sub-judul { 
        font-size: 17px; 
        color: #F8E287 !important; /* Emas Muda/Champagne */
        text-align: center; 
        margin-bottom: 35px; 
        font-weight: 600; 
        letter-spacing: 0.5px;
    }

    /* 4. KARTU METRIK CUACA (Dark Emerald + Gold Border) */
    div[data-testid="metric-container"] { 
        background-color: #013622; /* Emerald lebih pekat */
        border: 1px solid #AA771C;
        border-top: 5px solid #D4AF37; /* Garis atas emas */
        padding: 20px; 
        border-radius: 12px; 
        box-shadow: 0 10px 20px rgba(0,0,0,0.3); 
        transition: transform 0.2s ease-in-out;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 14px 28px rgba(212, 175, 55, 0.15); /* Shadow emas halus saat di-hover */
    }
    div[data-testid="stMetricValue"] {
        color: #FCF6BA !important; /* Angka warna emas terang */
        font-weight: 800 !important;
        font-size: 32px !important;
    }

    /* 5. KARTU HAMA */
    .hama-card { 
        background-color: #013622; 
        padding: 20px; 
        border-radius: 12px; 
        margin-bottom: 15px; 
        box-shadow: 0 6px 12px rgba(0,0,0,0.2); 
        border: 1px solid #4A5D23; 
    }
    /* Warna indikator disesuaikan agar cocok dengan background gelap */
    .hama-aman { border-left: 6px solid #2ECC71; }
    .hama-waspada { border-left: 6px solid #F1C40F; }
    .hama-card.hama-kritis { border-left: 6px solid #E74C3C; }

    /* 6. TABEL STREAMLIT */
    table {
        width: 100%;
        border-collapse: collapse;
        background-color: #013622;
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        font-size: 14px;
        border: 1px solid #D4AF37;
    }
    thead tr th {
        background-color: #1A2E1A !important;
        color: #D4AF37 !important;
        font-weight: 600 !important;
        padding: 12px 15px;
        border-bottom: 2px solid #D4AF37 !important;
    }
    tbody tr td {
        color: #E8D399 !important;
        padding: 12px 15px;
        border-bottom: 1px solid #035C3A !important;
    }
    tbody tr:hover td {
        background-color: #024B30 !important;
    }

    /* 7. EXPANDER (Menu Buka-Tutup - Garis Hitam) */
    div[data-testid="stExpander"] {
        border: 1px solid #000000 !important; /* Garis luar kotak menjadi hitam solid */
        border-radius: 8px !important;
        background-color: #013622 !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.4) !important; /* Bayangan hitam agar lebih menonjol */
        overflow: hidden;
    }
    
    .streamlit-expanderHeader {
        background-color: #013622 !important;
        border: none !important; /* Menghilangkan garis bawaan agar tidak nabrak */
        color: #D4AF37 !important;
        font-weight: 700 !important;
        padding: 12px 18px !important;
    }

    /* Mengubah garis pemisah bagian dalam (saat menu dibuka) menjadi hitam */
    div[data-testid="stExpanderDetails"] {
        border-top: 1px solid #000000 !important; 
        background-color: #013622 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Menampilkan Judul
st.markdown('<h1 class="judul-besar">AgriWarning</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-judul">Sistem Prediksi Hama & Agroklimatologi Modern (2019 - 2024)</p>', unsafe_allow_html=True)

# ====================================================
# 2. FUNGSI BAWAAN (PARSING EXCEL & MODUL GDD)
# ====================================================
def xlsx_sheet_to_df(path, sheet_xml='xl/worksheets/sheet1.xml'):
    ns = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    try:
        with zipfile.ZipFile(path) as z:
            shared = []
            if 'xl/sharedStrings.xml' in z.namelist():
                root_ss = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in root_ss.findall('.//a:si', ns):
                    texts = [t.text or '' for t in si.findall('.//a:t', ns)]
                    shared.append(''.join(texts))
            root = ET.fromstring(z.read(sheet_xml))
            rows = {}
            max_col = 0
            for row in root.findall('.//a:sheetData/a:row', ns):
                r = int(row.attrib['r'])
                cells = {}
                for c in row.findall('a:c', ns):
                    ref = c.attrib.get('r')
                    col_letters = ''.join(ch for ch in ref if ch.isalpha())
                    col_num = 0
                    for ch in col_letters: col_num = col_num * 26 + (ord(ch.upper()) - 64)
                    max_col = max(max_col, col_num)
                    val = c.find('a:v', ns)
                    v_text = val.text if val is not None else (c.find('a:is/a:t', ns).text if c.find('a:is/a:t', ns) is not None else None)
                    if val is not None and c.attrib.get('t') == 's': v_text = shared[int(v_text)]
                    cells[col_num] = v_text
                rows[r] = cells
            matrix = []
            for r in sorted(rows):
                matrix.append([rows[r].get(i) for i in range(1, max_col + 1)])
            return pd.DataFrame(matrix)
    except: return pd.DataFrame()

def excel_serial_to_datetime(series):
    serial = pd.to_numeric(series, errors='coerce')
    return pd.to_datetime('1899-12-30') + pd.to_timedelta(serial, unit='D')

def get_dasarian(hari):
    if hari <= 10: return 'I'
    elif hari <= 20: return 'II'
    else: return 'III'

def get_opt_config(nama_opt):
    opt = str(nama_opt).upper()
    if 'HYPOTHENEMUS' in opt or 'PBKO' in opt or 'BUAH KOPI' in opt: return {'t_base': 14.9, 'fase': {'Telur': 68, 'Larva (Merusak)': 254, 'Pupa': 318, 'Imago Dewasa': 400}, 'kritis': 254}
    elif 'SPODOPTERA' in opt or 'GRAYAK' in opt: return {'t_base': 10.0, 'fase': {'Telur': 54, 'Larva (Sangat Merusak)': 314, 'Pupa': 454, 'Imago Dewasa': 550}, 'kritis': 314}
    elif 'CONOPOMORPHA' in opt or 'PBK' in opt: return {'t_base': 12.0, 'fase': {'Telur': 50, 'Larva (Merusak)': 280, 'Pupa': 380, 'Imago Dewasa': 450}, 'kritis': 280}
    elif 'HELOPELTIS' in opt or 'KEPIK' in opt: return {'t_base': 12.0, 'fase': {'Telur': 115, 'Nimfa (Menghisap)': 295, 'Imago Dewasa': 450}, 'kritis': 295}
    elif 'ORYCTES' in opt or 'KUMBANG NYIUR' in opt: return {'t_base': 10.5, 'fase': {'Telur': 160, 'Larva': 1800, 'Pupa': 2100, 'Imago Dewasa (Merusak)': 2900}, 'kritis': 2900}
    elif 'BRONTISPA' in opt or 'KUMBANG JANUR' in opt: return {'t_base': 11.5, 'fase': {'Telur': 70, 'Larva (Merusak)': 350, 'Pupa': 450, 'Imago Dewasa': 600}, 'kritis': 350}
    elif 'HELICOVERPA' in opt: return {'t_base': 10.5, 'fase': {'Telur': 45, 'Larva (Merusak)': 280, 'Pupa': 430, 'Imago Dewasa': 500}, 'kritis': 280}
    elif 'PLANOCOCCUS' in opt or 'KUTU PUTIH' in opt: return {'t_base': 8.5, 'fase': {'Telur': 70, 'Nimfa (Menghisap)': 320, 'Imago Dewasa': 420}, 'kritis': 320}
    elif 'CHILO' in opt or 'PENGGEREK BATANG' in opt: return {'t_base': 12.5, 'fase': {'Telur': 90, 'Larva (Merusak Batang)': 420, 'Pupa': 550, 'Imago Dewasa': 650}, 'kritis': 420}
    elif 'ARTONA' in opt or 'ULAT DAUN' in opt: return {'t_base': 10.0, 'fase': {'Telur': 110, 'Larva (Merusak)': 850, 'Pupa': 1150, 'Imago Dewasa': 1350}, 'kritis': 850}
    elif 'WERENG' in opt: return {'t_base': 10.0, 'fase': {'Telur': 120, 'Nimfa (Menghisap)': 350, 'Imago Dewasa': 450}, 'kritis': 350}
    else: return {'t_base': 10.0, 'fase': {'Telur': 100, 'Larva/Nimfa': 400, 'Pupa': 600, 'Imago': 800}, 'kritis': 400}

def tentukan_fase_sekarang(acc_gdd, fase_dict):
    fase_urut = sorted(fase_dict.items(), key=lambda x: x[1])
    total_satu_siklus = fase_urut[-1][1]
    
    if acc_gdd <= 0:
        return fase_urut[0][0], fase_urut[0][1]
        
    gdd_dalam_siklus = acc_gdd % total_satu_siklus
    nomor_generasi = int(acc_gdd // total_satu_siklus) + 1
    
    for nama_fase, ambang_batas in fase_urut:
        if gdd_dalam_siklus <= ambang_batas:
            target_gdd_global = ambang_batas + (total_satu_siklus * (nomor_generasi - 1))
            return f"{nama_fase} (Gen-{nomor_generasi})", target_gdd_global
            
    return f"Persiapan Bertelur (Gen-{nomor_generasi})", total_satu_siklus * nomor_generasi

def klasifikasi_enso(anom):
    if pd.isna(anom): return "Data Tidak Tersedia"
    if anom >= 2.0: return "El Niño Sangat Kuat"
    elif anom >= 1.5: return "El Niño Kuat"
    elif anom >= 1.0: return "El Niño Sedang"
    elif anom <= -0.5 and anom > -1.0: return "La Niña Lemah"
    elif anom <= -1.0 and anom > -1.5: return "La Niña Sedang"
    elif anom <= -1.5 and anom > -2.0: return "La Niña Kuat"
    elif anom <= -2.0: return "La Niña Sangat Kuat"
    else: return "Netral"

# ====================================================
# 3. FUNGSI LOADING DATA OTOMATIS (CACHED)
# ====================================================
@st.cache_data
def load_all_data():
    def preprocess_iklim(path):
        if not os.path.exists(path): return pd.DataFrame()
        df = xlsx_sheet_to_df(path)
        if df.empty: return pd.DataFrame()
        
        if df.shape[1] > 4:
            try:
                df = df[[0, 1, 2, 4]].copy()
                df.columns = ['tgl_raw', 'tmax', 'tmin', 'kelembapan']
                
                sample_tgl = df['tgl_raw'].dropna().iloc[0] if not df['tgl_raw'].dropna().empty else ''
                if str(sample_tgl).replace('.0', '').isdigit():
                    df['tanggal'] = excel_serial_to_datetime(df['tgl_raw'])
                else:
                    df['tanggal'] = pd.to_datetime(df['tgl_raw'], errors='coerce', format='mixed')
                
                df = df.dropna(subset=['tanggal'])
                for c in ['tmax', 'tmin', 'kelembapan']:
                    df[c] = pd.to_numeric(df[c].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce')
                df.loc[df['tmax'] >= 100, 'tmax'] = np.nan
                df.loc[df['tmin'] >= 100, 'tmin'] = np.nan
                df.loc[df['kelembapan'] > 100, 'kelembapan'] = 100.0
                df = df.dropna(subset=['tmax', 'tmin', 'kelembapan'])
                df = df[(df['tanggal'].dt.year >= 2019) & (df['tanggal'].dt.year <= 2024)].sort_values('tanggal')
                return df
            except: return pd.DataFrame()
        return pd.DataFrame()

    path_jt = "data/Jawa Tengah_ Semarang_96839.xlsx"
    if not os.path.exists(path_jt):
        path_jt = "data/Jawa Tengah_Semarang_96839.xlsx"

    suhu_banten = preprocess_iklim("data/Jakarta_Soekarno-Hatta_96749.xlsx")
    suhu_jabar = pd.concat([preprocess_iklim("data/Jawa Barat_Cilacap_96805.xlsx"), preprocess_iklim("data/Jawa Barat_Jatiwangi_96791.xlsx")])
    if not suhu_jabar.empty: suhu_jabar = suhu_jabar.groupby('tanggal', as_index=False)[['tmax', 'tmin', 'kelembapan']].mean()
    suhu_jateng = preprocess_iklim(path_jt)
    dict_suhu_provinsi = {'Banten': suhu_banten, 'Jawa Barat': suhu_jabar, 'Jawa Tengah': suhu_jateng}

    def preprocess_hujan_dasarian(path, col_hujan):
        if not os.path.exists(path): return pd.DataFrame()
        df = xlsx_sheet_to_df(path)
        if df.empty: return pd.DataFrame()
        try:
            if col_hujan >= len(df.columns): col_hujan = len(df.columns) - 1
            df = df[[0, col_hujan]].copy()
            df.columns = ['tanggal_raw', 'curah_hujan']
            
            sample_tgl = df['tanggal_raw'].dropna().iloc[0] if not df['tanggal_raw'].dropna().empty else ''
            if str(sample_tgl).replace('.0', '').isdigit():
                df['tanggal'] = excel_serial_to_datetime(df['tanggal_raw'])
            else:
                df['tanggal'] = pd.to_datetime(df['tanggal_raw'], errors='coerce', format='mixed')
                
            df = df.dropna(subset=['tanggal'])
            df['curah_hujan'] = df['curah_hujan'].astype(str).str.strip().replace({'----': np.nan, '-': np.nan, 'Tr': 0, 'tr': 0, '8888': 0, '9999': np.nan})
            df['curah_hujan'] = df['curah_hujan'].replace(r'[^\d.]', '', regex=True)
            df['curah_hujan'] = pd.to_numeric(df['curah_hujan'], errors='coerce')
            df.loc[df['curah_hujan'] >= 9000, 'curah_hujan'] = np.nan
            df = df.dropna(subset=['curah_hujan'])
            
            df['tahun'] = df['tanggal'].dt.year.astype(int)
            df['bulan'] = df['tanggal'].dt.month.astype(int)
            df['dasarian'] = df['tanggal'].dt.day.apply(get_dasarian)
            df = df[(df['tahun'] >= 2019) & (df['tahun'] <= 2025)]
            return df.groupby(['tahun', 'bulan', 'dasarian'])['curah_hujan'].sum().reset_index()
        except: return pd.DataFrame()

    hujan_banten = preprocess_hujan_dasarian("data/Jakarta_Soekarno-Hatta_96749.xlsx", 9)
    if not hujan_banten.empty: hujan_banten['provinsi'] = 'Banten'
    
    hujan_jabar1 = preprocess_hujan_dasarian("data/Jawa Barat_Cilacap_96805.xlsx", 10)
    hujan_jabar2 = preprocess_hujan_dasarian("data/Jawa Barat_Jatiwangi_96791.xlsx", 8)
    hujan_jabar = pd.concat([hujan_jabar1, hujan_jabar2])
    if not hujan_jabar.empty: 
        hujan_jabar = hujan_jabar.groupby(['tahun', 'bulan', 'dasarian'], as_index=False)['curah_hujan'].mean()
        hujan_jabar['provinsi'] = 'Jawa Barat'
        
    hujan_jateng = preprocess_hujan_dasarian(path_jt, 8)
    if not hujan_jateng.empty: hujan_jateng['provinsi'] = 'Jawa Tengah'

    hujan_all_provinsi = pd.concat([hujan_banten, hujan_jabar, hujan_jateng], ignore_index=True)
    if not hujan_all_provinsi.empty: hujan_all_provinsi.rename(columns={'curah_hujan': 'hujan_regional'}, inplace=True)

    list_files_opt = [
        "data/Data serangan OPT (Jateng 19).xlsx", "data/Data serangan OPT (Jateng 20).xlsx", "data/Data serangan OPT (Jateng 21).xlsx", "data/Data serangan OPT (Jateng 22).xlsx", "data/Data serangan OPT (Jateng 23).xlsx", "data/Data serangan OPT (Jateng 24).xlsx",
        "data/Data serangan OPT (Jabar 24).xlsx", "data/Data serangan OPT (Jabar 23).xlsx", "data/Data serangan OPT (Jabar 22).xlsx", "data/Data serangan OPT (Jabar 21).xlsx", "data/Data serangan OPT (Jabar 20).xlsx", "data/Data serangan OPT (Jabar 19).xlsx",
        "data/Data serangan OPT (Banten 19).xlsx", "data/Data serangan OPT (Banten 20).xlsx", "data/Data serangan OPT (Banten 21).xlsx", "data/Data serangan OPT (Banten 22).xlsx", "data/Data serangan OPT (Banten 23).xlsx", "data/Data serangan OPT (Banten 24).xlsx"
    ]
    all_opt = []
    for path in list_files_opt:
        if not os.path.exists(path): continue
        prov = 'Jawa Tengah' if 'Jateng' in path else 'Jawa Barat' if 'Jabar' in path else 'Banten'
        df = xlsx_sheet_to_df(path)
        if df.empty: continue
        
        baris_awal = [str(x).strip().lower() for x in df.iloc[0].values]
        has_header = any('komoditas' in c or 'opt' in c or 'kab' in c or 'tanaman' in c for c in baris_awal)
        
        if has_header:
            df.columns = baris_awal
            df = df.iloc[1:].copy().reset_index(drop=True)
            rename_mapping = {}
            for c in baris_awal:
                if 'komoditas' in c or 'tanaman' in c: rename_mapping[c] = 'komoditas'
                elif 'jenis opt' in c or 'hama' in c: rename_mapping[c] = 'jenis_opt'
                elif 'kab' in c or 'kota' in c: rename_mapping[c] = 'kabupaten'
                elif 'tahun' in c: rename_mapping[c] = 'tahun'
                elif 'luas' in c or 'total' in c or '(ha)' in c: rename_mapping[c] = 'total_serangan'
            df = df.rename(columns=rename_mapping)
        else:
            rename_mapping = {0: 'komoditas', 1: 'jenis_opt', 2: 'kabupaten', 3: 'tahun', 4: 'total_serangan'}
            df = df.rename(columns=rename_mapping)
        
        required_cols = ['komoditas', 'jenis_opt', 'kabupaten', 'tahun', 'total_serangan']
        for req in required_cols:
            if req not in df.columns: df[req] = np.nan

        df = df[required_cols].assign(provinsi=prov)
        df = df.loc[:, ~df.columns.duplicated()].copy()

        if isinstance(df['total_serangan'], pd.DataFrame): df['total_serangan'] = df['total_serangan'].iloc[:, 0]
        if isinstance(df['tahun'], pd.DataFrame): df['tahun'] = df['tahun'].iloc[:, 0]

        df['tahun'] = pd.to_numeric(df['tahun'], errors='coerce')
        df['total_serangan'] = pd.to_numeric(df['total_serangan'], errors='coerce')
        
        df = df.dropna(subset=['kabupaten', 'total_serangan'])
        df['jenis_opt'] = df['jenis_opt'].astype(str).str.strip().str.upper()
        df['komoditas'] = df['komoditas'].astype(str).str.strip().str.upper()
        
        if df['tahun'].isna().all():
            match = re.search(r'(\d{2,4})', path)
            if match:
                th = int(match.group(1))
                df['tahun'] = th + 2000 if th < 100 else th
                
        df = df[(df['tahun'] >= 2019) & (df['tahun'] <= 2024)]
        all_opt.append(df)
        
    df_opt_global = pd.concat(all_opt, ignore_index=True) if all_opt else pd.DataFrame()

    try:
        df_enso = pd.read_csv("data/nino_index.csv")
        df_enso.columns = ['YR', 'MON', 'NINO12', 'ANOM12', 'NINO3', 'ANOM3', 'NINO4', 'ANOM4', 'NINO34', 'ANOM34']
        df_enso['YR'] = pd.to_numeric(df_enso['YR'], errors='coerce')
        df_enso['MON'] = pd.to_numeric(df_enso['MON'], errors='coerce')
        df_enso['ANOM34'] = pd.to_numeric(df_enso['ANOM34'], errors='coerce')
        df_enso = df_enso[['YR', 'MON', 'ANOM34']]
        df_enso = df_enso[(df_enso['YR'] >= 2019) & (df_enso['YR'] <= 2024)]
    except:
        df_enso = pd.DataFrame(columns=['YR', 'MON', 'ANOM34'])

    return dict_suhu_provinsi, hujan_all_provinsi, df_opt_global, df_enso


# ====================================================
# 4. MULAI MEMBACA DATA 
# ====================================================
with st.spinner("⏳ Membaca Database BMKG & Kementan..."):
    dict_suhu_provinsi, hujan_all_provinsi, df_opt_global, df_enso = load_all_data()

# ====================================================
# 5. UI INPUT LOKASI & TANGGAL
# ====================================================
st.write("### 📍 Informasi Lahan & Tanam")
col_in1, col_in2, col_in3 = st.columns(3)

with col_in1: 
    provinsi = st.selectbox("Pilih Wilayah Anda", ["Banten", "Jawa Barat", "Jawa Tengah"])
with col_in2: 
    tgl_tanam = st.date_input("Mulai Tanam", datetime.date(2023, 1, 1), min_value=datetime.date(2019, 1, 1), max_value=datetime.date(2024, 12, 31))
with col_in3: 
    tgl_monitor = st.date_input("Pengecekan (Hari Ini)", datetime.date(2023, 4, 15), min_value=datetime.date(2019, 1, 1), max_value=datetime.date(2024, 12, 31))

st.markdown("---")

if tgl_tanam > tgl_monitor:
    st.error("❌ Tanggal Pengecekan harus lebih maju dari Tanggal Mulai Tanam.")
    st.stop()

# ====================================================
# 6. PEMROSESAN DATA NYATA
# ====================================================
df_suhu = dict_suhu_provinsi.get(provinsi)
if df_suhu is None or df_suhu.empty:
    st.error(f"❌ Data cuaca untuk {provinsi} tidak ditemukan atau kosong. Cek penamaan file Excel Anda.")
    st.stop()

mask = (df_suhu['tanggal'].dt.date >= tgl_tanam) & (df_suhu['tanggal'].dt.date <= tgl_monitor)
df_filter = df_suhu[mask].copy()

if df_filter.empty:
    st.warning("⚠️ Tidak ada data cuaca yang tercatat pada rentang tanggal tersebut di database.")
    st.stop()

t_avg_regional = ((df_filter['tmax'] + df_filter['tmin']) / 2.0).mean()
rh_avg_regional = df_filter['kelembapan'].mean()

tahun_evaluasi = tgl_monitor.year
bulan_evaluasi = tgl_monitor.month
dasarian_evaluasi = get_dasarian(tgl_monitor.day)

if not hujan_all_provinsi.empty:
    hujan_filter = hujan_all_provinsi[
        (hujan_all_provinsi['provinsi'] == provinsi) & 
        (hujan_all_provinsi['tahun'] == tahun_evaluasi) & 
        (hujan_all_provinsi['bulan'] == bulan_evaluasi) & 
        (hujan_all_provinsi['dasarian'] == dasarian_evaluasi)
    ]
    if not hujan_filter.empty:
        ch_val = hujan_filter['hujan_regional'].values[0]
        curah_hujan_str = f"{ch_val:.1f} mm"
    else: curah_hujan_str = "Belum Tercatat"
else: curah_hujan_str = "Data Kosong"

bulan_akhir = pd.Timestamp(year=tahun_evaluasi, month=bulan_evaluasi, day=1)
bulan_awal = bulan_akhir - pd.DateOffset(months=2)

if not df_enso.empty:
    df_enso_tmp = df_enso.copy()
    df_enso_tmp['tanggal'] = pd.to_datetime(dict(year=df_enso_tmp['YR'], month=df_enso_tmp['MON'], day=1))
    enso_window = df_enso_tmp[(df_enso_tmp['tanggal'] >= bulan_awal) & (df_enso_tmp['tanggal'] <= bulan_akhir)]
    anom34 = enso_window['ANOM34'].mean() if not enso_window.empty else np.nan
    enso_info = {"status": klasifikasi_enso(anom34), "anomali": round(anom34, 2)}
else: enso_info = {"status": "Tidak tersedia", "anomali": np.nan}

# ====================================================
# 7. TAMPILAN DASHBOARD METRIK
# ====================================================
st.write("### 🌤️ Info Cuaca & Iklim Daerah Anda")
col_c1, col_c2, col_c3 = st.columns(3)
col_c1.metric("Suhu Rata-rata", f"{t_avg_regional:.1f} °C", "Memicu pertumbuhan hama")
col_c2.metric("Kelembapan (RH)", f"{rh_avg_regional:.1f} %", "Risiko penyakit jamur" if rh_avg_regional > 85 else "Stabil")
col_c3.metric("Curah Hujan", curah_hujan_str, f"Dasarian {dasarian_evaluasi} ({tgl_monitor.strftime('%b')})")

st.info(f"**Prediksi Musim Global:** Saat ini kondisi **{enso_info['status']}** (Anomali: {enso_info['anomali']}°C).")
st.markdown("---")

# ====================================================
# 8. TAMPILAN OPT & FORECASTING JANGKA PANJANG
# ====================================================
if df_opt_global.empty:
    st.error("❌ Data Hama Kosong atau Gagal Diparsing.")
    st.stop()

df_prov = df_opt_global[df_opt_global['provinsi'] == provinsi]
if df_prov.empty:
    st.warning(f"Belum ada data serangan hama untuk wilayah {provinsi} di dalam database Kementan.")
    st.stop()

st.write(f"### 🚨 Top 5 Komoditas Paling Rawan di {provinsi}")
st.write("Klik nama komoditas di bawah ini untuk melihat fase saat ini dan peramalan masa depan:")

kom_totals = df_prov.groupby('komoditas')['total_serangan'].sum().nlargest(5)
status_kritis = False
status_jamur = True if rh_avg_regional > 85.0 else False 

for i, kom_name in enumerate(kom_totals.index, 1):
    with st.expander(f"🌾 {i}. {kom_name}"):
        opt_totals = df_prov[df_prov['komoditas'] == kom_name].groupby('jenis_opt')['total_serangan'].sum().nlargest(3)
        
        list_serangga = [h for h in opt_totals.index if not any(x in h for x in ['PHYTOPHTHORA', 'JAMUR', 'CACAR', 'BERCAK', 'LAYU', 'TIKUS'])]
        
        if not list_serangga:
            st.write("Tanaman ini umumnya didominasi serangan penyakit/jamur (bukan serangga hama).")
            continue 

        for hama_name in list_serangga:
            cfg = get_opt_config(hama_name)
            target_kritis_merusak = cfg['kritis']
            
            acc_gdd = ((df_filter['tmax'] + df_filter['tmin']) / 2.0 - cfg['t_base']).clip(lower=0).sum()
            fase_saat_ini, target_fase_selanjutnya = tentukan_fase_sekarang(acc_gdd, cfg['fase'])
            persen = min((acc_gdd / target_kritis_merusak) * 100, 100) if target_kritis_merusak > 0 else 0

            total_hari_tanam = max((tgl_monitor - tgl_tanam).days, 1)
            rata_gdd_harian = acc_gdd / total_hari_tanam
            
            if acc_gdd >= target_kritis_merusak:
                tgl_kritis_str = "⚠️ **SUDAH TERLEWATI!** Hama saat ini telah aktif merusak lahan."
            elif rata_gdd_harian <= 0:
                tgl_kritis_str = "♾️ **Tidak Terprediksi** (Suhu terlalu dingin untuk perkembangan hama ini)."
            else:
                sisa_gdd_kritis = target_kritis_merusak - acc_gdd
                hari_menuju_kritis = int(np.ceil(sisa_gdd_kritis / rata_gdd_harian))
                hari_kritis_obj = tgl_monitor + datetime.timedelta(days=hari_menuju_kritis)
                tgl_kritis_str = f"📅 **{hari_kritis_obj.strftime('%d %B %Y')}** (Sekitar **{hari_menuju_kritis} hari lagi** dari hari pengecekan)"

            if "Merusak" in fase_saat_ini or "Menghisap" in fase_saat_ini:
                status_kritis = True
                css_class = "hama-card"
                icon = "🔴"
                saran = "Segera lakukan pengendalian intensif / kuratif."
            elif persen >= 75:
                css_class = "hama-card hama-waspada"
                icon = "🟡"
                saran = "Hama mulai aktif. Pasang perangkap & monitor berkala."
            else:
                css_class = "hama-card hama-aman"
                icon = "🟢"
                saran = "Hama belum merusak. Pertahankan kondisi sanitasi lingkungan."

            st.markdown(f"""
            <div class="{css_class}">
                <h5 style="margin-top:0px; margin-bottom:5px;">{icon} {hama_name}</h5>
                <p style="margin-bottom:2px; font-size:14px;"><b>Panas Diterima:</b> {acc_gdd:.1f} GDD / Target Kritis: {target_kritis_merusak} GDD</p>
                <p style="margin-bottom:2px; font-size:14px;"><b>Fase Saat Ini:</b> 🔎 {fase_saat_ini} (Butuh {max(0, target_fase_selanjutnya - acc_gdd):.1f} GDD untuk ganti fase)</p>
                <p style="margin-bottom:5px; font-size:14px; color:#c0392b;"><b>Prediksi Tanggal Ledakan Kritis:</b> {tgl_kritis_str}</p>
                <p style="margin-bottom:0px; font-size:14px; color:#555;"><i>Rekomendasi: {saran}</i></p>
            </div>
            """, unsafe_allow_html=True)

            st.write("🔮 **Tabel Peramalan Cuaca & Fase Hidup Hama ke Depan:**")
            
            list_horizon_hari = [7, 15, 30]
            forecast_list = []
            
            for jml_hari in list_horizon_hari:
                tgl_target_forecast = tgl_monitor + datetime.timedelta(days=jml_hari)
                
                f_thn = tgl_target_forecast.year
                f_bln = tgl_target_forecast.month
                f_das = get_dasarian(tgl_target_forecast.day)
                
                if not hujan_all_provinsi.empty:
                    f_hujan_df = hujan_all_provinsi[
                        (hujan_all_provinsi['provinsi'] == provinsi) & 
                        (hujan_all_provinsi['tahun'] == f_thn) & 
                        (hujan_all_provinsi['bulan'] == f_bln) & 
                        (hujan_all_provinsi['dasarian'] == f_das)
                    ]
                    pred_hujan_str = f"{f_hujan_df['hujan_regional'].values[0]:.1f} mm" if not f_hujan_df.empty else "0.0 mm"
                else:
                    pred_hujan_str = "Data Kosong"
                
                mask_fut = (df_suhu['tanggal'].dt.date > tgl_monitor) & (df_suhu['tanggal'].dt.date <= tgl_target_forecast)
                df_fut = df_suhu[mask_fut]
                
                if not df_fut.empty:
                    f_tmax = df_fut['tmax'].mean()
                    f_tmin = df_fut['tmin'].mean()
                    f_tavg = (f_tmax + f_tmin) / 2.0
                    f_rh = df_fut['kelembapan'].mean()
                    gdd_tambahan = ( ((df_fut['tmax'] + df_fut['tmin']) / 2.0) - cfg['t_base'] ).clip(lower=0).sum()
                    sumber_data = "Riil BMKG"
                else:
                    f_tavg = t_avg_regional
                    f_rh = rh_avg_regional
                    gdd_tambahan = rata_gdd_harian * jml_hari
                    sumber_data = "Estimasi Model"
                
                gdd_proyeksi_total = acc_gdd + gdd_tambahan
                fase_proyeksi, _ = tentukan_fase_sekarang(gdd_proyeksi_total, cfg['fase'])
                
                forecast_list.append({
                    "Horizon Waktu": f"{jml_hari} Hari ({tgl_target_forecast.strftime('%d %b')})",
                    "Prediksi Suhu": f"{f_tavg:.1f} °C",
                    "Prediksi RH": f"{f_rh:.0f} %",
                    "Prediksi Curah Hujan (Per Dasarian)": pred_hujan_str,
                    "Prediksi Fase Hidup OPT": fase_proyeksi,
                    "Metode": sumber_data
                })
            
            df_forecast_tabel = pd.DataFrame(forecast_list)
            st.table(df_forecast_tabel)

st.markdown("---")

# ====================================================
# 9. KESIMPULAN REKOMENDASI (CSA)
# ====================================================
st.write("### ✅ Kesimpulan & Rekomendasi Akhir")

if status_kritis:
    st.error("**🚨 PERINGATAN HAMA:** Beberapa hama utama diproyeksikan telah/akan masuk fase **DESTRUKTIF (Merusak)** dalam waktu dekat. Segera siapkan strategi PHT (Pengendalian Hama Terpadu).")
else:
    st.success("**✅ STATUS HAMA STABIL:** Belum ada indikasi lonjakan populasi hama ke fase kritis dalam jangka pendek.")
    
if status_jamur:
    st.error(f"**🍄 PERINGATAN PENYAKIT:** Tingkat kelembapan udara sangat tinggi ({rh_avg_regional:.1f}%). Sangat rawan serangan infeksi jamur/bakteri patogen.")
else:
    st.success(f"**☀️ STATUS PENYAKIT AMAN:** Kondisi kelembapan mikro ({rh_avg_regional:.1f}%) relatif aman dari risiko ledakan penyakit jamur.")

if "El Niño" in enso_info['status']:
    st.warning(f"**🏜️ ADAPTASI IKLIM MAKRO:** Berada dalam pola **{enso_info['status']}**. Optimalkan manajemen saluran irigasi tertutup untuk menghemat air.")
elif "La Niña" in enso_info['status']:
    st.info(f"**🌧️ ADAPTASI IKLIM MAKRO:** Berada dalam pola **{enso_info['status']}**. Tinggikan bedengan dan pastikan sistem drainase lahan lancar agar akar tidak busuk.")