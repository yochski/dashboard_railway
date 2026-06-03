"""
Ultimate Dashboard Ekspor-Impor Indonesia + Neraca Pembayaran SEKI BI
Versi: Streamlit (Local DB + Google Drive Downloader)
"""

import os, re, sqlite3, io
import gdown
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import traceback

# ─────────────────────────────────────────────────────────────────
#  KONFIGURASI UTAMA & GDRIVE
# ─────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Trade Intelligence | NEXOS", layout="wide", page_icon="📊")

# ID File Google Drive yang Anda berikan
GDRIVE_FILE_ID = "1IhKP7Jw7xhRPDzvw4CY7FVvK0lO_biy_"
GDRIVE_URL = f"https://drive.google.com/uc?id={GDRIVE_FILE_ID}"

DATA_DIR     = os.environ.get("DATA_DIR", os.path.dirname(__file__))
BPS_DB_PATH  = os.path.join(DATA_DIR, os.environ.get("BPS_DB_FILE", "ekspor_impor_bps.db"))
BOP_DB_PATH  = os.path.join(DATA_DIR, os.environ.get("BOP_DB_FILE", "bop_indonesia.db"))
TM_XLSX      = os.path.join(DATA_DIR, os.environ.get("TM_XLSX_FILE", "data_trademap.xlsx"))

# Nama tabel sesuai dengan database Anda
BPS_TABLE    = "exim_data" 

HS_ALL       = [str(i).zfill(2) for i in range(1, 100)]
TAHUN_SAAT_INI = datetime.now().year
TAHUN_TERSEDIA = list(range(2015, TAHUN_SAAT_INI + 1))

PERIODE_OPSI = {
    "Tahunan": "tahunan", "Januari": "1", "Februari": "2", "Maret": "3", 
    "April": "4", "Mei": "5", "Juni": "6", "Juli": "7", "Agustus": "8", 
    "September": "9", "Oktober": "10", "November": "11", "Desember": "12"
}

PARTNER_LIST = [
    "China", "Amerika Serikat", "Jepang", "Singapura",
    "India", "Malaysia", "Korea Selatan", "Australia",
    "Jerman", "Belanda", "Thailand", "Vietnam",
]

HS_DESC = {
    "01":"Binatang Hidup","02":"Daging & Produk Daging","03":"Ikan & Produk Ikan",
    "04":"Produk Susu & Telur","05":"Produk Hewani Lainnya","06":"Tanaman Hidup & Bunga",
    "07":"Sayuran","08":"Buah-buahan","09":"Kopi, Teh & Rempah","10":"Serealia",
    "11":"Produk Penggilingan","12":"Biji Minyak","13":"Getah & Resin",
    "14":"Bahan Nabati Lainnya","15":"Lemak & Minyak Nabati/Hewani","16":"Olahan Daging & Ikan",
    "17":"Gula & Kembang Gula","18":"Kakao & Olahannya","19":"Olahan Sereal & Tepung",
    "20":"Olahan Sayur & Buah","21":"Aneka Olahan Pangan","22":"Minuman & Cuka",
    "23":"Ampas Industri Pangan","24":"Tembakau","25":"Garam, Belerang & Batu",
    "26":"Bijih, Terak & Abu","27":"Bahan Bakar Mineral & Minyak Bumi",
    "28":"Kimia Anorganik","29":"Kimia Organik","30":"Produk Farmasi",
    "31":"Pupuk","32":"Cat, Tinta & Vernis","33":"Minyak Atsiri & Kosmetik",
    "34":"Sabun & Deterjen","35":"Albuminoid & Pati","36":"Bahan Peledak",
    "37":"Produk Foto","38":"Kimia Lainnya","39":"Plastik & Barang Plastik",
    "40":"Karet & Barang Karet","41":"Kulit Mentah & Samak","42":"Barang Kulit & Tas",
    "43":"Bulu Binatang","44":"Kayu & Produk Kayu","45":"Gabus",
    "46":"Produk Anyaman","47":"Bubur Kayu (Pulp)","48":"Kertas & Karton",
    "49":"Buku & Produk Cetak","50":"Sutra","51":"Wol & Bulu Halus",
    "52":"Kapas","53":"Serat Tekstil Lainnya","54":"Filamen Buatan",
    "55":"Serat Staple Buatan","56":"Benang & Kain Khusus","57":"Karpet & Alas Lantai",
    "58":"Kain Tenunan Khusus","59":"Kain Teknik","60":"Kain Rajut",
    "61":"Pakaian Rajut","62":"Pakaian Tenun","63":"Tekstil Rumah Tangga",
    "64":"Alas Kaki","65":"Topi & Aksesori Kepala","66":"Payung & Tongkat",
    "67":"Bulu Olahan & Bunga Buatan","68":"Barang dari Batu & Semen",
    "69":"Produk Keramik","70":"Kaca & Produk Kaca",
    "71":"Batu Permata & Logam Mulia","72":"Besi & Baja",
    "73":"Barang dari Besi/Baja","74":"Tembaga & Produknya",
    "75":"Nikel & Produknya","76":"Aluminium & Produknya",
    "78":"Timbal & Produknya","79":"Seng & Produknya","80":"Timah & Produknya",
    "81":"Logam Dasar Lainnya","82":"Perkakas & Peralatan Logam",
    "83":"Barang Logam Lainnya","84":"Mesin & Perlengkapan Mekanik",
    "85":"Mesin & Perlengkapan Listrik","86":"Kereta Api & Komponen",
    "87":"Kendaraan Bermotor","88":"Pesawat Udara & Komponen",
    "89":"Kapal & Perahu","90":"Instrumen Optik & Medis",
    "91":"Jam & Arloji","92":"Instrumen Musik","93":"Senjata & Amunisi",
    "94":"Furnitur & Perlengkapan","95":"Mainan & Perlengkapan Olahraga",
    "96":"Produk Manufaktur Lainnya","97":"Karya Seni & Antik","99":"Barang Lainnya",
}

BOP_MAIN_ITEMS = {
    1:"Transaksi Berjalan", 2:"Barang", 17:"Jasa",
    20:"Pendapatan Primer", 23:"Pendapatan Sekunder",
    26:"Transaksi Modal", 29:"Transaksi Finansial",
    32:"Investasi Langsung", 35:"Investasi Portofolio",
    40:"Derivatif Finansial", 41:"Investasi Lainnya",
    46:"Total (I+II+III)", 47:"Selisih Perhitungan",
    48:"Neraca Keseluruhan", 54:"Cadangan Devisa",
    56:"CA % PDB",
}

_NEGARA_KW = {
    "China":           ["china", "cina", "tiongkok", "zhongguo"],
    "Amerika Serikat": ["united states", "america", "u.s.a", "u.s.", "serikat", "usa", "u.s.a."],
    "Jepang":          ["japan", "jepang", "nippon"],
    "Singapura":       ["singapore", "singapura"],
    "India":           ["india"],
    "Malaysia":        ["malaysia"],
    "Korea Selatan":   ["korea"],
    "Australia":       ["australia"],
    "Jerman":          ["germany", "jerman", "deutschland"],
    "Belanda":         ["netherlands", "belanda", "holland"],
    "Thailand":        ["thailand"],
    "Vietnam":         ["vietnam", "viet nam", "viet-nam"],
}

# ─────────────────────────────────────────────────────────────────
#  FUNGSI DOWNLOAD DATABASE VIA GDOWN
# ─────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def download_bps_database():
    if not os.path.exists(BPS_DB_PATH):
        with st.spinner("Sedang mengunduh database awal (~600MB) dari Google Drive. Mohon tunggu..."):
            try:
                gdown.download(GDRIVE_URL, BPS_DB_PATH, quiet=False)
            except Exception as e:
                st.error(f"Gagal mengunduh database: {e}")

# Trigger download
download_bps_database()

# ─────────────────────────────────────────────────────────────────
#  FUNGSI BANTUAN & DATA FETCHER (Di-cache)
# ─────────────────────────────────────────────────────────────────
def normalize_negara(nama: str) -> str:
    if not nama: return nama
    low = str(nama).strip().lower()
    for display, kws in _NEGARA_KW.items():
        if low in kws or any(kw in low for kw in kws):
            return display
    return str(nama).strip()

def clean_hs(raw) -> str:
    if pd.isna(raw) or str(raw).strip() == "": return ""
    s = str(raw).strip()
    m = re.search(r'\d+', s)
    if m: return m.group(0)[:2].zfill(2)
    return s[:2].zfill(2)

def get_periode_params(pilihan):
    return ("2", "") if pilihan == "tahunan" else ("1", pilihan)

def check_bps_db():
    return os.path.exists(BPS_DB_PATH)

@st.cache_data(ttl=600)
def fetch_bps_db(sumber, tahun, tipe, bulan=""):
    """Tarik data dari SQLite dengan struktur tabel 'exim_data' Anda."""
    if not check_bps_db(): return pd.DataFrame()
    try:
        conn = sqlite3.connect(BPS_DB_PATH)
        jenis_transaksi = "Ekspor" if str(sumber) == "1" else "Impor"
        
        query = f"SELECT kode_hs as kodehs, ctr as negara, value, netweight as berat FROM {BPS_TABLE} WHERE jenis_transaksi = ? AND tahun = ?"
        params = [jenis_transaksi, str(tahun)]
        
        if tipe == "1" and bulan:
            query += " AND bulan_kode = ?"
            params.append(str(bulan).zfill(2)) 
            
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()

        if not df.empty:
            df["kodehs"] = df["kodehs"].apply(clean_hs)
            df["negara"] = df["negara"].apply(normalize_negara)
            df["value"] = pd.to_numeric(df["value"], errors="coerce").fillna(0)
            df["berat"] = pd.to_numeric(df["berat"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"Error Database BPS: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def fetch_hist_bps_db(sumber, hs, tipe, bulan=""):
    if not check_bps_db(): return pd.DataFrame()
    try:
        conn = sqlite3.connect(BPS_DB_PATH)
        jenis_transaksi = "Ekspor" if str(sumber) == "1" else "Impor"
        
        query = f"SELECT tahun, ctr as negara, value FROM {BPS_TABLE} WHERE jenis_transaksi = ? AND kode_hs = ?"
        params = [jenis_transaksi, str(hs).zfill(2)]
        
        if tipe == "1" and bulan:
            query += " AND bulan_kode = ?"
            params.append(str(bulan).zfill(2))
            
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()

        if not df.empty:
            df["negara"] = df["negara"].apply(normalize_negara)
            df["value"] = pd.to_numeric(df["value"], errors="coerce").fillna(0)
        return df
    except Exception as e:
        st.error(f"Error Database Historis BPS: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def load_trademap(mitra, tahun, sumber):
    """Pembaca Excel khusus Mirroring."""
    if not os.path.exists(TM_XLSX): return pd.DataFrame(), "FILE_NOT_FOUND"
    try:
        df = pd.read_excel(TM_XLSX)
        need = ["Tahun", "Mitra", "HS", "Impor_Mitra", "Ekspor_Mitra"]
        if not all(c in df.columns for c in need): return pd.DataFrame(), "INVALID_COLUMNS"
        
        df["HS"] = df["HS"].apply(clean_hs)
        df["Tahun"] = df["Tahun"].astype(str).str.strip()
        df["_mitra_n"] = df["Mitra"].apply(normalize_negara)
        mitra_n = normalize_negara(mitra)
        
        df_f = df[(df["_mitra_n"] == mitra_n) & (df["Tahun"] == str(tahun))].copy()
        if df_f.empty:
            tahun_ada = df[df["_mitra_n"] == mitra_n]["Tahun"].unique().tolist()
            if tahun_ada: return pd.DataFrame(), f"DATA_EMPTY_TAHUN|{','.join(sorted(tahun_ada))}"
            return pd.DataFrame(), "DATA_EMPTY"
            
        col = "Impor_Mitra" if str(sumber) == "1" else "Ekspor_Mitra"
        df_f[col] = pd.to_numeric(df_f[col], errors="coerce").fillna(0)
        df_out = df_f.groupby("HS", as_index=False)[col].sum().rename(columns={col: "Trademap_Value"})
        return df_out, "SUCCESS"
    except Exception as e:
        return pd.DataFrame(), str(e)

@st.cache_data(ttl=3600)
def bop_query(sql, params=()):
    if not os.path.exists(BOP_DB_PATH): return pd.DataFrame()
    try:
        with sqlite3.connect(BOP_DB_PATH) as conn:
            return pd.read_sql_query(sql, conn, params=params)
    except Exception:
        return pd.DataFrame()

def bop_series(item_ids, y1, y2, freq):
    ph = ",".join("?" * len(item_ids))
    sql = f"""SELECT item_id, keterangan, items_en, year, quarter, period, value_mn_usd
              FROM bop_quarterly WHERE item_id IN ({ph}) AND year >= ? AND year <= ? ORDER BY item_id, year, quarter"""
    df = bop_query(sql, tuple(item_ids) + (y1, y2))
    if df.empty or freq == "quarterly": return df
    
    parts = []
    ratio = {54, 55, 56, 57, 58}
    for iid, grp in df.groupby("item_id"):
        if iid in ratio: r = grp[grp["quarter"] == "Q4"].copy()
        else: r = grp.groupby(["item_id","keterangan","items_en","year"], as_index=False)["value_mn_usd"].sum()
        parts.append(r)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

def calculate_ews(df):
    if df.empty: return df
    df["harga"] = df.apply(lambda row: row["value"] / row["berat"] if row.get("berat", 0) > 0 else 0, axis=1)
    m_val, s_val = df["value"].mean(), df["value"].std()
    df["z_score"] = 0 if pd.isna(s_val) or s_val == 0 else (df["value"] - m_val) / s_val
    df_vh = df[df["harga"] > 0]
    m_h, s_h = (df_vh["harga"].mean(), df_vh["harga"].std()) if not df_vh.empty else (0, 0)
    df["z_score_harga"] = df.apply(lambda row: 0 if pd.isna(s_h) or s_h == 0 or row["harga"] == 0 else (row["harga"] - m_h) / s_h, axis=1)
    
    df["status_ews"] = "Normal"
    df.loc[df["z_score"] >  1.5, "status_ews"] = "🔴 Batas Atas Nilai"
    df.loc[df["z_score"] < -0.5, "status_ews"] = "🟡 Batas Bawah Nilai"
    df.loc[df["z_score_harga"] > 2.0, "status_ews"] = "🟣 Anomali Harga"
    mask = (df["z_score"] > 1.5) & (df["z_score_harga"] > 2.0)
    df.loc[mask, "status_ews"] = "🚨 KRITIS: Spike"
    return df

# ─────────────────────────────────────────────────────────────────
#  UI STREAMLIT
# ─────────────────────────────────────────────────────────────────
st.title("NEXOS | BPS · ITC Trade Map · SEKI Bank Indonesia")
st.caption(f"Waktu Saat Ini: {datetime.now().strftime('%d %b %Y %H:%M')}")

# ── Sidebar Filter BPS ──
with st.sidebar:
    st.header("1. KONTROL BPS DB LOKAL")
    
    if check_bps_db():
        st.success("✅ DB BPS Tersambung")
    else:
        st.error("❌ DB BPS Tidak Ditemukan")

    with st.form("bps_form"):
        f_tahun = st.selectbox("Tahun", reversed(TAHUN_TERSEDIA), index=1)
        f_periode_label = st.selectbox("Periode", list(PERIODE_OPSI.keys()))
        f_periode = PERIODE_OPSI[f_periode_label]
        f_sumber_label = st.radio("Jenis Perdagangan", ["Ekspor", "Impor"])
        f_sumber = "1" if f_sumber_label == "Ekspor" else "2"
        f_unit = st.radio("Satuan", ["USD", "Miliar USD"])
        
        submitted = st.form_submit_button("▶ MUAT BPS", use_container_width=True)
        if submitted:
            tipe, bulan = get_periode_params(f_periode)
            with st.spinner('Menarik data dari Database BPS...'):
                df_raw = fetch_bps_db(f_sumber, str(f_tahun), tipe, bulan)
                if not df_raw.empty:
                    st.session_state['bps_data'] = df_raw
                    st.session_state['bps_meta'] = {"tahun": f_tahun, "sumber": f_sumber_label, "sumber_kode": f_sumber, "unit": f_unit, "tipe": tipe, "bulan": bulan}
                    st.success(f"Berhasil memuat {len(df_raw):,} baris data.")
                else:
                    st.error(f"Tidak ada data BPS {f_sumber_label} {f_tahun} untuk filter tersebut.")

# ── Tabs Setup ──
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Ringkasan BPS", "🗄️ Data Lengkap", "⚠️ Early Warning System", 
    "🪞 Mirroring", "🏦 Neraca Pembayaran (SEKI BI)"
])

# Ambil data dari session_state
df_bps = st.session_state.get('bps_data', pd.DataFrame())
meta = st.session_state.get('bps_meta', {})

if not df_bps.empty:
    div = 1e9 if meta['unit'] == "Miliar USD" else 1
    df_bps_clean = df_bps.copy()
    df_bps_clean["value"] = df_bps_clean["value"] / div
    
    kmd = df_bps_clean.groupby("kodehs", as_index=False)[["value","berat"]].sum().sort_values("value", ascending=False)
    neg = df_bps_clean.groupby("negara", as_index=False)["value"].sum().sort_values("value", ascending=False)
    kmd["deskripsi"] = kmd["kodehs"].map(HS_DESC).fillna("Lainnya")

# ── TAB 1: Ringkasan ──
with tab1:
    if df_bps.empty:
        st.info("👈 Silakan tekan tombol 'MUAT BPS' pada sidebar untuk memulai.")
    else:
        col1, col2, col3 = st.columns(3)
        col1.metric(f"TOTAL {meta['sumber'].upper()}", f"{df_bps_clean['value'].sum():,.2f} {meta['unit']}")
        col2.metric("KOMODITAS TERBESAR (HS)", kmd.iloc[0]["kodehs"] if not kmd.empty else "-")
        col3.metric("NEGARA TUJUAN/ASAL UTAMA", neg.iloc[0]["negara"] if not neg.empty else "-")
        
        c_left, c_right = st.columns(2)
        with c_left:
            st.subheader("TOP 15 KOMODITAS (HS)")
            fig_kmd = px.bar(kmd.head(15), y="kodehs", x="value", orientation='h', 
                             title="Komoditas Terbesar").update_yaxes(categoryorder='total ascending')
            st.plotly_chart(fig_kmd, use_container_width=True, theme="streamlit")
            
        with c_right:
            st.subheader("TOP NEGARA MITRA")
            fig_neg = px.bar(neg.head(15), y="negara", x="value", orientation='h',
                             title="Mitra Dagang Utama").update_yaxes(categoryorder='total ascending')
            st.plotly_chart(fig_neg, use_container_width=True, theme="streamlit")
            
        st.divider()
        st.subheader(f"HISTORIS TREN ({TAHUN_TERSEDIA[0]}–{TAHUN_SAAT_INI})")
        
        hist_col1, hist_col2, hist_col3 = st.columns(3)
        hs_hist = hist_col1.selectbox("Pilih HS untuk Histori", options=HS_ALL, index=26)
        negara_hist = hist_col2.selectbox("Filter Negara (Opsional)", options=["Semua Negara"] + PARTNER_LIST)
        metric_hist = hist_col3.radio("Metrik Histori", ["Nilai", "YoY %"], horizontal=True)
        
        if st.button("Tampilkan Histori", type="primary"):
            with st.spinner("Menarik data historis..."):
                df_hist_raw = fetch_hist_bps_db(meta['sumber_kode'], hs_hist, meta['tipe'], meta['bulan'])
                if df_hist_raw.empty:
                    st.warning("Tidak ada data historis.")
                else:
                    if negara_hist != "Semua Negara":
                        df_hist_raw = df_hist_raw[df_hist_raw["negara"] == normalize_negara(negara_hist)]
                    
                    df_h = df_hist_raw.groupby("tahun", as_index=False)["value"].sum().sort_values("tahun")
                    df_h["Tahun"] = df_h["tahun"].astype(str)
                    df_h["Value"] = df_h["value"] / div
                    
                    if metric_hist == "YoY %":
                        df_h["Value"] = df_h["Value"].pct_change() * 100
                        fig_hist = px.line(df_h, x="Tahun", y="Value", markers=True, title=f"YoY (%) – HS {hs_hist}")
                    else:
                        fig_hist = px.line(df_h, x="Tahun", y="Value", markers=True, title=f"Tren Nilai ({meta['unit']}) – HS {hs_hist}")
                    
                    st.plotly_chart(fig_hist, use_container_width=True, theme="streamlit")

# ── TAB 2: Data Lengkap ──
with tab2:
    if not df_bps.empty:
        st.subheader("Tabel Data Ekspor/Impor")
        full_df = df_bps_clean.groupby(["negara","kodehs"], as_index=False)[["value","berat"]].sum()
        full_df["deskripsi"] = full_df["kodehs"].map(HS_DESC).fillna("Lainnya")
        
        csv = full_df.to_csv(index=False).encode('utf-8')
        st.download_button(label="⬇ Download CSV", data=csv, file_name=f"BPS_{meta['sumber']}_{meta['tahun']}.csv", mime='text/csv')
        
        st.dataframe(full_df, use_container_width=True)
    else:
        st.info("Data belum dimuat.")

# ── TAB 3: EWS ──
with tab3:
    if not df_bps.empty:
        st.markdown("### ⚠️ Deteksi Anomali (Early Warning System)")
        st.caption("Batas Atas = Konsentrasi berlebih. Batas Bawah = Underperforming. Anomali Harga = Indikasi kesalahan input/lonjakan harga ekstrem.")
        
        ews_df = calculate_ews(kmd.copy())
        
        def highlight_ews(val):
            color = ''
            if 'Atas' in str(val): color = '#f78166'
            elif 'Bawah' in str(val): color = '#e3b341'
            elif 'Harga' in str(val): color = '#bc8cff'
            elif 'KRITIS' in str(val): color = 'red'
            return f'background-color: {color}'
        
        st.dataframe(ews_df.style.map(highlight_ews, subset=['status_ews']), use_container_width=True)
    else:
        st.info("Data belum dimuat.")

# ── TAB 4: Mirroring ──
with tab4:
    st.header("🪞 Analisis Asimetri Pencatatan (BPS vs ITC Trade Map)")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        mitra_mirror = st.selectbox("Pilih Negara Mitra", PARTNER_LIST)
    with col_m2:
        unit_mirror = st.radio("Satuan Mirroring", ["USD", "Juta USD"])
        
    if st.button("▶ JALANKAN MIRRORING", type="primary"):
        if not df_bps.empty:
            div_m = 1e6 if unit_mirror == "Juta USD" else 1
            
            with st.spinner("Mencocokkan data..."):
                try:
                    df_tm, status = load_trademap(mitra_mirror, meta['tahun'], meta['sumber_kode'])
                    
                    if status == "SUCCESS":
                        df_bps_m = df_bps[df_bps["negara"] == normalize_negara(mitra_mirror)].copy()
                        df_bps_m = df_bps_m.groupby("kodehs", as_index=False)["value"].sum()
                        df_bps_m.rename(columns={"kodehs":"HS","value":"BPS_Value"}, inplace=True)
                        
                        df_merge = pd.merge(df_bps_m, df_tm, on="HS", how="outer").fillna(0)
                        df_merge[["BPS_Value", "Trademap_Value"]] /= div_m
                        df_merge["Selisih"] = df_merge["Trademap_Value"] - df_merge["BPS_Value"]
                        df_merge["Deskripsi"] = df_merge["HS"].map(HS_DESC).fillna("Lainnya")
                        
                        st.success("Mirroring berhasil dijalankan.")
                        
                        cm1, cm2, cm3 = st.columns(3)
                        cm1.metric("Total BPS", f"{df_merge['BPS_Value'].sum():,.1f}")
                        cm2.metric("Total Trade Map", f"{df_merge['Trademap_Value'].sum():,.1f}")
                        cm3.metric("Selisih Asimetri", f"{df_merge['Selisih'].sum():,.1f}")
                        
                        # 5 HS Asimetri Terbesar
                        df_merge["Abs_Diff"] = df_merge["Selisih"].abs()
                        top5 = df_merge.nlargest(5, "Abs_Diff")
                        fig_diff = px.bar(top5, x="HS", y=["BPS_Value", "Trademap_Value"], barmode="group",
                                          title="5 HS dengan Selisih (Asimetri) Terbesar")
                        st.plotly_chart(fig_diff, use_container_width=True, theme="streamlit")
                        
                        st.dataframe(df_merge[["HS", "Deskripsi", "BPS_Value", "Trademap_Value", "Selisih"]], use_container_width=True)
                    elif status == "FILE_NOT_FOUND":
                        st.error("❌ File 'data_trademap.xlsx' tidak ditemukan.")
                    elif status.startswith("DATA_EMPTY_TAHUN"):
                        tahun_ada = status.split("|")[1] if "|" in status else "-"
                        st.warning(f"⚠️ Data Trade Map untuk '{mitra_mirror}' tahun {meta['tahun']} belum ada. Tahun di Excel: {tahun_ada}")
                    else:
                        st.error(f"Gagal memuat Trade Map: {status}")
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat mirroring: {str(e)}")
                    st.code(traceback.format_exc())
        else:
            st.warning("Muat data BPS di Sidebar terlebih dahulu sebelum menjalankan Mirroring!")

# ── TAB 5: SEKI BI ──
with tab5:
    db_ok = os.path.exists(BOP_DB_PATH)
    if not db_ok:
        st.error("❌ Database bop_indonesia.db tidak ditemukan.")
    else:
        st.success("✅ Database SEKI tersambung.")
        
        with st.form("seki_form"):
            cs1, cs2, cs3, cs4 = st.columns(4)
            y1 = cs1.number_input("Tahun Awal", min_value=2004, max_value=2025, value=2015)
            y2 = cs2.number_input("Tahun Akhir", min_value=2004, max_value=2025, value=2024)
            freq = cs3.selectbox("Frekuensi", ["Kuartalan", "Tahunan"])
            f_val = "quarterly" if freq == "Kuartalan" else "annual"
            unit_s = cs4.selectbox("Satuan", ["Juta USD", "Miliar USD"])
            div_s = 1000 if unit_s == "Miliar USD" else 1
            
            sub_seki = st.form_submit_button("▶ TAMPILKAN NERACA")
            
        if sub_seki:
            needed_ids = [1,2,17,20,23,26,29,32,35,40,41,46,47,48,54,55,56]
            df_seki = bop_series(needed_ids, y1, y2, f_val)
            
            if not df_seki.empty:
                df_seki["nilai"] = df_seki["value_mn_usd"] / div_s
                
                st.subheader("Tren Transaksi Berjalan")
                ca_df = df_seki[df_seki['item_id'] == 1]
                xcol = "period" if f_val == "quarterly" else "year"
                
                fig_ca = px.line(ca_df, x=xcol, y="nilai", markers=True, title="Current Account (CA)")
                st.plotly_chart(fig_ca, use_container_width=True, theme="streamlit")
                
                st.dataframe(df_seki[["year", "period", "keterangan", "nilai"]], use_container_width=True)
            else:
                st.warning("Tidak ada data SEKI untuk rentang waktu tersebut.")
