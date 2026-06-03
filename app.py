"""
Ultimate Dashboard Ekspor-Impor Indonesia + Neraca Pembayaran SEKI BI
Versi: Streamlit (Auto-Load BPS, Full Features)
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

# ID File Google Drive
GDRIVE_FILE_ID = "1IhKP7Jw7xhRPDzvw4CY7FVvK0lO_biy_"
GDRIVE_URL = f"https://drive.google.com/uc?id={GDRIVE_FILE_ID}"

DATA_DIR     = os.environ.get("DATA_DIR", os.path.dirname(__file__))
BPS_DB_PATH  = os.path.join(DATA_DIR, os.environ.get("BPS_DB_FILE", "ekspor_impor_bps.db"))
BOP_DB_PATH  = os.path.join(DATA_DIR, os.environ.get("BOP_DB_FILE", "bop_indonesia.db"))
TM_XLSX      = os.path.join(DATA_DIR, os.environ.get("TM_XLSX_FILE", "data_trademap.xlsx"))

# Nama tabel sesuai dengan database BPS Anda
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
    "China": ["china", "cina", "tiongkok", "zhongguo"],
    "Amerika Serikat": ["united states", "america", "u.s.a", "u.s.", "serikat", "usa"],
    "Jepang": ["japan", "jepang", "nippon"],
    "Singapura": ["singapore", "singapura"],
    "India": ["india"], "Malaysia": ["malaysia"], "Korea Selatan": ["korea"],
    "Australia": ["australia"], "Jerman": ["germany", "jerman"],
    "Belanda": ["netherlands", "belanda"], "Thailand": ["thailand"], "Vietnam": ["vietnam"],
}

# ─────────────────────────────────────────────────────────────────
#  FUNGSI UI KHUSUS (REPLIKASI GAYA DASH)
# ─────────────────────────────────────────────────────────────────
def kpi_card(title, value, color, sub_text=""):
    st.markdown(f"""
        <div style="border: 1px solid rgba(128,128,128,0.2); border-top: 3px solid {color}; 
                    padding: 15px; border-radius: 6px; margin-bottom: 15px; background-color: transparent;">
            <div style="font-size: 11px; color: gray; letter-spacing: 1px;">{title}</div>
            <div style="font-size: 22px; font-weight: bold; color: {color}; margin-top: 5px;">{value}</div>
            <div style="font-size: 11px; color: gray; margin-top: 5px;">{sub_text}</div>
        </div>
    """, unsafe_allow_html=True)

def section_title(title):
    st.markdown(f"<p style='font-size:11px; font-weight:bold; letter-spacing:1px; margin-bottom:0px; text-transform:uppercase;'>{title}</p>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
#  FUNGSI DOWNLOAD & DATA BPS
# ─────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def download_bps_database():
    if not os.path.exists(BPS_DB_PATH):
        with st.spinner("Sedang mengunduh database awal (~600MB) dari Google Drive. Mohon tunggu..."):
            try:
                gdown.download(GDRIVE_URL, BPS_DB_PATH, quiet=False)
            except Exception as e:
                st.error(f"Gagal mengunduh database: {e}")

download_bps_database()

def normalize_negara(nama: str) -> str:
    if not nama: return nama
    low = str(nama).strip().lower()
    for display, kws in _NEGARA_KW.items():
        if low in kws or any(kw in low for kw in kws): return display
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
    if not check_bps_db(): return pd.DataFrame()
    try:
        conn = sqlite3.connect(BPS_DB_PATH)
        jenis = "Ekspor" if str(sumber) == "1" else "Impor"
        query = f"SELECT kode_hs as kodehs, ctr as negara, value, netweight as berat FROM {BPS_TABLE} WHERE jenis_transaksi = ? AND tahun = ?"
        params = [jenis, str(tahun)]
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
        st.error(f"Error DB BPS: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def fetch_hist_bps_db(sumber, hs, tipe, bulan=""):
    if not check_bps_db(): return pd.DataFrame()
    try:
        conn = sqlite3.connect(BPS_DB_PATH)
        jenis = "Ekspor" if str(sumber) == "1" else "Impor"
        query = f"SELECT tahun, ctr as negara, value FROM {BPS_TABLE} WHERE jenis_transaksi = ? AND kode_hs = ?"
        params = [jenis, str(hs).zfill(2)]
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
        return pd.DataFrame()

# ─────────────────────────────────────────────────────────────────
#  FUNGSI MIRRORING & SEKI BI
# ─────────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def load_trademap(mitra, tahun, sumber):
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

def bop_latest():
    df = bop_query("""SELECT period FROM bop_quarterly
                      WHERE value_mn_usd IS NOT NULL
                      ORDER BY year DESC, quarter DESC LIMIT 1""")
    return df["period"].iloc[0] if not df.empty else "-"

def bop_latest_val(item_id):
    df = bop_query("""SELECT value_mn_usd FROM bop_quarterly
                      WHERE item_id=? AND value_mn_usd IS NOT NULL
                      ORDER BY year DESC, quarter DESC LIMIT 1""", (item_id,))
    return float(df["value_mn_usd"].iloc[0]) if not df.empty else None

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
st.caption(f"Waktu Sistem: {datetime.now().strftime('%d %b %Y %H:%M')}")

# ── Sidebar Filter BPS ──
with st.sidebar:
    st.markdown("### 1. KONTROL BPS DB LOKAL")
    
    if check_bps_db():
        st.success("✅ DB BPS Tersambung")
    else:
        st.error("❌ DB BPS Tidak Ditemukan")

    # WIDGET LANGSUNG (TANPA st.form) AGAR AUTO-LOAD
    f_tahun = st.selectbox("Tahun", reversed(TAHUN_TERSEDIA), index=1)
    f_periode_label = st.selectbox("Periode", list(PERIODE_OPSI.keys()))
    f_periode = PERIODE_OPSI[f_periode_label]
    f_sumber_label = st.radio("Jenis Perdagangan", ["Ekspor", "Impor"])
    f_sumber = "1" if f_sumber_label == "Ekspor" else "2"
    f_unit = st.radio("Satuan", ["USD", "Miliar USD"])
    
    # OTOMATIS MEMUAT DATA KETIKA FILTER DIUBAH
    tipe, bulan = get_periode_params(f_periode)
    with st.spinner('Menyiapkan data BPS...'):
        df_raw = fetch_bps_db(f_sumber, str(f_tahun), tipe, bulan)
        if not df_raw.empty:
            st.session_state['bps_data'] = df_raw
            st.session_state['bps_meta'] = {"tahun": f_tahun, "sumber": f_sumber_label, "sumber_kode": f_sumber, "unit": f_unit, "tipe": tipe, "bulan": bulan}
            st.success(f"Berhasil memuat {len(df_raw):,} baris data.")
        else:
            st.session_state['bps_data'] = pd.DataFrame()
            st.session_state['bps_meta'] = {}
            st.error(f"Tidak ada data BPS {f_sumber_label} {f_tahun} untuk filter tersebut.")

# ── Tabs Setup ──
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Ringkasan BPS", "🗄️ Data Lengkap", "⚠️ Early Warning System", 
    "🪞 Mirroring", "🏦 Neraca Pembayaran"
])

df_bps = st.session_state.get('bps_data', pd.DataFrame())
meta = st.session_state.get('bps_meta', {})

if not df_bps.empty:
    div = 1e9 if meta['unit'] == "Miliar USD" else 1
    df_bps_clean = df_bps.copy()
    df_bps_clean["value"] = df_bps_clean["value"] / div
    
    kmd = df_bps_clean.groupby("kodehs", as_index=False)[["value","berat"]].sum().sort_values("value", ascending=False)
    neg = df_bps_clean.groupby("negara", as_index=False)["value"].sum().sort_values("value", ascending=False)
    kmd["deskripsi"] = kmd["kodehs"].map(HS_DESC).fillna("Lainnya")
    kmd["label"] = kmd["kodehs"].astype(str) + " - " + kmd["deskripsi"].str[:15]

# ── TAB 1: Ringkasan ──
with tab1:
    if df_bps.empty:
        st.info("👈 Silakan atur filter pada sidebar di sebelah kiri.")
    else:
        # Row 1: KPI
        k1, k2, k3 = st.columns(3)
        c_green = "#3fb950" if meta['sumber'] == "Ekspor" else "#f78166"
        with k1: kpi_card(f"TOTAL {meta['sumber'].upper()}", f"{df_bps_clean['value'].sum():,.2f} {meta['unit']}", c_green)
        with k2: kpi_card("KOMODITAS TERBESAR (HS)", kmd.iloc[0]["kodehs"] if not kmd.empty else "-", "#58a6ff")
        with k3: kpi_card("NEGARA TUJUAN/ASAL UTAMA", neg.iloc[0]["negara"] if not neg.empty else "-", "#e3b341")
        
        st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
        
        # Row 2: Historis & Top 15
        r1_left, r1_right = st.columns(2)
        with r1_left:
            section_title(f"HISTORIS TREN ({TAHUN_TERSEDIA[0]}–{TAHUN_SAAT_INI})")
            h1, h2, h3, h4 = st.columns([2, 2, 2, 2])
            hs_hist = h1.selectbox("Pilih HS untuk Histori", options=HS_ALL, index=26, label_visibility="collapsed")
            negara_hist = h2.selectbox("Filter Negara", options=["Semua Negara"] + PARTNER_LIST, label_visibility="collapsed")
            metric_hist = h3.radio("Metrik Histori", ["Nilai", "YoY %"], horizontal=True, label_visibility="collapsed")
            btn_hist = h4.button("Tampilkan Histori", use_container_width=True)
            
            if btn_hist:
                with st.spinner("Menarik data historis..."):
                    df_hist_raw = fetch_hist_bps_db(meta['sumber_kode'], hs_hist, meta['tipe'], meta['bulan'])
                    if df_hist_raw.empty:
                        st.warning("Tidak ada data historis.")
                    else:
                        if negara_hist != "Semua Negara": df_hist_raw = df_hist_raw[df_hist_raw["negara"] == normalize_negara(negara_hist)]
                        df_h = df_hist_raw.groupby("tahun", as_index=False)["value"].sum().sort_values("tahun")
                        df_h["Tahun"], df_h["Value"] = df_h["tahun"].astype(str), df_h["value"] / div
                        
                        if metric_hist == "YoY %":
                            df_h["Value"] = df_h["Value"].pct_change() * 100
                            fig_hist = px.line(df_h, x="Tahun", y="Value", markers=True, title=f"YoY (%) – HS {hs_hist}")
                        else:
                            fig_hist = px.line(df_h, x="Tahun", y="Value", markers=True, title=f"Tren Nilai ({meta['unit']}) – HS {hs_hist}")
                        
                        fig_hist.update_layout(xaxis=dict(type='category'), margin=dict(l=0, r=0, t=30, b=0))
                        st.plotly_chart(fig_hist, use_container_width=True)
        
        with r1_right:
            section_title("TOP 15 KOMODITAS (HS)")
            fig_kmd = px.bar(kmd.head(15), y="label", x="value", orientation='h')
            fig_kmd.update_yaxes(categoryorder='total ascending', type='category')
            fig_kmd.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350)
            fig_kmd.update_traces(marker_color=c_green)
            st.plotly_chart(fig_kmd, use_container_width=True)
        
        st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
        
        # Row 3: Top Negara & Share Pie
        r2_left, r2_right = st.columns(2)
        with r2_left:
            section_title("TOP NEGARA MITRA")
            fig_neg = px.bar(neg.head(15), y="negara", x="value", orientation='h')
            fig_neg.update_yaxes(categoryorder='total ascending', type='category')
            fig_neg.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350)
            fig_neg.update_traces(marker_color="#58a6ff")
            st.plotly_chart(fig_neg, use_container_width=True)
            
        with r2_right:
            section_title("SHARE KOMODITAS (TOP 8)")
            fig_pie = px.pie(kmd.head(8), values='value', names='label', hole=0.45)
            fig_pie.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350, legend=dict(orientation="v"))
            st.plotly_chart(fig_pie, use_container_width=True)

# ── TAB 2: Data Lengkap ──
with tab2:
    if not df_bps.empty:
        st.markdown("<div style='display: flex; justify-content: space-between; align-items: center;'>", unsafe_allow_html=True)
        section_title("TABEL DATA EKSPOR/IMPOR")
        full_df = df_bps_clean.groupby(["negara","kodehs"], as_index=False)[["value","berat"]].sum()
        full_df["deskripsi"] = full_df["kodehs"].map(HS_DESC).fillna("Lainnya")
        
        csv = full_df.to_csv(index=False).encode('utf-8')
        st.download_button(label="⬇ Download CSV", data=csv, file_name=f"BPS_{meta['sumber']}_{meta['tahun']}.csv", mime='text/csv')
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.dataframe(full_df, use_container_width=True, hide_index=True)

# ── TAB 3: EWS ──
with tab3:
    if not df_bps.empty:
        section_title("DETEKSI ANOMALI (EARLY WARNING SYSTEM)")
        st.caption("Batas Atas = Konsentrasi berlebih. Batas Bawah = Underperforming. Anomali Harga = Indikasi lonjakan harga ekstrem.")
        
        ews_df = calculate_ews(kmd.copy())
        
        def highlight_ews(val):
            color = ''
            if 'Atas' in str(val): color = '#ffcccb' 
            elif 'Bawah' in str(val): color = '#ffe8b5'
            elif 'Harga' in str(val): color = '#dcbdfb'
            elif 'KRITIS' in str(val): color = '#ff7b72'
            return f'background-color: {color}; color: black'
        
        cols_to_show = [c for c in ews_df.columns if c != "label"]
        st.dataframe(ews_df[cols_to_show].style.map(highlight_ews, subset=['status_ews']), use_container_width=True, hide_index=True)

# ── TAB 4: Mirroring ──
with tab4:
    section_title("ANALISIS ASIMETRI PENCATATAN (BPS VS ITC TRADE MAP)")
    st.markdown("<br>", unsafe_allow_html=True)
    
    if not df_bps.empty:
        st.info(f"ℹ️ Saat ini menganalisis asimetri **{meta.get('sumber', 'Ekspor/Impor')}** tahun **{meta.get('tahun', '-')}**. (Ubah pengaturan di sidebar kiri jika ingin mengganti sumber).")
    
    col_m1, col_m2, col_m3 = st.columns([2, 2, 1])
    with col_m1: mitra_mirror = st.selectbox("Mitra Dagang", PARTNER_LIST)
    with col_m2: unit_mirror = st.radio("Satuan Mirroring", ["USD", "Juta USD"], horizontal=True)
    with col_m3: 
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True) 
        btn_mirror = st.button("▶ JALANKAN MIRRORING", type="primary", use_container_width=True)
        
    if btn_mirror:
        if not df_bps.empty:
            div_m = 1e6 if unit_mirror == "Juta USD" else 1
            with st.spinner("Mencocokkan data..."):
                try:
                    df_tm, status = load_trademap(mitra_mirror, meta['tahun'], meta['sumber_kode'])
                    
                    if status == "SUCCESS":
                        df_bps_m = df_bps[df_bps["negara"] == normalize_negara(mitra_mirror)].copy()
                        df_bps_m = df_bps_m.groupby("kodehs", as_index=False)["value"].sum().rename(columns={"kodehs":"HS","value":"BPS_Value"})
                        
                        df_merge = pd.merge(df_bps_m, df_tm, on="HS", how="outer").fillna(0)
                        df_merge[["BPS_Value", "Trademap_Value"]] /= div_m
                        df_merge["Selisih"] = df_merge["Trademap_Value"] - df_merge["BPS_Value"]
                        df_merge["Deskripsi"] = df_merge["HS"].map(HS_DESC).fillna("Lainnya")
                        
                        st.success("✅ Mirroring selesai.")
                        
                        lbl_bps = "EKSPOR IDN KE" if str(meta['sumber_kode']) == "1" else "IMPOR IDN DARI"
                        lbl_tm  = "IMPOR MITRA DARI IDN" if str(meta['sumber_kode']) == "1" else "EKSPOR MITRA KE IDN"
                        
                        cm1, cm2, cm3 = st.columns(3)
                        with cm1: kpi_card(f"{lbl_bps} {mitra_mirror.upper()} (BPS)", f"{df_merge['BPS_Value'].sum():,.1f} {unit_mirror}", "#3fb950")
                        with cm2: kpi_card(f"{lbl_tm} (TRADE MAP)", f"{df_merge['Trademap_Value'].sum():,.1f} {unit_mirror}", "#58a6ff")
                        with cm3: kpi_card("TOTAL SELISIH ASIMETRI", f"{df_merge['Selisih'].sum():,.1f} {unit_mirror}", "#e3b341")
                        
                        cg1, cg2 = st.columns(2)
                        with cg1:
                            section_title("10 HS TERBESAR BPS")
                            top10 = df_merge.nlargest(10, "BPS_Value").copy()
                            top10["label"] = top10["HS"].astype(str) + " - " + top10["Deskripsi"].str[:15]
                            fig_cmp = go.Figure([
                                go.Bar(name="BPS", x=top10["label"], y=top10["BPS_Value"], marker_color="#3fb950"),
                                go.Bar(name="Trade Map", x=top10["label"], y=top10["Trademap_Value"], marker_color="#f78166"),
                            ]).update_layout(barmode="group", height=320, margin=dict(l=0, r=0, t=30, b=0), 
                                             legend=dict(orientation="h", y=1.1), xaxis=dict(type='category')) 
                            st.plotly_chart(fig_cmp, use_container_width=True)
                        with cg2:
                            section_title("5 HS ASIMETRI TERBESAR")
                            top5 = df_merge.assign(Abs_Diff=df_merge["Selisih"].abs()).nlargest(5, "Abs_Diff")
                            top5["label"] = top5["HS"].astype(str) + " - " + top5["Deskripsi"].str[:15]
                            fig_diff = go.Figure([
                                go.Bar(name="BPS", x=top5["label"], y=top5["BPS_Value"], marker_color="#3fb950"),
                                go.Bar(name="Trade Map", x=top5["label"], y=top5["Trademap_Value"], marker_color="#f78166"),
                            ]).update_layout(barmode="group", height=320, margin=dict(l=0, r=0, t=30, b=0), 
                                             legend=dict(orientation="h", y=1.1), xaxis=dict(type='category'))
                            st.plotly_chart(fig_diff, use_container_width=True)
                            
                        st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                        st.markdown("<div style='display: flex; justify-content: space-between; align-items: center;'>", unsafe_allow_html=True)
                        section_title("DATA LENGKAP ASIMETRI")
                        csv_mirror = df_merge[["HS", "Deskripsi", "BPS_Value", "Trademap_Value", "Selisih"]].to_csv(index=False).encode('utf-8')
                        st.download_button("⬇ Download CSV", data=csv_mirror, file_name=f"Mirroring_{mitra_mirror}_{meta['tahun']}.csv", mime='text/csv')
                        st.markdown("</div>", unsafe_allow_html=True)
                        st.dataframe(df_merge[["HS", "Deskripsi", "BPS_Value", "Trademap_Value", "Selisih"]], use_container_width=True, hide_index=True)
                            
                    elif status == "FILE_NOT_FOUND": st.error("❌ File 'data_trademap.xlsx' tidak ditemukan.")
                    elif status.startswith("DATA_EMPTY_TAHUN"):
                        tahun_ada = status.split("|")[1] if "|" in status else "-"
                        st.warning(f"⚠️ Data Trade Map untuk '{mitra_mirror}' tahun {meta['tahun']} belum ada. Tahun di Excel: {tahun_ada}")
                    else: st.error(f"Gagal memuat Trade Map: {status}")
                except Exception as e:
                    st.error(f"Kesalahan mirroring: {str(e)}")
                    st.code(traceback.format_exc())
        else:
            st.warning("Muat data BPS di Sidebar terlebih dahulu!")

# ── TAB 5: SEKI BI ──
with tab5:
    if not os.path.exists(BOP_DB_PATH): st.error("❌ Database bop_indonesia.db tidak ditemukan.")
    else:
        section_title("NERACA PEMBAYARAN — SEKI BANK INDONESIA (2004–2025)")
        st.markdown("<br>", unsafe_allow_html=True)
        
        with st.form("seki_form"):
            cs1, cs2, cs3, cs4, cs5 = st.columns([1, 1, 1, 1, 1])
            y1 = cs1.number_input("Tahun Awal", min_value=2004, max_value=2025, value=2015)
            y2 = cs2.number_input("Tahun Akhir", min_value=2004, max_value=2025, value=2024)
            freq = cs3.selectbox("Frekuensi", ["Kuartalan", "Tahunan"])
            unit_s = cs4.selectbox("Satuan", ["Juta USD", "Miliar USD"])
            
            cs5.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            sub_seki = cs5.form_submit_button("▶ TAMPILKAN NERACA", use_container_width=True)
            
        if sub_seki:
            f_val = "quarterly" if freq == "Kuartalan" else "annual"
            div_s = 1000 if unit_s == "Miliar USD" else 1
            needed_ids = [1,2,17,20,23,26,29,32,35,40,41,46,47,48,54,55,56]
            df_seki = bop_series(needed_ids, y1, y2, f_val)
            
            if not df_seki.empty:
                df_seki["nilai"] = df_seki["value_mn_usd"] / div_s
                
                ck1, ck2, ck3 = st.columns(3)
                ca_v, cad_v, ner_v, _BOP_LATEST = bop_latest_val(1), bop_latest_val(54), bop_latest_val(48), bop_latest()
                
                with ck1: kpi_card("TRANSAKSI BERJALAN", f"{(ca_v/div_s if ca_v else 0):,.1f} {unit_s}", "#3fb950" if (ca_v or 0) >= 0 else "#f78166", f"Periode: {_BOP_LATEST}")
                with ck2: kpi_card("CADANGAN DEVISA", f"{(cad_v/div_s if cad_v else 0):,.1f} {unit_s}", "#58a6ff", f"Periode: {_BOP_LATEST}")
                with ck3: kpi_card("NERACA KESELURUHAN", f"{(ner_v/div_s if ner_v else 0):,.1f} {unit_s}", "#bc8cff", f"Periode: {_BOP_LATEST}")

                def gs(iid):
                    s = df_seki[df_seki["item_id"] == iid].copy()
                    s = s.sort_values("year" if f_val == "annual" else ["year","quarter"])
                    s["v"] = s["nilai"]
                    return s
                xcol = "period" if f_val == "quarterly" else "year"

                cc1, cc2 = st.columns([3, 2])
                with cc1:
                    section_title("TREN TRANSAKSI BERJALAN & KOMPONEN")
                    ca_df = df_seki[df_seki['item_id'].isin([2, 17, 20, 23, 1])] 
                    fig_ca = px.bar(ca_df[ca_df['item_id']!=1], x=xcol, y="nilai", color="keterangan", barmode="relative")
                    fig_ca.add_scatter(x=ca_df[ca_df['item_id']==1][xcol], y=ca_df[ca_df['item_id']==1]["nilai"], name="Total CA", line=dict(color="#000000", width=2))
                    fig_ca.update_layout(height=320, margin=dict(l=0, r=0, t=30, b=0), legend=dict(orientation="h", y=1.1, title=""), xaxis=dict(type='category'))
                    st.plotly_chart(fig_ca, use_container_width=True)
                with cc2:
                    section_title("DEKOMPOSISI NERACA (TOTAL PERIODE)")
                    wf_ids = [1, 26, 29, 47, 48]
                    wf_lbl = ["Transaksi<br>Berjalan","Transaksi<br>Modal","Transaksi<br>Finansial","Selisih<br>Perhitungan","Neraca<br>Keseluruhan"]
                    wf_v = [float(gs(iid)["v"].sum()) if not gs(iid).empty else 0 for iid in wf_ids]
                    fig_wf = go.Figure(go.Waterfall(x=wf_lbl, measure=["relative","relative","relative","relative","total"], y=wf_v, textposition="outside", decreasing=dict(marker_color="#f78166"), increasing=dict(marker_color="#3fb950"), totals=dict(marker_color="#bc8cff")))
                    fig_wf.update_layout(height=320, margin=dict(l=0, r=0, t=30, b=0), showlegend=False)
                    st.plotly_chart(fig_wf, use_container_width=True)

                cd1, cd2 = st.columns(2)
                with cd1:
                    section_title("TRANSAKSI FINANSIAL: KOMPONEN INVESTASI")
                    fig_inv = go.Figure()
                    for iid, lbl_i, c in [(32,"FDI","#3fb950"),(35,"Portofolio","#58a6ff"),(41,"Lainnya","#ffa657"),(40,"Derivatif","#e3b341")]:
                        s = gs(iid)
                        if not s.empty: fig_inv.add_trace(go.Bar(x=s[xcol], y=s["v"], name=lbl_i, marker_color=c))
                    s_fin = gs(29)
                    if not s_fin.empty: fig_inv.add_trace(go.Scatter(x=s_fin[xcol], y=s_fin["v"], name="Total Fin.", line=dict(color="#bc8cff", width=2, dash="dot")))
                    fig_inv.update_layout(barmode="relative", height=320, margin=dict(l=0, r=0, t=30, b=0), legend=dict(orientation="h", y=1.1), xaxis=dict(type='category'))
                    st.plotly_chart(fig_inv, use_container_width=True)
                with cd2:
                    section_title("CADANGAN DEVISA")
                    cad_df = df_seki[df_seki['item_id'] == 54]
                    fig_cad = px.area(cad_df, x=xcol, y="nilai")
                    fig_cad.update_traces(line_color="#39d0d8", fillcolor="rgba(57,208,216,0.1)")
                    fig_cad.update_layout(height=320, margin=dict(l=0, r=0, t=30, b=0), xaxis=dict(type='category'))
                    st.plotly_chart(fig_cad, use_container_width=True)

                ce1, ce2 = st.columns(2)
                with ce1:
                    section_title("KOMPARASI INDIKATOR")
                    seki_cmp = st.multiselect("Pilih Indikator:", options=list(BOP_MAIN_ITEMS.values()), default=["Transaksi Finansial", "Neraca Keseluruhan"], label_visibility="collapsed")
                    inv_map = {v: k for k, v in BOP_MAIN_ITEMS.items()}
                    fig_cmp = go.Figure()
                    pal = ["#58a6ff", "#3fb950", "#f78166", "#e3b341", "#bc8cff"]
                    for i, ind in enumerate(seki_cmp):
                        s = gs(inv_map[ind])
                        if not s.empty: fig_cmp.add_trace(go.Scatter(x=s[xcol], y=s["v"], name=ind, mode="lines+markers", line=dict(color=pal[i%len(pal)], width=2)))
                    fig_cmp.update_layout(height=320, margin=dict(l=0, r=0, t=30, b=0), legend=dict(orientation="h", y=1.1), xaxis=dict(type='category'))
                    st.plotly_chart(fig_cmp, use_container_width=True)
                with ce2:
                    section_title("CURRENT ACCOUNT % PDB")
                    s_pct = df_seki[df_seki["item_id"] == 56].copy().sort_values("year" if f_val == "annual" else ["year","quarter"])
                    fig_pct = go.Figure()
                    if not s_pct.empty:
                        fig_pct.add_trace(go.Bar(x=s_pct[xcol], y=s_pct["value_mn_usd"], marker_color=["#3fb950" if v >= 0 else "#f78166" for v in s_pct["value_mn_usd"]]))
                    fig_pct.update_layout(height=320, margin=dict(l=0, r=0, t=30, b=0), xaxis=dict(type='category'))
                    st.plotly_chart(fig_pct, use_container_width=True)

                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                st.markdown("<div style='display: flex; justify-content: space-between; align-items: center;'>", unsafe_allow_html=True)
                section_title("DATA LENGKAP NERACA PEMBAYARAN")
                
                seki_tbl_filter = st.multiselect("Filter Indikator:", options=list(BOP_MAIN_ITEMS.values()), default=None, label_visibility="collapsed")
                df_seki_table = df_seki.copy()
                if seki_tbl_filter:
                    df_seki_table = df_seki_table[df_seki_table["keterangan"].isin(seki_tbl_filter)]
                
                csv_seki = df_seki_table[["year", "period", "keterangan", "items_en", "nilai"]].to_csv(index=False).encode('utf-8')
                st.download_button("⬇ Download CSV", data=csv_seki, file_name=f"SEKI_{y1}_{y2}.csv", mime='text/csv')
                st.markdown("</div>", unsafe_allow_html=True)
                
                st.dataframe(df_seki_table[["year", "period", "keterangan", "items_en", "nilai"]], use_container_width=True, hide_index=True)
            else:
                st.warning("Tidak ada data SEKI untuk rentang waktu tersebut.")
