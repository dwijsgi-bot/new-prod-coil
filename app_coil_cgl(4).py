"""
Coil CGL Production & Quality Dashboard
=======================================

Jalankan:
    pip install streamlit plotly pandas openpyxl numpy
    streamlit run app_coil_cgl.py

Butuh Streamlit >= 1.35 (fitur klik-sampling pada grafik).

Fitur:
- Monitoring proses: sumbu X = UPDAT_DT (datetime), batas atas/bawah,
  cursor hairline (crosshair). Klik titik -> cursor permanen (daftar cursor).
- Multi panel & event: beberapa parameter bertumpuk dengan sumbu waktu sama,
  garis target ± toleransi, garis event maintenance, dan pemisah tahun.
- Semua data (coil, cursor, event, panel, batas proses, pengaturan) disimpan di
  SQLite: coil_cgl.db di folder aplikasi. File baru dibandingkan dulu dengan
  database (baru / berubah / sama) sebelum disimpan.
- Sifat mekanik: pass/fail YP/TS/EL terhadap LL/UL, Cp/Cpk, control chart.
- Profil suhu strip, setpoint zona tungku, dan suhu galvanizing.
- Skin pass vs sifat mekanik.
- Detail per coil (traceability).
- Grade & defect, Target vs Actual.
- Validasi data otomatis.
- Tabel data + ekspor CSV/Excel.
"""

from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from contextlib import closing
import json
import re
import sqlite3
import textwrap

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st


# =====================================================
# KONFIGURASI
# =====================================================
st.set_page_config(
    page_title="Coil CGL",
    page_icon="🏭",
    layout="wide",
)

st.markdown(
    """
    <style>
    div[data-testid="stMetric"] {
        background: rgba(74,144,217,0.06);
        border: 1px solid rgba(74,144,217,0.18);
        padding: 10px 14px;
        border-radius: 8px;
    }
    div[data-testid="stMetricValue"] { font-size: 1.5rem; }
    section[data-testid="stSidebar"] div[role="radiogroup"] label { padding: 4px 0; font-size: 1.02rem; }
    h2 { margin-bottom: 0.2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

INVISIBLE_CHARS = "[\u200b\u200c\u200d\ufeff\xa0]"  # karakter asli, bukan escape regex
ID_COL = "INLET_COIL_NO"
UPDAT_DT_COL = "UPDAT_DT_DATETIME"      # hasil parse UPDAT_DT
SHEAR_DT_COL = "SHEAR_CUT_DATETIME"     # hasil parse SHEAR_CUT_DT
ROW_KEY = "_ROW"

# Format UPDAT_DT: d/m/yyyy h:mm  (contoh "8/1/2025 8:30" = 8 Januari 2025 08:30)
UPDAT_DT_FORMATS = [
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%d-%m-%Y %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%d/%m/%Y",
]

# Batas proses default.
# Aturan: nilai >= batas atas  -> keluar batas (atas)
#         nilai <= batas bawah -> keluar batas (bawah)
DEFAULT_PROCESS_LIMITS = {
    "ANLNG_S_TP_HSD_TP_STRIP": {"lower": 780.0, "upper": 820.0},
    "SS direct - Multi": {"lower": -10.0, "upper": 20.0},
}

# Sifat mekanik: (kolom nilai, kolom LL, kolom UL, satuan)
# Nilai LL/UL = 0 diartikan "tidak ada batas".
MECH_PROPS = {
    "YP": ("YP Tail", "YP LL", "YP UL", "MPa"),
    "TS": ("TS Tail", "TS LL", "TS UL", "MPa"),
    "EL": ("EL Tail", "EL LL", "EL UL", "%"),
    "BH": ("BH Tail", "BH LL", "BH UL", "MPa"),
}

STRIP_STAGES = [
    ("ANLNG_S_TP_HSE_TP_STRIP", "HSE"),
    ("ANLNG_S_TP_HSD_TP_STRIP", "HSD"),
    ("ANLNG_S_TP_SS_TP_STRIP", "SS"),
    ("ANLNG_S_TP_SS_MLTLEF_TP_STRIP", "SS multi"),
    ("ANLNG_S_TP_RCS_SCAN_TP_STRIP", "RCS scan"),
    ("ANLNG_S_TP_RCS_MLTLEF_TP_STRIP", "RCS multi"),
    ("ANLNG_S_TP_6BR_SCAN_TP_STRIP", "6BR scan"),
    ("ANLNG_S_TP_TDS_MLTLEF_TP_STRIP", "TDS multi"),
    ("ANLNG_S_TP_POT_CALC_TP_STRIP", "POT calc"),
]

FURNACE_ZONES = [
    ("ANLNG_F_TP_HS_1ZN_F_TP", "HS 1Z"),
    ("ANLNG_F_TP_HS_2ZN_F_TP", "HS 2Z"),
    ("ANLNG_F_TP_HS_3ZN_F_TP", "HS 3Z"),
    ("ANLNG_F_TP_HS_4ZN_F_TP", "HS 4Z"),
    ("ANLNG_F_TP_SS_1ZN_F_TP", "SS 1Z"),
    ("ANLNG_F_TP_SS_2ZN_F_TP", "SS 2Z"),
    ("ANLNG_F_TP_4BR_F_TP", "4BR"),
    ("ANLNG_F_TP_5BR_F_TP", "5BR"),
    ("ANLNG_F_TP_6BR_F_TP", "6BR"),
    ("ANLNG_F_TP_TDS_1ZN_F_TP", "TDS 1Z"),
    ("ANLNG_F_TP_TDS_2ZN_F_TP", "TDS 2Z"),
]

GALV_POINTS = [
    ("HD_GALV_R_GAIH_OTLET_TP_STRIP", "GAIH outlet"),
    ("HD_GALV_R_GASS_OTLET_TP_STRIP", "GASS outlet"),
    ("HD_GALV_R_TOPROLL_TP_STRIP", "Top roll"),
    ("HD_GALV_R_WCS_INLET_TP_STRIP", "WCS inlet"),
]

SKIN_PASS = [
    ("SKP_R_AVE_EXTEN_PC", "Elongation skin pass (%)"),
    ("SKP_R_ROLL_LOAD", "Roll load"),
    ("Rolling Force (x10 kN/mm)", "Rolling force (x10 kN/mm)"),
    ("SKP_R_ENT_TENSN_TTL", "Tension masuk"),
    ("SKP_R_DLV_TENSN_TTL", "Tension keluar"),
]

# Kolom selisih suhu di file beserta rumus aslinya (kolom1 - kolom2)
DELTA_COLUMNS = [
    ("HSE - HSD", "ANLNG_S_TP_HSE_TP_STRIP", "ANLNG_S_TP_HSD_TP_STRIP"),
    ("HSD - SS Direct", "ANLNG_S_TP_HSD_TP_STRIP", "ANLNG_S_TP_SS_TP_STRIP"),
    ("HSD - SS Multi", "ANLNG_S_TP_HSD_TP_STRIP", "ANLNG_S_TP_SS_MLTLEF_TP_STRIP"),
    ("SS direct - Multi", "ANLNG_S_TP_SS_TP_STRIP", "ANLNG_S_TP_SS_MLTLEF_TP_STRIP"),
    ("RCS Direct - Multi", "ANLNG_S_TP_RCS_SCAN_TP_STRIP", "ANLNG_S_TP_RCS_MLTLEF_TP_STRIP"),
    ("6BR Direct - Multi", "ANLNG_S_TP_6BR_SCAN_TP_STRIP", "ANLNG_S_TP_TDS_MLTLEF_TP_STRIP"),
]

TARGET_ACTUAL_PAIRS = [
    ("Target1", "Actual1", "Target 1 vs Actual 1"),
    ("Target2", "Actual 2", "Target 2 vs Actual 2"),
    ("Target3", "Actual3", "Target 3 vs Actual 3"),
]

# Kolom teks/kode: tidak boleh dikonversi otomatis ke angka
PROTECTED_TEXT = {
    "INLET_COIL_NO", "OTLET_COIL_NO", "JSG_DESI_NM", "UPDAT_DT",
    "Del Shear Cut Time", "SHEAR_CUT_DT", "WORK_DT",
    "NO1_TRG_GRADE_CD", "CGL UP", "TOP_SRFCE_GRADE_CD", "SHAPE_GRADE_CD",
    "BOTTM_SRFCE_GRADE_CD", "NO1_DEFCT_CD", "NO1_DEFCT_GRADE_CD",
    "NO2_DEFCT_CD", "NO2_DEFCT_GRADE_CD", "NO3_DEFCT_CD",
    "NO3_DEFCT_GRADE_CD", "REPET_ORDER_NO", "CAL/CGL", "NO1_CUSTM_SB",
    "C00_CHEM_CD", "Thickness group", "Sample", "Date",
}

# Kolom yang tidak relevan sebagai parameter tren
NON_TREND_COLUMNS = {
    "SHEAR_CUT_DT", "WORK_DT", "Year", "Month", "Day", "DEVIDER",
    "RSLT_CD", ROW_KEY,
}

DEFAULT_HOVER = [
    "OTLET_COIL_NO", "JSG_DESI_NM", "Thickness group", "INLET_COIL_TH",
    "INLET_COIL_WH", "COIL_WT", "STATUS_MEKANIK", "NO1_DEFCT_CD",
]

# Panel default untuk menu Multi Panel (mengikuti format grafik HSE-HSD & SS Direct-SS Multi)
PANEL_COLUMNS = ["Parameter", "Judul", "Target", "Batas bawah", "Batas atas", "Y min", "Y max"]
DEFAULT_PANELS = [
    {"Parameter": "ANLNG_S_TP_HSD_TP_STRIP", "Judul": "Suhu strip HSD", "Target": None,
     "Batas bawah": 780.0, "Batas atas": 820.0, "Y min": None, "Y max": None},
    {"Parameter": "HSE - HSD", "Judul": "HSE-HSD", "Target": -150.0,
     "Batas bawah": None, "Batas atas": None, "Y min": -200.0, "Y max": 0.0},
    {"Parameter": "SS direct - Multi", "Judul": "SS Direct - SS Multi", "Target": 5.0,
     "Batas bawah": -10.0, "Batas atas": 20.0, "Y min": -25.0, "Y max": 40.0},
]

# Di atas jumlah titik ini grafik memakai WebGL (jauh lebih ringan untuk ribuan coil)
GL_THRESHOLD = 1500

# Nama kolom yang mudah dibaca (draf - silakan sesuaikan dengan istilah di pabrik)
COLUMN_LABELS = {
    "INLET_COIL_NO": "No. coil masuk", "OTLET_COIL_NO": "No. coil keluar",
    "INLET_COIL_TH": "Tebal (mm)", "INLET_COIL_WH": "Lebar (mm)", "INLET_COIL_LG": "Panjang (m)",
    "COIL_WT": "Berat coil (kg)", "JSG_DESI_NM": "Grade", "UPDAT_DT_DATETIME": "Waktu (UPDAT_DT)",
    "SHEAR_CUT_DATETIME": "Waktu shear cut",
    "ANLNG_S_TP_HSE_TP_STRIP": "Suhu strip HSE (°C)", "ANLNG_S_TP_HSD_TP_STRIP": "Suhu strip HSD (°C)",
    "ANLNG_S_TP_SS_TP_STRIP": "Suhu strip SS direct (°C)", "ANLNG_S_TP_SS_MLTLEF_TP_STRIP": "Suhu strip SS multi (°C)",
    "ANLNG_S_TP_RCS_SCAN_TP_STRIP": "Suhu strip RCS scan (°C)", "ANLNG_S_TP_RCS_MLTLEF_TP_STRIP": "Suhu strip RCS multi (°C)",
    "ANLNG_S_TP_6BR_SCAN_TP_STRIP": "Suhu strip 6BR scan (°C)", "ANLNG_S_TP_TDS_MLTLEF_TP_STRIP": "Suhu strip TDS multi (°C)",
    "ANLNG_S_TP_POT_CALC_TP_STRIP": "Suhu strip pot (°C)",
    "HSE - HSD": "Selisih HSE − HSD (°C)", "HSD - SS Direct": "Selisih HSD − SS direct (°C)",
    "HSD - SS Multi": "Selisih HSD − SS multi (°C)", "SS direct - Multi": "Selisih SS direct − SS multi (°C)",
    "RCS Direct - Multi": "Selisih RCS direct − multi (°C)", "6BR Direct - Multi": "Selisih 6BR − TDS multi (°C)",
    "Cooling Rate": "Cooling rate", "SKP_R_AVE_EXTEN_PC": "Elongation skin pass (%)",
    "SKP_R_ROLL_LOAD": "Roll load skin pass", "Rolling Force (x10 kN/mm)": "Rolling force (x10 kN/mm)",
    "SKP_R_ENT_TENSN_TTL": "Tension masuk skin pass", "SKP_R_DLV_TENSN_TTL": "Tension keluar skin pass",
    "YP Tail": "Yield point (MPa)", "TS Tail": "Tensile strength (MPa)", "EL Tail": "Elongation (%)",
    "BH Tail": "Bake hardening (MPa)", "STATUS_MEKANIK": "Status mekanik",
    "TOP_SRFCE_GRADE_CD": "Grade permukaan atas", "BOTTM_SRFCE_GRADE_CD": "Grade permukaan bawah",
    "SHAPE_GRADE_CD": "Grade bentuk", "NO1_DEFCT_CD": "Defect utama", "NO2_DEFCT_CD": "Defect ke-2",
    "NO3_DEFCT_CD": "Defect ke-3", "Thickness group": "Kelompok tebal",
    "HD_GALV_R_GAIH_OTLET_TP_STRIP": "Suhu GAIH outlet (°C)", "HD_GALV_R_GASS_OTLET_TP_STRIP": "Suhu GASS outlet (°C)",
    "HD_GALV_R_TOPROLL_TP_STRIP": "Suhu top roll (°C)", "HD_GALV_R_WCS_INLET_TP_STRIP": "Suhu WCS inlet (°C)",
    "AVE_CE_SP": "CE speed rata-rata", "MIN_CE_SP": "CE speed min", "MAX_CE_SP": "CE speed max",
    "LINE_SPEED": "Line speed",
}


def label(col):
    """Nama kolom yang mudah dibaca, dengan kode asli bila berbeda."""
    return COLUMN_LABELS.get(col, col)


def label_code(col):
    name = COLUMN_LABELS.get(col)
    return f"{name}  ·  {col}" if name and name != col else col


def scatter_cls(n_points):
    return go.Scattergl if n_points > GL_THRESHOLD else go.Scatter

AGGREGATION_RULES = {
    "Per coil": None,
    "Rata-rata per jam": "h",
    "Rata-rata per hari": "D",
    "Rata-rata per minggu": "W-MON",
    "Rata-rata per bulan": "MS",
}

# Database SQLite (di folder yang sama dengan aplikasi)
try:
    APP_DIR = Path(__file__).resolve().parent
except NameError:
    APP_DIR = Path.cwd()
DB_PATH = APP_DIR / "coil_cgl.db"
LEGACY_STATE_FILE = APP_DIR / "coil_dashboard_state.json"   # dari versi sebelumnya (dimigrasi sekali)
KEY_COL = "INLET_COIL_NO"   # kunci unik coil di database

COLOR_CURSOR = "#7B1FA2"
COLOR_EVENT = "#222222"
COLOR_YEAR = "#D32F2F"

COLOR_OK = "#3A6EA5"
COLOR_HIGH = "#C62828"
COLOR_LOW = "#E08A00"
COLOR_NA = "#9AA3AD"
COLOR_LIMIT = "#C62828"
COLOR_BAND = "rgba(46,125,107,0.08)"

STATUS_PROCESS_STYLE = {
    "Dalam batas": dict(color=COLOR_OK, symbol="circle"),
    "≥ batas atas": dict(color=COLOR_HIGH, symbol="triangle-up"),
    "≤ batas bawah": dict(color=COLOR_LOW, symbol="triangle-down"),
}

STATUS_MECH_STYLE = {
    "Dalam spesifikasi": dict(color=COLOR_OK, symbol="circle"),
    "Di atas UL": dict(color=COLOR_HIGH, symbol="triangle-up"),
    "Di bawah LL": dict(color=COLOR_LOW, symbol="triangle-down"),
    "Tanpa spesifikasi": dict(color=COLOR_NA, symbol="circle-open"),
}


# =====================================================
# FUNGSI BACA & SIAPKAN DATA
# =====================================================
@st.cache_data(show_spinner=False)
def read_table(file_bytes, file_name, sheet_name, header_row):
    bio = BytesIO(file_bytes)
    if Path(file_name).suffix.lower() == ".csv":
        return pd.read_csv(bio, header=header_row - 1)
    return pd.read_excel(bio, sheet_name=sheet_name, header=header_row - 1)


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes):
    return pd.ExcelFile(BytesIO(file_bytes)).sheet_names


def clean_column_names(df):
    df = df.copy()
    df.columns = [
        re.sub(INVISIBLE_CHARS, "", str(c)).strip() for c in df.columns
    ]
    return df


def parse_updat_dt(series):
    """UPDAT_DT -> datetime. Format utama d/m/yyyy h:mm (hari dulu)."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    # Sel yang sudah berupa tanggal Excel
    is_dt = series.map(lambda v: isinstance(v, (datetime, pd.Timestamp)))
    if is_dt.any():
        result.loc[is_dt] = pd.to_datetime(series[is_dt], errors="coerce")

    text = series.astype("string").str.strip()
    rest = (~is_dt) & text.notna()

    for fmt in UPDAT_DT_FORMATS:
        todo = rest & result.isna()
        if not todo.any():
            break
        result.loc[todo] = pd.to_datetime(
            text[todo], format=fmt, errors="coerce"
        )

    todo = rest & result.isna()
    if todo.any():
        result.loc[todo] = pd.to_datetime(
            text[todo], dayfirst=True, errors="coerce", format="mixed"
        )
    return result


def parse_compact_datetime(series):
    """SHEAR_CUT_DT format yyyymmddHHMMSS (mis. 20250108083200)."""
    if pd.api.types.is_numeric_dtype(series):
        text = (
            pd.to_numeric(series, errors="coerce")
            .round().astype("Int64").astype("string")
        )
    else:
        text = series.astype("string").str.strip().str.replace(
            r"\.0$", "", regex=True
        )
    return pd.to_datetime(text, format="%Y%m%d%H%M%S", errors="coerce")


def limit_or_nan(series):
    """LL/UL: 0 atau kosong berarti tidak ada batas."""
    s = pd.to_numeric(series, errors="coerce")
    return s.where(s != 0)


def mech_status(value, ll, ul):
    """Status satu sifat mekanik per baris."""
    status = pd.Series("Tanpa spesifikasi", index=value.index, dtype="object")
    has_spec = ll.notna() | ul.notna()
    has_value = value.notna() & ~((value == 0) & ~has_spec)
    check = has_spec & has_value
    status[check] = "Dalam spesifikasi"
    status[check & ll.notna() & (value < ll)] = "Di bawah LL"
    status[check & ul.notna() & (value > ul)] = "Di atas UL"
    return status


def clean_frame(raw_df):
    """Bersihkan data mentah (nama kolom, karakter tersembunyi, angka dalam teks)."""
    df = clean_column_names(raw_df)
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    df = df.dropna(how="all").reset_index(drop=True)

    for col in df.select_dtypes(include=["object", "string"]).columns:
        cleaned = (
            df[col].astype("string")
            .str.replace(INVISIBLE_CHARS, "", regex=True)
            .str.strip()
            .replace("", pd.NA)
        )
        # Jangan rusak sel tanggal Excel di kolom objek
        if not df[col].map(lambda v: isinstance(v, (datetime, pd.Timestamp))).any():
            df[col] = cleaned

    # Konversi otomatis kolom angka yang tersimpan sebagai teks
    for col in df.columns:
        if col in PROTECTED_TEXT:
            continue
        if pd.api.types.is_numeric_dtype(df[col]) or pd.api.types.is_datetime64_any_dtype(df[col]):
            continue
        converted = pd.to_numeric(
            df[col].astype("string").str.replace(",", ".", regex=False),
            errors="coerce",
        )
        non_empty = df[col].notna().sum()
        if non_empty and converted.notna().sum() / non_empty >= 0.85:
            df[col] = converted
    return df


DATE_ORDERS = ["Otomatis (disarankan)", "Hari dulu (d/m/yyyy)", "Bulan dulu (m/d/yyyy)"]


def swap_day_month(ts):
    """Tukar hari <-> bulan (hanya bila hari <= 12)."""
    ok = ts.notna() & (ts.dt.day <= 12)
    out = pd.Series(pd.NaT, index=ts.index, dtype="datetime64[ns]")
    if ok.any():
        t = ts[ok]
        out[ok] = pd.to_datetime(dict(year=t.dt.year, month=t.dt.day, day=t.dt.month,
                                      hour=t.dt.hour, minute=t.dt.minute, second=t.dt.second),
                                 errors="coerce")
    return out


def reference_time(df):
    """Waktu pembanding yang formatnya tidak ambigu: SHEAR_CUT_DT, atau WORK_DT (yyyymmdd)."""
    ref = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    if "SHEAR_CUT_DT" in df.columns:
        ref = parse_compact_datetime(df["SHEAR_CUT_DT"]).astype("datetime64[ns]")
    if "WORK_DT" in df.columns:
        work = pd.to_datetime(pd.to_numeric(df["WORK_DT"], errors="coerce").round().astype("Int64").astype("string"),
                              format="%Y%m%d", errors="coerce") + pd.Timedelta(hours=12)
        ref = ref.fillna(work)
    return ref


def resolve_updat_dt(df, order=DATE_ORDERS[0]):
    """Baca UPDAT_DT dengan aman.

    Kasus yang ditangani (umum pada file Excel dari sistem produksi):
    - Teks asli hari dulu (d/m/yyyy), tetapi Excel dengan setelan regional bulan dulu
      mengubah sel yang harinya <= 12 menjadi tanggal Excel dengan hari/bulan TERTUKAR
      (mis. teks 12/06/2026 22:03 = 12 Juni -> disimpan Excel sebagai 6 Desember).
      Sel yang harinya > 12 tetap teks (14/01/2025 13:04) karena tidak bisa diubah.
    - Teks yang ambigu (6/12/2026) dicocokkan dengan SHEAR_CUT_DT/WORK_DT bila ada.
    Hasil: (Series datetime, Series keterangan per baris).
    """
    raw = df["UPDAT_DT"].astype("string").str.strip()
    a = parse_updat_dt(df["UPDAT_DT"]).astype("datetime64[ns]")   # tafsiran apa adanya / hari dulu
    b = swap_day_month(a)                                         # tafsiran ditukar
    ambiguous = a.notna() & b.notna() & (a.dt.day != a.dt.month)
    note = pd.Series("", index=df.index, dtype="object")
    result = a.copy()

    # Jenis sel: tanggal Excel asli (tersimpan sebagai yyyy-mm-dd) atau teks d/m/yyyy
    excel_cell = raw.str.match(r"^\d{4}-\d{2}-\d{2}").fillna(False)
    parts = raw.str.extract(r"^(\d{1,2})[/-](\d{1,2})[/-]\d{4}")
    first, second = pd.to_numeric(parts[0], errors="coerce"), pd.to_numeric(parts[1], errors="coerce")
    text_cell = first.notna()
    if order.startswith("Hari"):
        text_order = "hari"
    elif order.startswith("Bulan"):
        text_order = "bulan"
    else:
        text_order = "bulan" if int((second > 12).sum()) > int((first > 12).sum()) else "hari"

    # 1) Sel teks: baca sesuai urutan teks
    if text_order == "bulan":
        m = ambiguous & text_cell
        result[m] = b[m]
        note[m] = "teks dibaca bulan dulu"

    # 2) Sel tanggal Excel di kolom campuran: kalau masih ada sel teks yang tidak bisa diubah Excel,
    #    berarti setelan regional Excel berbeda dengan urutan teks -> hari/bulan sel Excel tertukar.
    mixed = bool(excel_cell.any() and text_cell.any())
    if mixed:
        m = ambiguous & excel_cell
        result[m] = b[m]
        note[m] = "sel tanggal Excel tertukar hari/bulan (diperbaiki)"

    # 3) Mode otomatis: bila ada SHEAR_CUT_DT/WORK_DT, pilih tafsiran yang paling dekat
    if order.startswith("Otomatis"):
        ref = reference_time(df)
        has_ref = ambiguous & ref.notna()
        if has_ref.any():
            alt = a.where(result != a, b)
            better = has_ref & ((alt - ref).abs() < (result - ref).abs())
            result[better] = alt[better]
            note[better] = "hari/bulan ditukar (cocok dengan SHEAR_CUT_DT/WORK_DT)"

        # Pengaman terakhir: jangan biarkan tanggal di masa depan bila tafsiran lain masuk akal
        future = pd.Timestamp.now() + pd.Timedelta(days=1)
        alt = a.where(result != a, b)
        fix = ambiguous & (result > future) & (alt <= future)
        if fix.any():
            result[fix] = alt[fix]
            note[fix] = "hari/bulan ditukar (tanggal masa depan)"
    return result, note


def derive_frame(df, date_order=DATE_ORDERS[0]):
    """Tambahkan kolom turunan: datetime, thickness group, status mekanik."""
    df = df.copy()

    # Kolom waktu
    if "UPDAT_DT" in df.columns:
        df[UPDAT_DT_COL], df["_DATE_NOTE"] = resolve_updat_dt(df, date_order)
    if "SHEAR_CUT_DT" in df.columns:
        df[SHEAR_DT_COL] = parse_compact_datetime(df["SHEAR_CUT_DT"])

    # Thickness group bila belum ada
    if "Thickness group" not in df.columns and "INLET_COIL_TH" in df.columns:
        df["Thickness group"] = df["INLET_COIL_TH"].map(thickness_group)

    # Status sifat mekanik
    status_cols = []
    for prop, (vcol, llcol, ulcol, _) in MECH_PROPS.items():
        if vcol in df.columns:
            ll = limit_or_nan(df[llcol]) if llcol in df.columns else pd.Series(np.nan, index=df.index)
            ul = limit_or_nan(df[ulcol]) if ulcol in df.columns else pd.Series(np.nan, index=df.index)
            df[f"_LL_{prop}"] = ll
            df[f"_UL_{prop}"] = ul
            df[f"_ST_{prop}"] = mech_status(pd.to_numeric(df[vcol], errors="coerce"), ll, ul)
            status_cols.append(f"_ST_{prop}")

    if status_cols:
        st_df = df[status_cols]
        ng = st_df.isin(["Di atas UL", "Di bawah LL"]).any(axis=1)
        ok = st_df.eq("Dalam spesifikasi").any(axis=1)
        df["STATUS_MEKANIK"] = np.where(ng, "NG", np.where(ok, "OK", "Tanpa spesifikasi"))

    df[ROW_KEY] = np.arange(len(df))

    if UPDAT_DT_COL in df.columns and df[UPDAT_DT_COL].notna().any():
        time_col = UPDAT_DT_COL
    elif SHEAR_DT_COL in df.columns and df[SHEAR_DT_COL].notna().any():
        time_col = SHEAR_DT_COL
    else:
        time_col = None

    if time_col:
        df = df.sort_values(time_col, kind="stable").reset_index(drop=True)

    return df, time_col


def thickness_group(th):
    if pd.isna(th):
        return pd.NA
    if th < 0.5:
        return "0.1 ~ 0.499"
    if th < 0.8:
        return "0.5 ~ 0.799"
    return "0.8 ~ 1.199"


# =====================================================
# FUNGSI ANALISIS
# =====================================================
def capability(values, ll=None, ul=None):
    v = pd.to_numeric(values, errors="coerce").dropna()
    out = {"n": len(v), "Rata-rata": np.nan, "Std": np.nan,
           "Min": np.nan, "Max": np.nan, "Cp": np.nan, "Cpk": np.nan}
    if len(v) == 0:
        return out
    m = v.mean()
    s = v.std(ddof=1) if len(v) > 1 else np.nan
    out.update({"Rata-rata": m, "Std": s, "Min": v.min(), "Max": v.max()})
    if pd.isna(s) or s == 0:
        return out
    ll = None if ll is None or pd.isna(ll) else float(ll)
    ul = None if ul is None or pd.isna(ul) else float(ul)
    parts = []
    if ul is not None:
        parts.append((ul - m) / (3 * s))
    if ll is not None:
        parts.append((m - ll) / (3 * s))
    if ll is not None and ul is not None:
        out["Cp"] = (ul - ll) / (6 * s)
    if parts:
        out["Cpk"] = min(parts)
    return out


def cpk_label(cpk):
    if pd.isna(cpk):
        return "-"
    if cpk >= 1.33:
        return "Capable (≥ 1.33)"
    if cpk >= 1.0:
        return "Marginal (1.00–1.33)"
    return "Tidak capable (< 1.00)"


def process_status(values, lower, upper):
    status = pd.Series("Dalam batas", index=values.index, dtype="object")
    if upper is not None:
        status[values >= upper] = "≥ batas atas"
    if lower is not None:
        status[values <= lower] = "≤ batas bawah"
    status[values.isna()] = pd.NA
    return status


def single_value(series):
    """Kembalikan nilai tunggal jika batas konstan, selain itu None."""
    s = series.dropna().unique()
    return float(s[0]) if len(s) == 1 else None


def fmt_num(v, digits=2):
    if v is None or pd.isna(v):
        return "-"
    return f"{v:,.{digits}f}"


def fmt_dt(v):
    if v is None or pd.isna(v):
        return "-"
    return pd.Timestamp(v).strftime("%d-%m-%Y %H:%M")


def to_excel_bytes(sheets):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in sheets:
            frame.to_excel(writer, index=False, sheet_name=name[:31])
    return buffer.getvalue()


def export_view(frame):
    """Data untuk ditampilkan/diekspor (tanpa kolom internal)."""
    return frame[[c for c in frame.columns if not c.startswith("_")]]


# =====================================================
# GRAFIK: TOOLTIP, HAIRLINE, KLIK-SAMPLING
# =====================================================
def make_customdata(frame, hover_cols):
    cols = [ROW_KEY, ID_COL] + hover_cols
    parts = []
    for c in cols:
        if c not in frame.columns:
            parts.append(pd.Series("-", index=frame.index))
            continue
        s = frame[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            parts.append(s.dt.strftime("%d-%m-%Y %H:%M").fillna("-"))
        elif pd.api.types.is_float_dtype(s):
            parts.append(s.map(lambda v: "-" if pd.isna(v) else f"{v:g}"))
        else:
            parts.append(s.astype(object).where(s.notna(), "-").astype(str))
    return np.column_stack([p.to_numpy() for p in parts]) if parts else None


def make_hovertemplate(x_label, y_label, hover_cols, x_is_time):
    x_fmt = "%{x|%d-%m-%Y %H:%M}" if x_is_time else "%{x}"
    lines = [
        f"<b>{ID_COL}: %{{customdata[1]}}</b>",
        f"{x_label}: {x_fmt}",
        f"{y_label}: <b>%{{y}}</b>",
    ]
    for i, c in enumerate(hover_cols, start=2):
        lines.append(f"{c}: %{{customdata[{i}]}}")
    return "<br>".join(lines) + "<extra>%{fullData.name}</extra>"


def apply_hairline(fig, mode, x_is_time):
    """mode: 'crosshair' (X+Y) atau 'vertical' (garis vertikal, tooltip gabungan)."""
    spike = dict(
        showspikes=True,
        spikemode="across",
        spikesnap="data",
        spikethickness=1,
        spikedash="dot",
        spikecolor="rgba(90,90,90,0.9)",
    )
    fig.update_xaxes(**spike)
    if mode == "crosshair":
        fig.update_yaxes(**spike)
        fig.update_layout(hovermode="closest")
    else:
        fig.update_yaxes(showspikes=False)
        fig.update_layout(hovermode="x unified")
    fig.update_layout(spikedistance=-1, hoverdistance=50)

    if x_is_time:
        fig.update_xaxes(
            hoverformat="%d-%m-%Y %H:%M",
            tickformatstops=[
                dict(dtickrange=[None, 3_600_000], value="%H:%M<br>%d-%b-%Y"),
                dict(dtickrange=[3_600_000, 86_400_000], value="%H:%M<br>%d-%b-%Y"),
                dict(dtickrange=[86_400_000, None], value="%d-%b-%Y"),
            ],
        )


def build_trend_figure(plot_df, x_col, y_col, status, status_style,
                       limit_lines, band, hover_cols, hairline_mode,
                       title, x_label, y_label, x_is_time):
    fig = go.Figure()
    hovertemplate = make_hovertemplate(x_label, y_label, hover_cols, x_is_time)
    Trace = scatter_cls(len(plot_df))
    dense = len(plot_df) > GL_THRESHOLD

    # Garis tren (penghubung antar coil)
    fig.add_trace(Trace(
        x=plot_df[x_col], y=plot_df[y_col], mode="lines",
        line=dict(color="rgba(120,130,140,0.45)", width=1 if dense else 1.5),
        name="Tren", hoverinfo="skip", showlegend=False,
    ))

    # Titik per status (bisa diklik)
    for label, style in status_style.items():
        mask = status == label
        if not mask.any():
            continue
        sub = plot_df[mask]
        normal = style["symbol"] in ("circle", "circle-open")
        fig.add_trace(Trace(
            x=sub[x_col], y=sub[y_col], mode="markers", name=label,
            marker=dict(color=style["color"], symbol=style["symbol"],
                        size=(4 if normal else 8) if dense else 9,
                        line=dict(width=0 if dense else 1, color="white")),
            customdata=make_customdata(sub, hover_cols),
            hovertemplate=hovertemplate,
        ))

    # Garis batas
    y_values = [plot_df[y_col].min(), plot_df[y_col].max()]
    for lim in limit_lines:
        value = lim["value"]
        if isinstance(value, pd.Series):
            fig.add_trace(Trace(
                x=plot_df[x_col], y=value, mode="lines", name=lim["name"],
                line=dict(color=lim["color"], dash=lim.get("dash", "dash"), width=1.5, shape="hv"),
                hoverinfo="skip",
            ))
            y_values += [value.min(), value.max()]
        elif value is not None:
            fig.add_hline(
                y=value, line_color=lim["color"], line_dash=lim.get("dash", "dash"),
                line_width=1.5,
                annotation_text=f"{lim['name']}: {value:g}",
                annotation_position=lim.get("pos", "top left"),
                annotation_font_color=lim["color"],
            )
            y_values.append(value)

    if band and band[0] is not None and band[1] is not None:
        fig.add_hrect(y0=band[0], y1=band[1], fillcolor=COLOR_BAND, line_width=0, layer="below")

    y_values = [v for v in y_values if v is not None and pd.notna(v)]
    if y_values:
        lo, hi = min(y_values), max(y_values)
        pad = (hi - lo) * 0.08 or 1
        fig.update_yaxes(range=[lo - pad, hi + pad])

    fig.update_layout(
        title=title, height=540, xaxis_title=x_label, yaxis_title=y_label,
        dragmode="zoom", clickmode="event+select",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=70, b=40, l=60, r=20),
        hoverlabel=dict(namelength=-1),
    )
    apply_hairline(fig, hairline_mode, x_is_time)
    if x_is_time:
        add_range_buttons(fig)
    return fig


def add_range_buttons(fig, axis="xaxis"):
    """Tombol zoom cepat di atas grafik waktu."""
    fig.update_layout({axis: dict(rangeselector=dict(
        buttons=[
            dict(count=1, label="1 hari", step="day", stepmode="backward"),
            dict(count=7, label="7 hari", step="day", stepmode="backward"),
            dict(count=1, label="1 bulan", step="month", stepmode="backward"),
            dict(count=3, label="3 bulan", step="month", stepmode="backward"),
            dict(step="all", label="Semua"),
        ],
        x=0, y=1.0, yanchor="bottom", font=dict(size=11),
    ))})


def profile_figure(frame, stages, selected_keys, title, y_title, stage_limits=None):
    stages = [(c, lbl) for c, lbl in stages if c in frame.columns]
    if not stages:
        return None
    cols = [c for c, _ in stages]
    labels = [lbl for _, lbl in stages]
    values = frame[cols].apply(pd.to_numeric, errors="coerce")

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=labels, y=values.max().values, mode="lines",
        line=dict(width=0), showlegend=False, hoverinfo="skip",
    ))
    fig.add_trace(go.Scatter(
        x=labels, y=values.min().values, mode="lines", line=dict(width=0),
        fill="tonexty", fillcolor="rgba(58,110,165,0.15)",
        name="Rentang min–max", hovertemplate="Min: %{y}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=labels, y=values.mean().round(1).values, mode="lines+markers",
        line=dict(color="#5B6770", dash="dash"), name="Rata-rata",
        hovertemplate="Rata-rata: %{y}<extra></extra>",
    ))

    palette = px.colors.qualitative.Bold
    sel = frame[frame[ROW_KEY].isin(selected_keys)]
    for i, (_, row) in enumerate(sel.iterrows()):
        fig.add_trace(go.Scatter(
            x=labels, y=[row[c] for c in cols], mode="lines+markers",
            name=str(row.get(ID_COL, row[ROW_KEY])),
            line=dict(color=palette[i % len(palette)], width=2.5),
            hovertemplate=f"{row.get(ID_COL, '')}: %{{y}}<extra></extra>",
        ))

    for col, (lower, upper) in (stage_limits or {}).items():
        if col not in cols:
            continue
        lbl = labels[cols.index(col)]
        ys = [v for v in (lower, upper) if v is not None]
        if ys:
            fig.add_trace(go.Scatter(
                x=[lbl] * len(ys), y=ys, mode="lines+markers",
                marker=dict(symbol="line-ew", size=18, line=dict(width=3, color=COLOR_LIMIT)),
                line=dict(color=COLOR_LIMIT, width=1, dash="dot"),
                name=f"Batas {lbl}",
                hovertemplate=f"Batas {lbl}: %{{y}}<extra></extra>",
            ))

    fig.update_layout(
        title=title, height=460, yaxis_title=y_title, hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=70, b=40, l=60, r=20),
    )
    fig.update_xaxes(showspikes=True, spikemode="across", spikethickness=1,
                     spikedash="dot", spikecolor="rgba(90,90,90,0.9)")
    return fig


# =====================================================
# DATABASE SQLITE
# =====================================================
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    imported_at TEXT, file_name TEXT, sheet TEXT,
    rows_in_file INTEGER, n_new INTEGER, n_changed INTEGER, n_same INTEGER,
    n_skipped INTEGER, new_columns TEXT, note TEXT
);
CREATE TABLE IF NOT EXISTS coil_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    import_id INTEGER, coil TEXT, kolom TEXT, nilai_lama TEXT, nilai_baru TEXT
);
CREATE INDEX IF NOT EXISTS idx_history_coil ON coil_history(coil);
CREATE TABLE IF NOT EXISTS cursors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    time TEXT, coil TEXT, param TEXT, value REAL,
    note TEXT DEFAULT '', created_at TEXT
);
CREATE TABLE IF NOT EXISTS events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tanggal TEXT, kejadian TEXT, panel TEXT
);
CREATE TABLE IF NOT EXISTS panels (
    pos INTEGER, parameter TEXT, judul TEXT,
    target REAL, toleransi REAL, y_min REAL, y_max REAL
);
CREATE TABLE IF NOT EXISTS chart_panels (
    pos INTEGER, parameter TEXT, judul TEXT, target REAL, y_min REAL, y_max REAL
);
CREATE TABLE IF NOT EXISTS process_limits (
    parameter TEXT PRIMARY KEY, lower REAL, upper REAL
);
"""

def convert_legacy_panel(rec):
    """Panel versi lama (Target ± Toleransi) -> format baru (Batas bawah/atas)."""
    rec = dict(rec)
    tol, target = rec.pop("Toleransi ±", None), rec.get("Target")
    if tol is not None and target is not None and rec.get("Batas bawah") is None and rec.get("Batas atas") is None:
        rec["Batas bawah"], rec["Batas atas"] = target - tol, target + tol
    return {c: rec.get(c) for c in PANEL_COLUMNS}


def db_connect():
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.execute("PRAGMA journal_mode=WAL")
    return con


def qid(name):
    """Quote nama kolom SQLite (aman untuk spasi, tanda kurung, #, dll)."""
    return '"' + str(name).replace('"', '""') + '"'


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def py_value(v):
    """Nilai pandas/numpy -> tipe Python yang diterima sqlite3."""
    if v is None or v is pd.NA or v is pd.NaT:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if isinstance(v, np.integer):
        return int(v)
    if isinstance(v, np.floating):
        return None if np.isnan(v) else float(v)
    if isinstance(v, (np.bool_, bool)):
        return int(v)
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v).strftime("%Y-%m-%d %H:%M:%S")
    return v


def get_setting(key, default=None):
    with closing(db_connect()) as con:
        row = con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    if row is None:
        return default
    try:
        return json.loads(row[0])
    except (TypeError, ValueError):
        return default


def set_setting(key, value):
    with closing(db_connect()) as con, con:
        con.execute("INSERT INTO settings(key, value) VALUES(?, ?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (key, json.dumps(value, default=str)))


def coil_table_exists(con=None):
    own = con is None
    con = con or db_connect()
    try:
        return con.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='coil_data'").fetchone() is not None
    finally:
        if own:
            con.close()


def coil_columns(con):
    return [r[1] for r in con.execute("PRAGMA table_info(coil_data)").fetchall()]


@st.cache_resource(show_spinner=False)
def init_db():
    """Buat tabel, isi default, dan migrasi file JSON versi lama (sekali saja)."""
    with closing(db_connect()) as con, con:
        con.executescript(SCHEMA_SQL)

    if get_setting("limits_initialized") is None:
        for param, lim in DEFAULT_PROCESS_LIMITS.items():
            save_limit(param, lim.get("lower"), lim.get("upper"))
        set_setting("limits_initialized", True)
    if get_setting("panels_v2_initialized") is None:
        with closing(db_connect()) as con:
            old = con.execute("SELECT parameter, judul, target, toleransi, y_min, y_max "
                              "FROM panels ORDER BY pos").fetchall()
        if old:
            records = [convert_legacy_panel({"Parameter": r[0], "Judul": r[1], "Target": r[2],
                                             "Toleransi ±": r[3], "Y min": r[4], "Y max": r[5]})
                       for r in old]
            limits = load_limits()
            for r in records:   # batas yang sudah ada di database tetap dipakai
                if r["Parameter"] in limits:
                    r["Batas bawah"] = limits[r["Parameter"]]["lower"]
                    r["Batas atas"] = limits[r["Parameter"]]["upper"]
            save_panels(records)
        else:
            save_panels(DEFAULT_PANELS)
        set_setting("panels_v2_initialized", True)

    if get_setting("legacy_json_migrated") is None and LEGACY_STATE_FILE.exists():
        try:
            legacy = json.loads(LEGACY_STATE_FILE.read_text(encoding="utf-8"))
            if legacy.get("events"):
                save_events(legacy["events"])
            if legacy.get("panels"):
                save_panels([convert_legacy_panel(r) for r in legacy["panels"]])
            with closing(db_connect()) as con, con:
                for items in legacy.get("cursors", {}).values():
                    for c in items:
                        con.execute("INSERT INTO cursors(time, coil, param, value, note, created_at) "
                                    "VALUES(?,?,?,?,?,?)",
                                    (c.get("time"), c.get("coil"), c.get("param"),
                                     c.get("value"), c.get("note", ""), now_text()))
        except Exception:
            pass
        set_setting("legacy_json_migrated", True)
    return True


def data_version():
    return int(get_setting("data_version", 0) or 0)


def bump_data_version(con):
    current = con.execute("SELECT value FROM settings WHERE key='data_version'").fetchone()
    new = (int(json.loads(current[0])) if current else 0) + 1
    con.execute("INSERT INTO settings(key, value) VALUES('data_version', ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (json.dumps(new),))


@st.cache_resource(show_spinner="Memuat data dari database...", max_entries=2)
def load_dataset(version, date_order=DATE_ORDERS[0]):
    """Baca coil_data dari SQLite lalu olah. Di-cache per versi data (berubah setiap import)."""
    with closing(db_connect()) as con:
        if not coil_table_exists(con):
            return None, None
        raw = pd.read_sql_query("SELECT * FROM coil_data", con)
    if raw.empty:
        return raw, None
    df = clean_frame(raw)
    return derive_frame(df, date_order)


def db_summary():
    with closing(db_connect()) as con:
        n = con.execute("SELECT COUNT(*) FROM coil_data").fetchone()[0] if coil_table_exists(con) else 0
        last = con.execute("SELECT imported_at, file_name FROM import_log ORDER BY id DESC LIMIT 1").fetchone()
    return n, last


# ---------- cursor ----------
def load_cursors():
    with closing(db_connect()) as con:
        rows = con.execute("SELECT id, time, coil, param, value, note FROM cursors ORDER BY id").fetchall()
    return [{"id": r[0], "time": r[1], "coil": r[2], "param": r[3], "value": r[4], "note": r[5] or ""}
            for r in rows]


def db_add_cursor(time, coil, param, value):
    t = None if time is None or pd.isna(time) else pd.Timestamp(time).isoformat()
    v = None if value is None or pd.isna(value) else float(value)
    with closing(db_connect()) as con, con:
        exists = con.execute("SELECT 1 FROM cursors WHERE time IS ? AND coil IS ? AND param IS ?",
                             (t, coil, param)).fetchone()
        if exists:
            return False
        con.execute("INSERT INTO cursors(time, coil, param, value, note, created_at) VALUES(?,?,?,?,?,?)",
                    (t, coil, param, v, "", now_text()))
    return True


def db_delete_cursors(ids):
    with closing(db_connect()) as con, con:
        con.executemany("DELETE FROM cursors WHERE id=?", [(int(i),) for i in ids])


def db_update_cursor_note(cursor_id, note):
    with closing(db_connect()) as con, con:
        con.execute("UPDATE cursors SET note=? WHERE id=?", (note, int(cursor_id)))


# ---------- event, panel, batas proses ----------
def load_events():
    with closing(db_connect()) as con:
        rows = con.execute("SELECT tanggal, kejadian, panel FROM events ORDER BY tanggal, id").fetchall()
    return [{"Tanggal": r[0], "Kejadian": r[1], "Panel": r[2] or "Semua"} for r in rows]


def save_events(records):
    with closing(db_connect()) as con, con:
        con.execute("DELETE FROM events")
        con.executemany("INSERT INTO events(tanggal, kejadian, panel) VALUES(?,?,?)",
                        [(r.get("Tanggal"), r.get("Kejadian"), r.get("Panel") or "Semua")
                         for r in records if r.get("Tanggal")])


def load_panels():
    """Panel grafik + batas proses (batas disimpan terpisah supaya dipakai di semua halaman)."""
    with closing(db_connect()) as con:
        rows = con.execute("SELECT parameter, judul, target, y_min, y_max FROM chart_panels ORDER BY pos").fetchall()
    limits = load_limits()
    out = []
    for r in rows:
        lim = limits.get(r[0], {})
        out.append({"Parameter": r[0], "Judul": r[1], "Target": r[2],
                    "Batas bawah": lim.get("lower"), "Batas atas": lim.get("upper"),
                    "Y min": r[3], "Y max": r[4]})
    return out


def save_panels(records):
    records = [r for r in records if r.get("Parameter")]
    with closing(db_connect()) as con, con:
        con.execute("DELETE FROM chart_panels")
        con.executemany(
            "INSERT INTO chart_panels(pos, parameter, judul, target, y_min, y_max) VALUES(?,?,?,?,?,?)",
            [(i, r["Parameter"], r.get("Judul"), py_value(r.get("Target")),
              py_value(r.get("Y min")), py_value(r.get("Y max"))) for i, r in enumerate(records)])
    for r in records:
        save_limit(r["Parameter"], py_value(r.get("Batas bawah")), py_value(r.get("Batas atas")))


def load_limits():
    with closing(db_connect()) as con:
        rows = con.execute("SELECT parameter, lower, upper FROM process_limits").fetchall()
    return {r[0]: {"lower": r[1], "upper": r[2]} for r in rows}


def save_limit(param, lower, upper):
    with closing(db_connect()) as con, con:
        if lower is None and upper is None:
            con.execute("DELETE FROM process_limits WHERE parameter=?", (param,))
        else:
            con.execute("INSERT INTO process_limits(parameter, lower, upper) VALUES(?,?,?) "
                        "ON CONFLICT(parameter) DO UPDATE SET lower=excluded.lower, upper=excluded.upper",
                        (param, lower, upper))


def clean_records(frame, columns):
    """DataFrame -> list of dict (NaN -> None, tanggal -> 'YYYY-MM-DD')."""
    out = []
    for rec in frame.to_dict("records"):
        item = {}
        for c in columns:
            v = rec.get(c)
            if v is None or v is pd.NaT or v is pd.NA or (isinstance(v, float) and np.isnan(v)):
                v = None
            elif isinstance(v, (pd.Timestamp, datetime)):
                v = pd.Timestamp(v).strftime("%Y-%m-%d")
            elif hasattr(v, "isoformat"):
                v = v.isoformat()
            elif isinstance(v, np.integer):
                v = int(v)
            elif isinstance(v, np.floating):
                v = float(v)
            item[c] = v
        out.append(item)
    return out


# =====================================================
# IMPORT: BANDINGKAN FILE DENGAN DATABASE
# =====================================================
def to_storage_frame(frame):
    """Samakan format sebelum dibandingkan/disimpan (tanggal -> teks, kunci -> teks)."""
    out = frame[[c for c in frame.columns if not str(c).startswith("_")]].copy()
    for c in out.columns:
        s = out[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            out[c] = s.dt.strftime("%Y-%m-%d %H:%M:%S").astype(object).where(s.notna(), None)
        elif s.dtype == object and s.map(lambda v: isinstance(v, (datetime, pd.Timestamp))).any():
            out[c] = s.map(py_value)
        elif pd.api.types.is_bool_dtype(s):
            out[c] = s.astype("Int64")
    out[KEY_COL] = out[KEY_COL].astype("string").str.strip()
    return out


def is_blank(s):
    return s.isna() | s.astype("string").str.strip().fillna("").eq("")


def values_equal(a, b):
    """Bandingkan dua kolom: angka dengan toleransi kecil, teks setelah di-trim."""
    an = pd.to_numeric(a, errors="coerce")
    bn = pd.to_numeric(b, errors="coerce")
    eq = pd.Series(False, index=a.index)
    both_num = an.notna() & bn.notna()
    if both_num.any():
        eq[both_num] = np.isclose(an[both_num].astype(float), bn[both_num].astype(float),
                                  rtol=1e-9, atol=1e-9)
    both_blank = is_blank(a) & is_blank(b)
    eq |= both_blank
    rest = ~both_num & ~both_blank
    if rest.any():
        eq[rest] = (a[rest].astype("string").str.strip().fillna("")
                    == b[rest].astype("string").str.strip().fillna("")).fillna(False)
    return eq


def show_value(v):
    v = py_value(v)
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


@st.cache_data(show_spinner="Membandingkan file dengan database...", max_entries=3)
def compare_with_db(file_bytes, file_name, sheet_name, header_row, version):
    raw = read_table(file_bytes, file_name, sheet_name, header_row)
    inc = clean_frame(raw)
    if KEY_COL not in inc.columns:
        return {"error": f"Kolom kunci {KEY_COL} tidak ditemukan di file/sheet ini."}
    inc = to_storage_frame(inc)

    result = {"error": None, "rows_in_file": len(inc)}
    no_key = inc[KEY_COL].isna() | inc[KEY_COL].eq("")
    result["no_key"] = int(no_key.sum())
    inc = inc[~no_key]

    dup_mask = inc[KEY_COL].duplicated(keep=False)
    result["duplicates"] = inc.loc[dup_mask, [KEY_COL]].value_counts().reset_index(name="Jumlah baris")
    inc = inc.drop_duplicates(subset=KEY_COL, keep="last").reset_index(drop=True)
    result["incoming"] = inc

    with closing(db_connect()) as con:
        exists = coil_table_exists(con)
        result["table_exists"] = exists
        if not exists:
            result.update(new_keys=inc[KEY_COL].tolist(), changed_keys=[], same_count=0,
                          diffs=pd.DataFrame(columns=[KEY_COL, "Kolom", "Nilai database", "Nilai file"]),
                          db_only_count=0, db_only_keys=[], new_cols=[], missing_cols=[],
                          db_rows=0)
            return result

        db_cols = coil_columns(con)
        data_cols = [c for c in db_cols if not c.startswith("_")]
        common = [c for c in inc.columns if c in db_cols]
        result["new_cols"] = [c for c in inc.columns if c not in db_cols]
        result["missing_cols"] = [c for c in data_cols if c not in inc.columns]
        result["db_rows"] = con.execute("SELECT COUNT(*) FROM coil_data").fetchone()[0]

        con.execute("CREATE TEMP TABLE incoming_keys(k TEXT PRIMARY KEY)")
        con.executemany("INSERT INTO incoming_keys(k) VALUES(?)", [(k,) for k in inc[KEY_COL]])
        cols_sql = ", ".join(f"c.{qid(c)}" for c in common)
        db_part = pd.read_sql_query(
            f"SELECT {cols_sql} FROM coil_data c JOIN incoming_keys k ON c.{qid(KEY_COL)} = k.k", con)
        result["db_only_count"] = con.execute(
            f"SELECT COUNT(*) FROM coil_data WHERE {qid(KEY_COL)} NOT IN (SELECT k FROM incoming_keys)"
        ).fetchone()[0]
        result["db_only_keys"] = [r[0] for r in con.execute(
            f"SELECT {qid(KEY_COL)} FROM coil_data WHERE {qid(KEY_COL)} NOT IN "
            f"(SELECT k FROM incoming_keys) LIMIT 1000").fetchall()]

    db_part[KEY_COL] = db_part[KEY_COL].astype("string")
    matched_keys = set(db_part[KEY_COL])
    result["new_keys"] = [k for k in inc[KEY_COL] if k not in matched_keys]

    a = inc[inc[KEY_COL].isin(matched_keys)].set_index(KEY_COL)
    b = db_part.set_index(KEY_COL).reindex(a.index)
    diff_frames = []
    for c in common:
        if c == KEY_COL:
            continue
        eq = values_equal(a[c], b[c])
        if (~eq).any():
            keys = a.index[~eq.values]
            diff_frames.append(pd.DataFrame({
                KEY_COL: keys, "Kolom": c,
                "Nilai database": [show_value(v) for v in b.loc[keys, c]],
                "Nilai file": [show_value(v) for v in a.loc[keys, c]],
            }))
    diffs = (pd.concat(diff_frames, ignore_index=True) if diff_frames
             else pd.DataFrame(columns=[KEY_COL, "Kolom", "Nilai database", "Nilai file"]))
    result["diffs"] = diffs
    result["changed_keys"] = diffs[KEY_COL].drop_duplicates().tolist()
    result["same_count"] = len(a) - len(result["changed_keys"])
    return result


def sql_type(series):
    if pd.api.types.is_integer_dtype(series) or pd.api.types.is_bool_dtype(series):
        return "INTEGER"
    if pd.api.types.is_float_dtype(series):
        return "REAL"
    return "TEXT"


def apply_import(result, file_name, sheet_name, do_new, do_changed, add_columns):
    inc = result["incoming"]
    stamp = now_text()
    with closing(db_connect()) as con, con:
        if not coil_table_exists(con):
            col_defs = [f"{qid(KEY_COL)} TEXT PRIMARY KEY"]
            col_defs += [f"{qid(c)} {sql_type(inc[c])}" for c in inc.columns if c != KEY_COL]
            col_defs += ['"_source_file" TEXT', '"_import_id" INTEGER',
                         '"_imported_at" TEXT', '"_updated_at" TEXT']
            con.execute(f"CREATE TABLE coil_data ({', '.join(col_defs)})")
            added_cols = []
        else:
            added_cols = []
            if add_columns:
                for c in result["new_cols"]:
                    con.execute(f"ALTER TABLE coil_data ADD COLUMN {qid(c)} {sql_type(inc[c])}")
                    added_cols.append(c)

        cols = [c for c in inc.columns if c in coil_columns(con)]
        cur = con.execute(
            "INSERT INTO import_log(imported_at, file_name, sheet, rows_in_file, n_new, n_changed, "
            "n_same, n_skipped, new_columns, note) VALUES(?,?,?,?,?,?,?,?,?,?)",
            (stamp, file_name, sheet_name, result["rows_in_file"], 0, 0, result["same_count"],
             result["no_key"], ", ".join(added_cols), ""))
        import_id = cur.lastrowid

        n_new = n_changed = 0
        if do_new and result["new_keys"]:
            rows = inc[inc[KEY_COL].isin(set(result["new_keys"]))]
            all_cols = cols + ["_source_file", "_import_id", "_imported_at", "_updated_at"]
            sql = (f"INSERT INTO coil_data ({', '.join(qid(c) for c in all_cols)}) "
                   f"VALUES ({', '.join('?' for _ in all_cols)})")
            values = [[py_value(v) for v in r] + [file_name, import_id, stamp, stamp]
                      for r in rows[cols].itertuples(index=False, name=None)]
            con.executemany(sql, values)
            n_new = len(values)

        if do_changed and result["changed_keys"]:
            rows = inc[inc[KEY_COL].isin(set(result["changed_keys"]))]
            set_cols = [c for c in cols if c != KEY_COL]
            sql = (f"UPDATE coil_data SET {', '.join(f'{qid(c)}=?' for c in set_cols)}, "
                   f"\"_source_file\"=?, \"_import_id\"=?, \"_updated_at\"=? WHERE {qid(KEY_COL)}=?")
            values = [[py_value(v) for v in r[1:]] + [file_name, import_id, stamp, r[0]]
                      for r in rows[[KEY_COL] + set_cols].itertuples(index=False, name=None)]
            con.executemany(sql, values)
            n_changed = len(values)
            hist = result["diffs"]
            con.executemany(
                "INSERT INTO coil_history(import_id, coil, kolom, nilai_lama, nilai_baru) VALUES(?,?,?,?,?)",
                [(import_id, r[0], r[1], r[2], r[3]) for r in hist.itertuples(index=False, name=None)])

        # Kolom baru: isi juga nilainya untuk coil lama yang ada di file (tanpa mengubah kolom lain)
        if added_cols:
            matched = inc[~inc[KEY_COL].isin(set(result["new_keys"]))]
            sql = (f"UPDATE coil_data SET {', '.join(f'{qid(c)}=?' for c in added_cols)} "
                   f"WHERE {qid(KEY_COL)}=?")
            con.executemany(sql, [[py_value(v) for v in r[1:]] + [r[0]]
                                  for r in matched[[KEY_COL] + added_cols].itertuples(index=False, name=None)])

        con.execute("UPDATE import_log SET n_new=?, n_changed=? WHERE id=?", (n_new, n_changed, import_id))
        bump_data_version(con)
    return {"import_id": import_id, "n_new": n_new, "n_changed": n_changed, "added_cols": added_cols}


def delete_all_coil_data(include_cursors):
    with closing(db_connect()) as con, con:
        con.execute("DROP TABLE IF EXISTS coil_data")
        con.execute("DELETE FROM coil_history")
        if include_cursors:
            con.execute("DELETE FROM cursors")
        con.execute("INSERT INTO import_log(imported_at, file_name, note) VALUES(?,?,?)",
                    (now_text(), "-", "Semua data coil dihapus"))
        bump_data_version(con)


def database_backup_bytes():
    with closing(db_connect()) as con:
        con.execute("PRAGMA wal_checkpoint(FULL)")
    return DB_PATH.read_bytes()


def lazy_download(where, label, builder, file_name, mime, key):
    """Download yang filenya baru dibuat saat tombol diklik (lebih ringan).
    Untuk Streamlit lama yang belum mendukung data callable, file dibuat langsung."""
    try:
        where.download_button(label, data=builder, file_name=file_name, mime=mime,
                              key=key, on_click="ignore")
    except Exception:
        where.download_button(label, data=builder(), file_name=file_name, mime=mime, key=f"{key}_eager")


# =====================================================
# CURSOR PERMANEN (hasil klik di grafik)
# =====================================================
def init_session():
    ss = st.session_state
    ss.setdefault("chart_version", 0)
    ss["cursors"] = load_cursors()          # ringan: dibaca setiap rerun supaya selalu sinkron
    if "process_limits" not in ss:
        ss["process_limits"] = load_limits()


def reset_chart_selection():
    ss = st.session_state
    ss["chart_version"] = ss.get("chart_version", 0) + 1
    for k in [k for k in ss.keys() if str(k).startswith("_lastsel_")]:
        del ss[k]


def cursor_time(c):
    return pd.Timestamp(c["time"]) if c.get("time") else pd.NaT


def delete_cursors(ids):
    db_delete_cursors(ids)
    st.session_state["cursors"] = load_cursors()
    reset_chart_selection()


def points_from_event(event):
    try:
        points = event["selection"]["points"]
    except (KeyError, TypeError, AttributeError):
        return []
    out = []
    for p in points or []:
        if not isinstance(p, dict):
            continue
        cd = p.get("customdata")
        row = -1
        if cd is not None and len(cd):
            try:
                row = int(float(cd[0]))
            except (TypeError, ValueError):
                row = -1
        out.append({"row": row, "x": p.get("x"), "y": p.get("y"), "curve": p.get("curve_number")})
    return out


def handle_click(chart_id, event, df_all, time_col, curve_params):
    """Klik titik -> simpan cursor ke database. curve_params: nama parameter atau {curve_number: parameter}."""
    ss = st.session_state
    points = points_from_event(event)
    signature = tuple(sorted((p["row"], str(p["x"]), p["curve"]) for p in points))
    key = f"_lastsel_{chart_id}"
    if ss.get(key) == signature:
        return
    ss[key] = signature
    lookup = df_all.set_index(ROW_KEY)
    added = False
    for p in points:
        param = curve_params if isinstance(curve_params, str) else curve_params.get(p["curve"])
        if param is None:
            continue
        if p["row"] >= 0 and p["row"] in lookup.index:
            r = lookup.loc[p["row"]]
            t = r[time_col] if time_col else None
            coil = str(r[ID_COL])
        else:
            t = pd.to_datetime(p["x"], errors="coerce")
            coil = None
        added |= db_add_cursor(t, coil, param, p["y"])
    if added:
        ss["cursors"] = load_cursors()
        st.rerun()   # gambar ulang supaya garis cursor langsung muncul


def interactive_chart(fig, chart_id, df_all, time_col, curve_params):
    key = f"{chart_id}_{st.session_state['chart_version']}"
    event = st.plotly_chart(
        fig, width="stretch", key=key,
        on_select="rerun", selection_mode="points",
        config={"displaylogo": False, "scrollZoom": True},
    )
    handle_click(chart_id, event, df_all, time_col, curve_params)


def cursor_rows(df_all, time_col):
    """Pasangkan setiap cursor dengan baris coil (berdasarkan coil, atau waktu terdekat)."""
    by_coil = df_all.drop_duplicates(ID_COL).set_index(df_all.drop_duplicates(ID_COL)[ID_COL].astype(str))
    result = []
    for c in st.session_state["cursors"]:
        row = None
        if c.get("coil") is not None and c["coil"] in by_coil.index:
            row = by_coil.loc[c["coil"]]
        if row is None and time_col and c.get("time"):
            diff = (df_all[time_col] - cursor_time(c)).abs()
            if diff.notna().any():
                row = df_all.loc[diff.idxmin()]
        result.append((c, row))
    return result


def format_timedelta(td):
    if td is None or pd.isna(td):
        return ""
    minutes = int(round(abs(td.total_seconds()) / 60))
    days, rem = divmod(minutes, 1440)
    hours, mins = divmod(rem, 60)
    parts = []
    if days:
        parts.append(f"{days} hr")
    if hours:
        parts.append(f"{hours} j")
    parts.append(f"{mins} mnt")
    return ("-" if td.total_seconds() < 0 else "") + " ".join(parts)


def render_cursor_panel(df_all, time_col, panel_id, value_params):
    ss = st.session_state
    cursors = ss["cursors"]
    st.markdown(f"##### 📍 Cursor tersimpan ({len(cursors)})")
    if not cursors:
        st.caption("Klik titik di grafik untuk menambahkan cursor. Cursor tersimpan di database "
                   "sampai dihapus dari daftar ini.")
        return

    value_params = [p for p in dict.fromkeys(value_params) if p and p in df_all.columns]
    rows = []
    for c, r in cursor_rows(df_all, time_col):
        item = {
            "Hapus": False, "Cursor": f"C{c['id']}", "_id": c["id"],
            "Waktu": cursor_time(c), "Coil": c.get("coil") or "(agregat)",
            "Parameter diklik": label(c.get("param")), "Nilai diklik": c.get("value"),
        }
        for p in value_params:
            item[label(p)] = r[p] if r is not None else None
        item["Catatan"] = c.get("note", "")
        rows.append(item)
    table = pd.DataFrame(rows).sort_values("Waktu", kind="stable").reset_index(drop=True)
    table["Δ waktu vs sebelumnya"] = [format_timedelta(v) for v in table["Waktu"].diff()]
    for p in value_params:
        table[f"Δ {label(p)}"] = pd.to_numeric(table[label(p)], errors="coerce").diff().round(3)

    edited = st.data_editor(
        table.drop(columns="_id"),
        key=f"cursor_editor_{panel_id}_{hash(tuple(c['id'] for c in cursors))}",
        hide_index=True, width="stretch",
        disabled=[c for c in table.columns if c not in ("Hapus", "Catatan")],
        column_config={
            "Hapus": st.column_config.CheckboxColumn("Hapus", width="small"),
            "Waktu": st.column_config.DatetimeColumn("Waktu", format="DD-MM-YYYY HH:mm"),
            "Catatan": st.column_config.TextColumn("Catatan", width="medium"),
        },
    )

    # Simpan perubahan catatan ke database
    notes = dict(zip(table["_id"], edited["Catatan"].fillna("")))
    for c in cursors:
        if c["id"] in notes and notes[c["id"]] != c.get("note", ""):
            db_update_cursor_note(c["id"], notes[c["id"]])
            c["note"] = notes[c["id"]]

    to_delete = table.loc[edited["Hapus"].fillna(False).astype(bool).values, "_id"].tolist()
    b1, b2, b3 = st.columns([1.2, 1.2, 1.4])
    if b1.button(f"🗑️ Hapus yang dicentang ({len(to_delete)})", key=f"del_{panel_id}",
                 disabled=not to_delete):
        delete_cursors(to_delete)
        st.rerun()
    with b2.popover("Hapus semua"):
        st.write(f"Hapus {len(cursors)} cursor?")
        if st.button("Ya, hapus semua", key=f"delall_{panel_id}", type="primary"):
            delete_cursors([c["id"] for c in cursors])
            st.rerun()

    export = table.drop(columns=["Hapus", "_id"]).copy()
    lazy_download(b3, "⬇️ Download daftar cursor (Excel)",
                  lambda: to_excel_bytes((("Cursor", export),)),
                  "daftar_cursor.xlsx",
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                  f"dlcur_{panel_id}")


# =====================================================
# OVERLAY: CURSOR, EVENT, PEMISAH TAHUN
# =====================================================
def axis_refs(row_index):
    suffix = "" if row_index == 1 else str(row_index)
    return f"x{suffix}", f"y{suffix} domain"


def add_vertical_marker(fig, x, rows, color, dash, width=1.5):
    for i in rows:
        xref, yref = axis_refs(i)
        fig.add_shape(type="line", x0=x, x1=x, y0=0, y1=1, xref=xref, yref=yref,
                      line=dict(color=color, dash=dash, width=width), layer="above")


def add_box_label(fig, x, row, text, y, bgcolor, bordercolor, font_color="#111", xanchor="left"):
    xref, yref = axis_refs(row)
    fig.add_annotation(x=x, y=y, xref=xref, yref=yref, text=text, showarrow=False,
                       xanchor=xanchor, yanchor="top", align="left",
                       bgcolor=bgcolor, bordercolor=bordercolor, borderwidth=1, borderpad=3,
                       font=dict(size=11, color=font_color))


def add_cursor_overlay(fig, cursors, rows, x_kind="time", x_range=None):
    """x_kind: 'time' (pakai waktu cursor) atau 'coil' (sumbu X = nomor coil)."""
    for c in cursors:
        if x_kind == "time":
            x = cursor_time(c)
            if pd.isna(x):
                continue
            if x_range and not (x_range[0] <= x <= x_range[1]):
                continue
        else:
            x = c.get("coil")
            if x is None:
                continue
        add_vertical_marker(fig, x, rows, COLOR_CURSOR, "solid", 2)
        xref, _ = axis_refs(rows[0])
        fig.add_annotation(x=x, y=1, xref=xref, yref=axis_refs(rows[0])[1], yanchor="bottom",
                           text=f"C{c['id']}", showarrow=False,
                           bgcolor=COLOR_CURSOR, font=dict(color="white", size=11), borderpad=2)


def wrap_label(text, width=14):
    return "<br>".join(textwrap.wrap(str(text), width=width)) or str(text)


def add_event_overlay(fig, events, panel_rows, x_range, levels=3, mode="box"):
    """events: DataFrame Tanggal/Kejadian/Panel. panel_rows: {parameter: nomor baris}.
    mode 'box' = kotak keterangan; 'compact' = penanda kecil, keterangan muncul saat kursor diarahkan."""
    if events is None or events.empty:
        return 0
    all_rows = list(panel_rows.values())
    shown = 0
    counters = {r: 0 for r in all_rows}
    for _, ev in events.sort_values("Tanggal").iterrows():
        t = ev["Tanggal"]
        if pd.isna(t) or not (x_range[0] <= t <= x_range[1]):
            continue
        target = ev.get("Panel") or "Semua"
        rows = all_rows if target == "Semua" else [panel_rows[target]] if target in panel_rows else []
        if not rows:
            continue
        add_vertical_marker(fig, t, rows, COLOR_EVENT, "dash", 1.5)
        text = f"{ev['Kejadian']} ({t:%d/%m/%y})"
        for r in rows if target != "Semua" else rows[:1]:
            if mode == "compact":
                xref, yref = axis_refs(r)
                fig.add_annotation(x=t, y=1, xref=xref, yref=yref, yanchor="bottom", text="⚑",
                                   hovertext=text, showarrow=False, font=dict(size=13, color=COLOR_EVENT))
            else:
                level = counters[r] % levels
                counters[r] += 1
                add_box_label(fig, t, r, f"{wrap_label(ev['Kejadian'])}<br>{t:%d/%m/%y}",
                              y=0.99 - level * 0.22, bgcolor="rgba(255,255,255,0.92)", bordercolor="#333")
        shown += 1
    return shown


def add_year_overlay(fig, rows, x_range):
    start, end = x_range
    years = range(start.year, end.year + 1)
    for y in years:
        boundary = pd.Timestamp(year=y, month=1, day=1)
        if start < boundary <= end:
            add_vertical_marker(fig, boundary, rows, COLOR_YEAR, "dashdot", 2)
        label_x = max(boundary, start)
        if label_x <= end:
            add_box_label(fig, label_x, rows[0], f"<b>{y}</b>", y=0.99,
                          bgcolor="#FFF200", bordercolor="#999", font_color="#111")


def prepare_events(records):
    ev = pd.DataFrame(records or [], columns=["Tanggal", "Kejadian", "Panel"])
    ev["Tanggal"] = pd.to_datetime(ev["Tanggal"], errors="coerce", dayfirst=False)
    ev["Kejadian"] = ev["Kejadian"].fillna("").astype(str)
    ev["Panel"] = ev["Panel"].fillna("Semua").replace("", "Semua")
    return ev.dropna(subset=["Tanggal"])


def aggregate_frame(frame, time_col, params, rule):
    if rule is None:
        out = frame.copy()
        return out
    base = frame.dropna(subset=[time_col]).set_index(time_col)
    means = base[params].resample(rule).mean()
    counts = base[ID_COL].resample(rule).count()
    out = means.reset_index()
    out["Jumlah coil"] = counts.values
    out = out[out["Jumlah coil"] > 0].copy()
    out[ROW_KEY] = -1
    out[ID_COL] = "rata-rata " + out["Jumlah coil"].astype(str) + " coil"
    return out


def build_multipanel_figure(plot_df, time_col, panels, hover_cols, events, event_mode,
                            cursors, x_range, title, smooth_window=None):
    """panels: list of dict Parameter/Judul/Target/Batas bawah/Batas atas/Y min/Y max."""
    n = len(panels)
    fig = make_subplots(rows=n, cols=1, shared_xaxes=True, vertical_spacing=0.07 if n > 1 else 0.02,
                        subplot_titles=[p.get("Judul") or label(p["Parameter"]) for p in panels])
    Trace = scatter_cls(len(plot_df))
    dense = len(plot_df) > GL_THRESHOLD
    curve_params, panel_rows, stats = {}, {}, []

    for i, p in enumerate(panels, start=1):
        param = p["Parameter"]
        panel_rows[param] = i
        sub = plot_df.dropna(subset=[param])
        y = pd.to_numeric(sub[param], errors="coerce")
        target, lower, upper = p.get("Target"), p.get("Batas bawah"), p.get("Batas atas")
        tips = [c for c in hover_cols if c != param and c in sub.columns]
        template = make_hovertemplate("Waktu", label(param), tips, True)

        fig.add_trace(Trace(
            x=sub[time_col], y=y, mode="lines+markers", name=p.get("Judul") or label(param),
            line=dict(color="#4A90D9", width=1 if dense else 1.6),
            marker=dict(size=3 if dense else 5, color="#4A90D9"),
            customdata=make_customdata(sub, tips), hovertemplate=template, showlegend=False,
        ), row=i, col=1)
        curve_params[len(fig.data) - 1] = param

        status = process_status(y, lower, upper)
        out_mask = status.isin(["≥ batas atas", "≤ batas bawah"])
        if out_mask.any():
            fig.add_trace(Trace(
                x=sub.loc[out_mask, time_col], y=y[out_mask], mode="markers", name="Keluar batas",
                marker=dict(size=8, color=COLOR_HIGH, symbol="diamond"),
                customdata=make_customdata(sub[out_mask], tips), hovertemplate=template, showlegend=False,
            ), row=i, col=1)
            curve_params[len(fig.data) - 1] = param

        if smooth_window and len(y) > smooth_window:
            fig.add_trace(Trace(
                x=sub[time_col], y=y.rolling(smooth_window, min_periods=1, center=True).mean(),
                mode="lines", line=dict(color="#1B2A3A", width=2), hoverinfo="skip", showlegend=False,
            ), row=i, col=1)

        xref, _ = axis_refs(i)
        yref = "y" if i == 1 else f"y{i}"
        for val, dash, width in [(target, "dash", 1.2), (upper, "dashdot", 1.4), (lower, "dashdot", 1.4)]:
            if val is not None:
                fig.add_shape(type="line", xref=f"{xref} domain", x0=0, x1=1, yref=yref, y0=val, y1=val,
                              line=dict(color=COLOR_YEAR, dash=dash, width=width))
        parts = []
        if target is not None:
            parts.append(f"target {target:g}")
        if lower is not None or upper is not None:
            parts.append(f"batas {'' if lower is None else f'{lower:g}'} … {'' if upper is None else f'{upper:g}'}")
        if parts:
            fig.add_annotation(xref=f"{xref} domain", x=1, yref=f"{yref} domain", y=1, text=" | ".join(parts),
                               showarrow=False, xanchor="right", yanchor="bottom",
                               font=dict(color=COLOR_YEAR, size=11), bgcolor="rgba(255,242,0,0.85)")

        y_min, y_max = p.get("Y min"), p.get("Y max")
        if y_min is not None and y_max is not None and y_min < y_max:
            fig.update_yaxes(range=[y_min, y_max], row=i, col=1)
        fig.update_yaxes(gridcolor="rgba(128,128,128,0.2)", row=i, col=1)

        cap = capability(y, lower, upper)
        stats.append({"Panel": p.get("Judul") or label(param), "Titik": len(y),
                      "Rata-rata": round(cap["Rata-rata"], 2) if len(y) else None,
                      "Min": cap["Min"], "Max": cap["Max"],
                      "Keluar batas": int(out_mask.sum()),
                      "% dalam batas": round((1 - out_mask.mean()) * 100, 1) if len(y) and (lower is not None or upper is not None) else None,
                      "Cpk": round(cap["Cpk"], 2) if pd.notna(cap["Cpk"]) else None})

    rows = list(range(1, n + 1))
    if x_range[0].year != x_range[1].year:
        add_year_overlay(fig, rows, x_range)
    n_events = add_event_overlay(fig, events, panel_rows, x_range, mode=event_mode)
    add_cursor_overlay(fig, cursors, rows, "time", x_range)

    pad = (x_range[1] - x_range[0]) * 0.01 or pd.Timedelta(hours=1)
    fig.update_xaxes(range=[x_range[0] - pad, x_range[1] + pad], gridcolor="rgba(128,128,128,0.2)")
    fig.update_layout(
        title=dict(text=title, font=dict(size=15), x=0, xanchor="left"),
        height=max(260 * n + 140, 420),
        margin=dict(t=120, b=30, l=60, r=20), clickmode="event+select", dragmode="zoom",
        hoverlabel=dict(namelength=-1), plot_bgcolor="white",
    )
    apply_hairline(fig, "crosshair", True)
    add_range_buttons(fig)
    return fig, curve_params, stats, n_events
# =====================================================
# APLIKASI
# =====================================================
init_db()
init_session()

PAGE_HOME = "🏠 Ringkasan"
PAGE_TREND = "📈 Tren & Event"
PAGE_QUALITY = "🧪 Kualitas"
PAGE_PROCESS = "🌡️ Proses"
PAGE_COIL = "🔍 Coil"
PAGE_DATA = "⚙️ Data & Pengaturan"
PAGES = [PAGE_HOME, PAGE_TREND, PAGE_QUALITY, PAGE_PROCESS, PAGE_COIL, PAGE_DATA]

PERIODS = {
    "Semua data": None,
    "24 jam terakhir": pd.Timedelta(days=1),
    "7 hari terakhir": pd.Timedelta(days=7),
    "30 hari terakhir": pd.Timedelta(days=30),
    "90 hari terakhir": pd.Timedelta(days=90),
    "Pilih tanggal…": "custom",
}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

n_db, last_import = db_summary()
if n_db == 0 and "page" not in st.session_state:
    st.session_state["page"] = PAGE_DATA

st.sidebar.markdown("### 🏭 Coil CGL")
page = st.sidebar.radio("Menu", PAGES, key="page", label_visibility="collapsed")

LIMITS = load_limits()
if "events_saved" not in st.session_state:
    st.session_state["events_saved"] = load_events()
EVENTS = prepare_events(st.session_state["events_saved"])
HOVER_COLS = get_setting("hover_cols", DEFAULT_HOVER)
SMOOTH_N = int(get_setting("smooth_window", 25) or 25)
DATE_ORDER = get_setting("date_order", DATE_ORDERS[0])
if DATE_ORDER not in DATE_ORDERS:
    DATE_ORDER = DATE_ORDERS[0]


def sub_nav(options, key):
    """Navigasi kecil di dalam halaman (hanya bagian yang dipilih yang dihitung)."""
    return st.radio(key, options, horizontal=True, key=key, label_visibility="collapsed")


def friendly_select(container, title, options, default=None, key=None, help=None):
    idx = options.index(default) if default in options else 0
    return container.selectbox(title, options, index=idx, format_func=label, key=key, help=help)


# =====================================================
# DATA + FILTER SIDEBAR
# =====================================================
df, time_col = (load_dataset(data_version(), DATE_ORDER) if n_db else (None, None))
filtered = None

if df is not None and not df.empty:
    st.sidebar.divider()
    st.sidebar.markdown("**Filter**")
    filtered = df

    if time_col and df[time_col].notna().any():
        t_last = df[time_col].max()
        period = st.sidebar.selectbox("Periode", list(PERIODS.keys()), index=3, key="period")
        span = PERIODS[period]
        if span == "custom":
            d_min, d_max = df[time_col].min().date(), t_last.date()
            picked = st.sidebar.date_input("Tanggal", value=(max(d_min, d_max - timedelta(days=30)), d_max),
                                           min_value=d_min, max_value=d_max, format="DD/MM/YYYY")
            if isinstance(picked, tuple) and len(picked) == 2:
                start, end = pd.Timestamp(picked[0]), pd.Timestamp(picked[1]) + pd.Timedelta(days=1)
                filtered = filtered[(filtered[time_col] >= start) & (filtered[time_col] < end)]
        elif span is not None:
            filtered = filtered[filtered[time_col] >= t_last - span]

    if "JSG_DESI_NM" in df.columns:
        grades = sorted(df["JSG_DESI_NM"].dropna().astype(str).unique())
        if len(grades) > 1:
            g_pick = st.sidebar.multiselect("Grade", grades, placeholder="Semua grade", key="f_grade")
            if g_pick:
                filtered = filtered[filtered["JSG_DESI_NM"].astype(str).isin(g_pick)]

    thickness_range = None
    if "INLET_COIL_TH" in df.columns and df["INLET_COIL_TH"].notna().any():
        th_min, th_max = float(df["INLET_COIL_TH"].min()), float(df["INLET_COIL_TH"].max())
        if th_min < th_max:
            thickness_range = st.sidebar.slider("Tebal (mm)", th_min, th_max, (th_min, th_max),
                                                step=0.001, format="%.3f", key="f_thick")
            if thickness_range != (th_min, th_max):
                filtered = filtered[filtered["INLET_COIL_TH"].between(*thickness_range)]

    st.sidebar.caption(f"{len(filtered):,} dari {len(df):,} coil  ·  📍 {len(st.session_state['cursors'])} cursor")

hover_cols = [c for c in HOVER_COLS if df is not None and c in df.columns]


def numeric_options(frame, show_all=False):
    cols = []
    for c in frame.select_dtypes(include="number").columns:
        if c.startswith("_") or c in NON_TREND_COLUMNS:
            continue
        s = frame[c]
        if not show_all and (s.notna().sum() == 0 or s.nunique(dropna=True) <= 1):
            continue
        cols.append(c)
    return cols


def coil_label_map(frame):
    t = frame[time_col].dt.strftime("%d-%m-%Y %H:%M") if time_col else ""
    return dict(zip(frame[ROW_KEY], frame[ID_COL].astype(str) + ("  ·  " + t if time_col else "")))


def cursor_coil_rows():
    by_coil = dict(zip(df[ID_COL].astype(str), df[ROW_KEY]))
    keys = [by_coil.get(c.get("coil")) for c in st.session_state["cursors"]]
    return list(dict.fromkeys(k for k in keys if k is not None))


def process_violations(frame):
    """Per coil: daftar parameter yang keluar batas proses."""
    issues = pd.Series("", index=frame.index)
    for param, lim in LIMITS.items():
        if param not in frame.columns:
            continue
        stt = process_status(pd.to_numeric(frame[param], errors="coerce"), lim.get("lower"), lim.get("upper"))
        bad = stt.isin(["≥ batas atas", "≤ batas bawah"])
        issues[bad] = issues[bad] + label(param) + " " + stt[bad].astype(str) + "; "
    return issues.str.rstrip("; ")


def empty_state():
    st.info("Belum ada data. Buka **⚙️ Data & Pengaturan → Import data** untuk mengimpor file Excel/CSV.")


# =====================================================
# HALAMAN: RINGKASAN
# =====================================================
def page_home():
    st.header("Ringkasan")
    if last_import:
        st.caption(f"Data terakhir diimpor {last_import[0]} dari {last_import[1]}.")
    if time_col:
        n_future = int((df[time_col] > pd.Timestamp.now() + pd.Timedelta(days=1)).sum())
        if n_future:
            st.warning(f"{n_future} coil bertanggal di masa depan. Cek ⚙️ Data & Pengaturan → Validasi data.")
    f = filtered

    issues = process_violations(f)
    n_proc = int((issues != "").sum())
    k = st.columns(5)
    k[0].metric("Coil", f"{len(f):,}")
    if "COIL_WT" in f.columns:
        k[1].metric("Total berat", f"{f['COIL_WT'].sum() / 1000:,.1f} t")
    if "STATUS_MEKANIK" in f.columns:
        judged = f["STATUS_MEKANIK"].isin(["OK", "NG"])
        rate = (f.loc[judged, "STATUS_MEKANIK"] == "OK").mean() * 100 if judged.any() else np.nan
        n_ng = int((f["STATUS_MEKANIK"] == "NG").sum())
        k[2].metric("Lolos mekanik", f"{fmt_num(rate, 1)} %", delta=f"{n_ng} NG",
                    delta_color="inverse" if n_ng else "off")
    k[3].metric("Keluar batas proses", f"{n_proc:,} coil",
                delta="perlu dicek" if n_proc else "aman", delta_color="inverse" if n_proc else "off")
    if time_col and f[time_col].notna().any():
        k[4].metric("Periode", f"{f[time_col].min():%d/%m/%y} – {f[time_col].max():%d/%m/%y}")

    # Status per parameter yang punya batas
    rows = []
    for param, lim in LIMITS.items():
        if param not in f.columns:
            continue
        v = pd.to_numeric(f[param], errors="coerce")
        stt = process_status(v, lim.get("lower"), lim.get("upper"))
        n_out = int(stt.isin(["≥ batas atas", "≤ batas bawah"]).sum())
        cap = capability(v, lim.get("lower"), lim.get("upper"))
        cpk = cap["Cpk"]
        status = ("🔴 Ada yang keluar batas" if n_out else
                  "🟡 Mepet batas (Cpk < 1)" if pd.notna(cpk) and cpk < 1 else "🟢 Aman")
        rows.append({"Parameter": label(param), "Batas": f"{fmt_num(lim.get('lower'), 1)} … {fmt_num(lim.get('upper'), 1)}",
                     "Rata-rata": round(cap["Rata-rata"], 2), "Keluar batas": n_out,
                     "Cpk": round(cpk, 2) if pd.notna(cpk) else None, "Status": status})
    for p, (vcol, *_r, unit) in MECH_PROPS.items():
        if f"_ST_{p}" in f.columns and f[f"_ST_{p}"].ne("Tanpa spesifikasi").any():
            ll, ul = single_value(f[f"_LL_{p}"]), single_value(f[f"_UL_{p}"])
            cap = capability(f[vcol], ll, ul)
            n_out = int(f[f"_ST_{p}"].isin(["Di atas UL", "Di bawah LL"]).sum())
            cpk = cap["Cpk"]
            status = ("🔴 Ada yang keluar spesifikasi" if n_out else
                      "🟡 Mepet batas (Cpk < 1)" if pd.notna(cpk) and cpk < 1 else "🟢 Aman")
            rows.append({"Parameter": label(vcol), "Batas": f"{fmt_num(ll, 0)} … {fmt_num(ul, 0)}",
                         "Rata-rata": round(cap["Rata-rata"], 2), "Keluar batas": n_out,
                         "Cpk": round(cpk, 2) if pd.notna(cpk) else None, "Status": status})
    st.markdown("##### Kondisi parameter")
    if rows:
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    else:
        st.caption("Belum ada parameter dengan batas. Atur batas di halaman Tren & Event.")

    left, right = st.columns([3, 2])
    with left:
        st.markdown("##### ⚠️ Coil yang perlu diperhatikan")
        problem = pd.DataFrame({
            "Waktu": f[time_col] if time_col else pd.NaT,
            "Coil": f[ID_COL].astype(str),
            "Proses": issues,
            "Mekanik": np.where(f.get("STATUS_MEKANIK", pd.Series("", index=f.index)) == "NG", "NG", ""),
        })
        problem = problem[(problem["Proses"] != "") | (problem["Mekanik"] != "")]
        if problem.empty:
            st.success("Tidak ada coil bermasalah pada periode ini.")
        else:
            problem = problem.sort_values("Waktu", ascending=False)
            st.dataframe(problem, width="stretch", hide_index=True, height=min(38 * len(problem) + 38, 360),
                         column_config={"Waktu": st.column_config.DatetimeColumn(format="DD-MM-YYYY HH:mm")})
            st.caption(f"{len(problem)} coil. Detail per coil ada di halaman 🔍 Coil.")
    with right:
        st.markdown("##### 🛠️ Event terakhir")
        if EVENTS.empty:
            st.caption("Belum ada event maintenance.")
        else:
            ev = EVENTS.sort_values("Tanggal", ascending=False).head(6)
            for _, e in ev.iterrows():
                st.markdown(f"**{e['Tanggal']:%d/%m/%Y}** — {e['Kejadian']}")
        st.markdown("##### 📍 Cursor terakhir")
        curs = sorted(st.session_state["cursors"], key=lambda c: c["id"], reverse=True)[:5]
        if not curs:
            st.caption("Belum ada cursor.")
        for c in curs:
            st.markdown(f"**C{c['id']}** · {c.get('coil') or 'agregat'} · {label(c.get('param'))} = "
                        f"{fmt_num(c.get('value'), 2)}" + (f" · _{c['note']}_" if c.get("note") else ""))

    if time_col and "COIL_WT" in f.columns and f[time_col].notna().any():
        daily = f.dropna(subset=[time_col]).set_index(time_col).resample("D").agg(
            Tonase=("COIL_WT", lambda s: s.sum() / 1000), Coil=(ID_COL, "count"))
        daily = daily[daily["Coil"] > 0].reset_index()
        if len(daily) > 1:
            fig = go.Figure(go.Bar(x=daily[time_col], y=daily["Tonase"], marker_color=COLOR_OK,
                                   customdata=daily["Coil"],
                                   hovertemplate="%{x|%d-%m-%Y}<br>%{y:.1f} t · %{customdata} coil<extra></extra>"))
            fig.update_layout(title="Produksi per hari (ton)", height=280, margin=dict(t=40, b=20, l=50, r=10),
                              plot_bgcolor="white")
            st.plotly_chart(fig, width="stretch", key="home_daily", config={"displaylogo": False})


# =====================================================
# HALAMAN: TREN & EVENT
# =====================================================
def page_trend():
    st.header("Tren & Event")
    ss = st.session_state
    if not time_col:
        st.info("Halaman ini butuh kolom waktu (UPDAT_DT atau SHEAR_CUT_DT).")
        return

    all_numeric = numeric_options(df, show_all=True)
    code_by_text = {label_code(c): c for c in all_numeric}

    # ---------- panel & batas ----------
    ss.setdefault("panels_ver", 0)
    if "panels_saved" not in ss:
        ss["panels_saved"] = load_panels()
    with st.expander("⚙️ Atur panel & batas", expanded=not ss["panels_saved"]):
        base = pd.DataFrame(ss["panels_saved"], columns=PANEL_COLUMNS)
        base["Parameter"] = base["Parameter"].map(lambda c: label_code(c) if c in all_numeric else c)
        edited = st.data_editor(
            base, key=f"panels_editor_{ss['panels_ver']}", num_rows="dynamic", hide_index=True, width="stretch",
            column_config={
                "Parameter": st.column_config.SelectboxColumn("Parameter", options=list(code_by_text), required=True,
                                                              width="large"),
                "Judul": st.column_config.TextColumn("Judul panel"),
                "Target": st.column_config.NumberColumn("Target", format="%.2f"),
                "Batas bawah": st.column_config.NumberColumn("Batas bawah (≤ keluar)", format="%.2f"),
                "Batas atas": st.column_config.NumberColumn("Batas atas (≥ keluar)", format="%.2f"),
                "Y min": st.column_config.NumberColumn("Sumbu Y min", format="%.1f"),
                "Y max": st.column_config.NumberColumn("Sumbu Y max", format="%.1f"),
            },
        )
        records = clean_records(edited, PANEL_COLUMNS)
        for r in records:
            r["Parameter"] = code_by_text.get(r["Parameter"], r["Parameter"])
        records = [r for r in records if r.get("Parameter")]
        if records != ss["panels_saved"]:
            bad = [r for r in records if r.get("Batas bawah") is not None and r.get("Batas atas") is not None
                   and r["Batas bawah"] >= r["Batas atas"]]
            if bad:
                st.error("Batas bawah harus lebih kecil dari batas atas.")
            else:
                removed = {r["Parameter"] for r in ss["panels_saved"]} - {r["Parameter"] for r in records}
                save_panels(records)
                for p in removed:   # panel dihapus -> batasnya ikut dihapus
                    save_limit(p, None, None)
                ss["panels_saved"] = records
                st.rerun()
        st.caption("Urutan baris = urutan panel. Batas dipakai juga di Ringkasan dan Detail Coil. "
                   "Nilai tepat di batas dihitung keluar batas.")

    panels = [r for r in ss["panels_saved"] if r["Parameter"] in df.columns]
    if not panels:
        st.info("Tambahkan minimal satu panel di **Atur panel & batas**.")
        return

    # ---------- opsi tampilan ----------
    c = st.columns([1.4, 1.4, 1.2, 1.2])
    shown = c[0].multiselect("Panel ditampilkan", [p["Parameter"] for p in panels],
                             default=[p["Parameter"] for p in panels], format_func=label, key="trend_shown")
    agg_label = c[1].selectbox("Data", list(AGGREGATION_RULES.keys()), key="trend_agg",
                               help="Untuk ribuan coil, rata-rata per jam/hari membuat tren lebih mudah dibaca.")
    event_mode = c[2].selectbox("Label event", ["Kotak", "Ikon ⚑ (hover)"], key="trend_evmode",
                                help="Pilih Ikon jika event banyak dan labelnya saling tumpuk.")
    smooth = c[3].toggle(f"Garis tren halus ({SMOOTH_N} coil)", value=False, key="trend_smooth")
    panels = [p for p in panels if p["Parameter"] in shown]
    if not panels:
        st.info("Pilih minimal satu panel.")
        return

    params = [p["Parameter"] for p in panels]
    base_df = filtered.dropna(subset=[time_col]).sort_values(time_col)
    plot_df = aggregate_frame(base_df, time_col, params, AGGREGATION_RULES[agg_label])
    if plot_df.empty:
        st.warning("Tidak ada data pada periode/filter ini.")
        return

    x_start, x_end = plot_df[time_col].min(), plot_df[time_col].max()
    if x_start == x_end:
        x_start, x_end = x_start - pd.Timedelta(hours=12), x_end + pd.Timedelta(hours=12)

    grades = filtered["JSG_DESI_NM"].dropna().astype(str).unique() if "JSG_DESI_NM" in filtered else []
    title = "   |   ".join(
        ([", ".join(sorted(grades)[:3])] if len(grades) else [])
        + [f"{len(base_df):,} coil", f"{x_start:%d-%m-%Y} s/d {x_end:%d-%m-%Y}"]
    )
    tips = hover_cols + (["Jumlah coil"] if "Jumlah coil" in plot_df.columns else [])
    fig, curve_params, stats, n_events = build_multipanel_figure(
        plot_df, time_col, panels, tips, EVENTS,
        "compact" if event_mode.startswith("Ikon") else "box",
        st.session_state["cursors"], (x_start, x_end), title,
        smooth_window=SMOOTH_N if smooth and agg_label == "Per coil" else None,
    )
    interactive_chart(fig, f"trend_{agg_label}", df, time_col, curve_params)
    hidden = len(EVENTS) - n_events
    st.caption("Klik titik untuk menyimpan cursor · scroll untuk zoom · klik ganda untuk reset"
               + (f" · {hidden} event di luar periode ini" if hidden > 0 else ""))

    st.dataframe(pd.DataFrame(stats), width="stretch", hide_index=True)

    a, b = st.columns(2)
    with a.expander("＋ Tambah event maintenance"):
        with st.form("quick_event", clear_on_submit=True):
            e_date = st.date_input("Tanggal", value=min(x_end.date(), datetime.now().date()), format="DD/MM/YYYY")
            e_text = st.text_input("Kejadian", placeholder="mis. Clean SS Multi")
            e_panel = st.selectbox("Tampil di panel", ["Semua"] + params,
                                   format_func=lambda v: v if v == "Semua" else label(v))
            if st.form_submit_button("Simpan event", type="primary") and e_text.strip():
                new = ss["events_saved"] + [{"Tanggal": e_date.isoformat(), "Kejadian": e_text.strip(),
                                             "Panel": e_panel}]
                save_events(new)
                ss["events_saved"] = load_events()
                ss.pop("events_base", None)
                ss["events_ver"] = ss.get("events_ver", 0) + 1
                st.rerun()
        st.caption("Edit/hapus/import event di ⚙️ Data & Pengaturan → Event maintenance.")

    with b.expander("Coil keluar batas"):
        out = base_df.assign(Masalah=process_violations(base_df))
        out = out[out["Masalah"] != ""]
        if out.empty:
            st.write("Tidak ada.")
        else:
            st.dataframe(out[[time_col, ID_COL, "Masalah"]].rename(columns={time_col: "Waktu", ID_COL: "Coil"}),
                         width="stretch", hide_index=True,
                         column_config={"Waktu": st.column_config.DatetimeColumn(format="DD-MM-YYYY HH:mm")})

    render_cursor_panel(df, time_col, "trend", params)


# =====================================================
# HALAMAN: KUALITAS
# =====================================================
def page_quality():
    st.header("Kualitas")
    view = sub_nav(["Sifat mekanik", "Grade & defect", "Target vs actual"], "quality_view")

    f = filtered
    fc = st.columns(4)
    for i, col in enumerate(["STATUS_MEKANIK", "NO1_DEFCT_CD", "TOP_SRFCE_GRADE_CD", "BOTTM_SRFCE_GRADE_CD"]):
        if col in f.columns:
            s = f[col].astype("string").fillna("(Kosong)")
            opts = sorted(s.unique())
            if len(opts) > 1:
                pick = fc[i].multiselect(label(col), opts, placeholder="Semua", key=f"qf_{col}")
                if pick:
                    f = f[s.isin(pick)]
    if f.empty:
        st.warning("Tidak ada coil yang cocok dengan filter.")
        return

    if view == "Sifat mekanik":
        props = [p for p, (v, *_r) in MECH_PROPS.items()
                 if v in f.columns and f"_ST_{p}" in f.columns and f[f"_ST_{p}"].ne("Tanpa spesifikasi").any()]
        if not props:
            st.info("Tidak ada sifat mekanik yang memiliki nilai dan spesifikasi (LL/UL).")
            return
        rows = []
        groups = f.groupby("JSG_DESI_NM", dropna=False) if "JSG_DESI_NM" in f.columns else [("Semua", f)]
        for gname, gdf in groups:
            for p in props:
                vcol, _, _, unit = MECH_PROPS[p]
                ll, ul = single_value(gdf[f"_LL_{p}"]), single_value(gdf[f"_UL_{p}"])
                cap = capability(gdf[vcol], ll, ul)
                rows.append({"Grade": gname, "Sifat": label(vcol), "Coil": cap["n"], "LL": ll, "UL": ul,
                             "Rata-rata": round(cap["Rata-rata"], 2), "Min": cap["Min"], "Max": cap["Max"],
                             "NG": int(gdf[f"_ST_{p}"].isin(["Di atas UL", "Di bawah LL"]).sum()),
                             "Cpk": round(cap["Cpk"], 2) if pd.notna(cap["Cpk"]) else None,
                             "Kapabilitas": cpk_label(cap["Cpk"])})
        cap_df = pd.DataFrame(rows)
        mcols = st.columns(len(props))
        for box, p in zip(mcols, props):
            stt = f[f"_ST_{p}"]
            judged = stt.ne("Tanpa spesifikasi")
            ng = int(stt.isin(["Di atas UL", "Di bawah LL"]).sum())
            rate = (1 - ng / judged.sum()) * 100 if judged.any() else np.nan
            box.metric(f"{p} lolos", f"{fmt_num(rate, 1)} %", delta=f"{ng} NG",
                       delta_color="inverse" if ng else "off")
        st.dataframe(cap_df, width="stretch", hide_index=True)

        p_sel = st.radio("Grafik", props, horizontal=True, format_func=lambda p: label(MECH_PROPS[p][0]),
                         key="mech_prop")
        vcol, _, _, unit = MECH_PROPS[p_sel]
        x_mech = time_col or ID_COL
        plot_df = f.dropna(subset=[vcol])
        if x_mech != ID_COL:
            plot_df = plot_df.dropna(subset=[x_mech]).sort_values(x_mech)
        ll_s, ul_s = plot_df[f"_LL_{p_sel}"], plot_df[f"_UL_{p_sel}"]
        ll_c, ul_c = single_value(ll_s), single_value(ul_s)
        fig = build_trend_figure(
            plot_df, x_mech, vcol, plot_df[f"_ST_{p_sel}"], STATUS_MECH_STYLE,
            limit_lines=[
                {"name": "UL", "value": ul_s if (ul_s.notna().any() and ul_c is None) else ul_c,
                 "color": COLOR_LIMIT, "pos": "top left"},
                {"name": "LL", "value": ll_s if (ll_s.notna().any() and ll_c is None) else ll_c,
                 "color": COLOR_LOW, "pos": "bottom left"},
            ],
            band=(ll_c, ul_c), hover_cols=[h for h in hover_cols if h != vcol], hairline_mode="crosshair",
            title=f"{label(vcol)} per coil", x_label=label(x_mech), y_label=label(vcol),
            x_is_time=x_mech != ID_COL,
        )
        if x_mech != ID_COL and len(plot_df):
            add_cursor_overlay(fig, st.session_state["cursors"], [1], "time",
                               (plot_df[x_mech].min(), plot_df[x_mech].max()))
        fig.update_layout(margin=dict(t=100), plot_bgcolor="white")
        interactive_chart(fig, "mech", df, time_col, vcol)

        h1, h2 = st.columns(2)
        hist = px.histogram(plot_df, x=vcol, nbins=30, title=f"Distribusi {label(vcol)}")
        hist.update_traces(marker_color=COLOR_OK)
        for v, name, color in [(ul_c, "UL", COLOR_LIMIT), (ll_c, "LL", COLOR_LOW)]:
            if v is not None:
                hist.add_vline(x=v, line_color=color, line_dash="dash", annotation_text=f"{name} {v:g}",
                               annotation_font_color=color)
        hist.update_layout(height=320, bargap=0.05, margin=dict(t=50, b=30), plot_bgcolor="white",
                           xaxis_title=label(vcol), yaxis_title="Jumlah coil")
        h1.plotly_chart(hist, width="stretch", key="mech_hist")
        with h2:
            ng_df = f[f["STATUS_MEKANIK"] == "NG"] if "STATUS_MEKANIK" in f.columns else f.iloc[0:0]
            st.markdown(f"##### Coil NG ({len(ng_df)})")
            if ng_df.empty:
                st.success("Semua coil dalam spesifikasi.")
            else:
                show = pd.DataFrame({
                    "Waktu": ng_df[time_col] if time_col else None, "Coil": ng_df[ID_COL],
                    "Penyebab": ng_df.apply(lambda r: ", ".join(
                        f"{p} {r[f'_ST_{p}']}" for p in props if r[f"_ST_{p}"] in ("Di atas UL", "Di bawah LL")),
                        axis=1)})
                st.dataframe(show.sort_values("Waktu", ascending=False), width="stretch", hide_index=True,
                             height=280,
                             column_config={"Waktu": st.column_config.DatetimeColumn(format="DD-MM-YYYY HH:mm")})
        render_cursor_panel(df, time_col, "mech", [vcol])

    elif view == "Grade & defect":
        cols = [c for c in ["NO1_DEFCT_CD", "NO2_DEFCT_CD", "NO3_DEFCT_CD", "TOP_SRFCE_GRADE_CD",
                            "BOTTM_SRFCE_GRADE_CD", "SHAPE_GRADE_CD", "STATUS_MEKANIK"]
                if c in f.columns and f[c].notna().any()]
        if not cols:
            st.info("Kolom grade/defect tidak ditemukan.")
            return
        q_col = st.selectbox("Kategori", cols, format_func=label, key="q_cat")
        counts = (f[q_col].astype("string").fillna("(Tidak ada)").value_counts()
                  .rename_axis("Nilai").reset_index(name="Jumlah coil"))
        counts["%"] = (counts["Jumlah coil"] / counts["Jumlah coil"].sum() * 100).round(1)
        fig = px.bar(counts, x="Nilai", y="Jumlah coil", text="Jumlah coil", title=f"Distribusi {label(q_col)}")
        fig.update_traces(marker_color=COLOR_OK)
        fig.update_layout(height=380, plot_bgcolor="white")
        a, b = st.columns([2, 1])
        a.plotly_chart(fig, width="stretch", key="q_bar")
        b.dataframe(counts, width="stretch", hide_index=True)

        if "Thickness group" in f.columns:
            st.markdown(f"##### {label(q_col)} per kelompok tebal")
            ct = pd.crosstab(f["Thickness group"], f[q_col].astype("string").fillna("(Tidak ada)"),
                             normalize="index").mul(100).round(1)
            st.dataframe(ct, width="stretch")
            st.caption("Angka = % coil di setiap kelompok tebal.")

    else:
        pairs = [p for p in TARGET_ACTUAL_PAIRS if p[0] in f.columns and p[1] in f.columns]
        if not pairs:
            st.info("Kolom Target/Actual tidak ditemukan.")
            return
        rows = []
        for t, a_, name in pairs:
            tv, av = pd.to_numeric(f[t], errors="coerce"), pd.to_numeric(f[a_], errors="coerce")
            ok = tv.notna() & av.notna()
            n = int(ok.sum())
            rows.append({"Pasangan": name, "Coil": n,
                         "% sama": round((av[ok] == tv[ok]).mean() * 100, 1) if n else None,
                         "% actual > target": round((av[ok] > tv[ok]).mean() * 100, 1) if n else None,
                         "% actual < target": round((av[ok] < tv[ok]).mean() * 100, 1) if n else None})
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        pick = st.selectbox("Pasangan", [p[2] for p in pairs], key="ta_pick")
        t, a_, _ = next(p for p in pairs if p[2] == pick)
        ct = pd.crosstab(f[t], f[a_]).reset_index().melt(id_vars=t, var_name="Actual", value_name="Jumlah")
        ct["Target"], ct["Actual"] = ct[t].astype(str), ct["Actual"].astype(str)
        fig = px.bar(ct, x="Target", y="Jumlah", color="Actual", barmode="group", text="Jumlah",
                     title=f"Nilai actual untuk setiap target ({pick})")
        fig.update_layout(height=380, plot_bgcolor="white")
        st.plotly_chart(fig, width="stretch", key="ta_bar")
        st.caption("Arti kode Target/Actual belum dikonfirmasi; halaman ini hanya membandingkan nilainya.")


# =====================================================
# HALAMAN: PROSES
# =====================================================
def page_process():
    st.header("Proses")
    view = sub_nav(["Profil suhu", "Skin pass"], "process_view")
    f = filtered
    limits_for_profile = {c: (v.get("lower"), v.get("upper")) for c, v in LIMITS.items()}

    if view == "Profil suhu":
        names = coil_label_map(f)
        options = list(names.keys())[::-1]
        cur = [k for k in cursor_coil_rows() if k in names][:6]
        selected = st.multiselect("Coil yang dibandingkan (default: coil dari cursor)", options,
                                  default=cur or options[:1], format_func=lambda k: names.get(k, str(k)),
                                  key=f"profile_sel_{hash(tuple(cur))}", max_selections=10)
        fig = profile_figure(f, STRIP_STAGES, selected, "Suhu strip per titik ukur (°C)",
                             "Suhu strip (°C)", limits_for_profile)
        if fig:
            fig.update_layout(plot_bgcolor="white")
            st.plotly_chart(fig, width="stretch", key="profile_strip")
            st.caption("Area biru = rentang min–max semua coil pada periode ini; garis putus = rata-rata.")
        a, b = st.columns(2)
        fig = profile_figure(f, FURNACE_ZONES, selected, "Temperatur zona tungku (°C)", "°C")
        if fig:
            fig.update_layout(plot_bgcolor="white")
            a.plotly_chart(fig, width="stretch", key="profile_furnace")
        fig = profile_figure(f, GALV_POINTS, selected, "Suhu area galvanizing (°C)", "°C")
        if fig:
            fig.update_layout(plot_bgcolor="white")
            b.plotly_chart(fig, width="stretch", key="profile_galv")
        deltas = [d for d, *_ in DELTA_COLUMNS if d in f.columns]
        if deltas:
            stats = f[deltas].describe().T[["mean", "std", "min", "max"]].round(2)
            stats.columns = ["Rata-rata", "Std", "Min", "Max"]
            stats.index = [label(i) for i in stats.index]
            st.markdown("##### Selisih suhu antar titik ukur")
            st.dataframe(stats, width="stretch")

    else:
        skp = [c for c, _ in SKIN_PASS if c in f.columns]
        mech = [MECH_PROPS[p][0] for p in MECH_PROPS
                if MECH_PROPS[p][0] in f.columns and f[MECH_PROPS[p][0]].fillna(0).ne(0).any()]
        if not skp or not mech:
            st.info("Kolom skin pass atau sifat mekanik tidak ditemukan.")
            return
        c = st.columns(3)
        x = friendly_select(c[0], "Parameter skin pass (X)", skp, key="skp_x")
        y = friendly_select(c[1], "Sifat mekanik (Y)", mech, key="skp_y")
        color_opts = [o for o in ["Thickness group", "STATUS_MEKANIK", "JSG_DESI_NM"] if o in f.columns]
        color_by = friendly_select(c[2], "Warna", color_opts, key="skp_color")

        sp = f.dropna(subset=[x, y])
        Trace = scatter_cls(len(sp))
        fig = go.Figure()
        tips = [h for h in hover_cols if h not in (x, y)]
        template = make_hovertemplate(label(x), label(y), tips, False)
        palette = px.colors.qualitative.Safe
        curve_map = {}
        for i, (g, gdf) in enumerate(sp.groupby(color_by, dropna=False)):
            fig.add_trace(Trace(x=gdf[x], y=gdf[y], mode="markers", name=str(g),
                                marker=dict(size=6 if len(sp) > GL_THRESHOLD else 9, color=palette[i % len(palette)],
                                            opacity=0.75),
                                customdata=make_customdata(gdf, tips), hovertemplate=template))
            curve_map[len(fig.data) - 1] = y
        xn, yn = pd.to_numeric(sp[x], errors="coerce"), pd.to_numeric(sp[y], errors="coerce")
        ok = xn.notna() & yn.notna()
        r = np.nan
        if ok.sum() >= 3 and xn[ok].nunique() > 1:
            slope, intercept = np.polyfit(xn[ok], yn[ok], 1)
            r = np.corrcoef(xn[ok], yn[ok])[0, 1]
            xs = np.linspace(xn[ok].min(), xn[ok].max(), 50)
            fig.add_trace(go.Scatter(x=xs, y=slope * xs + intercept, mode="lines", name="Garis tren",
                                     line=dict(color="#1B2A3A", dash="dash"), hoverinfo="skip"))
        fig.update_layout(title=f"{label(y)} vs {label(x)}", height=480, xaxis_title=label(x), yaxis_title=label(y),
                          clickmode="event+select", plot_bgcolor="white",
                          legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        apply_hairline(fig, "crosshair", False)
        interactive_chart(fig, "skp", df, time_col, curve_map)
        m = st.columns(3)
        m[0].metric("Korelasi (r)", fmt_num(r))
        m[1].metric("Coil", f"{int(ok.sum()):,}")
        m[2].caption("r mendekati ±1 = hubungan kuat. Korelasi menunjukkan kecenderungan, bukan sebab-akibat.")
        if "Thickness group" in f.columns:
            summary = f.groupby("Thickness group")[skp + mech].mean().round(3)
            summary.insert(0, "Coil", f.groupby("Thickness group").size())
            summary.columns = [label(c) for c in summary.columns]
            st.markdown("##### Rata-rata per kelompok tebal")
            st.dataframe(summary, width="stretch")


# =====================================================
# HALAMAN: COIL
# =====================================================
def page_coil():
    st.header("Coil")
    names = coil_label_map(df)
    options = list(names.keys())[::-1]          # terbaru di atas
    cur = cursor_coil_rows()
    default = cur[-1] if cur else options[0]
    key = st.selectbox("Cari nomor coil", options, index=options.index(default),
                       format_func=lambda k: names.get(k, str(k)), key=f"coil_pick_{default}",
                       help="Ketik nomor coil untuk mencari. Default: coil dari cursor terakhir.")
    row = df[df[ROW_KEY] == key].iloc[0]
    coil_id = str(row[ID_COL])

    status = row.get("STATUS_MEKANIK", "-")
    proc = process_violations(df[df[ROW_KEY] == key]).iloc[0]
    badge = "🔴 NG mekanik" if status == "NG" else "🟢 Mekanik OK" if status == "OK" else "⚪ Tanpa spesifikasi"
    badge += "  ·  " + ("🔴 Keluar batas proses" if proc else "🟢 Proses dalam batas")
    st.markdown(f"### {coil_id}  \n{badge}")

    k = st.columns(6)
    k[0].metric("Waktu", fmt_dt(row.get(UPDAT_DT_COL)))
    k[1].metric("Grade", str(row.get("JSG_DESI_NM", "-")))
    k[2].metric("Tebal", f"{fmt_num(row.get('INLET_COIL_TH'), 3)} mm")
    k[3].metric("Lebar", f"{fmt_num(row.get('INLET_COIL_WH'), 0)} mm")
    k[4].metric("Berat", f"{fmt_num(row.get('COIL_WT'), 0)} kg")
    k[5].metric("Coil keluar", str(row.get("OTLET_COIL_NO", "-")))
    if proc:
        st.warning(f"Keluar batas proses: {proc}")
    if row.get("_DATE_NOTE"):
        st.caption(f"ℹ️ Teks UPDAT_DT asli: {row.get('UPDAT_DT')} → {row['_DATE_NOTE']}.")

    left, right = st.columns(2)
    with left:
        st.markdown("##### Sifat mekanik")
        mech_rows = [{"Sifat": label(v), "Nilai": row[v], "LL": row[f"_LL_{p}"], "UL": row[f"_UL_{p}"],
                      "Status": row[f"_ST_{p}"]}
                     for p, (v, *_r) in MECH_PROPS.items()
                     if f"_ST_{p}" in row.index and row[f"_ST_{p}"] != "Tanpa spesifikasi"]
        st.dataframe(pd.DataFrame(mech_rows), width="stretch", hide_index=True)
        st.markdown("##### Parameter dengan batas")
        lim_rows = []
        for param, lim in LIMITS.items():
            if param in row.index:
                v = pd.to_numeric(pd.Series([row[param]]), errors="coerce")
                lim_rows.append({"Parameter": label(param), "Nilai": row[param],
                                 "Batas": f"{fmt_num(lim.get('lower'), 1)} … {fmt_num(lim.get('upper'), 1)}",
                                 "Status": process_status(v, lim.get("lower"), lim.get("upper")).iloc[0]})
        st.dataframe(pd.DataFrame(lim_rows), width="stretch", hide_index=True)
    with right:
        st.markdown("##### Grade & defect")
        q_cols = ["TOP_SRFCE_GRADE_CD", "BOTTM_SRFCE_GRADE_CD", "SHAPE_GRADE_CD", "NO1_DEFCT_CD",
                  "NO1_DEFCT_GRADE_CD", "NO2_DEFCT_CD", "NO2_DEFCT_GRADE_CD", "NO3_DEFCT_CD"]
        st.dataframe(pd.DataFrame([{"Item": label(c), "Nilai": "-" if pd.isna(row[c]) else row[c]}
                                   for c in q_cols if c in row.index]), width="stretch", hide_index=True)
        st.markdown("##### Skin pass vs rata-rata semua coil")
        sk = []
        for c, _ in SKIN_PASS:
            if c in row.index:
                mean = pd.to_numeric(df[c], errors="coerce").mean()
                sk.append({"Parameter": label(c), "Nilai": row[c], "Rata-rata": round(mean, 3),
                           "Selisih": round(row[c] - mean, 3) if pd.notna(row[c]) else None})
        st.dataframe(pd.DataFrame(sk), width="stretch", hide_index=True)

    fig = profile_figure(df, STRIP_STAGES, [key], f"Profil suhu strip {coil_id} vs semua coil",
                         "Suhu strip (°C)", {c: (v.get("lower"), v.get("upper")) for c, v in LIMITS.items()})
    if fig:
        fig.update_layout(plot_bgcolor="white")
        st.plotly_chart(fig, width="stretch", key="coil_profile")

    a, b = st.columns(2)
    with a:
        st.markdown("##### Cursor pada coil ini")
        cc = [c for c in st.session_state["cursors"] if c.get("coil") == coil_id]
        if cc:
            st.dataframe(pd.DataFrame([{"Cursor": f"C{c['id']}", "Parameter": label(c["param"]),
                                        "Nilai": c["value"], "Catatan": c.get("note", "")} for c in cc]),
                         width="stretch", hide_index=True)
        else:
            st.caption("Tidak ada.")
    with b:
        st.markdown("##### Riwayat perubahan data")
        with closing(db_connect()) as con:
            hist = pd.read_sql_query(
                "SELECT h.import_id AS 'Import', l.imported_at AS 'Waktu import', h.kolom AS 'Kolom', "
                "h.nilai_lama AS 'Lama', h.nilai_baru AS 'Baru' FROM coil_history h "
                "LEFT JOIN import_log l ON l.id = h.import_id WHERE h.coil = ? ORDER BY h.id DESC",
                con, params=(coil_id,))
        if hist.empty:
            st.caption("Belum pernah berubah sejak pertama diimpor.")
        else:
            hist["Kolom"] = hist["Kolom"].map(label)
            st.dataframe(hist, width="stretch", hide_index=True)

    with st.expander("Semua kolom"):
        full = export_view(df[df[ROW_KEY] == key]).T.reset_index()
        full.columns = ["Kolom", "Nilai"]
        full.insert(1, "Nama", full["Kolom"].map(lambda c: COLUMN_LABELS.get(c, "")))
        full["Nilai"] = full["Nilai"].astype(str)
        st.dataframe(full, width="stretch", hide_index=True, height=420)


# =====================================================
# HALAMAN: DATA & PENGATURAN
# =====================================================
def section_import():
    st.caption(f"Kunci unik coil: {KEY_COL}. File dibandingkan dulu dengan database; "
               "tidak ada yang disimpan sebelum Anda menekan **Simpan ke database**.")
    src = st.radio("Sumber file", ["Upload file", "Ambil dari folder"], horizontal=True, key="imp_src")
    file_bytes = file_name = None
    if src == "Upload file":
        up = st.file_uploader("File Excel atau CSV", type=["xlsx", "xls", "xlsm", "csv"], key="imp_up")
        if up is not None:
            file_bytes, file_name = up.getvalue(), up.name
    else:
        folder = Path(st.text_input("Folder", value=get_setting("import_folder", str(Path.cwd())),
                                    key="imp_folder")).expanduser()
        if folder.is_dir():
            if str(folder) != get_setting("import_folder"):
                set_setting("import_folder", str(folder))
            files = sorted((p for p in folder.iterdir()
                            if p.is_file() and p.suffix.lower() in {".xlsx", ".xls", ".xlsm", ".csv"}),
                           key=lambda p: p.stat().st_mtime, reverse=True)
            if files:
                chosen = st.selectbox("File (terbaru di atas)", files, format_func=lambda p: p.name, key="imp_file")
                file_bytes, file_name = chosen.read_bytes(), chosen.name
            else:
                st.warning("Tidak ada file Excel/CSV di folder ini.")
        else:
            st.error("Folder tidak ditemukan.")

    if st.session_state.get("_import_msg"):
        st.success(st.session_state.pop("_import_msg"))
    if file_bytes is None:
        return

    c1, c2 = st.columns(2)
    sheet_name = None
    if Path(file_name).suffix.lower() != ".csv":
        sheets = list_sheets(file_bytes)
        sheet_name = c1.selectbox("Sheet", sheets, index=sheets.index("Data") if "Data" in sheets else 0,
                                  key="imp_sheet")
    header_row = int(c2.number_input("Baris judul kolom", 1, 100, 1, key="imp_header"))
    try:
        res = compare_with_db(file_bytes, file_name, sheet_name, header_row, data_version())
    except Exception as e:
        res = {"error": f"File tidak bisa dibaca: {e}"}
    if res.get("error"):
        st.error(res["error"])
        return

    st.markdown("##### Hasil perbandingan dengan database")
    m = st.columns(5)
    m[0].metric("🆕 Coil baru", f"{len(res['new_keys']):,}")
    m[1].metric("✏️ Berubah", f"{len(res['changed_keys']):,}")
    m[2].metric("✅ Sama", f"{res['same_count']:,}")
    m[3].metric("Duplikat di file", f"{len(res['duplicates']):,}",
                help="Nomor coil muncul lebih dari sekali di file; yang dipakai baris terakhir.")
    m[4].metric("Hanya di database", f"{res['db_only_count']:,}",
                help="Coil di database yang tidak ada di file ini. Tidak dihapus.")
    if not res["table_exists"]:
        st.info("Database masih kosong: semua coil akan disimpan sebagai data baru.")
    if res["no_key"]:
        st.warning(f"{res['no_key']} baris tanpa nomor coil akan dilewati.")
    if res["new_cols"]:
        st.warning(f"Kolom baru di file: {', '.join(res['new_cols'])}")

    inc = res["incoming"]
    if "UPDAT_DT" in inc.columns:
        t_inc, note_inc = resolve_updat_dt(inc, DATE_ORDER)
        n_fixed = int((note_inc != "").sum())
        n_future = int((t_inc > pd.Timestamp.now() + pd.Timedelta(days=1)).sum())
        n_bad = int((t_inc.isna() & inc["UPDAT_DT"].notna()).sum())
        msg = (f"Tanggal UPDAT_DT di file ini: {t_inc.min():%d-%m-%Y} s/d {t_inc.max():%d-%m-%Y} "
               f"(cara baca: {DATE_ORDER}).") if t_inc.notna().any() else "UPDAT_DT tidak terbaca."
        if n_fixed:
            msg += f" {n_fixed:,} tanggal ambigu dibaca dengan hari/bulan ditukar."
        (st.warning if (n_future or n_bad) else st.info)(
            msg + (f" ⚠️ {n_future} tanggal di masa depan." if n_future else "")
            + (f" ⚠️ {n_bad} tidak terbaca." if n_bad else ""))
        with st.expander("Lihat contoh cara baca tanggal"):
            sample = pd.DataFrame({KEY_COL: inc[KEY_COL], "Teks UPDAT_DT": inc["UPDAT_DT"].astype(str),
                                   "Dibaca sebagai": t_inc.dt.strftime("%d %b %Y %H:%M"), "Keterangan": note_inc})
            fixed = sample[note_inc != ""]
            st.dataframe(pd.concat([fixed.head(10), sample[note_inc == ""].head(5)]),
                         width="stretch", hide_index=True)
    preview = [c for c in [KEY_COL, "UPDAT_DT", "JSG_DESI_NM", "INLET_COIL_TH", "COIL_WT"] if c in inc.columns]
    if res["new_keys"]:
        with st.expander(f"Lihat coil baru ({len(res['new_keys']):,})"):
            st.dataframe(inc[inc[KEY_COL].isin(set(res["new_keys"]))][preview], width="stretch", hide_index=True)
    if len(res["diffs"]):
        with st.expander(f"Lihat perubahan ({len(res['diffs']):,} nilai di {len(res['changed_keys']):,} coil)",
                         expanded=len(res["diffs"]) <= 30):
            d = res["diffs"].copy()
            per_col = d["Kolom"].value_counts().rename_axis("Kolom").reset_index(name="Coil")
            per_col["Kolom"] = per_col["Kolom"].map(label_code)
            a, b = st.columns([1, 2])
            a.dataframe(per_col, width="stretch", hide_index=True)
            d["Kolom"] = d["Kolom"].map(label)
            b.dataframe(d, width="stretch", hide_index=True, height=300)
    if len(res["duplicates"]):
        with st.expander("Lihat duplikat"):
            st.dataframe(res["duplicates"], width="stretch", hide_index=True)

    o = st.columns(3)
    do_new = o[0].checkbox(f"Tambah coil baru ({len(res['new_keys']):,})", value=True, key="imp_do_new")
    do_chg = o[1].checkbox(f"Perbarui coil berubah ({len(res['changed_keys']):,})", value=True, key="imp_do_chg",
                           help="Nilai lama dicatat di riwayat perubahan.")
    add_cols = o[2].checkbox(f"Tambah kolom baru ({len(res['new_cols'])})", value=True,
                             disabled=not res["new_cols"], key="imp_add_cols")
    nothing = not ((do_new and res["new_keys"]) or (do_chg and res["changed_keys"]) or (add_cols and res["new_cols"]))
    if nothing:
        st.caption("Isi file sudah sama dengan database — tidak ada yang perlu disimpan.")
    if st.button("💾 Simpan ke database", type="primary", disabled=nothing, key="imp_save"):
        with st.spinner("Menyimpan..."):
            s = apply_import(res, file_name, sheet_name, do_new, do_chg, add_cols)
        st.session_state["_import_msg"] = (
            f"Tersimpan: {s['n_new']:,} coil baru, {s['n_changed']:,} coil diperbarui"
            + (f", kolom baru {', '.join(s['added_cols'])}" if s["added_cols"] else "") + ".")
        st.rerun()


def section_validation():
    if df is None:
        empty_state()
        return
    c = st.columns(2)
    density = c[0].number_input("Densitas baja (g/cm³)", value=7.85, step=0.01, format="%.2f")
    tol = c[1].number_input("Toleransi selisih berat (%)", value=2.0, step=0.5, min_value=0.0)
    checks, details = [], {}

    def add(name, n, note, detail=None, info=False):
        checks.append({"Pemeriksaan": name, "Hasil": "ℹ️ Info" if info else ("✅ OK" if n == 0 else "⚠️ Cek"),
                       "Temuan": n, "Keterangan": note})
        if detail is not None and len(detail):
            details[name] = detail

    if "UPDAT_DT" in df.columns:
        bad = df[df[UPDAT_DT_COL].isna() & df["UPDAT_DT"].notna()]
        add("UPDAT_DT terbaca sebagai tanggal-jam", len(bad), "Format d/m/yyyy h:mm", bad[[ID_COL, "UPDAT_DT"]])
        dup = df[df[UPDAT_DT_COL].duplicated(keep=False) & df[UPDAT_DT_COL].notna()]
        add("Waktu UPDAT_DT kembar", len(dup), "Dua coil dengan waktu sama", dup[[ID_COL, "UPDAT_DT"]])
        future = df[df[UPDAT_DT_COL] > pd.Timestamp.now() + pd.Timedelta(days=1)]
        add("Tanggal di masa depan", len(future), "UPDAT_DT lebih dari hari ini",
            future[[ID_COL, "UPDAT_DT", UPDAT_DT_COL]])
        if SHEAR_DT_COL in df.columns:
            diff_h = (df[UPDAT_DT_COL] - df[SHEAR_DT_COL]).abs().dt.total_seconds() / 3600
            far = df[diff_h > 24]
            add("UPDAT_DT jauh dari SHEAR_CUT_DT (> 1 hari)", len(far), "Kemungkinan tanggal salah",
                far[[ID_COL, "UPDAT_DT", UPDAT_DT_COL, "SHEAR_CUT_DT"]])
        if "_DATE_NOTE" in df.columns:
            fixed = df[df["_DATE_NOTE"] != ""]
            add("Tanggal ambigu yang dibaca ulang", len(fixed), f"Cara baca: {DATE_ORDER}",
                fixed[[ID_COL, "UPDAT_DT", UPDAT_DT_COL, "_DATE_NOTE"]].rename(columns={"_DATE_NOTE": "Keterangan"}),
                info=True)
        gap = df[UPDAT_DT_COL].sort_values().diff().dt.total_seconds() / 60
        if gap.notna().any():
            add("Jarak antar coil", 0, f"rata-rata {gap.mean():.1f} menit, terlama {gap.max():,.0f} menit", info=True)
    if "OTLET_COIL_NO" in df.columns:
        mm = df[df["OTLET_COIL_NO"].astype(str) != df[ID_COL].astype(str) + "1"]
        add("No. coil keluar = coil masuk + '1'", len(mm), "Pola penomoran", mm[[ID_COL, "OTLET_COIL_NO"]])
    if {"INLET_COIL_TH", "INLET_COIL_WH", "INLET_COIL_LG", "COIL_WT"} <= set(df.columns):
        est = df["INLET_COIL_TH"] * df["INLET_COIL_WH"] * df["INLET_COIL_LG"] * density / 1000
        dev = (df["COIL_WT"] - est) / est * 100
        w = df[[ID_COL, "COIL_WT"]].assign(**{"Berat hitung": est.round(0), "Selisih %": dev.round(2)})
        add("Berat coil vs dimensi", int((dev.abs() > tol).sum()),
            f"tebal × lebar × panjang × {density}; rata-rata selisih {dev.abs().mean():.2f}%", w[dev.abs() > tol])
    rf = "Rolling Force (x10 kN/mm)"
    if {rf, "SKP_R_ROLL_LOAD", "INLET_COIL_WH"} <= set(df.columns):
        calc = df["SKP_R_ROLL_LOAD"] / df["INLET_COIL_WH"] / 10
        badm = (calc - df[rf]).abs() > 1e-6 * calc.abs().clip(lower=1)
        add("Rumus rolling force", int(badm.sum()), "roll load / lebar / 10", df.loc[badm, [ID_COL, rf]])
    for name, a_, b_ in DELTA_COLUMNS:
        if {name, a_, b_} <= set(df.columns):
            badm = ((df[a_] - df[b_]) - df[name]).abs() > 1e-6
            add(f"Rumus {label(name)}", int(badm.sum()), f"{label(a_)} − {label(b_)}", df.loc[badm, [ID_COL, name]])
    vis = [c for c in df.columns if not c.startswith("_")]
    empty_cols = [c for c in vis if df[c].isna().all()]
    add("Kolom kosong", len(empty_cols), ", ".join(empty_cols) or "-", info=True)

    st.dataframe(pd.DataFrame(checks), width="stretch", hide_index=True)
    for name, frame in details.items():
        with st.expander(f"{name} ({len(frame)})"):
            st.dataframe(frame, width="stretch", hide_index=True)


def section_events():
    ss = st.session_state
    ss.setdefault("events_ver", 0)
    panel_params = [p["Parameter"] for p in load_panels()]
    opts = list(dict.fromkeys(["Semua"] + panel_params + EVENTS["Panel"].dropna().tolist()))
    if "events_base" not in ss:
        ev0 = prepare_events(ss["events_saved"])
        ev0["Tanggal"] = ev0["Tanggal"].dt.date
        ss["events_base"] = ev0.sort_values("Tanggal", ascending=False).reset_index(drop=True)
    st.caption(f"{len(EVENTS)} event tersimpan di database. Tambah baris di bawah tabel, "
               "atau pilih baris lalu tekan Delete untuk menghapus. Perubahan tersimpan otomatis.")
    edited = st.data_editor(
        ss["events_base"], key=f"events_editor_{ss['events_ver']}", num_rows="dynamic", hide_index=True,
        width="stretch",
        column_config={
            "Tanggal": st.column_config.DateColumn("Tanggal", format="DD/MM/YYYY", required=True),
            "Kejadian": st.column_config.TextColumn("Kejadian", required=True, width="large"),
            "Panel": st.column_config.SelectboxColumn("Tampil di panel", options=opts, default="Semua",
                                                      help="'Semua' atau kode parameter panel"),
        },
    )
    records = [r for r in clean_records(edited, ["Tanggal", "Kejadian", "Panel"]) if r.get("Tanggal")]
    saved = clean_records(prepare_events(ss["events_saved"]), ["Tanggal", "Kejadian", "Panel"])
    if sorted(records, key=str) != sorted(saved, key=str):
        save_events(records)
        ss["events_saved"] = load_events()
        st.rerun()

    a, b = st.columns(2)
    with a:
        up = st.file_uploader("Import event dari Excel/CSV (kolom: Tanggal, Kejadian, Panel)",
                              type=["xlsx", "xls", "csv"], key="ev_up")
        mode = st.radio("Cara import", ["Tambahkan", "Ganti semua"], horizontal=True, key="ev_mode")
        if up is not None and st.button("Import event", key="ev_import"):
            try:
                raw = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                raw = clean_column_names(raw)
                if not {"Tanggal", "Kejadian"} <= set(raw.columns):
                    st.error("Kolom wajib: Tanggal dan Kejadian.")
                else:
                    if "Panel" not in raw.columns:
                        raw["Panel"] = "Semua"
                    raw["Tanggal"] = pd.to_datetime(raw["Tanggal"], dayfirst=True, errors="coerce")
                    new = prepare_events(clean_records(raw, ["Tanggal", "Kejadian", "Panel"]))
                    merged = new if mode == "Ganti semua" else pd.concat([EVENTS, new])
                    merged = merged.drop_duplicates().sort_values("Tanggal")
                    save_events(clean_records(merged, ["Tanggal", "Kejadian", "Panel"]))
                    ss["events_saved"] = load_events()
                    ss.pop("events_base", None)
                    ss["events_ver"] += 1
                    st.rerun()
            except Exception as e:
                st.error(f"File event tidak bisa dibaca: {e}")
    with b:
        exp = EVENTS.copy()
        exp["Tanggal"] = exp["Tanggal"].dt.date
        if exp.empty:
            exp = pd.DataFrame({"Tanggal": [datetime.now().date()], "Kejadian": ["Contoh: Clean SS Multi"],
                                "Panel": ["Semua"]})
        lazy_download(st, "⬇️ Download event / template", lambda: to_excel_bytes((("Event", exp),)),
                      "event_maintenance.xlsx", XLSX_MIME, "dl_events")


def section_export():
    if df is None:
        empty_state()
        return
    view = export_view(filtered)
    st.write(f"{len(view):,} coil sesuai filter di sidebar.")
    st.dataframe(view.head(500), width="stretch", height=360, hide_index=True)
    if len(view) > 500:
        st.caption("Pratinjau 500 baris pertama. File unduhan berisi semua baris.")
    cursor_df = pd.DataFrame([{"Cursor": f"C{c['id']}", "Waktu": cursor_time(c), "Coil": c.get("coil"),
                               "Parameter": label(c.get("param")), "Nilai": c.get("value"),
                               "Catatan": c.get("note", "")} for c in st.session_state["cursors"]])
    lim_df = pd.DataFrame([{"Parameter": label(k), "Kode": k, "Batas bawah": v.get("lower"),
                            "Batas atas": v.get("upper")} for k, v in LIMITS.items()])
    a, b = st.columns(2)
    lazy_download(a, "⬇️ CSV", lambda: view.to_csv(index=False).encode("utf-8-sig"),
                  "coil_data.csv", "text/csv", "dl_csv")
    lazy_download(b, "⬇️ Excel (data + cursor + event + batas)",
                  lambda: to_excel_bytes((("Data", view), ("Cursor", cursor_df), ("Event", EVENTS),
                                          ("Batas", lim_df))), "coil_data.xlsx", XLSX_MIME, "dl_xlsx")


def section_display():
    order = st.selectbox(
        "Cara membaca tanggal UPDAT_DT", DATE_ORDERS, index=DATE_ORDERS.index(DATE_ORDER),
        help="Otomatis: tanggal ambigu (mis. 6/12/2026) dicocokkan dengan SHEAR_CUT_DT/WORK_DT pada baris "
             "yang sama. Pilih manual hanya jika semua file memakai satu format yang pasti.")
    if order != DATE_ORDER:
        set_setting("date_order", order)
        st.rerun()
    st.caption("Mengubah cara baca tidak perlu import ulang: teks asli UPDAT_DT tersimpan di database.")
    st.divider()
    opts = [c for c in df.columns if not c.startswith("_") and c != ID_COL] if df is not None else []
    cur = [c for c in HOVER_COLS if c in opts]
    pick = st.multiselect("Informasi di tooltip grafik", opts, default=cur, format_func=label_code,
                          help="Muncul saat kursor diarahkan ke titik grafik. Lebih sedikit = grafik lebih ringan.")
    if pick != cur:
        set_setting("hover_cols", pick)
        st.rerun()
    n = st.slider("Jumlah coil untuk garis tren halus", 5, 200, SMOOTH_N, step=5)
    if n != SMOOTH_N:
        set_setting("smooth_window", n)
        st.rerun()


def section_database():
    st.caption(f"File database: {DB_PATH}")
    with closing(db_connect()) as con:
        counts = {t: con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                  for t in ["cursors", "events", "chart_panels", "process_limits", "import_log"]}
        log = pd.read_sql_query("SELECT id AS 'No', imported_at AS 'Waktu', file_name AS 'File', "
                                "rows_in_file AS 'Baris', n_new AS 'Baru', n_changed AS 'Diperbarui', "
                                "n_same AS 'Sama', new_columns AS 'Kolom baru', note AS 'Catatan' "
                                "FROM import_log ORDER BY id DESC LIMIT 200", con)
    m = st.columns(5)
    m[0].metric("Coil", f"{n_db:,}")
    m[1].metric("Cursor", counts["cursors"])
    m[2].metric("Event", counts["events"])
    m[3].metric("Panel", counts["chart_panels"])
    m[4].metric("Import", counts["import_log"])
    st.markdown("##### Riwayat import")
    st.dataframe(log, width="stretch", hide_index=True)
    a, b = st.columns(2)
    lazy_download(a, "⬇️ Backup database (.db)", database_backup_bytes,
                  f"coil_cgl_backup_{datetime.now():%Y%m%d_%H%M}.db", "application/octet-stream", "dl_db")
    with b.popover("⚠️ Hapus semua data coil"):
        st.write("Data coil dan riwayat perubahan akan dihapus. Event, panel, dan batas tetap ada.")
        also = st.checkbox("Hapus juga semua cursor", key="del_cursor_too")
        if st.button("Ya, hapus data coil", type="primary", key="del_all_data"):
            delete_all_coil_data(also)
            st.session_state["cursors"] = load_cursors()
            st.rerun()


def page_data():
    st.header("Data & Pengaturan")
    sections = {
        "Import data": section_import, "Event maintenance": section_events, "Validasi data": section_validation,
        "Ekspor": section_export, "Tampilan": section_display, "Database": section_database,
    }
    view = sub_nav(list(sections.keys()), "data_view")
    sections[view]()


# =====================================================
# JALANKAN HALAMAN YANG DIPILIH
# =====================================================
if page == PAGE_DATA:
    page_data()
elif df is None or df.empty:
    empty_state()
elif filtered is None or filtered.empty:
    st.warning("Tidak ada coil pada periode/filter ini. Ubah filter di sidebar.")
else:
    {PAGE_HOME: page_home, PAGE_TREND: page_trend, PAGE_QUALITY: page_quality,
     PAGE_PROCESS: page_process, PAGE_COIL: page_coil}[page]()
