"""
Dashboard Monitoring Kualitas Produksi Rokok SKT - PT. XYZ
Six Sigma (DMAI): Define - Measure - Analyze - Improve

Prinsip visualisasi mengikuti Cole Nussbaumer Knaflic, "Storytelling with Data":
  1. Pilih visual yang tepat (bar horizontal terurut, garis untuk tren; tanpa pie/donut, tanpa sumbu ganda).
  2. Buang clutter (tanpa gridline, tanpa bingkai, tanpa legenda: label langsung pada data).
  3. Fokuskan perhatian (abu-abu sebagai dasar, satu warna aksen royal blue; merah hanya untuk peringatan).
  4. Judul berupa pesan/kesimpulan, subjudul menjelaskan apa yang ditampilkan.
"""
import io
import re
import copy
import hashlib
import textwrap
from datetime import date
from statistics import NormalDist

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ============================================================
# KONFIGURASI & PALET WARNA (royal blue + abu-abu)
# ============================================================
st.set_page_config(page_title="Dashboard Pengendalian Kualitas Proses Produksi Rokok SKT", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

NAVY = "#0E1F3F"        # navy gelap: sidebar & header
NAVY_DEEP = "#08142B"
BURGUNDY = "#7B1E2E"    # merah marun: tombol, kartu, penekanan antarmuka
GOLD = "#C9A96E"        # aksen emas lembut: garis tipis & judul frame
IVORY = "#FAF6EE"       # putih gading
ROYAL = "#16305F"       # (nama lama dipertahankan) satu-satunya warna aksen pada grafik: navy
ROYAL_LIGHT = "#F1E9DA"
INK = "#1F2937"         # teks utama
GRAY_DARK = "#4B5563"   # data utama non-sorotan
GRAY = "#9CA3AF"        # data pendukung
GRAY_LIGHT = "#D1D5DB"  # garis bantu
BG = IVORY
LINE = "#E7DFCF"        # garis tepi kartu
RED = "#C0392B"         # hanya untuk titik di luar batas kendali / peringatan

APP_TITLE = "Dashboard Pengendalian Kualitas Proses Produksi Rokok SKT"

st.markdown(f"""
<style>
  .stApp {{ background-color: {BG}; color: {INK}; }}
  [data-testid="stHeader"] {{ background: transparent; }}
  .block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; }}
  h1, h2, h3, h4 {{ color: {NAVY}; font-family: Georgia, 'Times New Roman', serif; }}

  /* ---------- SIDEBAR ---------- */
  [data-testid="stSidebar"] {{ background: linear-gradient(180deg, {NAVY} 0%, {NAVY_DEEP} 100%);
                               border-right: 1px solid rgba(201,169,110,0.35); }}
  [data-testid="stSidebar"] * {{ color: {IVORY} !important; }}
  [data-testid="stSidebar"] [data-baseweb="select"] * {{ color: {INK} !important; }}
  [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {{ color: {INK} !important; }}
  [data-testid="stSidebar"] [data-baseweb="input"] * {{ color: {INK} !important; }}
  [data-testid="stSidebar"] [data-testid="stSidebarContent"] {{ padding-top: 0.5rem; }}
  [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"],
  [data-testid="stSidebar"] [class*="st-key-frame_"] {{
      background: rgba(255,255,255,0.06) !important; border: 1px solid rgba(201,169,110,0.40) !important;
      border-radius: 12px !important; }}
  [data-testid="stSidebar"] [data-testid="stExpander"] details {{ background: rgba(255,255,255,0.06);
      border: 1px solid rgba(201,169,110,0.40); border-radius: 12px; }}
  [data-testid="stSidebar"] .sb-brand {{ border: 1px solid rgba(201,169,110,0.55); border-radius: 12px;
      padding: 16px 16px 14px 16px; margin-bottom: 4px;
      background: linear-gradient(135deg, rgba(123,30,46,0.55) 0%, rgba(255,255,255,0.05) 100%); }}
  [data-testid="stSidebar"] .sb-brand .sb-kicker {{ color: {GOLD} !important; font-size: 12px; font-weight: 700;
      letter-spacing: 3px; }}
  [data-testid="stSidebar"] .sb-brand .sb-title {{ color: {IVORY} !important; font-family: Georgia, serif;
      font-size: 16px; font-weight: 700; line-height: 1.35; margin-top: 6px; }}
  [data-testid="stSidebar"] .fr-title {{ color: {GOLD} !important; font-size: 12px; font-weight: 700;
      letter-spacing: 1.6px; text-transform: uppercase; margin: 0 0 6px 0; }}
  [data-testid="stSidebar"] [role="radiogroup"] label {{ padding: 6px 10px; border-radius: 8px; width: 100%; }}
  [data-testid="stSidebar"] [role="radiogroup"] label:hover {{ background: rgba(255,255,255,0.08); }}
  [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{ background: rgba(123,30,46,0.70); }}
  [data-testid="stSidebar"] .stDownloadButton button {{ background: {BURGUNDY}; border: 1px solid {GOLD};
      border-radius: 8px; }}
  [data-testid="stSidebar"] .stDownloadButton button:hover {{ background: #96283B; }}

  /* ---------- KARTU METRIK ---------- */
  div[data-testid="stMetric"] {{ background: #FFFFFF; border: 1px solid {LINE}; border-left: 4px solid {BURGUNDY};
                                 padding: 14px 16px; border-radius: 10px; }}
  div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {{ color: {GRAY_DARK};
      white-space: normal !important; overflow: visible !important; text-overflow: clip !important; }}
  div[data-testid="stMetricValue"], div[data-testid="stMetricValue"] * {{ color: {NAVY};
      white-space: normal !important; overflow: visible !important; text-overflow: clip !important; }}
  .kpi {{ background: #FFFFFF; border: 1px solid {LINE}; border-left: 4px solid {BURGUNDY}; border-radius: 10px;
          padding: 16px 18px; margin-bottom: 6px; }}
  .kpi .kpi-l {{ color: {GRAY_DARK}; font-size: 13px; line-height: 1.35; }}
  .kpi .kpi-v {{ color: {NAVY}; font-size: 30px; font-weight: 700; line-height: 1.2; margin-top: 4px;
                 font-family: Georgia, serif; }}
  .kpi .kpi-s {{ color: {GRAY}; font-size: 12px; margin-top: 2px; }}
  .hero {{ background: linear-gradient(135deg, {NAVY} 0%, {NAVY_DEEP} 100%); border: 1px solid {GOLD};
           border-radius: 14px; padding: 22px 28px; margin-bottom: 14px; }}
  .hero .hero-l {{ color: {GOLD}; font-size: 13px; font-weight: 700; letter-spacing: 3px; }}
  .hero .hero-v {{ color: {IVORY}; font-family: Georgia, serif; font-size: 64px; font-weight: 700; line-height: 1.1;
                   margin: 6px 0 2px 0; }}
  .hero .hero-s {{ color: #E8DFC9; font-size: 15px; }}

  /* ---------- KARTU, CATATAN, SIPOC ---------- */
  .card {{ background: #FFFFFF; border: 1px solid {LINE}; border-top: 3px solid {GOLD}; border-radius: 10px;
           padding: 16px 18px; height: 100%; }}
  .card .num {{ color: {BURGUNDY}; font-weight: 700; font-size: 13px; letter-spacing: 1px; }}
  .card .head {{ color: {NAVY}; font-weight: 700; font-size: 17px; margin: 4px 0 6px 0; font-family: Georgia, serif; }}
  .card .body {{ color: {GRAY_DARK}; font-size: 14px; line-height: 1.5; }}
  .note {{ background: {ROYAL_LIGHT}; border-left: 4px solid {BURGUNDY}; padding: 10px 14px; border-radius: 6px;
           color: {INK}; font-size: 14px; margin-bottom: 10px; }}
  .sipoc-h {{ background: {NAVY}; color: #FFFFFF; font-weight: 700; padding: 8px 12px; border-radius: 8px 8px 0 0;
              border-bottom: 2px solid {GOLD}; }}
  .sipoc-b {{ background: #FFFFFF; border: 1px solid {LINE}; border-top: 0; padding: 10px 12px;
              border-radius: 0 0 8px 8px; font-size: 14px; color: {INK}; min-height: 290px; }}
  button[data-baseweb="tab"][aria-selected="true"] {{ color: {BURGUNDY}; }}
  [data-baseweb="tab-highlight"] {{ background-color: {BURGUNDY}; }}
</style>
""", unsafe_allow_html=True)

st.markdown(
    f"<div style='background:linear-gradient(120deg,{NAVY} 0%,{NAVY_DEEP} 62%,{BURGUNDY} 160%);"
    f"padding:22px 28px;border-radius:14px;margin-bottom:18px;border-bottom:3px solid {GOLD}'>"
    f"<div style='color:{GOLD};font-size:12px;letter-spacing:3px;font-weight:700'>QUALITY CONTROL · SIX SIGMA DMAIC</div>"
    f"<div style='color:{IVORY};font-family:Georgia,serif;font-size:28px;font-weight:700;line-height:1.3;margin-top:4px'>"
    f"{APP_TITLE}</div>"
    f"<div style='color:#D9D2C0;font-size:14px;margin-top:2px'>Pengendalian kualitas Sigaret Kretek Tangan · PT. XYZ</div></div>",
    unsafe_allow_html=True)

# ============================================================
# STRUKTUR DATA
# ============================================================
NORM = NormalDist()
DATE, N = "Tanggal", "Jumlah Diperiksa"
DEFECTS = ["Gembos", "Lem Ambri Tengah", "Lem Ambri Ujung", "Guntingan", "Korep", "Besar Kepala"]
D, P = "Jumlah Cacat", "Proporsi Cacat"
FMEA_COLS = ["No", "Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan", "S", "O", "D",
             "Usulan Perbaikan"]
FAKTOR_LIST = ["Manusia (Man)", "Mesin (Machine)", "Metode (Method)", "Material", "Pengukuran (Measurement)",
               "Lingkungan (Environment)"]

# Data bawaan: data pemeriksaan QA PT. XYZ, 30 Maret - 30 Juni 2026 (66 subgrup), sumber Proyek Akhir.
EMBEDDED_CSV = """Tanggal,Jumlah Diperiksa,Gembos,Lem Ambri Tengah,Lem Ambri Ujung,Guntingan,Korep,Besar Kepala
2026-03-30,350,1,11,6,2,3,3
2026-03-31,260,1,7,1,2,1,1
2026-04-01,500,2,8,5,3,1,4
2026-04-02,300,1,11,6,1,0,1
2026-04-03,500,1,12,2,1,2,2
2026-04-06,50,2,2,2,0,0,0
2026-04-07,100,3,1,0,2,0,1
2026-04-08,200,2,0,4,1,1,4
2026-04-09,300,1,8,8,0,1,2
2026-04-10,150,1,2,2,1,0,3
2026-04-13,700,0,14,2,2,3,3
2026-04-14,300,1,8,8,0,1,2
2026-04-15,100,2,1,1,0,0,2
2026-04-16,100,1,4,0,1,0,0
2026-04-17,150,0,5,4,1,0,0
2026-04-20,300,1,8,8,0,1,2
2026-04-21,500,3,16,2,5,1,7
2026-04-22,500,1,12,2,1,2,2
2026-04-23,500,1,13,5,1,1,4
2026-04-24,700,1,6,11,2,5,5
2026-04-27,800,11,15,17,2,3,10
2026-04-28,500,9,8,4,2,1,4
2026-04-29,500,6,8,8,2,4,4
2026-04-30,350,4,3,8,2,1,5
2026-05-04,550,3,11,6,2,8,3
2026-05-05,550,3,13,8,1,4,7
2026-05-06,550,2,14,6,2,4,6
2026-05-07,600,5,13,5,2,6,4
2026-05-08,800,8,13,7,2,9,7
2026-05-09,500,4,11,2,6,3,5
2026-05-11,700,2,19,9,5,2,4
2026-05-12,650,3,21,12,4,2,1
2026-05-13,650,8,12,7,3,7,2
2026-05-14,700,2,17,8,5,6,3
2026-05-15,650,1,19,11,2,1,5
2026-05-18,800,9,24,23,2,2,10
2026-05-19,800,11,12,14,5,2,7
2026-05-20,400,5,3,6,1,2,3
2026-05-21,500,6,3,5,1,4,8
2026-05-22,100,6,8,12,2,5,6
2026-05-25,600,7,3,12,1,2,4
2026-05-26,700,0,14,2,2,3,3
2026-05-27,700,1,10,0,3,0,6
2026-05-28,700,0,4,6,1,4,0
2026-05-29,700,4,5,10,3,0,0
2026-06-02,400,5,15,8,2,2,3
2026-06-03,100,6,6,4,3,0,5
2026-06-04,300,0,2,2,3,3,7
2026-06-05,200,0,10,3,1,1,6
2026-06-08,500,7,8,13,4,4,8
2026-06-09,100,3,11,6,4,0,1
2026-06-10,100,7,11,8,0,0,1
2026-06-11,600,3,1,1,2,2,3
2026-06-12,700,0,22,10,0,3,3
2026-06-15,800,7,19,1,0,2,5
2026-06-16,300,1,23,9,2,0,0
2026-06-17,200,7,7,3,1,5,8
2026-06-18,500,3,12,9,4,4,2
2026-06-19,600,7,18,8,2,5,6
2026-06-22,300,3,22,5,4,2,6
2026-06-23,700,4,6,9,4,2,6
2026-06-24,700,1,6,11,2,5,5
2026-06-25,700,6,13,8,2,5,1
2026-06-26,600,4,13,2,1,1,0
2026-06-29,300,5,6,5,3,0,0
2026-06-30,300,3,9,3,3,4,1"""

EMBEDDED_FMEA = [
    {
        "No": 1,
        "Faktor": "Manusia (Man)",
        "Penyebab Kegagalan": "Ketelitian pekerja berbeda-beda dan pekerja dapat mengalami kelelahan selama proses produksi.",
        "Potensi Kegagalan": "Pelintingan rokok tidak konsisten.",
        "Dampak Kegagalan": "Bentuk rokok tidak seragam dan berpotensi tidak memenuhi standar.",
        "S": 6,
        "O": 6,
        "D": 5,
        "Usulan Perbaikan": "Memberikan pengarahan tentang standar pelintingan, mengatur waktu istirahat, dan melakukan pengawasan berkala."
    },
    {
        "No": 2,
        "Faktor": "Mesin (Machine)",
        "Penyebab Kegagalan": "Alat pemotong kurang tajam atau kondisi peralatan kurang baik.",
        "Potensi Kegagalan": "Pemotongan ujung rokok tidak seragam.",
        "Dampak Kegagalan": "Ujung rokok tidak sesuai standar sehingga produk berpotensi menjadi cacat.",
        "S": 6,
        "O": 5,
        "D": 5,
        "Usulan Perbaikan": "Memeriksa kondisi alat sebelum digunakan dan melakukan perawatan serta penggantian alat pemotong bila diperlukan."
    },
    {
        "No": 3,
        "Faktor": "Metode (Method)",
        "Penyebab Kegagalan": "Prosedur pelintingan dan pemberian lem belum diterapkan secara konsisten.",
        "Potensi Kegagalan": "Pemberian lem pada ambri tidak merata.",
        "Dampak Kegagalan": "Ambri tidak merekat dengan baik sehingga menimbulkan kecacatan lem ambri tengah atau lem ambri ujung.",
        "S": 7,
        "O": 8,
        "D": 6,
        "Usulan Perbaikan": "Menetapkan cara pemberian lem yang seragam, memberikan arahan kepada pekerja, dan memeriksa hasil perekatan secara berkala."
    },
    {
        "No": 4,
        "Faktor": "Material",
        "Penyebab Kegagalan": "Kadar air tembakau, kualitas kertas ambri, atau daya rekat lem tidak sesuai standar.",
        "Potensi Kegagalan": "Bahan baku atau lem tidak sesuai standar.",
        "Dampak Kegagalan": "Hasil pelintingan atau perekatan ambri berpotensi tidak memenuhi standar kualitas.",
        "S": 7,
        "O": 6,
        "D": 6,
        "Usulan Perbaikan": "Memeriksa kondisi bahan baku sebelum digunakan dan memastikan kualitas lem serta kertas ambri sesuai standar."
    },
    {
        "No": 5,
        "Faktor": "Pengukuran (Measurement)",
        "Penyebab Kegagalan": "Pemeriksaan kurang teliti dan standar kecacatan belum dipahami secara seragam.",
        "Potensi Kegagalan": "Produk cacat tidak terdeteksi saat pemeriksaan.",
        "Dampak Kegagalan": "Produk yang tidak sesuai standar dapat lolos ke proses berikutnya.",
        "S": 7,
        "O": 5,
        "D": 7,
        "Usulan Perbaikan": "Menyeragamkan kriteria kecacatan, menggunakan panduan pemeriksaan, dan melakukan pemeriksaan ulang secara berkala."
    },
    {
        "No": 6,
        "Faktor": "Lingkungan (Environment)",
        "Penyebab Kegagalan": "Suhu dan kelembapan berubah-ubah serta pencahayaan ruang produksi kurang memadai.",
        "Potensi Kegagalan": "Kondisi lingkungan kerja kurang sesuai.",
        "Dampak Kegagalan": "Kondisi bahan atau hasil produksi berpotensi menjadi tidak konsisten dan cacat lebih sulit terlihat saat pemeriksaan.",
        "S": 5,
        "O": 4,
        "D": 5,
        "Usulan Perbaikan": "Memantau kondisi suhu dan kelembapan serta memastikan pencahayaan area kerja memadai."
    }
]

# Diagram sebab-akibat (Fishbone) hasil penelitian: 6 kategori (5M+1E), 18 penyebab.
# Redaksi mengikuti gambar Fishbone penelitian. Urutan kategori: 3 cabang atas (kiri ke kanan), lalu 3 cabang bawah.
FISHBONE_PROBLEM = "Proses Produksi Rokok Sigaret Kretek Tangan (SKT) Belum Kapabel di PT. XYZ"
FISHBONE_CATS = ["Measurement", "Machine", "Man", "Material", "Environment", "Method"]
FISHBONE_RESEARCH = {
    "Measurement": ["Pemeriksaan produk kurang teliti",
                    "Standar Kualitas belum dipahami secara seragam",
                    "Produk Cacat tidak teridentifikasi saat pemeriksaan"],
    "Machine": ["Alat pemotong kurang tajam",
                "Kondisi peralatan kurang baik",
                "Peralatan tidak sesuai dengan kebutuhan produksi"],
    "Man": ["Pekerja kurang teliti saat melinting rokok",
            "Keterampilan pekerja berbeda-beda",
            "Pekerja mengalami kelelahan saat produksi"],
    "Material": ["Kadar air tembakau tidak sesuai",
                 "Kualitas kertas ambri kurang baik",
                 "Daya rekat lem kurang optimal"],
    "Environment": ["Suhu ruang produksi tidak sesuai",
                    "Kelembapan ruangan berubah-ubah",
                    "Pencahayaan area kerja kurang memadai"],
    "Method": ["Teknik pelintingan tidak konsisten",
               "Pemotongan ujung rokok tidak seragam",
               "Pemberian lem pada ambri tidak merata"],
}

SIPOC = {
    "Supplier": ["Pemasok bahan baku dan bahan pendukung produksi"],
    "Input": ["Bahan baku: tembakau, cengkeh, dan bahan tambahan sesuai formulasi perusahaan",
              "Bahan produksi: kertas rokok",
              "Bahan pengemasan: ambri, plastik pembungkus, lem, etiket, pita cukai, label, kertas bale, "
              "segel bale, karton box"],
    "Process": ["Persiapan produksi", "Pencampuran tembakau, cengkeh, dan bahan tambahan",
                "Pembentukan & pemotongan batang (giling-gunting)", "Pengeringan batang (oven batangan)",
                "Pemeriksaan kualitas (QA)", "Pembungkusan bagian dalam (inner wrap)", "Pengemasan (pack)",
                "Pemasangan pita cukai (bandrol)"],
    "Output": ["Produk rokok SKT yang memenuhi standar kualitas dan siap didistribusikan"],
    "Customer": ["Gudang produk jadi", "Konsumen"],
}

# Nilai acuan dari workbook Excel olah data terbaru (untuk verifikasi bila memakai data bawaan)
NASKAH = {"p̄ peta kendali awal": 0.060470, "Subgrup di luar batas": 12, "p̄ peta kendali revisi": 0.057468,
          "Statistik uji Z (runs test)": 0.150684, "p-value runs test": 0.880225, "Indeks kapabilitas": 0.525463,
          "DPMO": 60470.434, "Level sigma": 3.050837}


class WorkbookError(Exception):
    """Pesan kesalahan yang aman ditampilkan kepada pengguna."""


# ============================================================
# FORMAT ANGKA (gaya Indonesia)
# ============================================================
def fmt_int(x):
    return "—" if x is None or pd.isna(x) else f"{x:,.0f}".replace(",", ".")


def fmt_dec(x, n=3):
    if x is None or pd.isna(x):
        return "—"
    return f"{x:,.{n}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def fmt_pct(x, n=1):
    return "—" if x is None or pd.isna(x) else fmt_dec(100 * x, n) + "%"


# ============================================================
# PEMBACAAN DATA
# ============================================================
def norm(text):
    t = re.sub(r"\(.*?\)", "", str(text).replace("\xa0", " "))
    return re.sub(r"\s+", " ", t).strip().lower().rstrip(".")


DATA_ALIASES = {"tanggal": DATE, "day": DATE, "jumlah diperiksa": N, "total sampel": N,
                **{norm(c): c for c in DEFECTS}}
FMEA_ALIASES = {"no": "No", "faktor": "Faktor", "penyebab kegagalan": "Penyebab Kegagalan",
                "potensi kegagalan": "Potensi Kegagalan", "dampak kegagalan": "Dampak Kegagalan",
                "s": "S", "o": "O", "d": "D", "usulan perbaikan": "Usulan Perbaikan"}


def table_from_sheet(raw, aliases, must_have, sheet):
    for i in range(min(len(raw), 20)):
        cells = {norm(v) for v in raw.iloc[i].tolist() if pd.notna(v)}
        if all(m in cells for m in must_have):
            names = [aliases.get(norm(v)) if pd.notna(v) else None for v in raw.iloc[i].tolist()]
            keep, seen = [], set()
            for j, name in enumerate(names):
                if name and name not in seen:
                    keep.append(j)
                    seen.add(name)
            df = raw.iloc[i + 1:, keep].copy()
            df.columns = [names[j] for j in keep]
            return df.reset_index(drop=True)
    raise WorkbookError(f"Sheet '{sheet}': baris judul kolom tidak ditemukan. Pastikan memakai template "
                        f"dashboard (kolom {', '.join(must_have)}).")


def to_date(v):
    if isinstance(v, (pd.Timestamp, date)):
        return pd.Timestamp(v)
    if v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() == "":
        return pd.NaT
    txt = str(v).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}( .*)?", txt):  # format ISO (yyyy-mm-dd)
        return pd.to_datetime(txt[:10], format="%Y-%m-%d", errors="coerce")
    return pd.to_datetime(txt, dayfirst=True, errors="coerce")  # format Indonesia (dd/mm/yyyy)


def prepare_data(df, strict=True):
    """Bersihkan data harian, hitung d dan p. Mengembalikan (data, daftar peringatan)."""
    missing = [c for c in [DATE, N, *DEFECTS] if c not in df.columns]
    if missing:
        raise WorkbookError(f"Kolom berikut tidak ditemukan: {', '.join(missing)}.")
    df = df[[DATE, N, *DEFECTS]].copy()
    df[DATE] = df[DATE].map(to_date)
    for c in [N, *DEFECTS]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    blank = df[DATE].isna() & df[[N, *DEFECTS]].isna().all(axis=1)
    df = df[~blank]
    warns = []
    bad_date = df[DATE].isna()
    if bad_date.any() and strict:
        warns.append(f"{int(bad_date.sum())} baris tanpa tanggal yang valid diabaikan.")
    df = df[~bad_date]
    bad_n = df[N].isna() | (df[N] <= 0)
    if bad_n.any():
        warns.append("Baris dengan Jumlah Diperiksa kosong/nol diabaikan: "
                     + ", ".join(df.loc[bad_n, DATE].dt.strftime("%d/%m/%Y")) + ".")
    df = df[~bad_n]
    df[DEFECTS] = df[DEFECTS].fillna(0)
    neg = (df[[N, *DEFECTS]] < 0).any(axis=1)
    if neg.any():
        raise WorkbookError("Terdapat nilai negatif pada tanggal: "
                            + ", ".join(df.loc[neg, DATE].dt.strftime("%d/%m/%Y")) + ".")
    df = df.sort_values(DATE).reset_index(drop=True)
    df[D] = df[DEFECTS].sum(axis=1)
    over = df[D] > df[N]
    if over.any():
        raise WorkbookError("Jumlah cacat melebihi jumlah diperiksa pada tanggal: "
                            + ", ".join(df.loc[over, DATE].dt.strftime("%d/%m/%Y")) + ".")
    dup = df[DATE].duplicated(keep=False)
    if dup.any():
        warns.append("Ada tanggal ganda (tetap dihitung sebagai subgrup terpisah): "
                     + ", ".join(sorted(set(df.loc[dup, DATE].dt.strftime("%d/%m/%Y")))) + ".")
    if df.empty:
        raise WorkbookError("Tidak ada baris data yang valid pada sheet Data Harian.")
    df.insert(0, "Subgrup", np.arange(1, len(df) + 1))
    df[P] = df[D] / df[N]
    return df, warns


def prepare_fmea(df):
    for c in FMEA_COLS:
        if c not in df.columns:
            df[c] = np.nan if c in ("No", "S", "O", "D") else ""
    df = df[FMEA_COLS].copy()
    for c in ["No", "S", "O", "D"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df["Faktor"].notna() & (df["Faktor"].astype(str).str.strip() != "")]
    df = df[df[["S", "O", "D"]].notna().all(axis=1)].copy()
    if df.empty:
        return None
    df["RPN"] = df["S"] * df["O"] * df["D"]
    for c in ["Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan", "Usulan Perbaikan"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    return df.reset_index(drop=True)


def embedded():
    data = pd.read_csv(io.StringIO(EMBEDDED_CSV))
    data, _ = prepare_data(data)
    return data, prepare_fmea(pd.DataFrame(EMBEDDED_FMEA)), []


@st.cache_data(show_spinner=False)
def read_upload(raw_bytes):
    try:
        xls = pd.ExcelFile(io.BytesIO(raw_bytes), engine="openpyxl")
    except Exception as exc:
        raise WorkbookError("File tidak dapat dibuka sebagai workbook Excel (.xlsx). "
                            f"Detail: {type(exc).__name__}: {exc}")
    by_norm = {norm(s): s for s in xls.sheet_names}
    notes = []
    if "data harian" in by_norm:
        sheet = by_norm["data harian"]
    elif "peta kendali 1" in by_norm:  # format workbook olah data lama
        sheet = by_norm["peta kendali 1"]
        notes.append("Workbook format lama terdeteksi; data dibaca dari sheet 'Peta Kendali 1'.")
    else:
        raise WorkbookError("Sheet 'Data Harian' tidak ditemukan. Sheet yang ada: "
                            + ", ".join(repr(s) for s in xls.sheet_names) + ".")
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)
    legacy = "data harian" not in by_norm
    data, warns = prepare_data(table_from_sheet(raw, DATA_ALIASES, ["day" if legacy else "tanggal", "gembos"], sheet),
                               strict=not legacy)
    fmea = None
    if "fmea" in by_norm:
        rawf = pd.read_excel(xls, sheet_name=by_norm["fmea"], header=None)
        try:
            fmea = prepare_fmea(table_from_sheet(rawf, FMEA_ALIASES, ["faktor", "s", "o", "d"], "FMEA"))
        except WorkbookError as exc:
            notes.append(str(exc))
    return data, fmea, notes + warns


# ============================================================
# PERHITUNGAN STATISTIK
# ============================================================
def with_limits(df):
    """Batas kendali peta P dengan ukuran sampel bervariasi (Persamaan 2.2-2.5)."""
    out = df.copy()
    pbar = out[D].sum() / out[N].sum()
    se = np.sqrt(pbar * (1 - pbar) / out[N])
    out["GT"] = pbar
    out["BKA"] = pbar + 3 * se
    out["BKB"] = (pbar - 3 * se).clip(lower=0)
    out["Status"] = np.where(out[P] > out["BKA"], "Di atas BKA",
                             np.where(out[P] < out["BKB"], "Di bawah BKB", "Terkendali"))
    return out, pbar


def auto_exclude(df, above_only=False, max_iter=25):
    """Keluarkan titik di luar batas secara berulang sampai seluruh titik terkendali."""
    kept, excluded = df.copy(), []
    for _ in range(max_iter):
        lim, _ = with_limits(kept)
        bad = lim["Status"].eq("Di atas BKA") if above_only else lim["Status"].ne("Terkendali")
        if not bad.any() or bad.sum() >= len(kept) - 1:
            break
        excluded += lim.loc[bad, "Subgrup"].tolist()
        kept = kept[~kept["Subgrup"].isin(excluded)]
    return sorted(excluded)


def run_test(df):
    """Runs test terhadap median (Persamaan 2.6-2.8), mengikuti workbook Excel (sheet Uji Keacakan).
    Proporsi dan median dibulatkan 4 desimal; p > median = 'Atas', p < median = 'Bawah';
    p = median TIDAK diikutsertakan. Jumlah run V dihitung dari urutan titik yang tersisa.
    Dipanggil dengan data peta P revisi (setelah subgrup di luar batas dikeluarkan)."""
    d = df.sort_values("Subgrup")
    raw = d[P].astype(float).to_numpy()
    if len(raw) < 3:
        return None
    med = float(np.median(raw))
    pr, mr = np.round(raw, 4), round(med, 4)
    state = np.where(pr > mr, 1, np.where(pr < mr, -1, 0))  # 1 = atas, -1 = bawah, 0 = sama dengan median
    seq = state[state != 0] > 0
    n1, n2 = int(seq.sum()), int((~seq).sum())
    if n1 == 0 or n2 == 0 or n1 + n2 < 3:
        return None
    v = 1 + int((seq[1:] != seq[:-1]).sum())
    n = n1 + n2
    mu = 2 * n1 * n2 / n + 1
    sd = np.sqrt(2 * n1 * n2 * (2 * n1 * n2 - n) / (n ** 2 * (n - 1)))
    z = (v - mu) / sd
    return {"median": med, "n1": n1, "n2": n2, "V": v, "mu": float(mu), "sigma": float(sd), "z": float(z),
            "p": float(2 * (1 - NORM.cdf(abs(z)))), "state": state, "ties": int((state == 0).sum()),
            "subgrup": d["Subgrup"].to_numpy(), "y": raw}


def capability(df):
    """Indeks kapabilitas atribut = Φ⁻¹(1 − p̄) / 3 (Persamaan 2.9)."""
    if df.empty:
        return None
    pbar = df[D].sum() / df[N].sum()
    if not 0 < pbar < 1:
        return None
    z = NORM.inv_cdf(1 - pbar)
    return {"pbar": pbar, "z": z, "index": z / 3}


def six_sigma(df):
    """DPU, DPO, DPMO, dan level sigma = Φ⁻¹(1 − DPO) + 1,5 (Persamaan 2.10-2.13).
    Memakai 1 opportunity (DPO = DPU), sesuai workbook Excel (sheet Level Sigma)."""
    if df.empty or not df[N].sum():
        return None
    dpu = df[D].sum() / df[N].sum()
    dpo = dpu  # 1 opportunity
    sigma = NORM.inv_cdf(1 - dpo) + 1.5 if 0 < dpo < 1 else np.nan
    return {"dpu": dpu, "dpo": dpo, "dpmo": dpo * 1e6, "sigma": sigma}


def sigma_category(s):
    if s is None or pd.isna(s):
        return "—"
    if s < 2:
        return "Sangat tidak kompetitif"
    if s < 4:
        return "Rata-rata industri Indonesia"
    if s < 6:
        return "Rata-rata industri USA"
    return "Industri kelas dunia"


def pareto_table(df):
    tot = df[DEFECTS].sum().sort_values(ascending=False)
    out = pd.DataFrame({"Jenis Cacat": tot.index, "Frekuensi": tot.values.astype(int)})
    grand = out["Frekuensi"].sum()
    out["Persentase"] = out["Frekuensi"] / grand if grand else 0.0
    out["Kumulatif"] = out["Persentase"].cumsum()
    return out


def join_id(items):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " dan " + items[-1]


# ============================================================
# TEMPLATE EXCEL
# ============================================================
def build_template(data=None, fmea=None, blank_rows=300):
    navy = PatternFill("solid", fgColor=NAVY[1:])
    white_b = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D1D5DB")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Petunjuk"
    lines = [
        ("TEMPLATE DATA DASHBOARD KUALITAS ROKOK SKT", True),
        ("", False),
        ("Cara mengisi:", True),
        ("1. Sheet 'Data Harian': satu baris = satu hari pemeriksaan (satu subgrup).", False),
        ("2. Isi Tanggal, Jumlah Diperiksa (n), dan jumlah rokok cacat untuk tiap jenis cacat.", False),
        ("3. Setiap rokok yang ditolak dicatat pada SATU jenis cacat utama saja, sehingga "
         "Jumlah Cacat = jumlah rokok cacat.", False),
        ("4. Kolom berwarna abu-abu (Jumlah Cacat, Proporsi Cacat, RPN) terisi otomatis. Jangan diubah.", False),
        ("5. Sheet 'FMEA': isi faktor penyebab, mode kegagalan, nilai S, O, D (skala 1-10), dan usulan perbaikan.", False),
        ("6. Jangan mengubah nama sheet atau judul kolom. Baris kosong boleh dibiarkan.", False),
        ("7. Simpan sebagai .xlsx, lalu unggah melalui sidebar dashboard (Sumber data > Unggah data perusahaan).", False),
        ("", False),
        ("Pedoman skala S, O, D mengikuti Tabel 2.3-2.5 pada naskah Proyek Akhir (AIAG, 2008).", False),
    ]
    for i, (t, b) in enumerate(lines, 1):
        ws.cell(i, 1, t).font = Font(bold=b, size=13 if i == 1 else 11, color=NAVY[1:] if b else "1F2937")
    ws.column_dimensions["A"].width = 110

    ws = wb.create_sheet("Data Harian")
    heads = ["Tanggal", "Jumlah Diperiksa (n)", *DEFECTS, "Jumlah Cacat (otomatis)", "Proporsi Cacat (otomatis)"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h)
        c.fill, c.font, c.alignment, c.border = navy, white_b, center, box
    ws.row_dimensions[1].height = 34
    widths = [13, 14, 10, 12, 12, 11, 9, 11, 14, 14]
    for j, wd in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + j)].width = wd
    n_data = 0 if data is None else len(data)
    for i in range(2, n_data + blank_rows + 2):
        if data is not None and i - 2 < n_data:
            r = data.iloc[i - 2]
            ws.cell(i, 1, r[DATE].to_pydatetime())
            ws.cell(i, 2, int(r[N]))
            for j, c in enumerate(DEFECTS, 3):
                ws.cell(i, j, int(r[c]))
        ws.cell(i, 9, f'=IF(B{i}="","",SUM(C{i}:H{i}))')
        ws.cell(i, 10, f'=IF(OR(B{i}="",B{i}=0),"",I{i}/B{i})')
        ws.cell(i, 1).number_format = "dd/mm/yyyy"
        ws.cell(i, 10).number_format = "0.00%"
        for j in range(1, 11):
            ws.cell(i, j).border = box
        ws.cell(i, 9).fill = ws.cell(i, 10).fill = PatternFill("solid", fgColor="F3F4F6")
    last = n_data + blank_rows + 1
    dv_num = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                            error="Isi bilangan bulat ≥ 0", errorTitle="Input tidak valid", showErrorMessage=True)
    dv_date = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                             error="Isi tanggal yang valid (dd/mm/yyyy)", showErrorMessage=True)
    ws.add_data_validation(dv_num)
    ws.add_data_validation(dv_date)
    dv_num.add(f"B2:H{last}")
    dv_date.add(f"A2:A{last}")
    ws.freeze_panes = "B2"

    ws = wb.create_sheet("FMEA")
    heads = ["No", "Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan", "S", "O", "D",
             "RPN (otomatis)", "Usulan Perbaikan"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h)
        c.fill, c.font, c.alignment, c.border = navy, white_b, center, box
    for col, wd in zip("ABCDEFGHIJ", [5, 22, 38, 30, 38, 5, 5, 5, 10, 42]):
        ws.column_dimensions[col].width = wd
    rows = [] if fmea is None else fmea.to_dict("records")
    for i in range(2, max(len(rows), 0) + 12):
        if i - 2 < len(rows):
            r = rows[i - 2]
            vals = [r["No"], r["Faktor"], r["Penyebab Kegagalan"], r["Potensi Kegagalan"], r["Dampak Kegagalan"],
                    r["S"], r["O"], r["D"]]
            for j, v in enumerate(vals, 1):
                ws.cell(i, j, int(v) if j in (1, 6, 7, 8) else v)
            ws.cell(i, 10, r["Usulan Perbaikan"])
        ws.cell(i, 9, f'=IF(COUNT(F{i}:H{i})=3,F{i}*G{i}*H{i},"")')
        ws.cell(i, 9).fill = PatternFill("solid", fgColor="F3F4F6")
        for j in range(1, 11):
            ws.cell(i, j).border = box
            ws.cell(i, j).alignment = wrap
    dv_sod = DataValidation(type="whole", operator="between", formula1="1", formula2="10", allow_blank=True,
                            error="Nilai S, O, D harus 1-10", showErrorMessage=True)
    dv_fak = DataValidation(type="list", formula1='"' + ",".join(FAKTOR_LIST) + '"', allow_blank=True)
    ws.add_data_validation(dv_sod)
    ws.add_data_validation(dv_fak)
    dv_sod.add("F2:H200")
    dv_fak.add("B2:B200")
    ws.freeze_panes = "C2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def template_blank():
    return build_template(None, None)


@st.cache_data(show_spinner=False)
def template_example():
    data, fm_, _ = embedded()
    return build_template(data, fm_)


# ============================================================
# GAYA GRAFIK (Storytelling with Data)
# ============================================================
def style(fig, title, subtitle=None, height=380, right=110):
    head = f"<b>{title}</b>"
    if subtitle:
        head += f"<br><span style='font-size:13px;color:{GRAY_DARK}'>{subtitle}</span>"
    fig.update_layout(
        title=dict(text=head, x=0, xanchor="left", y=0.97, yanchor="top", font=dict(size=17, color=INK)),
        height=height, plot_bgcolor="white", paper_bgcolor="white", showlegend=False,
        font=dict(color=GRAY_DARK, size=13), margin=dict(l=10, r=right, t=86, b=40),
        hoverlabel=dict(bgcolor="white", font=dict(color=INK)))
    fig.update_xaxes(showgrid=False, zeroline=False, showline=True, linecolor=GRAY_LIGHT, ticks="outside",
                     tickcolor=GRAY_LIGHT, title_font=dict(color=GRAY_DARK))
    fig.update_yaxes(showgrid=False, zeroline=False, showline=False, title_font=dict(color=GRAY_DARK))
    return fig


def show(fig):
    st.plotly_chart(fig, width="stretch", config={"displaylogo": False})


def end_label(fig, x, y, text, color):
    fig.add_annotation(x=x, y=y, text=text, showarrow=False, xanchor="left", xshift=8,
                       font=dict(color=color, size=12))


def fig_pchart(lim, pbar, title, subtitle, excluded=()):
    x = lim["Subgrup"]
    fig = go.Figure()
    for col in ("BKA", "BKB"):
        fig.add_trace(go.Scatter(x=x, y=lim[col], mode="lines", hoverinfo="skip",
                                 line=dict(color=GRAY_LIGHT, width=1.5, shape="hvh")))
    fig.add_trace(go.Scatter(x=x, y=lim["GT"], mode="lines", hoverinfo="skip",
                             line=dict(color=ROYAL, width=2, dash="dash")))
    fig.add_trace(go.Scatter(
        x=x, y=lim[P], mode="lines+markers", line=dict(color=GRAY_DARK, width=1.5),
        marker=dict(size=5, color=GRAY_DARK),
        customdata=list(zip(lim[DATE].dt.strftime("%d %b %Y"), lim[N].astype(int), lim[D].astype(int),
                            lim["BKA"].astype(float), lim["BKB"].astype(float))),
        hovertemplate="Subgrup %{x} · %{customdata[0]}<br>p = %{y:.4f} (%{customdata[2]}/%{customdata[1]})"
                      "<br>BKA %{customdata[3]:.4f} · BKB %{customdata[4]:.4f}<extra></extra>"))
    ooc = lim[lim["Status"] != "Terkendali"]
    if not ooc.empty:
        fig.add_trace(go.Scatter(x=ooc["Subgrup"], y=ooc[P], mode="markers+text", text=ooc["Subgrup"].astype(str),
                                 textposition="top center", textfont=dict(color=RED, size=11), hoverinfo="skip",
                                 marker=dict(color=RED, size=9)))
    last = lim.iloc[-1]
    end_label(fig, last["Subgrup"], last["BKA"], "BKA", GRAY)
    end_label(fig, last["Subgrup"], last["BKB"], "BKB", GRAY)
    end_label(fig, last["Subgrup"], pbar, f"p̄ = {fmt_dec(pbar, 4)}", ROYAL)
    if len(excluded):
        fig.add_annotation(x=0, y=-0.16, xref="paper", yref="paper", showarrow=False, xanchor="left",
                           text=f"Subgrup dikeluarkan: {', '.join(map(str, excluded))}",
                           font=dict(size=11, color=GRAY))
    style(fig, title, subtitle, height=420)
    fig.update_xaxes(title="Subgrup (hari pemeriksaan)")
    fig.update_yaxes(title="Proporsi cacat", tickformat=".0%", rangemode="tozero")
    return fig


def fig_hbar(labels, values, highlight, title, subtitle, text, xtitle=None, height=None):
    colors = [ROYAL if h else GRAY_LIGHT for h in highlight]
    fig = go.Figure(go.Bar(y=labels, x=values, orientation="h", marker_color=colors, text=text,
                           textposition="outside", cliponaxis=False, hoverinfo="skip",
                           textfont=dict(color=[ROYAL if h else GRAY_DARK for h in highlight])))
    style(fig, title, subtitle, height=height or 90 + 46 * len(labels), right=170)
    fig.update_yaxes(autorange="reversed", ticks="", tickfont=dict(color=INK))
    fig.update_xaxes(showticklabels=False, showline=False, ticks="", title=xtitle)
    return fig


def card(num, head, body):
    st.markdown(f"<div class='card'><div class='num'>{num}</div><div class='head'>{head}</div>"
                f"<div class='body'>{body}</div></div>", unsafe_allow_html=True)


def note(text):
    st.markdown(f"<div class='note'>{text}</div>", unsafe_allow_html=True)


def kpi(label, value, sub=None):
    s = f"<div class='kpi-s'>{sub}</div>" if sub else ""
    st.markdown(f"<div class='kpi'><div class='kpi-l'>{label}</div><div class='kpi-v'>{value}</div>{s}</div>",
                unsafe_allow_html=True)


def hero(label, value, sub):
    st.markdown(f"<div class='hero'><div class='hero-l'>{label}</div><div class='hero-v'>{value}</div>"
                f"<div class='hero-s'>{sub}</div></div>", unsafe_allow_html=True)


# ---------- Fishbone ----------
def fb_clean(text):
    return " ".join(str(text).split())


def fb_add(data, cat, text):
    t = fb_clean(text)
    if not t or cat not in data or t in data[cat]:
        return False
    data[cat].append(t)
    return True


def fb_edit(data, cat, i, text):
    t = fb_clean(text)
    if not t or cat not in data or not 0 <= i < len(data[cat]) or (t in data[cat] and data[cat][i] != t):
        return False
    data[cat][i] = t
    return True


def fb_delete(data, cat, i):
    if cat not in data or not 0 <= i < len(data[cat]):
        return False
    del data[cat][i]
    return True


def fb_wrap(text, width=24):
    return textwrap.wrap(fb_clean(text), width=width) or ["—"]


def fishbone_layout(data, problem):
    """Posisi garis dan teks Fishbone (koordinat 0-1 pada sumbu x, -1,3 s.d. 1,3 pada sumbu y).
    Tiga cabang atas (kategori 1-3) dan tiga cabang bawah (kategori 4-6); penyebab diletakkan di kiri cabang."""
    xb, dx, x_spine = [0.19, 0.45, 0.71], 0.11, 0.83
    lines = [(0.0, 0, x_spine, 0, "spine")]
    boxes, texts = [], []
    for row, sgn in ((FISHBONE_CATS[:3], 1), (FISHBONE_CATS[3:], -1)):
        for xi, cat in zip(xb, row):
            lines.append((xi, sgn * 1.0, xi + dx, 0, "branch"))
            boxes.append((xi, sgn * 1.0, cat))
            causes = data.get(cat, [])
            ys = [0.57] if len(causes) == 1 else list(np.linspace(0.80, 0.34, len(causes))) if causes else []
            if sgn < 0:  # cabang bawah: penyebab pertama paling dekat tulang punggung (seperti gambar penelitian)
                ys = ys[::-1]
            for y, c in zip(ys, causes):
                xbr = xi + dx * (1 - y)
                lines.append((xbr - 0.04, sgn * y, xbr, sgn * y, "cause"))
                texts.append((xbr - 0.045, sgn * y, fb_wrap(c, 28), cat))
    return {"lines": lines, "boxes": boxes, "texts": texts, "problem": [l.replace("¤", "\u00a0") for l in fb_wrap(problem.replace("PT. ", "PT.¤"), 15)],
            "x_spine": x_spine}


def fig_fishbone(data, problem, title, subtitle):
    lay = fishbone_layout(data, problem)
    fig = go.Figure()
    for x0, y0, x1, y1, kind in lay["lines"]:
        fig.add_shape(type="line", x0=x0, y0=y0, x1=x1, y1=y1,
                      line=dict(color={"spine": NAVY, "branch": BURGUNDY, "cause": GRAY_DARK}[kind],
                                width={"spine": 3, "branch": 2, "cause": 1.2}[kind]))
    for x, y, cat in lay["boxes"]:
        fig.add_annotation(x=x, y=y, text=f"<b>{cat}</b>", showarrow=False, yanchor="bottom" if y > 0 else "top",
                           yshift=4 if y > 0 else -4, bgcolor="white", bordercolor=BURGUNDY, borderwidth=1.5, borderpad=6,
                           font=dict(color=BURGUNDY, size=13))
    for x, y, lines_, cat in lay["texts"]:
        esc = "<br>".join(l.replace("<", "&lt;").replace(">", "&gt;") for l in lines_)
        fig.add_annotation(x=x, y=y, text=esc, showarrow=False, xanchor="right", yanchor="bottom" if y > 0 else "top",
                           yshift=2 if y > 0 else -2, align="right",
                           font=dict(color=INK, size=11))
    head = "<b>Problem</b><br>" + "<br>".join(l.replace("<", "&lt;").replace(">", "&gt;") for l in lay["problem"])
    fig.add_annotation(x=lay["x_spine"] + 0.012, y=0, text=head, showarrow=False, xanchor="left", align="center",
                       font=dict(color="white", size=12), bgcolor=BURGUNDY, borderpad=8)
    maxk = max([len(v) for v in data.values()] + [3])
    style(fig, title, subtitle, height=680 + 120 * max(0, maxk - 3), right=10)
    fig.update_xaxes(range=[-0.04, 1.0], visible=False)
    fig.update_yaxes(range=[-1.4, 1.4], visible=False)
    return fig


# ---------- FMEA ----------
def show_fmea(fmx):
    """Grafik RPN dan daftar usulan perbaikan; fmx sudah memuat kolom RPN dan terurut menurun."""
    hl = [i < 3 for i in range(len(fmx))]
    show(fig_hbar(fmx["Potensi Kegagalan"].str.rstrip("."), fmx["RPN"], hl,
                  f"Prioritas utama: {fmx.iloc[0]['Potensi Kegagalan'].rstrip('.').lower()} "
                  f"(RPN {fmt_int(fmx.iloc[0]['RPN'])})",
                  "Risk Priority Number = Severity × Occurrence × Detection; biru = tiga prioritas teratas",
                  text=[f"{fmt_int(r)}  (S{int(s_)} × O{int(o)} × D{int(d_)})"
                        for r, s_, o, d_ in zip(fmx["RPN"], fmx["S"], fmx["O"], fmx["D"])]))
    st.markdown("#### Usulan perbaikan")
    for i, r in fmx.iterrows():
        with st.expander(f"{i + 1}. {r['Faktor']} · RPN {fmt_int(r['RPN'])}", expanded=i < 3):
            st.markdown(f"**Penyebab:** {r['Penyebab Kegagalan']}  \n**Mode kegagalan:** {r['Potensi Kegagalan']}"
                        f"  \n**Dampak:** {r['Dampak Kegagalan']}  \n**Usulan perbaikan:** {r['Usulan Perbaikan']}")


def fmea_xlsx(fmx):
    out = fmx[["No", "Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan", "S", "O", "D", "RPN",
               "Usulan Perbaikan"]]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        out.to_excel(xw, index=False, sheet_name="FMEA")
    return buf.getvalue()


# ============================================================
# SIDEBAR: SUMBER DATA, FILTER, PENGATURAN
# ============================================================
def fr_title(text):
    st.markdown(f"<div class='fr-title'>{text}</div>", unsafe_allow_html=True)


notes, using_embedded = [], True
data_all, fmea, _ = embedded()
fmea_source = "bawaan"
upload_msgs = []

with st.sidebar:
    # A. Header
    st.markdown(
        "<div class='sb-brand'><div class='sb-kicker'>QUALITY CONTROL</div>"
        f"<div class='sb-title'>{APP_TITLE}</div></div>", unsafe_allow_html=True)

    # B. Frame sumber data
    with st.container(border=True, key="frame_data"):
        fr_title("Sumber Data")
        source = st.radio("Sumber data", ["Data bawaan (Proyek Akhir)", "Unggah data perusahaan"],
                          label_visibility="collapsed")
        if source == "Unggah data perusahaan":
            up = st.file_uploader("Workbook sesuai template (.xlsx)", type=["xlsx"])
            if up is None:
                st.info("Belum ada file. Dashboard sementara menampilkan data bawaan.")
            else:
                try:
                    data_up, fmea_up, notes = read_upload(up.getvalue())
                    data_all, using_embedded = data_up, False
                    if fmea_up is not None:
                        fmea, fmea_source = fmea_up, "unggahan"
                    else:
                        notes = notes + ["Sheet FMEA kosong/tidak ditemukan; halaman Improve memakai FMEA bawaan."]
                except WorkbookError as exc:
                    upload_msgs.append(("error", f"File '{up.name}' belum bisa dibaca: {exc}"))
                    upload_msgs.append(("info", "Unduh template di sidebar, salin data ke template tersebut, lalu unggah ulang."))
                except Exception as exc:
                    upload_msgs.append(("error", f"Kesalahan tak terduga saat membaca '{up.name}': {type(exc).__name__}: {exc}"))

    # C. Frame unduh template
    with st.container(border=True, key="frame_tpl"):
        fr_title("Unduh Template")
        st.download_button("⬇ Unduh Template Kosong", template_blank(),
                           file_name="Template_Data_SKT.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")

    # D. Frame periode pengamatan
    min_d, max_d = data_all[DATE].min().date(), data_all[DATE].max().date()
    with st.container(border=True, key="frame_period"):
        fr_title("Periode Pengamatan")
        rng = st.date_input("Periode Pengamatan", value=(min_d, max_d), min_value=min_d, max_value=max_d,
                            label_visibility="collapsed")
    start, end = (rng if isinstance(rng, (tuple, list)) and len(rng) == 2 else (min_d, max_d))
    data = data_all[(data_all[DATE].dt.date >= start) & (data_all[DATE].dt.date <= end)].copy()
    full_period = (start, end) == (min_d, max_d)
    too_few = len(data) < 3

    # E. Frame navigasi
    with st.container(border=True, key="frame_nav"):
        fr_title("Navigasi")
        page = st.radio("Navigasi", ["Ringkasan", "Define", "Measure", "Analyze", "Improve", "Data & Template"],
                        label_visibility="collapsed")

    # F. Pengaturan analisis
    above_only, auto, excluded = False, [], []
    with st.expander("Pengaturan analisis"):
        if too_few:
            st.caption("Perlebar periode pengamatan (minimal 3 subgrup) untuk membuka pengaturan ini.")
        else:
            rule = st.radio("Titik yang dikeluarkan pada peta revisi",
                            ["Di atas BKA dan di bawah BKB (sesuai naskah)", "Hanya di atas BKA"])
            above_only = rule.startswith("Hanya")
            auto = auto_exclude(data, above_only)
            key = "excl_" + hashlib.md5(f"{data[DATE].tolist()}{data[D].tolist()}{rule}".encode()).hexdigest()
            excluded = sorted(st.multiselect("Subgrup dikeluarkan (otomatis, bisa diubah)", data["Subgrup"].tolist(),
                                             default=auto, key=key))

    # G. Informasi data
    with st.container(border=True, key="frame_info"):
        fr_title("Informasi Data")
        st.caption(f"Data: {'bawaan Proyek Akhir' if using_embedded else 'unggahan perusahaan'}")
        st.caption(f"Periode tersedia: {min_d:%d %b %Y} – {max_d:%d %b %Y}")

    for msg in notes:
        st.warning(msg)

for kind_, msg_ in upload_msgs:
    (st.error if kind_ == "error" else st.info)(msg_)
if too_few:
    st.warning("Periode terpilih berisi kurang dari 3 subgrup. Perlebar filter periode di sidebar.")
    st.stop()

# ---------- hasil perhitungan bersama ----------
lim1, pbar1 = with_limits(data)
kept = data[~data["Subgrup"].isin(excluded)]
lim2, pbar2 = with_limits(kept) if len(kept) >= 2 else (None, None)
ooc1 = lim1[lim1["Status"] != "Terkendali"]
ooc2 = lim2[lim2["Status"] != "Terkendali"] if lim2 is not None else pd.DataFrame()
rt = run_test(kept) if len(kept) >= 3 else None
cap = capability(kept) if len(kept) else None
ss = six_sigma(data)
par = pareto_table(data)
tot_n, tot_d = int(data[N].sum()), int(data[D].sum())
pr_all = tot_d / tot_n
fm = fmea.sort_values("RPN", ascending=False).reset_index(drop=True) if fmea is not None else None
fm_research = prepare_fmea(pd.DataFrame(EMBEDDED_FMEA)).sort_values("RPN", ascending=False).reset_index(drop=True)

# ============================================================
# RINGKASAN
# ============================================================
if page == "Ringkasan":
    st.subheader("Ringkasan: jawaban atas tujuan penelitian")
    st.caption(f"Periode {start:%d %b %Y} – {end:%d %b %Y} · {len(data)} subgrup")
    k = st.columns(4)
    k[0].metric("Unit diperiksa", fmt_int(tot_n))
    k[1].metric("Produk cacat", fmt_int(tot_d))
    k[2].metric("Proporsi cacat", fmt_pct(pr_all, 2))
    k[3].metric("Level sigma", fmt_dec(ss["sigma"], 2) if ss else "—")

    c1, c2 = st.columns(2)
    with c1:
        if len(ooc1) == 0:
            stab = f"Seluruh {len(data)} subgrup berada dalam batas kendali peta P. "
        else:
            stab = (f"{len(ooc1)} dari {len(data)} subgrup berada di luar batas kendali pada peta P awal; "
                    f"setelah {len(excluded)} subgrup dikeluarkan, "
                    + ("seluruh titik terkendali" if lim2 is not None and ooc2.empty else
                       f"masih ada {len(ooc2)} titik di luar batas") + ". ")
        capt = (f"Indeks kapabilitas {fmt_dec(cap['index'], 3)} "
                + ("&lt; 1, sehingga proses <b>belum kapabel</b>" if cap['index'] < 1 else
                   "≥ 1, proses kapabel") + ". ") if cap else ""
        card("TUJUAN 1 · STABILITAS & KAPABILITAS",
             "Proses belum kapabel" if cap and cap["index"] < 1 else "Kondisi proses",
             stab + capt + (f"DPMO {fmt_dec(ss['dpmo'], 0)} dan level sigma {fmt_dec(ss['sigma'], 2)} "
                            f"({sigma_category(ss['sigma'])})." if ss else ""))
    with c2:
        top_f = fm.iloc[0] if fm is not None else None
        top_d = par.iloc[0]
        urut = ", ".join(f"{r['Jenis Cacat']} ({fmt_pct(r['Persentase'], 1)})" for _, r in par.head(3).iterrows())
        card("TUJUAN 2 · CACAT DOMINAN & PENYEBAB",
             f"{top_d['Jenis Cacat']} = {fmt_pct(top_d['Persentase'], 1)} dari seluruh cacat",
             f"Dari {fmt_int(par['Frekuensi'].sum())} cacat, tiga jenis terbesar menurut diagram Pareto: {urut}. "
             f"Penyebab dengan risiko tertinggi berdasarkan FMEA: "
             + (f"{top_f['Penyebab Kegagalan'][0].lower()}{top_f['Penyebab Kegagalan'][1:]}" if top_f is not None
                else "—"))
    st.write("")
    c3, c4 = st.columns(2)
    with c3:
        if fm is not None:
            body = "<br>".join(f"{i + 1}. {r['Potensi Kegagalan']} <b>(RPN {fmt_int(r['RPN'])})</b>"
                               for i, r in fm.head(3).iterrows())
            card("TUJUAN 3 · PRIORITAS PERBAIKAN (FMEA)", f"Prioritas utama: RPN {fmt_int(fm.iloc[0]['RPN'])}",
                 body)
    with c4:
        if fm is not None:
            card("TINDAK LANJUT", "Usulan perbaikan prioritas", fm.iloc[0]["Usulan Perbaikan"]
                 + "<br><span style='color:#9CA3AF'>Tujuan 4 (dashboard monitoring) dijawab oleh aplikasi ini.</span>")
    st.write("")
    fig = fig_pchart(lim1, pbar1, f"{len(ooc1)} subgrup keluar dari batas kendali pada peta P awal" if len(ooc1)
                     else "Seluruh subgrup berada dalam batas kendali",
                     "Proporsi cacat harian dengan batas kendali 3σ; titik merah berada di luar batas")
    show(fig)

# ============================================================
# DEFINE
# ============================================================
elif page == "Define":
    st.subheader("Define · Diagram SIPOC proses produksi rokok SKT")
    note("SIPOC memetakan proses produksi SKT secara umum, dari pemasok hingga pelanggan, sebagai ruang lingkup "
         "proyek perbaikan kualitas.")
    cols = st.columns([1, 1.5, 1.6, 1.1, 0.9])
    for col, (k_, items) in zip(cols, SIPOC.items()):
        with col:
            li = "".join(f"<li style='margin-bottom:4px'>{t}</li>" for t in items)
            tag = "ol" if k_ == "Process" else "ul"
            st.markdown(f"<div class='sipoc-h'>{k_}</div><div class='sipoc-b'><{tag} style='padding-left:18px;"
                        f"margin:0'>{li}</{tag}></div>", unsafe_allow_html=True)
    st.write("")
    note("Karakteristik cacat (CTQ) yang dipantau: " + ", ".join(DEFECTS) + ".")

# ============================================================
# MEASURE
# ============================================================
elif page == "Measure":
    st.subheader("Measure · Kinerja awal proses")
    k = st.columns(4)
    k[0].metric("Subgrup (hari)", fmt_int(len(data)))
    k[1].metric("Unit diperiksa", fmt_int(tot_n))
    k[2].metric("Produk cacat", fmt_int(tot_d))
    k[3].metric("Proporsi cacat", fmt_pct(pr_all, 2))
    st.caption(f"Ukuran sampel harian bervariasi antara {fmt_int(data[N].min())} dan {fmt_int(data[N].max())} unit.")

    wk = data.set_index(DATE)[[N, D]].resample("W-MON", label="left", closed="left").sum()
    wk = wk[wk[N] > 0]
    wk["p"] = wk[D] / wk[N]
    hi, lo = wk["p"].idxmax(), wk["p"].idxmin()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=wk.index, y=wk["p"], mode="lines+markers", line=dict(color=GRAY_DARK, width=2),
                             marker=dict(size=6, color=GRAY_DARK),
                             hovertemplate="Minggu %{x|%d %b %Y}<br>%{y:.2%}<extra></extra>"))
    fig.add_hline(y=pr_all, line=dict(color=ROYAL, dash="dash", width=1.5))
    fig.add_trace(go.Scatter(x=[hi], y=[wk.loc[hi, "p"]], mode="markers+text", marker=dict(size=10, color=ROYAL),
                             text=[f"tertinggi {fmt_pct(wk.loc[hi, 'p'])}"], textposition="top center",
                             textfont=dict(color=ROYAL), hoverinfo="skip"))
    end_label(fig, wk.index[-1], pr_all, f"rata-rata {fmt_pct(pr_all, 2)}", ROYAL)
    style(fig, f"Proporsi cacat mingguan berkisar {fmt_pct(wk['p'].min())}–{fmt_pct(wk['p'].max())}",
          "Total produk cacat ÷ total unit diperiksa per minggu (minggu dimulai Senin)", height=380)
    fig.update_yaxes(tickformat=".0%", rangemode="tozero", title="Proporsi cacat")
    show(fig)
    with st.expander("Data pemeriksaan harian"):
        tbl = data[["Subgrup", DATE, N, *DEFECTS, D, P]].copy()
        tbl[DATE] = tbl[DATE].dt.date
        st.dataframe(tbl, width="stretch", hide_index=True,
                     column_config={P: st.column_config.NumberColumn(format="%.4f")})

# ============================================================
# ANALYZE
# ============================================================
elif page == "Analyze":
    st.subheader("Analyze · Kestabilan, kapabilitas, dan penyebab cacat")
    tabs = st.tabs(["Peta Kendali P", "Uji Keacakan", "Kapabilitas Proses", "Level Sigma", "Diagram Pareto",
                    "Diagram Sebab-Akibat"])

    with tabs[0]:
        show(fig_pchart(lim1, pbar1,
                        f"Peta P awal: {len(ooc1)} subgrup di luar batas kendali" if len(ooc1)
                        else "Peta P awal: seluruh subgrup terkendali",
                        f"{len(data)} subgrup · {int((ooc1['Status'] == 'Di atas BKA').sum())} di atas BKA, "
                        f"{int((ooc1['Status'] == 'Di bawah BKB').sum())} di bawah BKB"))
        if lim2 is not None:
            show(fig_pchart(lim2, pbar2,
                            "Peta P revisi: seluruh subgrup terkendali" if ooc2.empty
                            else f"Peta P revisi: masih {len(ooc2)} subgrup di luar batas",
                            f"{len(kept)} subgrup · p̄ berubah dari {fmt_dec(pbar1, 4)} menjadi {fmt_dec(pbar2, 4)}"
                            if excluded else f"{len(kept)} subgrup · tidak ada subgrup yang dikeluarkan",
                            excluded=excluded))
        note("Subgrup yang dikeluarkan dipilih otomatis (berulang hingga terkendali) dan dapat diubah melalui "
             "<b>Pengaturan analisis</b> di sidebar. Pengeluaran subgrup harus didukung bukti penyebab khusus.")
        if len(ooc1):
            t = ooc1[["Subgrup", DATE, N, D, P, "BKA", "BKB", "Status"]].copy()
            t[DATE] = t[DATE].dt.date
            st.dataframe(t, width="stretch", hide_index=True,
                         column_config={c: st.column_config.NumberColumn(format="%.4f") for c in (P, "BKA", "BKB")})

    with tabs[1]:
        if rt is None:
            st.warning("Uji keacakan memerlukan data di atas dan di bawah median.")
        else:
            acak = rt["p"] > 0.05
            x, y = rt["subgrup"], rt["y"]
            col_pts = [ROYAL if s_ > 0 else GRAY if s_ < 0 else GRAY_LIGHT for s_ in rt["state"]]
            fig = go.Figure()
            fig.add_hline(y=rt["median"], line=dict(color=GRAY, width=1.5, dash="dash"))
            fig.add_trace(go.Scatter(x=x, y=y, mode="lines", line=dict(color=GRAY_LIGHT, width=1), hoverinfo="skip"))
            fig.add_trace(go.Scatter(x=x, y=y, mode="markers", hovertemplate="Subgrup %{x}: %{y:.4f}<extra></extra>",
                                     marker=dict(size=7, color=col_pts)))
            end_label(fig, x[-1], rt["median"], f"median {fmt_dec(rt['median'], 3)}", GRAY_DARK)
            style(fig, ("Data tersusun acak" if acak else "Data tidak acak")
                  + f": {rt['V']} run, sesuai harapan {fmt_dec(rt['mu'], 0)} run",
                  f"<span style='color:{ROYAL}'><b>biru</b></span> = di atas median, abu-abu = di bawah median, "
                  f"abu muda = sama dengan median (tidak dihitung) · p-value {fmt_dec(rt['p'], 3)}", height=380)
            fig.update_yaxes(tickformat=".0%", title="Proporsi cacat")
            fig.update_xaxes(title="Subgrup")
            show(fig)
            m = st.columns(7)
            for col, (lab, val) in zip(m, [("n₁ (atas)", str(rt["n1"])), ("n₂ (bawah)", str(rt["n2"])),
                                           ("Run (V)", str(rt["V"])), ("μV", fmt_dec(rt["mu"], 3)),
                                           ("σV", fmt_dec(rt["sigma"], 3)),
                                           ("Z", fmt_dec(rt["z"], 3)), ("p-value", fmt_dec(rt["p"], 3))]):
                col.metric(lab, val)
            note(f"H₀ (data acak) {'tidak ditolak' if acak else 'ditolak'} pada α = 0,05. Uji dilakukan pada "
                 f"{len(kept)} subgrup peta P revisi (setelah {len(excluded)} subgrup di luar batas dikeluarkan). "
                 f"{rt['ties']} subgrup yang proporsinya sama dengan median tidak diikutsertakan, sehingga "
                 f"n₁ + n₂ = {rt['n1'] + rt['n2']}.")

    with tabs[2]:
        if cap is None:
            st.warning("Kapabilitas tidak dapat dihitung.")
        else:
            idx = cap["index"]
            xmax = max(1.6, idx * 1.25)
            fig = go.Figure(go.Bar(x=[idx], y=["Indeks"], orientation="h", width=0.45, hoverinfo="skip",
                                   marker_color=ROYAL if idx < 1 else GRAY_DARK, text=[fmt_dec(idx, 3)],
                                   textposition="outside", cliponaxis=False, textfont=dict(color=ROYAL, size=16)))
            for xv, lab, anc, sh in [(1.0, "1,00 batas minimum", "right", -4), (1.33, "1,33 standar industri", "left", 4)]:
                fig.add_vline(x=xv, line=dict(color=GRAY, dash="dot"))
                fig.add_annotation(x=xv, y=1, yref="paper", text=lab, showarrow=False, yanchor="bottom",
                                   xanchor=anc, xshift=sh, font=dict(color=GRAY_DARK, size=12))
            style(fig, ("Proses belum kapabel" if idx < 1 else "Proses cukup kapabel" if idx < 1.33
                        else "Proses kapabel") + f": indeks {fmt_dec(idx, 3)}",
                  f"Indeks kapabilitas atribut = Φ⁻¹(1 − p̄) / 3 dari peta P revisi ({len(kept)} subgrup)",
                  height=340, right=70)
            fig.update_layout(margin=dict(l=20, r=70, t=140, b=40))
            fig.update_xaxes(range=[0, xmax], showline=False, showticklabels=False, ticks="")
            fig.update_yaxes(showticklabels=False)
            show(fig)
            st.write("")
            m = st.columns(3, gap="medium")
            with m[0]:
                kpi("p̄ revisi", fmt_dec(cap["pbar"], 4))
            with m[1]:
                kpi("Φ⁻¹(1 − p̄)", fmt_dec(cap["z"], 3))
            with m[2]:
                kpi("Indeks kapabilitas", fmt_dec(idx, 3))

    with tabs[3]:
        if ss is None:
            st.warning("Level sigma tidak dapat dihitung.")
        else:
            s = ss["sigma"]
            hero("LEVEL SIGMA", fmt_dec(s, 3) if not pd.isna(s) else "—", sigma_category(s))
            m = st.columns(3, gap="medium")
            with m[0]:
                kpi("DPU", fmt_dec(ss["dpu"], 4))
            with m[1]:
                kpi("DPO", fmt_dec(ss["dpo"], 5))
            with m[2]:
                kpi("DPMO", fmt_dec(ss["dpmo"], 2))
            st.caption("DPO = DPU (1 opportunity: proporsi cacat); "
                       "level sigma = Φ⁻¹(1 − DPO) + 1,5.")
            fig = go.Figure()
            bands = [(1, 2, "Sangat tidak<br>kompetitif"), (2, 4, "Rata-rata industri<br>Indonesia"),
                     (4, 6, "Rata-rata industri<br>USA"), (6, 6.6, "Kelas<br>dunia")]
            shades = ["#F3F4F6", "#E5E7EB", "#F3F4F6", "#E5E7EB"]
            for (a, b, lab), sh in zip(bands, shades):
                fig.add_shape(type="rect", x0=a, x1=b, y0=0, y1=1, fillcolor=sh, line_width=0, layer="below")
                fig.add_annotation(x=(a + b) / 2, y=0.18, text=lab, showarrow=False, font=dict(color=GRAY_DARK, size=12))
            fig.add_trace(go.Scatter(x=[s], y=[0.62], mode="markers+text", marker=dict(size=18, color=ROYAL,
                                     symbol="diamond"), text=[f"{fmt_dec(s, 2)}σ"], textposition="top center",
                                     textfont=dict(color=ROYAL, size=16), hoverinfo="skip"))
            style(fig, f"Level sigma {fmt_dec(s, 2)}: setara {sigma_category(s)[0].lower() + sigma_category(s)[1:]}",
                  f"DPMO {fmt_dec(ss['dpmo'], 0)} · 1 opportunity · {len(data)} subgrup · kategori Gaspersz (2002)",
                  height=260, right=30)
            fig.update_xaxes(range=[1, 6.6], tickvals=[1, 2, 3, 4, 5, 6], ticktext=[f"{i}σ" for i in range(1, 7)])
            fig.update_yaxes(range=[0, 1], showticklabels=False)
            show(fig)

    with tabs[4]:
        txt = [f"{fmt_int(f)}  ·  {fmt_pct(p_)}  (kum. {fmt_pct(c)})"
               for f, p_, c in zip(par["Frekuensi"], par["Persentase"], par["Kumulatif"])]
        top_d = par.iloc[0]
        show(fig_hbar(par["Jenis Cacat"], par["Frekuensi"], [True] * len(par),
                      f"{top_d['Jenis Cacat']} menyumbang {fmt_pct(top_d['Persentase'])} dari seluruh cacat",
                      f"Frekuensi cacat per jenis, diurutkan dari terbesar · frekuensi, persentase, dan persentase "
                      f"kumulatif · total {fmt_int(par['Frekuensi'].sum())} cacat dari {len(data)} subgrup", text=txt))

    with tabs[5]:
        if "fishbone" not in st.session_state:
            st.session_state["fishbone"] = copy.deepcopy(FISHBONE_RESEARCH)
        fb = st.session_state["fishbone"]
        n_cause = sum(len(v) for v in fb.values())
        is_research = fb == FISHBONE_RESEARCH
        show(fig_fishbone(fb, FISHBONE_PROBLEM,
                          f"{n_cause} penyebab dalam {len(FISHBONE_CATS)} kategori mengarah pada proses SKT yang belum kapabel",
                          "Diagram sebab-akibat (Fishbone) hasil penelitian, kategori 5M + 1E" if is_research else
                          "Diagram sebab-akibat (Fishbone) yang telah disesuaikan · berbeda dari data penelitian"))
        note("Isi awal diagram adalah <b>18 penyebab hasil penelitian</b>. Perusahaan dapat menambah, mengubah, atau "
             "menghapus penyebab di bawah ini. Perubahan berlaku selama sesi dan tidak mengubah data penelitian; "
             "tombol <b>Kembalikan</b> memulihkan 18 penyebab asli.")
        with st.expander("Kelola penyebab (tambah, ubah, hapus)"):
            ca, cb = st.columns(2)
            with ca:
                st.markdown("**Tambah penyebab**")
                with st.form("fb_add_form", clear_on_submit=True):
                    cat_add = st.selectbox("Kategori", FISHBONE_CATS)
                    txt_add = st.text_input("Penyebab baru", max_chars=120)
                    if st.form_submit_button("Tambah"):
                        if fb_add(fb, cat_add, txt_add):
                            st.rerun()
                        else:
                            st.warning("Penyebab kosong atau sudah ada pada kategori tersebut.")
            with cb:
                st.markdown("**Ubah atau hapus penyebab**")
                flat = [(c_, i_, t_) for c_ in FISHBONE_CATS for i_, t_ in enumerate(fb[c_])]
                if flat:
                    sel = st.selectbox("Pilih penyebab", list(range(len(flat))),
                                       format_func=lambda k: f"{flat[k][0]} · {flat[k][2]}", key="fb_sel")
                    c_sel, i_sel, t_sel = flat[sel]
                    new_txt = st.text_input("Redaksi penyebab", value=t_sel, max_chars=120,
                                            key=f"fb_edit_{c_sel}_{i_sel}_{hashlib.md5(t_sel.encode()).hexdigest()[:6]}")
                    b1, b2 = st.columns(2)
                    if b1.button("Simpan perubahan", key="fb_save"):
                        if fb_edit(fb, c_sel, i_sel, new_txt):
                            st.rerun()
                        else:
                            st.warning("Redaksi kosong atau sama dengan penyebab lain pada kategori itu.")
                    if b2.button("Hapus penyebab", key="fb_del"):
                        fb_delete(fb, c_sel, i_sel)
                        st.rerun()
                else:
                    st.info("Belum ada penyebab. Tambahkan melalui kotak di sebelah kiri.")
            if st.button("Kembalikan ke 18 penyebab penelitian", key="fb_reset", disabled=is_research):
                st.session_state["fishbone"] = copy.deepcopy(FISHBONE_RESEARCH)
                st.rerun()
        tbl_fb = pd.DataFrame([(c_, i_ + 1, t_) for c_ in FISHBONE_CATS for i_, t_ in enumerate(fb[c_])],
                              columns=["Kategori", "No", "Penyebab"])
        buf_fb = io.BytesIO()
        with pd.ExcelWriter(buf_fb, engine="openpyxl") as xw:
            tbl_fb.to_excel(xw, index=False, sheet_name="Fishbone")
        st.download_button("⬇ Unduh daftar penyebab (Excel)", buf_fb.getvalue(), "Fishbone_penyebab.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# IMPROVE
# ============================================================
elif page == "Improve":
    st.subheader("Improve · Prioritas perbaikan berdasarkan FMEA")
    t_ex, t_own = st.tabs(["FMEA Penelitian (contoh pengisian)", "FMEA Perusahaan (isi sendiri)"])
    with t_ex:
        note("Ini adalah <b>FMEA hasil penelitian</b> di PT. XYZ, ditampilkan sebagai contoh pengisian. Datanya tetap "
             "dan tidak dipengaruhi data unggahan. Untuk mengisi FMEA perusahaan, buka tab di sebelah kanan.")
        show_fmea(fm_research)
        st.markdown("#### Tabel FMEA penelitian")
        st.dataframe(fm_research[["No", "Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan",
                                  "S", "O", "D", "RPN", "Usulan Perbaikan"]], width="stretch", hide_index=True)
    with t_own:
        note("Isi tabel di bawah dengan kondisi produksi aktual. <b>Faktor, S, O, dan D wajib diisi</b> (S, O, D berskala "
             "1–10); RPN dihitung otomatis. Tabel ini terpisah dari FMEA penelitian.")
        cols_own = ["Faktor", "Penyebab Kegagalan", "Potensi Kegagalan", "Dampak Kegagalan", "S", "O", "D",
                    "Usulan Perbaikan"]
        if fmea_source == "unggahan" and fmea is not None:
            init = fmea[cols_own].copy()
            note("Tabel terisi dari sheet FMEA pada file perusahaan yang diunggah. Anda tetap dapat mengubahnya di sini.")
        else:
            init = pd.DataFrame({c: pd.Series([None] * 5, dtype="object") for c in cols_own[:4]}
                                | {c: pd.Series([np.nan] * 5, dtype="float") for c in ("S", "O", "D")}
                                | {"Usulan Perbaikan": pd.Series([None] * 5, dtype="object")})[cols_own]
        faktor_opts = FAKTOR_LIST + [f for f in init["Faktor"].dropna().astype(str).unique() if f not in FAKTOR_LIST]
        num = lambda lab: st.column_config.NumberColumn(lab, min_value=1, max_value=10, step=1, format="%d")
        edited = st.data_editor(
            init, num_rows="dynamic", hide_index=True, width="stretch",
            key="fmea_own_" + hashlib.md5(init.to_json().encode()).hexdigest()[:8],
            column_config={"Faktor": st.column_config.SelectboxColumn("Faktor", options=faktor_opts),
                           "Penyebab Kegagalan": st.column_config.TextColumn("Penyebab Kegagalan"),
                           "Potensi Kegagalan": st.column_config.TextColumn("Potensi Kegagalan"),
                           "Dampak Kegagalan": st.column_config.TextColumn("Dampak Kegagalan"),
                           "S": num("S"), "O": num("O"), "D": num("D"),
                           "Usulan Perbaikan": st.column_config.TextColumn("Usulan Perbaikan")})
        own = prepare_fmea(edited.assign(No=np.arange(1, len(edited) + 1)))
        n_rows = int(edited.dropna(how="all").shape[0])
        if own is None:
            st.info("Isi minimal satu baris lengkap (Faktor, S, O, dan D) untuk menampilkan hasil FMEA perusahaan.")
        else:
            if n_rows > len(own):
                st.caption(f"{n_rows - len(own)} baris belum lengkap (Faktor, S, O, D wajib) dan belum dihitung.")
            own = own.sort_values("RPN", ascending=False).reset_index(drop=True)
            show_fmea(own)
            st.download_button("⬇ Unduh FMEA perusahaan (Excel)", fmea_xlsx(own), "FMEA_perusahaan.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# DATA & TEMPLATE
# ============================================================
elif page == "Data & Template":
    st.subheader("Data & Template")
    note("Perusahaan dapat memperbarui dashboard tanpa mengubah kode: unduh template, isi data harian dan FMEA, "
         "lalu unggah melalui sidebar (<b>Sumber data → Unggah data perusahaan</b>).")
    a, b = st.columns(2)
    with a:
        card("LANGKAH 1", "Unduh template", "Template berisi sheet Petunjuk, Data Harian, dan FMEA. Kolom jumlah "
             "cacat, proporsi, dan RPN terhitung otomatis di Excel.")
        st.download_button("⬇ Template kosong", template_blank(), "Template_Data_SKT.xlsx",
                           width="stretch")
        st.download_button("⬇ Template berisi data Proyek Akhir (contoh)", template_example(),
                           "Template_Data_SKT_contoh.xlsx", width="stretch")
    with b:
        card("LANGKAH 2", "Isi dan unggah", "Satu baris = satu hari pemeriksaan. Setiap rokok cacat dicatat pada "
             "satu jenis cacat utama. Nilai S, O, D memakai skala 1–10. Simpan sebagai .xlsx lalu unggah.")
    st.write("")
    st.markdown(f"**Sumber data aktif:** {'data bawaan Proyek Akhir' if using_embedded else 'unggahan perusahaan'}"
                f" · {len(data_all)} subgrup · FMEA {fmea_source}")
    if notes:
        for msg in notes:
            st.warning(msg)
    if using_embedded and full_period and auto == excluded and not above_only:
        st.markdown("#### Verifikasi terhadap workbook Excel olah data")
        mine = {"p̄ peta kendali awal": pbar1, "Subgrup di luar batas": len(ooc1), "p̄ peta kendali revisi": pbar2,
                "Statistik uji Z (runs test)": rt["z"] if rt else None, "p-value runs test": rt["p"] if rt else None,
                "Indeks kapabilitas": cap["index"] if cap else None, "DPMO": ss["dpmo"], "Level sigma": ss["sigma"]}
        rows = []
        for k_, ref in NASKAH.items():
            v = mine[k_]
            nd = 0 if k_ == "Subgrup di luar batas" else 3 if k_ in ("DPMO",) else 4
            ok = v is not None and abs(v - ref) <= (0.5 if k_ == "DPMO" else 0.0006 if isinstance(ref, float) else 0)
            rows.append({"Indikator": k_, "Dashboard": fmt_dec(v, nd), "Excel": fmt_dec(ref, nd),
                         "Status": "Sesuai" if ok else "Berbeda"})
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        st.caption("Dasar data: peta P awal, level sigma, dan Pareto memakai 66 subgrup; peta P revisi, kapabilitas "
                   "proses, dan runs test memakai 54 subgrup (setelah 12 subgrup di luar batas dikeluarkan).")
    with st.expander("Lihat data aktif"):
        t = data_all.copy()
        t[DATE] = t[DATE].dt.date
        st.dataframe(t, width="stretch", hide_index=True)
