"""
RKP Monitor — Dashboard Monitoring Rencana Kerja Proyek
Jalankan lokal:  streamlit run app.py
Deploy: push ke GitHub lalu hubungkan repo di https://share.streamlit.io
"""

import re
import io
import hashlib
from pathlib import Path
from datetime import datetime

import openpyxl
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# Folder di repo GitHub tempat file .xlsx RKP disimpan.
# Update dashboard = tambah/ganti file .xlsx di folder ini lalu commit ke GitHub.
DATA_DIR = Path(__file__).parent / "data"

# Set True lagi setelah file/sheet Realisasi resmi mulai diupload — untuk sementara
# dikosongkan dulu semua (semua proyek akan tampil status "Belum ada").
ENABLE_REALISASI = False

# Koreksi manual nama perusahaan untuk file yang nama perusahaannya tidak tertulis
# jelas di baris atas sheet (jadi tidak bisa dibaca otomatis oleh extract_meta).
COMPANY_OVERRIDES = {
    "RKP_BIOGAS_SYSTEM__19_Feb_2026_.xlsx": "PT. Kumai Sentosa",
    "RKP_Proyek_Dermaga_-_PT__BKB_Tahap_II.xlsx": "PT. Buana Karya Bhakti",
    "Timeline_RKP_Kandang_Fattening___RPH_site.xlsx": "PT. Siskaranch",
}

# Penulisan nama perusahaan di berbagai file sumber tidak konsisten (ada/tidaknya titik,
# huruf besar semua, typo "FASS" vs "FAST", dst) — disamakan di sini supaya tidak muncul
# sebagai baris terpisah di filter Perusahaan padahal perusahaannya sama.
def _company_key(s):
    return re.sub(r"[.\s]+", " ", str(s).upper()).strip()


CANONICAL_COMPANIES = {
    _company_key("PT BUANA KARYA BHAKTI"): "PT. Buana Karya Bhakti",
    _company_key("PT FAST FOREST DEVELOPMENT"): "PT. Fast Forest Development",
    _company_key("PT FASS FOREST DEVELOPMENT"): "PT. Fast Forest Development",  # typo di file sumber
    _company_key("PT KUMAI SENTOSA"): "PT. Kumai Sentosa",
    _company_key("PT SISKARANCH"): "PT. Siskaranch",
}

# Nama tampilan proyek yang diminta pengguna (menggantikan nama hasil parsing filename).
PROJECT_NAME_OVERRIDES = {
    "RKP_BIOGAS_SYSTEM__19_Feb_2026_.xlsx": "Biogas System",
    "RKP_Teluk_Pulai_Kumai_Sentosa__06_Maret_2026_.xlsx": "Pembukaan Lahan Teluk Pulai",
    "RKP_Replanting_BKB_Inti__PT__Buana_Karya_Bhakti_.xlsx": "Replanting BKB Inti",
    "RKP_Reklamasi_FFD_Inti__PT__Fast_Forest_Development_.xlsx": "Reklamasi FFD Inti",
    "RKP_Reklamasi_BKB_Inti__PT__Buana_Karya_Bhakti_.xlsx": "Reklamasi BKB Inti",
    "RKP_Proyek_Restorasi_PKS_Batulaki_PT__BKB_Tahun_2026-2030.xlsx": "Restorasi PKS Batulaki",
    "RKP_Proyek_Dermaga_-_PT__BKB_Tahap_II.xlsx": "Dermaga Tahap II BKB",
    "RKP_Pembukaan_Lahan_Satui_Timur__PT__Buana_Karya_Bhakti_.xlsx": "Pembukaan Lahan Satui Timur",
    "RKP_PLASMA_MANDIRI_TR_200_Ha.xlsx": "Pembukaan Lahan Plasma Mandiri TR",
    "RKP_PLASMA_MANDIRI_SUCAB__300_Ha.xlsx": "Pembukaan Lahan Plasma Mandiri Sucab",
    "RKP_PASTURA_KEBUN_BKB_INTI_-_2026.xlsx": "Pastura Kebun BKB Inti",
    "RKP_PASTURA_KEBUN_FFD_INTI_-_2026.xlsx": "Pastura Kebun FFD Inti",
    "Timeline_RKP_Kandang_Fattening___RPH_site.xlsx": "Kandang Fattening & RPH",
}

# ============================================================
# KONFIGURASI HALAMAN & TEMA
# ============================================================
st.set_page_config(
    page_title="RKP Monitor",
    page_icon="🌴",
    layout="wide",
    initial_sidebar_state="expanded",
)

FOREST = "#123524"
FOREST_LIGHT = "#3C7A5A"
GOLD = "#C08A2E"
GOLD_LIGHT = "#E8C776"
RUST = "#B25330"
INK = "#16231A"
IVORY = "#F6F4EE"
PALETTE = ["#123524", "#C08A2E", "#3C7A5A", "#B25330", "#7C9A85", "#8C6A2E", "#4E6B57", "#D9AE63"]

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700;9..144,800&family=Inter:wght@400;500;600;700;800&display=swap');

    /* Paksa tema terang ini terlepas dari mode gelap/terang browser/OS, supaya
       tidak bergantung pada .streamlit/config.toml ikut ter-upload atau tidak. */

    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"],
    [data-testid="stMain"], [data-testid="stBottomBlockContainer"] {{
        background-color: {IVORY} !important;
        color: {INK} !important;
        font-family: 'Inter', sans-serif;
    }}

    h1, h2, h3, h4, h5, h6,
    [data-testid="stMarkdownContainer"] h1, [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3, [data-testid="stMarkdownContainer"] h4 {{
        font-family: 'Fraunces', serif !important; color: {FOREST} !important;
        letter-spacing: -0.01em;
    }}
    [data-testid="stMarkdownContainer"] h1 {{ font-weight: 700 !important; }}
    [data-testid="stMarkdownContainer"] h4 {{ font-weight: 600 !important; font-size: 19px !important; margin-top: 6px !important; }}

    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li,
    [data-testid="stMarkdownContainer"] span,
    .stCaption, [data-testid="stCaptionContainer"],
    label, .stRadio label, .stRadio span {{
        color: {INK} !important;
    }}
    [data-testid="stCaptionContainer"] p {{ color: #6C7566 !important; }}

    /* ---------- KPI metric cards: aksen tepi kiri berwarna per kategori ---------- */
    div[data-testid="stMetric"] {{
        background: #FFFFFF !important; border: 1px solid #E4E1D6; border-left: 4px solid {FOREST};
        border-radius: 14px; padding: 18px 20px 16px;
        box-shadow: 0 1px 2px rgba(18,53,36,0.05), 0 8px 20px rgba(18,53,36,0.04);
    }}
    div[data-testid="stMetricLabel"] p {{
        color: #6C7566 !important; font-size: 11.5px !important; font-weight: 700 !important;
        text-transform: uppercase; letter-spacing: 0.06em;
    }}
    div[data-testid="stMetricValue"] {{
        color: {FOREST} !important; font-family: 'Fraunces', serif !important; font-weight: 700 !important;
    }}
    div[data-testid="stMetricDelta"] {{ color: {FOREST_LIGHT} !important; font-weight: 600 !important; }}

    /* Kartu KPI ke-2/3/4 dalam satu baris: variasi aksen supaya tidak seragam total */
    [data-testid="stHorizontalBlock"] div[data-testid="stMetric"]:nth-of-type(4n+2) {{ border-left-color: {GOLD}; }}
    [data-testid="stHorizontalBlock"] div[data-testid="stMetric"]:nth-of-type(4n+3) {{ border-left-color: {FOREST_LIGHT}; }}
    [data-testid="stHorizontalBlock"] div[data-testid="stMetric"]:nth-of-type(4n+4) {{ border-left-color: {RUST}; }}

    /* ---------- Tabs ---------- */
    button[data-baseweb="tab"] p {{ color: #6C7566 !important; font-weight: 600; }}
    button[data-baseweb="tab"][aria-selected="true"] p {{ color: {FOREST} !important; }}
    [data-testid="stTabs"] {{ background: transparent !important; }}
    div[data-baseweb="tab-highlight"] {{ background-color: {GOLD} !important; }}
    div[data-baseweb="tab-border"] {{ background-color: #E1E3D9 !important; }}

    /* ---------- Tombol biasa: lebih premium, tidak kotak polos ---------- */
    .stButton button {{
        border-radius: 10px !important; font-weight: 600 !important; border: 1px solid #E4E1D6 !important;
        transition: transform .08s ease, box-shadow .12s ease;
    }}
    .stButton button:hover {{ transform: translateY(-1px); box-shadow: 0 4px 14px rgba(18,53,36,0.12); }}
    [data-testid="stMain"] .stButton button {{
        background: #FFFFFF !important; color: {FOREST} !important;
    }}
    [data-testid="stDownloadButton"] button {{
        background: {GOLD} !important; color: #2A1D06 !important; border: none !important; font-weight: 700 !important;
    }}

    /* ---------- Sidebar: hijau tua + navigasi ala segmented pill ---------- */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {FOREST} 0%, #0C2318 100%) !important;
    }}
    section[data-testid="stSidebar"] * {{ color: {IVORY} !important; }}
    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{ color: #A9C9B6 !important; }}
    section[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,0.12) !important; }}

    section[data-testid="stSidebar"] .stButton button {{
        background-color: {GOLD} !important; color: #2A1D06 !important; border: none !important;
        border-radius: 10px !important; font-weight: 700 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stExpander"] {{
        border: 1px solid rgba(255,255,255,0.16) !important; border-radius: 12px; background: rgba(255,255,255,0.03);
    }}
    section[data-testid="stSidebar"] [data-testid="stAlertContainer"] {{ color: {INK} !important; border-radius: 10px; }}
    section[data-testid="stSidebar"] [data-testid="stAlertContainer"] * {{ color: {INK} !important; }}

    /* Radio "Halaman" jadi pil segmented, bukan bulatan radio bawaan */
    section[data-testid="stSidebar"] div[role="radiogroup"] {{
        display: flex; flex-direction: column; gap: 4px; background: rgba(255,255,255,0.06);
        border-radius: 12px; padding: 4px;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] label {{
        border-radius: 9px; padding: 8px 12px !important; margin: 0 !important; transition: background .12s ease;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {{
        background: {GOLD} !important;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) p {{
        color: #2A1D06 !important; font-weight: 700 !important;
    }}
    section[data-testid="stSidebar"] div[role="radiogroup"] input {{ display: none !important; }}
    section[data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child {{ display: none !important; }}

    /* ---------- Dataframe / table ---------- */
    [data-testid="stDataFrame"] {{ color: {INK} !important; border-radius: 12px; overflow: hidden; }}

    /* ---------- Container berbatas (panel) ---------- */
    [data-testid="stVerticalBlockBorderWrapper"] {{
        border-radius: 16px !important; border-color: #E4E1D6 !important;
        box-shadow: 0 1px 2px rgba(18,53,36,0.04), 0 10px 24px rgba(18,53,36,0.05);
    }}

    .badge-wait {{ background:#F5E4DA; color:{RUST}; padding:3px 10px; border-radius:99px; font-size:11.5px; font-weight:700; }}
    .badge-ok {{ background:#E4EEE7; color:{FOREST}; padding:3px 10px; border-radius:99px; font-size:11.5px; font-weight:700; }}
    .footnote {{ background:#E4EEE7; border-radius:12px; padding:14px 16px; font-size:13px; color:#3d4a40 !important; }}
    .footnote * {{ color:#3d4a40 !important; }}

    /* ---------- Hero band (Beranda) ---------- */
    .hero-band {{
        background: linear-gradient(120deg, {FOREST} 0%, #0C2318 100%);
        border-radius: 22px; padding: 36px 40px; color: {IVORY};
        display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 24px;
        box-shadow: 0 20px 40px rgba(12,35,24,0.25);
        position: relative; overflow: hidden;
    }}
    .hero-band::after {{
        content: ""; position: absolute; right: -60px; top: -60px; width: 220px; height: 220px;
        border-radius: 50%; background: radial-gradient(circle, rgba(232,199,118,0.18), transparent 70%);
    }}
    .hero-eyebrow {{ font-size: 12px; letter-spacing: 0.1em; text-transform: uppercase; color: #A9C9B6; font-weight: 700; margin: 0 0 10px; }}
    .hero-number {{ font-family: 'Fraunces', serif; font-weight: 700; font-size: 52px; color: #FFFFFF; margin: 0; line-height: 1; }}
    .hero-sub {{ font-size: 14px; color: #CBE0D3; margin: 10px 0 0; max-width: 440px; }}
    .hero-split {{ display: flex; gap: 28px; flex-wrap: wrap; }}
    .hero-split-item {{ text-align: right; }}
    .hero-split-num {{ font-family: 'Fraunces', serif; font-size: 24px; font-weight: 700; color: #FFFFFF; }}
    .hero-split-lbl {{ font-size: 11.5px; color: #A9C9B6; text-transform: uppercase; letter-spacing: 0.05em; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# HELPER FORMAT
# ============================================================
def fmt_rp(n):
    if n is None or pd.isna(n):
        return "—"
    if abs(n) >= 1e9:
        return f"Rp {n/1e9:,.2f} M".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs(n) >= 1e6:
        return f"Rp {n/1e6:,.1f} Jt".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"Rp {n:,.0f}".replace(",", ".")


def fmt_rp_full(n):
    if n is None or pd.isna(n):
        return "—"
    return f"Rp {n:,.0f}".replace(",", ".")


def fmt_ha(n):
    if n is None or pd.isna(n):
        return "—"
    return f"{n:,.2f} Ha".replace(",", "X").replace(".", ",").replace("X", ".")


# ============================================================
# PARSER — mendukung beberapa "keluarga" format RKP:
#   Tipe "periode"   : breakdown Tahun > Catur Wulan > Fisik/Biaya (format asli)
#   Tipe "tahunan"   : breakdown per Tahun langsung (kolom tahun 4-digit) + kolom TOTAL
#   Tipe "sederhana" : daftar Uraian/Pekerjaan + Volume(opsional) + Biaya, tanpa jadwal waktu
# ============================================================
def norm(v):
    return "" if v is None else str(v).strip().lower()


def _is_leaf_label(name):
    """False untuk baris subtotal/total/grand total (bukan item asli, untuk hindari double count)."""
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    return not re.match(r"^(total|tatal|sub\s*total|grand\s*total|jumlah)\b", s, re.I)


def _extract_reconciled_items(rows, start, end, name_col, biaya_col, vol_col, grand_biaya, period_builder=None):
    """Bangun daftar item pekerjaan untuk ditampilkan di rekap. File dengan subtotal berlapis
    (mis. leaf -> Sub Total -> Total -> GRAND TOTAL) akan double/under-count kalau langsung
    dijumlah di level rincian paling bawah. Di sini dicoba beberapa level granularitas — Sub
    Total saja, Total saja, gabungan keduanya, atau leaf polos — dan dipilih yang levelnya
    paling ringkas (sesuai permintaan: level 'sub', bukan rincian sampai ke akar) SELAMA
    jumlahnya masih merekonsiliasi (mendekati) Grand Total; kalau tidak ada yang cocok,
    fallback ke leaf item (perilaku lama)."""
    def mk_item(row, name):
        vol = row[vol_col] if (vol_col is not None and vol_col < len(row)) else None
        vol = vol if isinstance(vol, (int, float)) else None
        biaya = row[biaya_col] if biaya_col < len(row) else None
        biaya = biaya if isinstance(biaya, (int, float)) else None
        periods = period_builder(row) if period_builder else []
        return dict(no=None, nama=str(name).strip(), volume_ha=vol, biaya_rencana=biaya,
                    rp_per_ha=(biaya / vol) if (biaya and vol) else None, periods=periods)

    leaf_items, subtotal_items, total_items = [], [], []
    for r in range(start, end):
        row = rows[r]
        name = row[name_col] if name_col < len(row) else None
        biaya = row[biaya_col] if biaya_col < len(row) else None
        vol = row[vol_col] if (vol_col is not None and vol_col < len(row)) else None
        if name is None or str(name).strip() == "":
            continue
        if not isinstance(biaya, (int, float)) and not isinstance(vol, (int, float)):
            continue
        s = str(name).strip()
        if re.match(r"^sub\s*total\b", s, re.I):
            subtotal_items.append(mk_item(row, name))
        elif re.match(r"^(grand\s*total|jumlah)\b", s, re.I):
            continue
        elif re.match(r"^(total|tatal)\b", s, re.I):
            total_items.append(mk_item(row, name))
        else:
            leaf_items.append(mk_item(row, name))

    def diff_of(items):
        if not items or not grand_biaya:
            return None
        s = sum(it["biaya_rencana"] for it in items if it["biaya_rencana"])
        return abs(s - grand_biaya) / abs(grand_biaya)

    # Urutan preferensi: level "sub" paling diutamakan (sesuai permintaan), baru "total",
    # gabungan keduanya, dan leaf polos paling akhir — dipilih kandidat pertama yang
    # merekonsiliasi dalam toleransi 3%; kalau tidak ada, pilih yang selisihnya terkecil.
    candidates = [
        ("sub_total", subtotal_items),
        ("total", total_items),
        ("sub_total+total", subtotal_items + total_items if (subtotal_items and total_items) else []),
        ("leaf", leaf_items),
    ]
    within_tol = [(label, items, d) for label, items, in candidates for d in [diff_of(items)] if d is not None and d <= 0.03]
    if within_tol:
        label, items, diff = within_tol[0]
        return items, True
    scored = [(label, items, diff_of(items)) for label, items in candidates if items]
    scored = [t for t in scored if t[2] is not None]
    if scored:
        label, items, diff = min(scored, key=lambda t: t[2])
        return items, diff <= 0.03
    return leaf_items, _items_reliable(leaf_items, grand_biaya)


def find_header_row(rows, names=("pekerjaan", "item", "uraian"), limit=25):
    for r in range(min(len(rows), limit)):
        row = rows[r]
        for c, v in enumerate(row):
            if norm(v) in names:
                return r, c
    return -1, -1


def parse_rkp_rows(rows):
    """Tipe 'periode': Tahun > Catur Wulan > Fisik/Biaya (RKP asli & RKP 08 / RKP FFD / RKP BKB)."""
    hr, name_col = find_header_row(rows, names=("pekerjaan", "item"))
    if hr == -1:
        return None

    tahun_row = rows[hr] if hr < len(rows) else []
    periode_row = rows[hr + 1] if hr + 1 < len(rows) else []
    sub_row = rows[hr + 2] if hr + 2 < len(rows) else []

    vol_col = next((c for c, v in enumerate(sub_row) if norm(v) == "volume"), -1)
    if vol_col == -1:
        return None
    biaya_col = vol_col + 1

    fisik_cols = [c for c, v in enumerate(sub_row) if norm(v) == "fisik"]
    if not fisik_cols:
        return None  # bukan tipe periode — biarkan dicoba tipe lain oleh dispatcher

    maxc = max(len(tahun_row), len(periode_row), len(sub_row))
    last_t, last_p = None, None
    filled_t, filled_p = {}, {}
    for c in range(fisik_cols[0], maxc):
        if c < len(tahun_row) and tahun_row[c] not in (None, ""):
            last_t = tahun_row[c]
        if c < len(periode_row) and periode_row[c] not in (None, ""):
            last_p = periode_row[c]
        filled_t[c] = last_t
        filled_p[c] = last_p

    periods = []
    for fc in fisik_cols:
        bc = fc + 1
        tl = str(filled_t.get(fc) or "")
        pl = str(filled_p.get(fc) or "")
        is_cw = bool(re.search(r"catur\s*wulan", pl, re.I))
        year_m = re.search(r"(\d{4})", tl)
        cw_m = re.search(r"catur\s*wulan\s*(\d+)", pl, re.I)
        year = int(year_m.group(1)) if year_m else None
        cw = int(cw_m.group(1)) if cw_m else None
        display = f"{year or '?'} CW{cw or '?'}" if is_cw else pl.strip()
        periods.append(
            dict(fisik_col=fc, biaya_col=bc, is_cw=is_cw,
                 sort_key=(year or 0) * 10 + (cw or 0), display=display)
        )
    cw_periods = [p for p in periods if p["is_cw"]]

    grand_row = -1
    for r in range(hr + 3, len(rows)):
        row = rows[r]
        label = row[1] if len(row) > 1 else None
        label2 = row[2] if len(row) > 2 else None
        if (label and re.search(r"grand.?total", str(label), re.I)) or (
            label2 and re.search(r"grand.?total", str(label2), re.I)
        ):
            grand_row = r
            break

    end = grand_row if grand_row != -1 else len(rows)

    def build_periods_for_row(row):
        per = []
        for p in cw_periods:
            f = row[p["fisik_col"]] if p["fisik_col"] < len(row) else None
            b = row[p["biaya_col"]] if p["biaya_col"] < len(row) else None
            per.append(dict(key=p["display"], sort_key=p["sort_key"],
                             fisik=f if isinstance(f, (int, float)) else 0,
                             biaya=b if isinstance(b, (int, float)) else 0))
        return per

    grand = None
    if grand_row != -1:
        row = rows[grand_row]
        gvol = row[vol_col] if vol_col < len(row) else None
        gbiaya = row[biaya_col] if biaya_col < len(row) else None
        grand = dict(volume_ha=gvol if isinstance(gvol, (int, float)) else None,
                     biaya_rencana=gbiaya if isinstance(gbiaya, (int, float)) else None,
                     periods=build_periods_for_row(row))

    items, items_reliable = _extract_reconciled_items(
        rows, hr + 3, end, name_col, biaya_col, vol_col,
        grand["biaya_rencana"] if grand else None, period_builder=build_periods_for_row,
    )
    if not items:
        return None

    if grand is None:
        # fallback: jumlahkan item kalau tidak ada baris GRAND TOTAL eksplisit
        grand = dict(
            volume_ha=sum(it["volume_ha"] for it in items if it["volume_ha"]) or None,
            biaya_rencana=sum(it["biaya_rencana"] for it in items if it["biaya_rencana"]) or None,
            periods=[],
        )

    return dict(items=items, grand=grand, items_reliable=items_reliable)


def parse_simple_list_rows(rows):
    """Tipe 'sederhana': Uraian/Pekerjaan/Item + Volume(opsional) + Biaya (atau 'Budget RKP'),
    tanpa breakdown waktu. Mendukung kolom Realisasi inline kalau ada (mis. template Dermaga)."""
    hr, name_col = find_header_row(rows, names=("pekerjaan", "item", "uraian"))
    if hr == -1:
        return None

    search_rows = [r for r in rows[hr:hr + 4] if r]

    def find_col(patterns, exact=True):
        for rr in search_rows:
            for c, v in enumerate(rr):
                t = norm(v)
                if not t:
                    continue
                for p in patterns:
                    if (exact and t == p) or (not exact and p in t):
                        return c
        return None

    vol_col = find_col(("volume",), exact=True)
    biaya_col = find_col(("biaya",), exact=True)
    if biaya_col is None:
        biaya_col = find_col(("budget rkp",), exact=False)
    if biaya_col is None:
        return None

    realisasi_cols = []
    for rr in search_rows:
        for c, v in enumerate(rr):
            t = norm(v)
            if t and "realisasi" in t and c not in realisasi_cols:
                realisasi_cols.append(c)

    items = []
    for r in range(hr + 1, len(rows)):
        row = rows[r]
        name = row[name_col] if name_col < len(row) else None
        biaya = row[biaya_col] if biaya_col < len(row) else None
        if not _is_leaf_label(name) or not isinstance(biaya, (int, float)):
            continue
        vol = row[vol_col] if (vol_col is not None and vol_col < len(row)) else None
        vol = vol if isinstance(vol, (int, float)) else None
        real_val, found_real = 0, False
        for rc in realisasi_cols:
            v = row[rc] if rc < len(row) else None
            if isinstance(v, (int, float)):
                real_val += v
                found_real = True
        items.append(dict(
            no=None, nama=str(name).strip(), volume_ha=vol, biaya_rencana=biaya,
            rp_per_ha=(biaya / vol) if (vol and biaya) else None, periods=[],
            _realisasi=(real_val if found_real else None),
        ))
    if not items:
        return None

    # Grand biaya: baris "total/subtotal/grand total" dgn biaya TERBESAR = grand total sebenarnya.
    total_rows = []
    for r in range(hr + 1, len(rows)):
        row = rows[r]
        name = row[name_col] if name_col < len(row) else None
        biaya = row[biaya_col] if biaya_col < len(row) else None
        if name and isinstance(biaya, (int, float)) and re.match(r"^(total|tatal|sub\s*total|grand\s*total|jumlah)\b", str(name).strip(), re.I):
            total_rows.append(biaya)
    grand_biaya = max(total_rows) if total_rows else None
    if grand_biaya is None:
        grand_biaya = sum(it["biaya_rencana"] for it in items if it["biaya_rencana"])

    # Grand luas: JANGAN jumlah semua item — kalau beberapa aktivitas berbeda diterapkan pada
    # plot yang sama, volume Ha-nya akan berulang identik di tiap baris. Jumlahkan nilai
    # volume yang UNIK saja (lebih tahan banting daripada mengandalkan label subtotal, yang
    # kadang typo di file sumber, mis. "Tatal Biaya X" alih-alih "Total Biaya X").
    grand_vol = None
    if vol_col is not None:
        distinct_vols = sorted({it["volume_ha"] for it in items if it["volume_ha"]})
        grand_vol = sum(distinct_vols) if distinct_vols else None

    inline_realisasi_total = None
    if realisasi_cols:
        vals = [it["_realisasi"] for it in items if it["_realisasi"]]
        inline_realisasi_total = sum(vals) if vals else 0

    # Untuk ditampilkan (rekap pekerjaan): pakai level yang paling ringkas yang masih
    # merekonsiliasi ke grand_biaya (Sub Total/Total), bukan rincian leaf paling detail.
    display_items, items_reliable = _extract_reconciled_items(
        rows, hr + 1, len(rows), name_col, biaya_col, vol_col, grand_biaya,
    )

    return dict(
        items=display_items,
        grand=dict(volume_ha=grand_vol, biaya_rencana=grand_biaya, periods=[]),
        inline_realisasi_total=inline_realisasi_total,
        items_reliable=items_reliable,
    )


def parse_yearly_coa_rows(rows):
    """Tipe 'tahunan': tabel biaya per Tahun (kolom tahun 4-digit langsung) + kolom TOTAL
    (mis. sheet REKAP TAHUNAN pada proyek konstruksi/restorasi)."""
    hr, name_col = find_header_row(rows, names=("pekerjaan", "item", "uraian"))
    if hr == -1:
        return None

    search_rows = [r for r in rows[hr:hr + 3] if r]
    year_cols = []
    for rr in search_rows:
        for c, v in enumerate(rr):
            if isinstance(v, (int, float)) and not isinstance(v, bool) and 1990 <= v <= 2100 and float(v).is_integer():
                if not any(c == cc for cc, _ in year_cols):
                    year_cols.append((c, int(v)))
    if len(year_cols) < 2:
        return None

    total_col = None
    for rr in search_rows:
        for c, v in enumerate(rr):
            if norm(v) == "total":
                total_col = c
                break
        if total_col is not None:
            break

    def row_periods(row):
        return [dict(key=f"Tahun {yr}", sort_key=yr, fisik=0,
                     biaya=(row[c] if c < len(row) and isinstance(row[c], (int, float)) else 0))
                for c, yr in year_cols]

    grand_biaya, grand_periods = None, None
    for r in range(hr + 1, len(rows)):
        row = rows[r]
        name = row[name_col] if name_col < len(row) else None
        if name and re.match(r"^(total|tatal|sub\s*total|grand\s*total|jumlah)\b", str(name).strip(), re.I):
            tot_val = row[total_col] if (total_col is not None and total_col < len(row)) else None
            if isinstance(tot_val, (int, float)) and (grand_biaya is None or tot_val > grand_biaya):
                grand_biaya = tot_val
                grand_periods = row_periods(row)

    items, items_reliable = _extract_reconciled_items(
        rows, hr + 1, len(rows), name_col, total_col, None, grand_biaya, period_builder=row_periods,
    )
    if not items:
        return None

    if grand_biaya is None:
        grand_biaya = sum(it["biaya_rencana"] for it in items if it["biaya_rencana"])
        grand_periods = [dict(key=f"Tahun {yr}", sort_key=yr, fisik=0,
                               biaya=sum(it["periods"][i]["biaya"] for it in items))
                          for i, (c, yr) in enumerate(year_cols)]

    return dict(items=items, grand=dict(volume_ha=None, biaya_rencana=grand_biaya, periods=grand_periods),
                items_reliable=items_reliable)


def _items_reliable(items, grand_biaya):
    """False kalau jumlah biaya semua item menyimpang >3% dari grand total — pertanda struktur
    subtotal berlapis di file sumber tidak terbaca bersih (dobel hitung atau ada yang kelewat),
    supaya dashboard tidak menampilkan rincian per-pekerjaan seolah pasti akurat."""
    if not grand_biaya:
        return True
    s = sum(it["biaya_rencana"] for it in items if it["biaya_rencana"])
    return abs(s - grand_biaya) <= 0.03 * abs(grand_biaya)


def parse_timeline_rkp_rows(rows):
    """Tipe 'timeline': daftar Pekerjaan (kolom A) + Biaya Rupiah (kolom di sebelahnya
    yang headernya persis kata 'Rupiah'), tanpa header 'Pekerjaan/Item/Uraian' baku
    (mis. template 'TIMELINE RKP' — kolom breakdown bulanannya biasanya kosong/template)."""
    header_row, biaya_col = None, None
    for r in range(min(len(rows), 15)):
        row = rows[r]
        for c, v in enumerate(row):
            if norm(v) == "rupiah":
                header_row, biaya_col = r, c
                break
        if header_row is not None:
            break
    if header_row is None:
        return None
    name_col = 0

    items = []
    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        name = row[name_col] if name_col < len(row) else None
        biaya = row[biaya_col] if biaya_col < len(row) else None
        if not _is_leaf_label(name) or not isinstance(biaya, (int, float)):
            continue
        items.append(dict(no=None, nama=str(name).strip(), volume_ha=None,
                           biaya_rencana=biaya, rp_per_ha=None, periods=[]))
    if not items:
        return None

    total_rows = []
    for r in range(header_row + 1, len(rows)):
        row = rows[r]
        name = row[name_col] if name_col < len(row) else None
        biaya = row[biaya_col] if biaya_col < len(row) else None
        if name and isinstance(biaya, (int, float)) and re.match(r"^(total|tatal|sub\s*total|grand\s*total|jumlah)\b", str(name).strip(), re.I):
            total_rows.append(biaya)
    grand_biaya = max(total_rows) if total_rows else sum(it["biaya_rencana"] for it in items)

    return dict(items=items, grand=dict(volume_ha=None, biaya_rencana=grand_biaya, periods=[]),
                items_reliable=_items_reliable(items, grand_biaya))


def extract_meta(rows, file_name):
    company, desc, luas_text, periode_text = None, None, None, None
    for r in range(min(len(rows), 10)):
        row = rows[r]
        for c, v in enumerate(row):
            if v is None:
                continue
            s = str(v).strip()
            if not s:
                continue
            if re.search(r"rencana kerja proyek", s, re.I):
                continue
            if re.match(r"^pt[\s.]", s, re.I) and not company:
                company = s
                continue
            if re.match(r"^luas\s*:?\s*$", s, re.I):
                for c2 in range(c + 1, len(row)):
                    if row[c2] not in (None, "") and str(row[c2]).strip() != "":
                        luas_text = str(row[c2]).strip()
                        break
            if re.search(r"tahun|periode", s, re.I) and re.search(r"\d{4}", s) and not periode_text:
                periode_text = s
            if 2 <= r <= 4 and not desc and not re.search(r"luas|tahun", s, re.I) and not re.match(r"^pt[\s.]", s, re.I):
                desc = s

    luas_num = None
    if luas_text:
        m = re.search(r"[\d.]+", luas_text.replace(".", "").replace(",", "."))
        if m:
            try:
                luas_num = float(m.group(0))
            except ValueError:
                luas_num = None

    name = re.sub(r"\.xlsx$", "", file_name, flags=re.I)
    name = re.sub(r"^RKP[_\s]?", "", name, flags=re.I)
    name = re.sub(r"_+", " ", name).strip()

    return dict(name=name, company=company or "—", desc=desc or "", luas_num=luas_num,
                luas_text=luas_text, periode_text=periode_text)


def try_parse_realisasi(wb):
    """Cari sheet terpisah bernama mengandung 'realisasi'/'aktual' dengan struktur tipe 'periode'
    yang SAMA seperti sheet RKP-nya. Sengaja tidak mencoba tipe lain di sini, supaya sheet yang
    ternyata milik proyek lain (leftover template) tidak ikut kebaca sebagai angka valid."""
    real_sheet = next((n for n in wb.sheetnames if re.search(r"realisasi|aktual", n, re.I)), None)
    if not real_sheet:
        return dict(status="none")
    try:
        ws = wb[real_sheet]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        parsed = parse_rkp_rows(rows)
        if not parsed:
            return dict(status="unrecognized", sheet=real_sheet)
        return dict(status="ok", data=parsed, sheet=real_sheet)
    except Exception:
        return dict(status="unrecognized", sheet=real_sheet)


def _gather_rkp_sheet_candidates(sheetnames):
    """Urutan prioritas kandidat sheet utama: 'RKP' persis, lalu sheet berawalan 'RKP',
    lalu sheet 'REKAP TAHUNAN' / 'REKAP PEKERJAAN' (format proyek konstruksi). Kalau tidak
    ada satu pun yang cocok (mis. sheet cuma bernama 'Sheet1'), coba semua sheet sebagai
    upaya terakhir — aman karena tetap harus lolos salah satu detektor format di bawah."""
    cands = []
    for n in sheetnames:
        if n.strip().upper() == "RKP" and n not in cands:
            cands.append(n)
    for n in sheetnames:
        if n.strip().upper().startswith("RKP") and n not in cands:
            cands.append(n)
    for n in sheetnames:
        if re.search(r"rekap\s*tahunan", n, re.I) and n not in cands:
            cands.append(n)
    for n in sheetnames:
        if re.search(r"rekap\s*pekerjaan", n, re.I) and n not in cands:
            cands.append(n)
    if not cands:
        cands = list(sheetnames)
    return cands


def parse_workbook(file_bytes, file_name):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    candidates = _gather_rkp_sheet_candidates(wb.sheetnames)
    if not candidates:
        wb.close()
        return dict(error='Tidak ditemukan sheet RKP/REKAP di file ini.')

    chosen_sheet, chosen_rows, rencana, fmt = None, None, None, None
    for sheet_name in candidates:
        rows = [list(r) for r in wb[sheet_name].iter_rows(values_only=True)]
        parsed = parse_rkp_rows(rows)
        if parsed:
            chosen_sheet, chosen_rows, rencana, fmt = sheet_name, rows, parsed, "periode"
            break
        parsed = parse_yearly_coa_rows(rows)
        if parsed:
            chosen_sheet, chosen_rows, rencana, fmt = sheet_name, rows, parsed, "tahunan"
            break
        parsed = parse_simple_list_rows(rows)
        if parsed:
            chosen_sheet, chosen_rows, rencana, fmt = sheet_name, rows, parsed, "sederhana"
            break
        parsed = parse_timeline_rkp_rows(rows)
        if parsed:
            chosen_sheet, chosen_rows, rencana, fmt = sheet_name, rows, parsed, "sederhana"
            break

    if rencana is None:
        wb.close()
        return dict(error=(
            f"Format tidak dikenali pada sheet yang dicoba ({', '.join(candidates)}). "
            "Bukan format periode (Catur Wulan), tahunan, atau daftar biaya sederhana yang didukung."
        ))

    meta = extract_meta(chosen_rows, file_name)
    if file_name in COMPANY_OVERRIDES:
        meta["company"] = COMPANY_OVERRIDES[file_name]
    meta["company"] = CANONICAL_COMPANIES.get(_company_key(meta["company"]), meta["company"])
    if file_name in PROJECT_NAME_OVERRIDES:
        meta["name"] = PROJECT_NAME_OVERRIDES[file_name]

    # Realisasi dinonaktifkan sementara atas permintaan — dikosongkan dulu semua,
    # nanti diaktifkan lagi setelah file/sheet realisasi resmi diupload terpisah.
    if ENABLE_REALISASI and fmt == "sederhana" and rencana.get("inline_realisasi_total") is not None:
        total_real = rencana["inline_realisasi_total"]
        realisasi = dict(
            status="total_only",
            data=dict(items=[], grand=dict(volume_ha=None, biaya_rencana=total_real, periods=[])),
        )
    elif ENABLE_REALISASI:
        realisasi = try_parse_realisasi(wb)
    else:
        realisasi = dict(status="none")

    wb.close()
    return dict(id=file_name, file_name=file_name, updated_at=datetime.now().isoformat(),
                meta=meta, rencana=rencana, realisasi=realisasi, format=fmt, sheet=chosen_sheet)


# ============================================================
# GABUNG PROYEK — beberapa file yang sebenarnya satu proyek (mis. Plasma Mandiri
# TR + Sucab) digabung jadi satu entri: Luas, Biaya, dan rincian pekerjaan dijumlahkan.
# ============================================================
MERGE_GROUPS = [
    dict(
        id="__MERGED_PLASMA_MANDIRI__",
        name="Pembukaan Lahan Plasma Mandiri",
        files=["RKP_PLASMA_MANDIRI_TR_200_Ha.xlsx", "RKP_PLASMA_MANDIRI_SUCAB__300_Ha.xlsx"],
    ),
]


def _merge_period_lists(period_lists):
    merged = {}
    for periods in period_lists:
        for pd_ in periods:
            k = pd_["key"]
            if k not in merged:
                merged[k] = dict(key=k, sort_key=pd_["sort_key"], fisik=0, biaya=0)
            merged[k]["fisik"] += pd_.get("fisik") or 0
            merged[k]["biaya"] += pd_.get("biaya") or 0
    return sorted(merged.values(), key=lambda x: x["sort_key"])


def _merge_items(items_lists):
    merged, order = {}, []
    for items in items_lists:
        for it in items:
            k = it["nama"]
            if k not in merged:
                merged[k] = dict(volume_ha=0, biaya_rencana=0, periods_lists=[])
                order.append(k)
            if it.get("volume_ha"):
                merged[k]["volume_ha"] += it["volume_ha"]
            if it.get("biaya_rencana"):
                merged[k]["biaya_rencana"] += it["biaya_rencana"]
            merged[k]["periods_lists"].append(it.get("periods") or [])
    result = []
    for k in order:
        m = merged[k]
        vol = m["volume_ha"] or None
        biaya = m["biaya_rencana"] or None
        result.append(dict(
            no=None, nama=k, volume_ha=vol, biaya_rencana=biaya,
            rp_per_ha=(biaya / vol) if (biaya and vol) else None,
            periods=_merge_period_lists(m["periods_lists"]),
        ))
    return result


def merge_projects(project_list, merged_id, merged_name):
    base = project_list[0]
    grand_biaya = sum((p["rencana"]["grand"]["biaya_rencana"] or 0) for p in project_list) or None
    grand_vol = sum((p["rencana"]["grand"]["volume_ha"] or 0) for p in project_list) or None
    grand_periods = _merge_period_lists([p["rencana"]["grand"]["periods"] for p in project_list])
    merged_items = _merge_items([p["rencana"]["items"] for p in project_list])

    desc_parts = sorted({p["meta"]["desc"] for p in project_list if p["meta"]["desc"]})
    merged_meta = dict(
        name=merged_name,
        company=base["meta"]["company"],
        desc=" + ".join(desc_parts),
        luas_num=grand_vol,
        luas_text=None,
        periode_text=base["meta"]["periode_text"],
    )
    return dict(
        id=merged_id,
        file_name=" + ".join(p["file_name"] for p in project_list),
        updated_at=datetime.now().isoformat(),
        meta=merged_meta,
        rencana=dict(
            items=merged_items,
            grand=dict(volume_ha=grand_vol, biaya_rencana=grand_biaya, periods=grand_periods),
            items_reliable=_items_reliable(merged_items, grand_biaya),
        ),
        realisasi=dict(status="none"),
        format=base["format"],
        sheet="(gabungan " + str(len(project_list)) + " file)",
    )


def apply_merge_groups(projects):
    """projects: dict id -> project. Mengembalikan dict baru dengan grup di MERGE_GROUPS digabung
    jadi satu entri (hanya kalau SEMUA file anggota grup itu ada di data yang dimuat)."""
    result = dict(projects)
    for group in MERGE_GROUPS:
        members = [result[f] for f in group["files"] if f in result]
        if len(members) == len(group["files"]) and len(members) >= 2:
            for f in group["files"]:
                del result[f]
            result[group["id"]] = merge_projects(members, group["id"], group["name"])
    return result


# ============================================================
# DERIVED HELPERS
# ============================================================
def total_rencana(p):
    return p["rencana"]["grand"]["biaya_rencana"] if p["rencana"]["grand"] else None


def luas_proj(p):
    g = p["rencana"]["grand"]
    return g["volume_ha"] if g and g["volume_ha"] else p["meta"]["luas_num"]


def rp_per_ha(p):
    t, l = total_rencana(p), luas_proj(p)
    return t / l if t and l else None


def has_realisasi(p):
    """True hanya untuk realisasi yang cocok item-per-item (status 'ok')."""
    return p["realisasi"]["status"] == "ok"


def has_any_realisasi(p):
    """True kalau ada angka realisasi sama sekali, walau cuma total (status 'ok' atau 'total_only')."""
    return p["realisasi"]["status"] in ("ok", "total_only")


def realisasi_total(p):
    if not has_any_realisasi(p):
        return None
    g = p["realisasi"]["data"]["grand"]
    return g["biaya_rencana"] if g else None


def capaian_biaya_pct(p):
    """% capaian realisasi biaya terhadap rencana (biaya), atau None kalau data belum ada."""
    rt = realisasi_total(p)
    rc = total_rencana(p)
    if rt is None or not rc:
        return None
    return rt / rc * 100


def capaian_fisik_pct(p):
    """% capaian realisasi fisik (Ha) terhadap rencana fisik — hanya untuk proyek dengan
    realisasi per-periode yang cocok (status 'ok') dan format berjadwal (Catur Wulan/Tahun)."""
    if not has_realisasi(p):
        return None
    rg = p["realisasi"]["data"]["grand"]
    pg = p["rencana"]["grand"]
    if not rg or not pg or not pg.get("periods"):
        return None
    rencana_fisik = sum(pd_["fisik"] for pd_ in pg["periods"])
    real_fisik = sum(pd_["fisik"] for pd_ in rg["periods"])
    if not rencana_fisik:
        return None
    return real_fisik / rencana_fisik * 100


def all_period_keys(projects):
    seen = {}
    for p in projects.values():
        g = p["rencana"]["grand"]
        if not g:
            continue
        for pd_ in g["periods"]:
            if pd_["key"] not in seen or pd_["sort_key"] < seen[pd_["key"]]:
                seen[pd_["key"]] = pd_["sort_key"]
    return [k for k, _ in sorted(seen.items(), key=lambda x: x[1])]


# ---------------- Jenis proyek (untuk pengelompokan) ----------------
JENIS_COLORS = {
    "Tanaman": FOREST,
    "Infrastruktur": GOLD,
}


def project_jenis(p):
    fmt = p.get("format")
    if fmt == "periode":
        return "Tanaman"
    if fmt == "sederhana" and luas_proj(p):
        return "Tanaman"
    return "Infrastruktur"


# ---------------- Periode "sekarang" & rentang tahun proyek ----------------
def current_period_label(fmt):
    """Label periode berjalan saat ini, dalam bentuk yang sama seperti kunci periode
    proyek (mis. '2026 CW3' untuk tipe periode, 'Tahun 2026' untuk tipe tahunan)."""
    now = datetime.now()
    if fmt == "tahunan":
        return f"Tahun {now.year}"
    cw = 1 if now.month <= 4 else (2 if now.month <= 8 else 3)
    return f"{now.year} CW{cw}"


def add_now_marker(fig, keys, fmt):
    """Tambahkan garis vertikal putus-putus penanda periode berjalan saat ini, kalau
    periode itu ada di daftar kunci sumbu-x grafik (kalau tidak ada, dilewati saja)."""
    label = current_period_label(fmt)
    if label in keys:
        fig.add_shape(
            type="line", x0=label, x1=label, xref="x", y0=0, y1=1, yref="paper",
            line=dict(color=RUST, width=2, dash="dot"),
        )
        fig.add_annotation(
            x=label, y=1, xref="x", yref="paper", yanchor="bottom",
            text="Sekarang", showarrow=False, font=dict(color=RUST, size=11),
        )


def project_year_range(p):
    """(tahun_mulai, tahun_selesai) dari kunci periode proyek, atau None kalau proyek
    ini tidak punya jadwal periode sama sekali (mis. Pastura, Biogas, Dermaga)."""
    g = p["rencana"]["grand"]
    if not g or not g["periods"]:
        return None
    years = [int(m.group(1)) for pd_ in g["periods"] if (m := re.search(r"(\d{4})", pd_["key"]))]
    if not years:
        return None
    return min(years), max(years)


def project_time_progress_pct(p):
    """Persentase waktu yang sudah lewat dari total durasi rencana proyek (berdasarkan
    kalender), atau None kalau proyek tidak punya rentang tahun yang jelas."""
    yr = project_year_range(p)
    if not yr:
        return None
    start_y, end_y = yr
    now = datetime.now()
    frac_now = now.year + (now.month - 1) / 12
    total_span = (end_y + 1) - start_y
    if total_span <= 0:
        return None
    return max(0, min(100, (frac_now - start_y) / total_span * 100))


SECTIONS = ["🏠 Beranda", "📋 Portofolio", "🏢 Per Perusahaan", "📁 Detail Proyek"]


# ============================================================
# DATA DARI REPO GITHUB (folder data/)
# ============================================================
def _repo_cache_key():
    """Kunci cache berbasis nama+ukuran+waktu-modifikasi file di folder data/.
    Berubah otomatis begitu file di-update lewat GitHub -> cache re-parse."""
    if not DATA_DIR.exists():
        return "no-data-dir"
    parts = []
    for fp in sorted(DATA_DIR.glob("*.xlsx")):
        stat = fp.stat()
        parts.append(f"{fp.name}:{stat.st_mtime}:{stat.st_size}")
    return hashlib.md5("|".join(parts).encode()).hexdigest()


@st.cache_data(show_spinner="Membaca file RKP dari repo GitHub...")
def load_repo_projects(_cache_key):
    result = {}
    errs = []
    if not DATA_DIR.exists():
        return result, errs
    for fp in sorted(DATA_DIR.glob("*.xlsx")):
        try:
            parsed = parse_workbook(fp.read_bytes(), fp.name)
            if "error" in parsed:
                errs.append(f"{fp.name}: {parsed['error']}")
                continue
            result[parsed["id"]] = parsed
        except Exception as e:
            errs.append(f"{fp.name}: gagal dibaca ({e})")
    return result, errs


# ============================================================
# STATE
# ============================================================
if "session_projects" not in st.session_state:
    st.session_state.session_projects = {}  # file uji coba sementara (tidak permanen)

repo_projects, repo_errs = load_repo_projects(_repo_cache_key())

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("### 🌴 RKP Monitor")

    if repo_projects:
        st.success(f"📁 {len(repo_projects)} file RKP dimuat dari repo GitHub (folder `data/`)")
    elif DATA_DIR.exists():
        st.warning("Folder `data/` ada tapi belum berisi file .xlsx.")
    else:
        st.info("Folder `data/` belum ada di repo. Lihat README untuk cara menambahkannya.")
    for e in repo_errs:
        st.warning(e)

    if st.button("🔄 Muat ulang data dari GitHub", use_container_width=True):
        load_repo_projects.clear()
        st.rerun()

    st.divider()
    st.caption("Untuk update data secara permanen: ganti/tambah file `.xlsx` di folder `data/` pada repo GitHub, lalu commit — dashboard otomatis membaca versi terbaru dalam beberapa menit.")

    with st.expander("🧪 Uji coba file lain (sementara, tidak permanen)"):
        uploaded = st.file_uploader("Upload file RKP", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
        if uploaded:
            ok, errs = 0, []
            for f in uploaded:
                try:
                    result = parse_workbook(f.read(), f.name)
                    if "error" in result:
                        errs.append(f"{f.name}: {result['error']}")
                        continue
                    st.session_state.session_projects[result["id"]] = result
                    ok += 1
                except Exception as e:
                    errs.append(f"{f.name}: gagal dibaca ({e})")
            if ok:
                st.success(f"{ok} file berhasil diproses (sesi ini saja).")
            for e in errs:
                st.warning(e)
        if st.session_state.session_projects:
            st.caption(f"{len(st.session_state.session_projects)} file uji coba aktif")
            if st.button("Hapus file uji coba", use_container_width=True):
                st.session_state.session_projects = {}
                st.rerun()

# Gabungkan: data dari repo GitHub (utama) + file uji coba sesi (opsional, menimpa nama file yang sama)
projects_all = apply_merge_groups({**repo_projects, **st.session_state.session_projects})

# ============================================================
# SIDEBAR — FILTER
# ============================================================
projects = projects_all
if projects_all:
    with st.sidebar:
        st.divider()
        st.markdown("#### 🔎 Filter")

        jenis_list = sorted({project_jenis(p) for p in projects_all.values()})
        selected_jenis = st.multiselect(
            "Jenis Proyek", jenis_list, default=jenis_list, key="filter_jenis"
        )
        projects_by_jenis = {
            k: v for k, v in projects_all.items() if project_jenis(v) in selected_jenis
        }

        companies = sorted({p["meta"]["company"] for p in projects_by_jenis.values()})
        selected_companies = st.multiselect(
            "Perusahaan", companies, default=companies, key="filter_company"
        )

        projects_by_company = {
            k: v for k, v in projects_by_jenis.items() if v["meta"]["company"] in selected_companies
        }

        project_names = sorted({p["meta"]["name"] for p in projects_by_company.values()})
        selected_projects = st.multiselect(
            "Proyek", project_names, default=project_names, key="filter_project"
        )

        projects = {
            k: v for k, v in projects_by_company.items() if v["meta"]["name"] in selected_projects
        }

        if projects_all and not projects:
            st.warning("Tidak ada proyek yang cocok dengan filter ini.")

        if projects:
            st.markdown("#### 🧭 Navigasi")
            if st.session_state.get("section") not in SECTIONS:
                st.session_state["section"] = SECTIONS[0]
            st.radio("Halaman", SECTIONS, key="section", label_visibility="collapsed")

            if st.session_state["section"] == "📁 Detail Proyek":
                proj_names = [p["meta"]["name"] for p in projects.values()]
                if st.session_state.get("detail_project") not in proj_names:
                    st.session_state["detail_project"] = proj_names[0]
                st.selectbox("Pilih proyek", proj_names, key="detail_project")


# ============================================================
# MAIN
# ============================================================
if not projects_all:
    st.title("🌴 Monitoring RKP")
    st.info(
        "⬅️ Belum ada data. Tambahkan file `.xlsx` RKP ke folder **`data/`** di repo GitHub lalu "
        "commit, atau upload file uji coba lewat panel kiri untuk mulai memantau."
    )
    st.stop()

if not projects:
    st.title("🌴 Monitoring RKP")
    st.warning("Tidak ada proyek yang cocok dengan filter Perusahaan/Proyek yang dipilih di sidebar. Coba longgarkan filternya.")
    st.stop()

section = st.session_state.get("section", SECTIONS[0])


def _go_section(name):
    st.session_state["section"] = name


# ================================================================
# BERANDA (ringkasan eksekutif)
# ================================================================
if section == "🏠 Beranda":
    total_biaya = sum((total_rencana(p) or 0) for p in projects.values())
    total_luas = sum((luas_proj(p) or 0) for p in projects.values())
    n_real = sum(1 for p in projects.values() if has_any_realisasi(p))
    by_jenis = {}
    for p in projects.values():
        by_jenis[project_jenis(p)] = by_jenis.get(project_jenis(p), 0) + (total_rencana(p) or 0)
    tanaman_pct = (by_jenis.get("Tanaman", 0) / total_biaya * 100) if total_biaya else 0

    st.markdown(
        f"""
        <div class="hero-band">
          <div>
            <p class="hero-eyebrow">🌴 Monitoring RKP &middot; {len(projects)} proyek terfilter</p>
            <p class="hero-number">{fmt_rp(total_biaya)}</p>
            <p class="hero-sub">Total investasi portofolio pada {datetime.now().strftime('%d %B %Y')} — {fmt_rp_full(total_biaya)}, dari {len(projects)} proyek di {len({p['meta']['company'] for p in projects.values()})} perusahaan.</p>
          </div>
          <div class="hero-split">
            <div class="hero-split-item">
              <div class="hero-split-num">{tanaman_pct:.0f}%</div>
              <div class="hero-split-lbl">Tanaman</div>
            </div>
            <div class="hero-split-item">
              <div class="hero-split-num">{100-tanaman_pct:.0f}%</div>
              <div class="hero-split-lbl">Infrastruktur</div>
            </div>
            <div class="hero-split-item">
              <div class="hero-split-num">{n_real}/{len(projects)}</div>
              <div class="hero-split-lbl">Ada Realisasi</div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.write("")

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Luas Tanam", fmt_ha(total_luas))
    c2.metric("Jumlah Proyek", len(projects))
    c3.metric("Cakupan Realisasi", f"{n_real} / {len(projects)}", "menunggu upload data realisasi" if n_real == 0 else None)

    unreliable_names = [p["meta"]["name"] for p in projects.values() if not p["rencana"].get("items_reliable", True)]
    no_real_count = len(projects) - n_real
    alerts = []
    if unreliable_names:
        alerts.append(f"<b>{len(unreliable_names)} proyek</b> rincian pekerjaannya belum bisa direkonsiliasi otomatis — {', '.join(unreliable_names)}. Total biaya proyeknya sendiri tetap akurat.")
    if no_real_count:
        alerts.append(f"<b>{no_real_count} dari {len(projects)} proyek</b> belum ada data realisasi. Upload sheet Realisasi resmi untuk mulai membandingkan Rencana vs Realisasi.")
    if alerts:
        st.markdown(
            '<div class="footnote" style="background:#F5E4DA;">⚠️ &nbsp;' + '<br><br>⚠️ &nbsp;'.join(alerts) + '</div>',
            unsafe_allow_html=True,
        )

    projects_with_real = [p for p in projects.values() if has_any_realisasi(p)]
    if projects_with_real:
        st.markdown("#### Capaian Realisasi vs RKP")
        total_rencana_real = sum(total_rencana(p) or 0 for p in projects_with_real)
        total_real_biaya = sum(realisasi_total(p) or 0 for p in projects_with_real)
        capaian_biaya_agg = (total_real_biaya / total_rencana_real * 100) if total_rencana_real else None

        fisik_projects = [p for p in projects_with_real if capaian_fisik_pct(p) is not None]
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("Proyek dengan Realisasi", f"{len(projects_with_real)} / {len(projects)}")
        rc2.metric("Capaian Biaya (agregat)", f"{capaian_biaya_agg:.1f}%" if capaian_biaya_agg is not None else "—",
                   f"{fmt_rp(total_real_biaya)} dari {fmt_rp(total_rencana_real)}")
        if fisik_projects:
            avg_fisik = sum(capaian_fisik_pct(p) for p in fisik_projects) / len(fisik_projects)
            rc3.metric("Capaian Fisik (rata-rata)", f"{avg_fisik:.1f}%", f"{len(fisik_projects)} proyek terukur")
        else:
            rc3.metric("Capaian Fisik (rata-rata)", "—", "belum ada proyek dengan data fisik cocok")

    col_l, col_r = st.columns([1.4, 1])
    with col_l:
        st.markdown("#### Peringkat Semua Proyek Berdasarkan Nilai")
        plist = sorted(projects.values(), key=lambda p: -(total_rencana(p) or 0))
        names = [p["meta"]["name"] for p in plist]
        vals = [total_rencana(p) or 0 for p in plist]
        bar_colors = [JENIS_COLORS.get(project_jenis(p), "#7C9A85") for p in plist]
        fig_rank = go.Figure(go.Bar(
            x=vals, y=names, orientation="h", marker_color=bar_colors,
            text=[fmt_rp(v) for v in vals], textposition="outside",
        ))
        fig_rank.update_layout(height=max(280, 42 * len(names)), margin=dict(l=10, r=60, t=10, b=10),
                                xaxis_title="Rp", plot_bgcolor="white", paper_bgcolor="white",
                                font=dict(color="#1B2A1E", size=12),
                                xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                                yaxis=dict(color="#1B2A1E", automargin=True, categoryorder="array", categoryarray=names[::-1]))
        st.plotly_chart(fig_rank, use_container_width=True, theme=None)

    with col_r:
        st.markdown("#### Distribusi per Jenis")
        fig_donut = go.Figure(go.Pie(
            labels=list(by_jenis.keys()), values=list(by_jenis.values()), hole=0.6,
            marker=dict(colors=[JENIS_COLORS.get(j, "#7C9A85") for j in by_jenis]),
            textinfo="label+percent", textfont=dict(color="#1B2A1E", size=12),
        ))
        fig_donut.update_layout(height=260, margin=dict(l=10, r=10, t=10, b=10),
                                 showlegend=False, paper_bgcolor="white")
        st.plotly_chart(fig_donut, use_container_width=True, theme=None)

        st.markdown("#### Per Perusahaan")
        by_company = {}
        for p in projects.values():
            by_company[p["meta"]["company"]] = by_company.get(p["meta"]["company"], 0) + (total_rencana(p) or 0)
        fig_comp = go.Figure(go.Bar(
            x=list(by_company.values()), y=list(by_company.keys()), orientation="h",
            marker_color=PALETTE[:len(by_company)],
            text=[fmt_rp(v) for v in by_company.values()], textposition="outside",
        ))
        fig_comp.update_layout(height=max(140, 50 * len(by_company)), margin=dict(l=10, r=60, t=10, b=10),
                                plot_bgcolor="white", paper_bgcolor="white",
                                font=dict(color="#1B2A1E", size=12),
                                xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                                yaxis=dict(color="#1B2A1E", automargin=True))
        st.plotly_chart(fig_comp, use_container_width=True, theme=None)

    st.markdown("#### Jelajahi Lebih Lanjut")
    b1, b2, b3 = st.columns(3)
    with b1:
        st.button("📋 Portofolio", use_container_width=True, on_click=_go_section, args=("📋 Portofolio",))
        st.caption("Tabel lengkap semua proyek, bisa diurut & difilter")
    with b2:
        st.button("🏢 Per Perusahaan", use_container_width=True, on_click=_go_section, args=("🏢 Per Perusahaan",))
        st.caption("Rekap investasi & daftar proyek tiap perusahaan")
    with b3:
        st.button("📁 Detail Proyek", use_container_width=True, on_click=_go_section, args=("📁 Detail Proyek",))
        st.caption("Jadwal, Kurva-S, rincian pekerjaan per proyek")

# ================================================================
# PORTOFOLIO (dulu "Ringkasan": tabel + grafik semua proyek terfilter)
# ================================================================
elif section == "📋 Portofolio":
    st.title("📋 Portofolio Proyek")
    total_biaya = sum((total_rencana(p) or 0) for p in projects.values())
    total_luas = sum((luas_proj(p) or 0) for p in projects.values())
    avg_rp_ha = total_biaya / total_luas if total_luas else None
    n_real = sum(1 for p in projects.values() if has_any_realisasi(p))

    top_l, top_r = st.columns([3, 1])
    with top_l:
        st.caption(f"{len(projects)} proyek terfilter · diperbarui {datetime.now().strftime('%d %b %Y')}")
    with top_r:
        export_df = pd.DataFrame([
            {
                "Proyek": p["meta"]["name"],
                "Perusahaan": p["meta"]["company"],
                "Jenis": project_jenis(p),
                "Luas (Ha)": luas_proj(p),
                "Total Biaya Rencana": total_rencana(p),
                "Biaya / Ha": rp_per_ha(p),
                "Status Realisasi": "Tersedia" if has_realisasi(p) else ("Total saja" if has_any_realisasi(p) else "Belum ada"),
            }
            for p in projects.values()
        ])
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
            export_df.to_excel(writer, sheet_name="Ringkasan Proyek", index=False)
        st.download_button(
            "📥 Export Excel", data=xbuf.getvalue(), file_name="ringkasan_rkp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Jumlah Proyek", len(projects), f"{n_real} dengan realisasi")
    c2.metric("Total Luas", fmt_ha(total_luas))
    c3.metric("Total Biaya Rencana", fmt_rp(total_biaya), fmt_rp_full(total_biaya))
    c4.metric("Rata-rata Biaya / Ha", fmt_rp(avg_rp_ha))

    st.markdown("#### Tabel Proyek")
    st.caption("Klik header kolom untuk mengurutkan.")
    table_df = pd.DataFrame([
        {
            "Proyek": p["meta"]["name"],
            "Perusahaan": p["meta"]["company"],
            "Jenis": project_jenis(p),
            "Luas (Ha)": luas_proj(p),
            "Total Biaya": total_rencana(p),
            "Biaya / Ha": rp_per_ha(p),
            "Progres Waktu (%)": project_time_progress_pct(p),
            "Capaian Biaya (%)": capaian_biaya_pct(p),
            "Capaian Fisik (%)": capaian_fisik_pct(p),
            "Status Realisasi": "Tersedia" if has_realisasi(p) else ("Total saja" if has_any_realisasi(p) else "Belum ada"),
        }
        for p in projects.values()
    ])
    st.dataframe(
        table_df, use_container_width=True, hide_index=True,
        column_config={
            "Luas (Ha)": st.column_config.NumberColumn(format="%.2f"),
            "Total Biaya": st.column_config.NumberColumn(format="Rp %d"),
            "Biaya / Ha": st.column_config.NumberColumn(format="Rp %d"),
            "Progres Waktu (%)": st.column_config.ProgressColumn(format="%.0f%%", min_value=0, max_value=100),
            "Capaian Biaya (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "Capaian Fisik (%)": st.column_config.NumberColumn(format="%.1f%%"),
        },
    )

    st.markdown("#### Perbandingan Biaya Antar Proyek")
    st.caption("Warna batang menandakan Jenis Proyek — lihat legenda di bawah grafik.")
    mode = st.radio("Mode", ["Total Biaya", "Biaya / Ha"], horizontal=True, label_visibility="collapsed")
    plist = sorted(projects.values(), key=lambda p: (project_jenis(p), -(total_rencana(p) or 0)))
    names = [p["meta"]["name"] for p in plist]
    vals = [
        (total_rencana(p) or 0) if mode == "Total Biaya" else (rp_per_ha(p) or 0)
        for p in plist
    ]
    bar_colors = [JENIS_COLORS.get(project_jenis(p), "#7C9A85") for p in plist]
    fig = go.Figure(go.Bar(
        x=vals, y=names, orientation="h", marker_color=bar_colors,
        text=[fmt_rp(v) for v in vals], textposition="outside",
    ))
    for jenis, color in JENIS_COLORS.items():
        if jenis in {project_jenis(p) for p in plist}:
            fig.add_bar(x=[None], y=[None], marker_color=color, name=jenis, showlegend=True)
    fig.update_layout(height=100 + 60 * len(names), margin=dict(l=10, r=40, t=10, b=10),
                       xaxis_title="Rp", plot_bgcolor="white", paper_bgcolor="white",
                       font=dict(color="#1B2A1E", size=13), barmode="overlay",
                       legend=dict(orientation="h", y=-0.12, font=dict(color="#1B2A1E", size=11)),
                       xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                       yaxis=dict(color="#1B2A1E", automargin=True, categoryorder="array", categoryarray=names[::-1]))
    st.plotly_chart(fig, use_container_width=True, theme=None)

    st.markdown("#### Target Fisik per Periode (Catur Wulan) — Semua Proyek")
    st.caption("Hanya proyek yang punya jadwal periode (perkebunan/konstruksi) yang muncul di sini.")
    keys = all_period_keys(projects)
    fig2 = go.Figure()
    for i, p in enumerate(projects.values()):
        g = p["rencana"]["grand"]
        m = {pd_["key"]: pd_["fisik"] for pd_ in (g["periods"] if g else [])}
        fig2.add_bar(name=p["meta"]["name"], x=keys, y=[m.get(k, 0) for k in keys],
                     marker_color=PALETTE[i % len(PALETTE)])
    fig2.update_layout(barmode="stack", height=380, margin=dict(l=10, r=10, t=10, b=10),
                        yaxis_title="Ha", plot_bgcolor="white", paper_bgcolor="white",
                        legend=dict(orientation="h", y=-0.3, font=dict(color="#1B2A1E", size=11)),
                        font=dict(color="#1B2A1E", size=13),
                        xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                        yaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E", automargin=True))
    add_now_marker(fig2, keys, "periode")
    add_now_marker(fig2, keys, "tahunan")
    st.plotly_chart(fig2, use_container_width=True, theme=None)

    st.markdown("#### Kurva-S Kumulatif Biaya Rencana — Semua Proyek")
    st.caption("Akumulasi biaya rencana dari waktu ke waktu. Hanya proyek berjadwal periode yang tampil; garis putus-putus menandai periode berjalan saat ini.")
    fig_s = go.Figure()
    any_scurve = False
    for i, p in enumerate(projects.values()):
        g = p["rencana"]["grand"]
        if not g or not g["periods"]:
            continue
        any_scurve = True
        m = {pd_["key"]: pd_["biaya"] for pd_ in g["periods"]}
        vals_s, running = [], 0
        for k in keys:
            running += m.get(k, 0)
            vals_s.append(running)
        fig_s.add_scatter(x=keys, y=vals_s, mode="lines+markers", name=p["meta"]["name"],
                           line=dict(color=PALETTE[i % len(PALETTE)], width=2.5))
    if any_scurve:
        fig_s.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10),
                             yaxis_title="Rp (kumulatif)", plot_bgcolor="white", paper_bgcolor="white",
                             legend=dict(orientation="h", y=-0.3, font=dict(color="#1B2A1E", size=11)),
                             font=dict(color="#1B2A1E", size=13),
                             xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                             yaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E", tickformat=",.0f"))
        add_now_marker(fig_s, keys, "periode")
        add_now_marker(fig_s, keys, "tahunan")
        st.plotly_chart(fig_s, use_container_width=True, theme=None)
    else:
        st.info("Tidak ada proyek berjadwal periode pada filter saat ini.")

    st.markdown("#### Pekerjaan & Biaya — Semua Proyek")
    st.caption("Semua jenis pekerjaan (level sama seperti tabel Rincian Pekerjaan) digabung lintas proyek yang terfilter.")

    unreliable_names = [p["meta"]["name"] for p in projects.values() if not p["rencana"].get("items_reliable", True)]

    work_rows = []
    for p in projects.values():
        if not p["rencana"].get("items_reliable", True):
            continue
        for it in p["rencana"]["items"]:
            work_rows.append({
                "Pekerjaan": it["nama"],
                "Proyek": p["meta"]["name"],
                "Volume (Ha)": it["volume_ha"] or 0,
                "Biaya Rencana": it["biaya_rencana"] or 0,
            })
    work_df = pd.DataFrame(work_rows)

    if unreliable_names:
        st.caption(
            "⚠️ " + ", ".join(unreliable_names) + " tidak disertakan di sini karena rincian per "
            "pekerjaannya tidak bisa direkonsiliasi otomatis dengan total rencananya (struktur subtotal "
            "berlapis di file sumber). Total biaya proyek itu sendiri tetap akurat — lihat di halaman "
            "Ringkasan atau detail proyeknya masing-masing."
        )

    if work_df.empty:
        st.info("Belum ada rincian pekerjaan pada proyek yang terfilter.")
    else:
        agg = (
            work_df.groupby("Pekerjaan", as_index=False)
            .agg(**{
                "Total Biaya": ("Biaya Rencana", "sum"),
                "Total Volume (Ha)": ("Volume (Ha)", "sum"),
                "Jumlah Proyek": ("Proyek", "nunique"),
            })
            .sort_values("Total Biaya", ascending=False)
        )
        agg["Rp / Ha"] = agg["Total Biaya"] / agg["Total Volume (Ha)"].replace(0, pd.NA)

        fig5 = go.Figure(go.Bar(
            x=agg["Total Biaya"], y=agg["Pekerjaan"], orientation="h",
            marker_color=GOLD, text=[fmt_rp(v) for v in agg["Total Biaya"]], textposition="outside",
        ))
        fig5.update_layout(
            height=max(220, 46 * len(agg)), margin=dict(l=10, r=40, t=10, b=10),
            xaxis_title="Rp", plot_bgcolor="white", paper_bgcolor="white",
            font=dict(color="#1B2A1E", size=13),
            xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
            yaxis=dict(color="#1B2A1E", automargin=True, categoryorder="total ascending"),
        )
        st.plotly_chart(fig5, use_container_width=True, theme=None)

        with st.expander("📋 Lihat tabel rincian per pekerjaan & proyek"):
            st.markdown("**Ringkasan per jenis pekerjaan (digabung semua proyek)**")
            st.dataframe(
                agg, use_container_width=True, hide_index=True,
                column_config={
                    "Total Biaya": st.column_config.NumberColumn(format="Rp %d"),
                    "Total Volume (Ha)": st.column_config.NumberColumn(format="%.2f"),
                    "Rp / Ha": st.column_config.NumberColumn(format="Rp %d"),
                },
            )
            st.markdown("**Rincian per proyek**")
            detail_df = work_df.sort_values(["Pekerjaan", "Proyek"])
            st.dataframe(
                detail_df, use_container_width=True, hide_index=True,
                column_config={
                    "Volume (Ha)": st.column_config.NumberColumn(format="%.2f"),
                    "Biaya Rencana": st.column_config.NumberColumn(format="Rp %d"),
                },
            )

    st.markdown("#### Daftar Proyek")
    cols = st.columns(3)
    for i, p in enumerate(projects.values()):
        with cols[i % 3]:
            real = has_any_realisasi(p)
            badge = '<span class="badge-ok">Ada realisasi</span>' if real else '<span class="badge-wait">Rencana saja</span>'
            st.markdown(
                f"""
                <div style="background:#fff;border:1px solid #E1E3D9;border-radius:12px;padding:14px 16px;margin-bottom:12px;">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;">
                  <div>
                    <b>{p['meta']['name']}</b><br>
                    <span style="color:#6C7566;font-size:12.5px;">{p['meta']['company']}</span>
                  </div>
                  {badge}
                </div>
                <div style="margin-top:8px;font-size:13px;color:#6C7566;">📐 {fmt_ha(luas_proj(p))} &nbsp;·&nbsp; 📅 {p['meta']['periode_text'] or '—'}</div>
                <div style="margin-top:8px;display:flex;justify-content:space-between;">
                  <div><b style="color:{FOREST};">{fmt_rp(total_rencana(p))}</b><br><span style="font-size:11px;color:#6C7566;">total rencana</span></div>
                  <div style="text-align:right;"><b style="color:{FOREST};">{fmt_rp(rp_per_ha(p))}</b><br><span style="font-size:11px;color:#6C7566;">per Ha</span></div>
                </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

# ================================================================
# PER PERUSAHAAN
# ================================================================
elif section == "🏢 Per Perusahaan":
    st.title("🏢 Per Perusahaan")
    companies = sorted({p["meta"]["company"] for p in projects.values()})
    st.caption(f"{len(companies)} perusahaan pada filter saat ini.")

    for c in companies:
        comp_projects = [p for p in projects.values() if p["meta"]["company"] == c]
        total_c = sum(total_rencana(p) or 0 for p in comp_projects)
        luas_c = sum(luas_proj(p) or 0 for p in comp_projects)
        n_real_c = sum(1 for p in comp_projects if has_any_realisasi(p))

        st.markdown(f"### {c}")
        c1, c2, c3 = st.columns(3)
        c1.metric("Jumlah Proyek", len(comp_projects), f"{n_real_c} dengan realisasi")
        c2.metric("Total Biaya", fmt_rp(total_c), fmt_rp_full(total_c))
        c3.metric("Total Luas", fmt_ha(luas_c))

        comp_df = pd.DataFrame([
            {
                "Proyek": p["meta"]["name"],
                "Jenis": project_jenis(p),
                "Luas (Ha)": luas_proj(p),
                "Total Biaya": total_rencana(p),
                "Biaya / Ha": rp_per_ha(p),
                "Status Realisasi": "Tersedia" if has_realisasi(p) else ("Total saja" if has_any_realisasi(p) else "Belum ada"),
            }
            for p in sorted(comp_projects, key=lambda p: -(total_rencana(p) or 0))
        ])
        st.dataframe(
            comp_df, use_container_width=True, hide_index=True,
            column_config={
                "Luas (Ha)": st.column_config.NumberColumn(format="%.2f"),
                "Total Biaya": st.column_config.NumberColumn(format="Rp %d"),
                "Biaya / Ha": st.column_config.NumberColumn(format="Rp %d"),
            },
        )
        st.divider()

# ================================================================
# DETAIL SATU PROYEK (dipilih lewat dropdown sidebar)
# ================================================================
elif section == "📁 Detail Proyek":
    selected_view = st.session_state.get("detail_project")
    p = next((pp for pp in projects.values() if pp["meta"]["name"] == selected_view), None)
    if p is None:
        st.warning("Proyek tidak ditemukan pada hasil filter saat ini. Kembali ke Ringkasan.")
        st.stop()

    real = has_realisasi(p)
    any_real = has_any_realisasi(p)
    total = total_rencana(p)
    luas = luas_proj(p)

    st.title(f"📁 {p['meta']['name']}")
    st.caption(f"{p['meta']['company']}" + (f" · {p['meta']['desc']}" if p["meta"]["desc"] else "") +
               (f" · {p['meta']['periode_text']}" if p["meta"]["periode_text"] else ""))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Luas", fmt_ha(luas))
    c2.metric("Total Biaya Rencana", fmt_rp(total), fmt_rp_full(total))
    c3.metric("Biaya / Ha", fmt_rp(rp_per_ha(p)))
    status_label = "Tersedia" if real else ("Total saja" if any_real else "Belum ada")
    c4.metric("Status Realisasi", status_label)

    if any_real:
        cb, cf = capaian_biaya_pct(p), capaian_fisik_pct(p)
        d1, d2 = st.columns(2)
        d1.metric("Capaian Realisasi Biaya vs RKP", f"{cb:.1f}%" if cb is not None else "—",
                  f"{fmt_rp(realisasi_total(p))} dari {fmt_rp(total)}" if cb is not None else None)
        d2.metric("Capaian Realisasi Fisik vs RKP", f"{cf:.1f}%" if cf is not None else "—",
                  None if cf is not None else "butuh realisasi per-periode yang cocok")

    yr_range = project_year_range(p)
    if yr_range:
        start_y, end_y = yr_range
        now = datetime.now()
        frac_now = now.year + (now.month - 1) / 12
        total_span = (end_y + 1) - start_y
        pct = max(0, min(100, (frac_now - start_y) / total_span * 100)) if total_span > 0 else None
        if pct is not None:
            st.markdown(f"**⏱️ Progres Waktu Proyek** — Tahun {start_y} s.d {end_y}")
            st.progress(int(pct))
            st.caption(f"Proyek ini seharusnya sudah berjalan sekitar **{pct:.0f}%** dari total durasi rencananya (berdasarkan kalender, bukan realisasi biaya).")

    if any_real and not real:
        rt = realisasi_total(p)
        capaian = (rt / total * 100) if (rt is not None and total) else None
        st.markdown(
            f"""<div class="footnote">📊 Ada angka realisasi (total) untuk proyek ini: <b>{fmt_rp_full(rt)}</b>
            {f"— sekitar <b>{capaian:.1f}%</b> dari rencana." if capaian is not None else ""}
            Rinciannya belum bisa dipecah per pekerjaan karena struktur sheet realisasi berbeda dari sheet
            rencananya.</div>""",
            unsafe_allow_html=True,
        )
    elif not any_real:
        st.markdown(
            """<div class="footnote">📋 File ini belum berisi data realisasi (kolom/sheet "Realisasi" tidak
            terdeteksi, atau sheet tersebut ternyata bukan untuk proyek ini). Tambahkan sheet baru bernama
            mengandung kata <b>Realisasi</b> dengan struktur tabel yang mirip sheet rencananya, lalu upload
            ulang — dashboard otomatis akan menampilkan perbandingan Rencana vs Realisasi di sini.</div>""",
            unsafe_allow_html=True,
        )

    g = p["rencana"]["grand"]
    keys = [pd_["key"] for pd_ in g["periods"]] if g else []

    if not keys:
        st.info("Proyek ini tidak memiliki breakdown per periode (Catur Wulan/Tahun) — hanya total biaya per pekerjaan di bawah.")
    else:
        period_label = "Tahun" if keys[0].startswith("Tahun ") else "Catur Wulan"
        st.markdown(f"#### Biaya per Periode ({period_label})")
        rencana_vals = [pd_["biaya"] for pd_ in g["periods"]]
        fig3 = go.Figure()
        fig3.add_bar(name="Rencana", x=keys, y=rencana_vals, marker_color=FOREST)
        if real and p["realisasi"]["data"]["grand"]:
            rmap = {pd_["key"]: pd_["biaya"] for pd_ in p["realisasi"]["data"]["grand"]["periods"]}
            fig3.add_bar(name="Realisasi", x=keys, y=[rmap.get(k, 0) for k in keys], marker_color=GOLD)
        fig3.update_layout(barmode="group", height=320, margin=dict(l=10, r=10, t=10, b=10),
                            plot_bgcolor="white", paper_bgcolor="white",
                            legend=dict(orientation="h", y=-0.2, font=dict(color="#1B2A1E")),
                            font=dict(color="#1B2A1E", size=13),
                            xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                            yaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"))
        add_now_marker(fig3, keys, p["format"])
        st.plotly_chart(fig3, use_container_width=True, key=f"biaya_{p['id']}", theme=None)

        st.markdown("#### Kurva-S Kumulatif Biaya Rencana")
        cum_vals, running = [], 0
        for v in rencana_vals:
            running += v
            cum_vals.append(running)
        fig_s1 = go.Figure()
        fig_s1.add_scatter(x=keys, y=cum_vals, mode="lines+markers", name="Rencana (kumulatif)",
                            line=dict(color=FOREST, width=3))
        if real and p["realisasi"]["data"]["grand"]:
            rmap = {pd_["key"]: pd_["biaya"] for pd_ in p["realisasi"]["data"]["grand"]["periods"]}
            cum_real, running_r = [], 0
            for k in keys:
                running_r += rmap.get(k, 0)
                cum_real.append(running_r)
            fig_s1.add_scatter(x=keys, y=cum_real, mode="lines+markers", name="Realisasi (kumulatif)",
                                line=dict(color=GOLD, width=3))
        fig_s1.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10),
                              plot_bgcolor="white", paper_bgcolor="white",
                              legend=dict(orientation="h", y=-0.2, font=dict(color="#1B2A1E")),
                              font=dict(color="#1B2A1E", size=13),
                              xaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E"),
                              yaxis=dict(gridcolor="#EEF0E8", color="#1B2A1E", tickformat=",.0f"))
        add_now_marker(fig_s1, keys, p["format"])
        st.plotly_chart(fig_s1, use_container_width=True, key=f"scurve_{p['id']}", theme=None)

    items_reliable = p["rencana"].get("items_reliable", True)
    if not items_reliable:
        st.warning(
            "⚠️ Rincian per pekerjaan di bawah ini kemungkinan **tidak sepenuhnya akurat** — struktur "
            "subtotal berlapis pada file sumber membuat sebagian baris berpotensi terhitung dobel atau "
            "malah terlewat. Total Biaya Rencana pada kartu di atas tetap akurat (diambil langsung dari "
            "baris Grand Total/Total di file, bukan dari penjumlahan baris di bawah)."
        )

    st.markdown("#### Komposisi Biaya per Pekerjaan")
    items = [it for it in p["rencana"]["items"] if it["biaya_rencana"]]
    items = sorted(items, key=lambda x: -x["biaya_rencana"])
    fig4 = px.pie(
        names=[it["nama"] for it in items],
        values=[it["biaya_rencana"] for it in items],
        color_discrete_sequence=PALETTE, hole=0.55,
    )
    fig4.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10),
                        plot_bgcolor="white", paper_bgcolor="white",
                        font=dict(color="#1B2A1E", size=13),
                        legend=dict(font=dict(color="#1B2A1E")))
    fig4.update_traces(textfont=dict(color="#1B2A1E"))
    st.plotly_chart(fig4, use_container_width=True, key=f"comp_{p['id']}", theme=None)

    st.markdown("#### Rincian Pekerjaan" + ("" if items_reliable else " ⚠️"))
    real_map = {}
    if real:
        for it in p["realisasi"]["data"]["items"]:
            real_map[it["nama"]] = it["biaya_rencana"]

    rows_table = []
    for it in p["rencana"]["items"]:
        r_biaya = real_map.get(it["nama"])
        capaian = (r_biaya / it["biaya_rencana"] * 100) if (real and r_biaya and it["biaya_rencana"]) else None
        rows_table.append({
            "No": it["no"],
            "Pekerjaan": it["nama"],
            "Volume (Ha)": it["volume_ha"],
            "Biaya Rencana": it["biaya_rencana"],
            "Rp / Ha": it["rp_per_ha"],
            "Realisasi Biaya": r_biaya if real else None,
            "% Capaian": round(capaian, 1) if capaian is not None else None,
        })
    df = pd.DataFrame(rows_table)
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Volume (Ha)": st.column_config.NumberColumn(format="%.2f"),
            "Biaya Rencana": st.column_config.NumberColumn(format="Rp %d"),
            "Rp / Ha": st.column_config.NumberColumn(format="Rp %d"),
            "Realisasi Biaya": st.column_config.NumberColumn(format="Rp %d"),
            "% Capaian": st.column_config.NumberColumn(format="%.1f%%"),
        },
    )
