"""
Aplikasi Absensi Peserta PKL - Bank Sulselbar (versi siap deploy)
=================================================================
Fitur:
- Layar login di awal: pilih peran. Peserta (nama + PIN) -> halaman absen saja; Admin (password) -> area admin
- Absen MASUK saja (tidak ada absen pulang): foto selfie + lokasi GPS
- Login PIN per peserta (dibuat admin)
- Cek lokasi: harus di sekitar gedung yang terdaftar untuk peserta (A / B)
- Setiap peserta juga punya Divisi (dipilih dari daftar, Gedung ikut otomatis)
- Batas waktu: absen masuk hanya bisa sebelum jam BATAS_JAM_MASUK
- Uang makan Rp15.000 otomatis saat absen masuk berhasil
- Blokir absen dobel di hari yang sama
- Admin (pakai password): kelola peserta, rekap + grafik + export Excel, bukti foto
- Database: SQLite (lokal) atau Postgres/Supabase (cloud) - otomatis

Jalanin lokal:   streamlit run app.py
"""

import hashlib
import hmac
import io
from datetime import datetime, date, time as dtime, timedelta, timezone
from math import radians, sin, cos, sqrt, atan2

import pandas as pd
import streamlit as st
from PIL import Image
from sqlalchemy import (
    Column, Integer, String, Float, LargeBinary, MetaData, Table,
    create_engine, select, insert, update, delete, func, and_,
)
from sqlalchemy.exc import IntegrityError
from streamlit_js_eval import get_geolocation

# ------------------------------------------------------------------
# KONFIGURASI - SESUAIKAN BAGIAN INI
# ------------------------------------------------------------------
UANG_MAKAN = 15000

# Koordinat asli Gedung A dan B Bank Sulselbar
GEDUNG_COORD = {
    "A": (-5.1164911547868, 119.51684816101282),  # (latitude, longitude)
    "B": (-5.147453606311076, 119.4293639296109),  # (latitude, longitude)
}

RADIUS_METER = 150  # area longgar di sekitar gedung (bukan angka ketat)

# Daftar Divisi asli beserta gedungnya ("A" atau "B").
# Peserta tinggal pilih Divisi, Gedung otomatis ikut sesuai daftar ini.
DIVISI_GEDUNG = {
    # Gedung A
    "Divisi Korporasi dan Komersial": "A",
    "Divisi Ritel dan Konsumer": "A",
    "Divisi Corporate Secretary (CORSEC)": "A",
    "Divisi Human Capital": "A",
    "Divisi SIMO": "A",
    "Divisi Keuangan dan Akuntansi": "A",
    "Divisi Treasury": "A",
    "Divisi Teknologi dan Informasi": "A",
    "Divisi International Banking": "A",
    "Divisi Digitalisasi dan Layanan": "A",
    "Divisi Umum": "A",
    # Gedung B
    "Divisi Kepatuhan": "B",
    "Divisi Dana dan Jasa": "B",
    "Divisi Syariah": "B",
    "Divisi Manajemen Risiko": "B",
    "Divisi SKAI": "B",
}

BATAS_JAM_MASUK = dtime(23, 0)  # absen masuk harus sebelum jam ini (WITA), lewat = ditolak

WITA = timezone(timedelta(hours=8))  # server cloud pakai UTC, jadi jam dipaksa WITA


# ------------------------------------------------------------------
# HELPER: SECRETS
# ------------------------------------------------------------------
def secret(key, default=None):
    """Ambil nilai dari Streamlit Secrets; aman kalau file secrets belum ada."""
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


# ------------------------------------------------------------------
# DATABASE (SQLAlchemy: jalan di SQLite lokal DAN Postgres cloud)
# ------------------------------------------------------------------
metadata = MetaData()

peserta_t = Table(
    "peserta", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("nama", String(100), nullable=False, unique=True),
    Column("gedung", String(1), nullable=False),
    Column("divisi", String(100), nullable=False, default=""),
    Column("pin_hash", String(64), nullable=False),
)

absensi_t = Table(
    "absensi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("nama", String(100), nullable=False, index=True),
    Column("tanggal", String(10), nullable=False, index=True),  # YYYY-MM-DD
    Column("jam", String(8), nullable=False),                   # HH:MM:SS
    Column("jenis", String(10), nullable=False, default="masuk"),  # selalu "masuk"
    Column("lat", Float),
    Column("lon", Float),
    Column("akurasi", Float),
    Column("jarak", Float),
    Column("status", String(20), nullable=False),               # berhasil / gagal_lokasi
    Column("uang_makan", Integer, nullable=False, default=0),
    Column("foto", LargeBinary),
)


@st.cache_resource
def get_engine():
    url = secret("database_url")
    if url:
        url = url.replace("postgres://", "postgresql://", 1)
        # Paksa driver psycopg2 secara eksplisit, supaya SQLAlchemy tidak
        # mencoba driver lain (mis. psycopg v3) yang belum tentu terinstall.
        if url.startswith("postgresql://"):
            url = url.replace("postgresql://", "postgresql+psycopg2://", 1)
        eng = create_engine(url, pool_pre_ping=True, pool_recycle=300)
    else:
        # nama file baru (v2) supaya tidak bentrok dengan database versi lama
        eng = create_engine(
            "sqlite:///absensi_pkl_v2.db",
            connect_args={"check_same_thread": False},
        )
    metadata.create_all(eng)
    return eng


def hash_pin(nama, pin):
    garam = secret("salt", "pkl-sulselbar")
    return hashlib.sha256(f"{nama.strip().lower()}|{pin}|{garam}".encode()).hexdigest()


def list_peserta(eng):
    with eng.connect() as c:
        return pd.read_sql(
            select(peserta_t.c.nama, peserta_t.c.gedung, peserta_t.c.divisi).order_by(peserta_t.c.nama), c
        )


def get_peserta(eng, nama):
    with eng.connect() as c:
        return c.execute(select(peserta_t).where(peserta_t.c.nama == nama)).mappings().first()


def add_peserta(eng, nama, gedung, divisi, pin):
    with eng.begin() as c:
        c.execute(insert(peserta_t).values(
            nama=nama.strip(), gedung=gedung, divisi=divisi.strip(), pin_hash=hash_pin(nama, pin)
        ))


def reset_pin(eng, nama, pin_baru):
    with eng.begin() as c:
        c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(pin_hash=hash_pin(nama, pin_baru)))


def set_divisi(eng, nama, divisi, gedung):
    with eng.begin() as c:
        c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(divisi=divisi.strip(), gedung=gedung))


def hapus_peserta(eng, nama, dengan_absensi=False):
    """Hapus peserta. Kalau dengan_absensi=True, riwayat absensinya ikut dihapus."""
    with eng.begin() as c:
        c.execute(delete(peserta_t).where(peserta_t.c.nama == nama))
        if dengan_absensi:
            c.execute(delete(absensi_t).where(absensi_t.c.nama == nama))


def daftar_nama_absensi(eng):
    """Semua nama yang punya riwayat absensi (termasuk peserta yang sudah dihapus)."""
    q = select(absensi_t.c.nama).distinct().order_by(absensi_t.c.nama)
    with eng.connect() as c:
        return [r[0] for r in c.execute(q).all()]


def hapus_absensi(eng, nama=None):
    """Hapus riwayat absensi (beserta fotonya). nama=None berarti SEMUA peserta.
    Mengembalikan jumlah baris yang terhapus."""
    stmt = delete(absensi_t)
    if nama is not None:
        stmt = stmt.where(absensi_t.c.nama == nama)
    with eng.begin() as c:
        return c.execute(stmt).rowcount


def sudah_berhasil(eng, nama, tanggal, jenis):
    q = select(func.count()).select_from(absensi_t).where(and_(
        absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal,
        absensi_t.c.jenis == jenis, absensi_t.c.status == "berhasil",
    ))
    with eng.connect() as c:
        return c.execute(q).scalar() > 0


def simpan_absen(eng, **nilai):
    with eng.begin() as c:
        c.execute(insert(absensi_t).values(**nilai))


def get_absensi_bulan(eng, tahun, bulan):
    prefix = f"{tahun:04d}-{bulan:02d}-"
    q = (
        select(
            absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jam, absensi_t.c.jenis,
            absensi_t.c.status, absensi_t.c.jarak, absensi_t.c.uang_makan,
        )
        .where(absensi_t.c.tanggal.like(prefix + "%"))
        .order_by(absensi_t.c.tanggal, absensi_t.c.jam)
    )
    with eng.connect() as c:
        return pd.read_sql(q, c)


def get_absensi_hari_foto(eng, tanggal):
    q = (
        select(
            absensi_t.c.nama, absensi_t.c.jam, absensi_t.c.jenis, absensi_t.c.status,
            absensi_t.c.jarak, absensi_t.c.foto,
        )
        .where(absensi_t.c.tanggal == tanggal)
        .order_by(absensi_t.c.jam)
    )
    with eng.connect() as c:
        return c.execute(q).mappings().all()


# ------------------------------------------------------------------
# LOGIKA: LOKASI & FOTO
# ------------------------------------------------------------------
def jarak_meter(lat1, lon1, lat2, lon2):
    R = 6371000
    p1, p2 = radians(lat1), radians(lat2)
    dp, dl = radians(lat2 - lat1), radians(lon2 - lon1)
    a = sin(dp / 2) ** 2 + cos(p1) * cos(p2) * sin(dl / 2) ** 2
    return R * 2 * atan2(sqrt(a), sqrt(1 - a))


def cek_lokasi(lat, lon, gedung):
    target = GEDUNG_COORD[gedung]
    jarak = jarak_meter(lat, lon, target[0], target[1])
    return jarak <= RADIUS_METER, jarak


def beri_nomor(df):
    """Nomor urut tabel mulai dari 1 (bukan 0)."""
    df = df.copy()
    df.index = range(1, len(df) + 1)
    df.index.name = "No"
    return df


def excel_bytes(rekap, detail):
    """Bangun file Excel (2 sheet) dari tabel rekap & detail, kolom No mulai dari 1."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        beri_nomor(rekap).to_excel(w, sheet_name="Rekap Uang Makan")
        beri_nomor(detail).to_excel(w, sheet_name="Detail Absensi")
    return buf.getvalue()


def kompres_foto(data: bytes) -> bytes:
    """Kecilkan foto (maks 640px, JPEG q70) supaya database tidak cepat penuh."""
    img = Image.open(io.BytesIO(data)).convert("RGB")
    img.thumbnail((640, 640))
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=70, optimize=True)
    return out.getvalue()


# ------------------------------------------------------------------
# ADMIN LOGIN
# ------------------------------------------------------------------
def password_admin():
    """Password admin yang berlaku. None = belum diatur di Secrets (mode cloud)."""
    pw = secret("admin_password")
    if pw:
        return str(pw)
    if secret("database_url"):
        return None
    return "admin123"  # mode lokal (tanpa database cloud)


def cek_pin(eng, nama, pin):
    peserta = get_peserta(eng, nama)
    return peserta is not None and hmac.compare_digest(peserta["pin_hash"], hash_pin(nama, pin))


def admin_login():
    """Penjaga halaman admin. Login normal lewat layar login di awal; ini cadangan
    kalau halaman admin terbuka tanpa lewat login."""
    if st.session_state.get("admin_ok"):
        return True

    pw_benar = password_admin()
    if pw_benar is None:
        st.error("admin_password belum diatur di Secrets. Hubungi pengelola aplikasi.")
        return False
    if not secret("admin_password") and not secret("database_url"):
        st.info("Mode lokal: password admin default = admin123")

    pw = st.text_input("Password admin", type="password")
    if st.button("Masuk sebagai admin"):
        if hmac.compare_digest(pw, str(pw_benar)):
            st.session_state.admin_ok = True
            st.rerun()
        else:
            st.error("Password salah.")
    return False


# ------------------------------------------------------------------
# TAMPILAN PESERTA (gaya aplikasi "Presensi" mobile)
# ------------------------------------------------------------------
HARI_ID = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BULAN_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
            "Agustus", "September", "Oktober", "November", "Desember"]


def tgl_indonesia(d):
    return f"{HARI_ID[d.weekday()]}, {d.day} {BULAN_ID[d.month - 1]} {d.year}"


def salam(jam):
    if jam < 11:
        return "Selamat Pagi"
    if jam < 15:
        return "Selamat Siang"
    if jam < 18:
        return "Selamat Sore"
    return "Selamat Malam"


def rupiah(n):
    return f"Rp{int(n):,}".replace(",", ".")


def _svg(path, size=24, stroke=2):
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" '
        f'fill="none" stroke="currentColor" stroke-width="{stroke}" stroke-linecap="round" '
        f'stroke-linejoin="round">{path}</svg>'
    )


ICON = {
    "fingerprint": '<path d="M12 10a2 2 0 0 0-2 2c0 1.02-.1 2.51-.26 4"/><path d="M14 13.12c0 2.38 0 6.38-1 8.88"/><path d="M17.29 21.02c.12-.6.43-2.3.5-3.02"/><path d="M2 12a10 10 0 0 1 18-6"/><path d="M2 16h.01"/><path d="M21.8 16c.2-2 .131-5.354 0-6"/><path d="M5 19.5C5.5 18 6 15 6 12a6 6 0 0 1 .34-2"/><path d="M8.65 22c.21-.66.45-1.32.57-2"/><path d="M9 6.8a6 6 0 0 1 9 5.2v2"/>',
    "logout": '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" x2="9" y1="12" y2="12"/>',
    "login": '<path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" x2="3" y1="12" y2="12"/>',
    "calendar": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    "calendar_days": '<rect width="18" height="18" x="3" y="4" rx="2"/><path d="M16 2v4M8 2v4M3 10h18M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01"/>',
    "clock": '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
    "inbox": '<polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
    "plane": '<path d="M17.8 19.2 16 11l3.5-3.5C21 6 21.5 4 21 3c-1-.5-3 0-4.5 1.5L13 8 4.8 6.2c-.5-.1-.9.1-1.1.5l-.3.5c-.2.5-.1 1 .3 1.3L9 12l-2 3H4l-1 1 3 2 2 3 1-1v-3l3-2 3.5 5.3c.3.4.8.5 1.3.3l.5-.2c.4-.3.6-.7.5-1.2z"/>',
    "file": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M10 9H8M16 13H8M16 17H8"/>',
    "home": '<path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>',
    "history": '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M12 7v5l4 2"/>',
    "user": '<path d="M19 21v-2a4 4 0 0 0-4-4H9a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>',
    "chart": '<line x1="12" x2="12" y1="20" y2="10"/><line x1="18" x2="18" y1="20" y2="4"/><line x1="6" x2="6" y1="20" y2="16"/>',
    "back": '<path d="m15 18-6-6 6-6"/>',
}


def ic(nama, size=24, stroke=2):
    return _svg(ICON[nama], size, stroke)


CSS_PRESENSI = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"] {
  background: #eef2fb !important; font-family: 'Inter', system-ui, sans-serif;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], footer { display: none !important; }
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 430px !important; padding: 0 0 130px 0 !important; position: relative;
  background: #eef2fb; min-height: 100vh;
}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] { gap: 0 !important; }

/* ---------- header biru ---------- */
.pr-head {
  background: linear-gradient(135deg, #2563eb 0%, #3b3fd8 100%);
  color: #fff; padding: 34px 22px 86px; border-radius: 0 0 0 0;
  position: relative; overflow: hidden;
}
.pr-head::after {
  content: ""; position: absolute; left: -10%; right: -10%; bottom: -34px; height: 70px;
  background: #eef2fb; border-radius: 50% 50% 0 0 / 100% 100% 0 0;
}
.pr-top { display: flex; align-items: center; gap: 14px; }
.pr-logo { width: 46px; height: 46px; border-radius: 14px; background: rgba(255,255,255,.18);
  display: flex; align-items: center; justify-content: center; }
.pr-title { font-size: 25px; font-weight: 600; letter-spacing: .2px; flex: 1; }
.pr-avatar { width: 46px; height: 46px; border-radius: 12px; background: rgba(255,255,255,.9);
  color: #2b3fd0; font-weight: 700; font-size: 18px; display: flex; align-items: center; justify-content: center; }
.pr-out { width: 52px; height: 52px; border-radius: 50%; background: rgba(255,255,255,.18);
  display: flex; align-items: center; justify-content: center; }
.pr-hello { margin-top: 26px; font-size: 16px; color: rgba(255,255,255,.82); }
.pr-name { font-size: 20px; font-weight: 600; margin-top: 4px; line-height: 1.25; }
.pr-sub { font-size: 13px; color: rgba(255,255,255,.75); margin-top: 3px; }

/* ---------- kartu presensi ---------- */
.pr-card { background: #fff; border-radius: 26px; margin: -52px 22px 0; padding: 20px 18px 18px;
  box-shadow: 0 8px 28px rgba(37,52,140,.10); position: relative; z-index: 2; }
.pr-date { display: flex; align-items: center; gap: 12px; color: #374151; font-size: 17px; font-weight: 500; }
.pr-date svg { color: #4b5563; }
.pr-badge { margin-left: auto; background: #d9f7e8; color: #12a15f; font-weight: 600; font-size: 15px;
  padding: 7px 18px; border-radius: 999px; }
.pr-tiles { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 18px; }
.pr-tile { border-radius: 18px; text-align: center; padding: 16px 4px 14px; }
.pr-tile .lb { font-size: 14px; font-weight: 500; margin-top: 6px; }
.pr-tile .vl { font-size: 24px; font-weight: 700; margin-top: 6px; line-height: 1.1; }
.pr-tile .vl.sm { font-size: 17px; margin-top: 10px; }
.t-green { background: #e6f9f0; color: #14a068; }
.t-yellow { background: #fff6dc; color: #c58a05; }
.t-red { background: #fdeaea; color: #e04a4a; }
.pr-info { margin-top: 16px; background: #eaf1ff; border-radius: 16px; padding: 14px 16px; color: #2f5fd0; }
.pr-info .h { display: flex; align-items: center; gap: 10px; font-size: 17px; font-weight: 500; }
.pr-info .s { font-size: 15px; margin-top: 4px; color: #3b6fdc; }

/* kolom tidak boleh turun ke bawah di layar sempit */
[data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; }
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { min-width: 0 !important; flex: 1 1 0 !important; width: auto !important; }

/* ---------- menu ---------- */
.st-key-menucard { width: calc(100% - 44px) !important; background: #fff; border-radius: 26px; margin: 22px 22px 0; padding: 22px 12px 18px;
  box-shadow: 0 8px 28px rgba(37,52,140,.08); }
.pr-menu-title { font-size: 21px; font-weight: 600; color: #1f2937; padding: 0 8px 22px; }
.mn { text-align: center; }
.mn .bx { width: 66px; height: 66px; border-radius: 20px; margin: 0 auto; display: flex;
  align-items: center; justify-content: center; }
.mn .tx { font-size: 15px; color: #6b7280; margin-top: 10px; line-height: 1.25; }
.b-pink { background: #fde8ec; color: #e0436a; }
.b-sky { background: #e3f3fd; color: #1b9bd8; }
.b-mint { background: #dff7ec; color: #15a36a; }
.b-lav { background: #ece9fd; color: #5a47d6; }

/* ---------- tombol transparan di atas tampilan custom ---------- */
[class*="st-key-ov_"] { position: relative; }
[class*="st-key-ov_"] [data-testid="stElementContainer"]:has(button) { position: absolute; inset: 0; z-index: 5; }
[class*="st-key-ov_"] [data-testid="stButton"], [class*="st-key-ov_"] .stButton { height: 100%; }
[class*="st-key-ov_"] button { width: 100% !important; height: 100% !important; opacity: 0; cursor: pointer; min-height: 0; }
.st-key-ov_keluar { position: absolute !important; top: 30px; right: 22px; width: 52px; height: 52px; z-index: 10; }
.st-key-ov_keluar > div:first-child { height: 52px; }

/* ---------- navigasi bawah ---------- */
.st-key-bottomnav { position: fixed; bottom: 0; left: 50%; transform: translateX(-50%); width: min(430px, 100%);
  background: #fff; z-index: 999; padding: 8px 6px 12px; border-radius: 22px 22px 0 0;
  box-shadow: 0 -6px 24px rgba(37,52,140,.10); }
.st-key-bottomnav [data-testid="stHorizontalBlock"] { gap: 0 !important; align-items: flex-end; }
.nv { text-align: center; color: #9aa3b2; font-size: 12px; padding-top: 6px; min-height: 52px; }
.nv svg { display: block; margin: 0 auto 3px; }
.nv.on { color: #2563eb; font-weight: 600; }
.nv-fab { width: 64px; height: 64px; border-radius: 50%; background: linear-gradient(135deg,#2563eb,#3b3fd8);
  color: #fff; margin: -34px auto 0; display: flex; align-items: center; justify-content: center;
  box-shadow: 0 8px 20px rgba(37,63,216,.45); border: 5px solid #fff; box-sizing: content-box; }
.st-key-ov_fab { height: 64px; }

/* ---------- halaman dalam ---------- */
.pr-page-head { background: linear-gradient(135deg, #2563eb 0%, #3b3fd8 100%); color: #fff;
  padding: 30px 22px 56px; font-size: 22px; font-weight: 600; border-radius: 0 0 28px 28px; }
.pr-page-head small { display: block; font-size: 13px; font-weight: 400; opacity: .8; margin-top: 4px; }
.st-key-pagebody { width: calc(100% - 36px) !important; margin: -34px 18px 0; background: #fff; border-radius: 24px; padding: 18px 16px;
  box-shadow: 0 8px 28px rgba(37,52,140,.10); position: relative; z-index: 2; }
.pr-row { display: flex; align-items: center; gap: 12px; padding: 12px 0; border-bottom: 1px solid #eef0f5; }
.pr-row:last-child { border-bottom: 0; }
.pr-row .d { font-weight: 600; color: #1f2937; font-size: 15px; }
.pr-row .m { font-size: 13px; color: #6b7280; margin-top: 2px; }
.pr-row .r { margin-left: auto; text-align: right; font-size: 13px; font-weight: 600; }
.ok { color: #14a068; } .bad { color: #e04a4a; }
.pr-kv { display: flex; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid #eef0f5; font-size: 15px; }
.pr-kv:last-child { border-bottom: 0; }
.pr-kv span:first-child { color: #6b7280; } .pr-kv span:last-child { font-weight: 600; color: #1f2937; }
.pr-stat { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 8px; }
.pr-stat div { background: #eaf1ff; border-radius: 16px; padding: 14px; color: #2f5fd0; }
.pr-stat b { display: block; font-size: 22px; margin-top: 4px; }
.pr-stat span { font-size: 13px; }
.st-key-pagebody [data-testid="stButton"] button[kind="primary"] { background: #2563eb; border: 0; border-radius: 14px; }

/* === Perbaikan kontras dan fokus: login + admin === */
.st-key-loginbody, .st-key-adminbody {
  color: #111827 !important;
  color-scheme: light !important;
}
.st-key-loginbody :is(p, label, span, small, .stMarkdown, [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"]),
.st-key-adminbody :is(p, label, span, small, .stMarkdown, [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"]) {
  color: #1f2937;
}
.st-key-loginbody h1, .st-key-loginbody h2, .st-key-loginbody h3,
.st-key-adminbody h1, .st-key-adminbody h2, .st-key-adminbody h3 {
  color: #111827 !important;
}
.st-key-loginbody input, .st-key-adminbody input,
.st-key-loginbody textarea, .st-key-adminbody textarea,
.st-key-loginbody [data-baseweb="select"] > div,
.st-key-adminbody [data-baseweb="select"] > div {
  background: #ffffff !important;
  color: #111827 !important;
  border-color: #d1d5db !important;
  border-radius: 12px !important;
  transition: border-color .16s ease, box-shadow .16s ease, transform .16s ease;
}
.st-key-loginbody input::placeholder, .st-key-adminbody input::placeholder,
.st-key-loginbody textarea::placeholder, .st-key-adminbody textarea::placeholder {
  color: #6b7280 !important;
  opacity: 1 !important;
}
.st-key-loginbody input:focus, .st-key-adminbody input:focus,
.st-key-loginbody textarea:focus, .st-key-adminbody textarea:focus,
.st-key-loginbody [data-baseweb="select"] > div:focus-within,
.st-key-adminbody [data-baseweb="select"] > div:focus-within {
  border-color: #60a5fa !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.22) !important;
  outline: none !important;
}
.st-key-loginbody [data-testid="stRadio"] label,
.st-key-adminbody [data-testid="stRadio"] label {
  color: #374151 !important;
  cursor: pointer;
}
.st-key-loginbody [data-testid="stRadio"] label > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label > div:first-child {
  transition: transform .16s ease, box-shadow .16s ease;
  border-radius: 50%;
}
.st-key-loginbody [data-testid="stRadio"] label:has(input:checked) > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label:has(input:checked) > div:first-child {
  transform: scale(1.10);
  box-shadow: 0 0 0 4px rgba(59,130,246,.18);
}
.st-key-loginbody [data-testid="stRadio"] input:checked + div,
.st-key-adminbody [data-testid="stRadio"] input:checked + div {
  border-color: #3b82f6 !important;
}
.st-key-loginbody [data-testid="stRadio"] input:focus-visible + div,
.st-key-adminbody [data-testid="stRadio"] input:focus-visible + div {
  box-shadow: 0 0 0 4px rgba(59,130,246,.22) !important;
}
.st-key-loginbody button[kind="primary"], .st-key-adminbody button[kind="primary"] {
  color: #ffffff !important;
}
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  color: #1f2937 !important;
}
.st-key-adminbody [data-testid="stDataFrame"] *, .st-key-adminbody [data-testid="stTable"] * {
  color: inherit;
}
/* Header tetap menggunakan teks putih di atas gradasi biru. */
.auth-head, .admin-head, .auth-head *, .admin-head * { color: #ffffff; }

</style>
"""



CSS_AUTH_ADMIN = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"] {
  background: #eef2fb !important; font-family: 'Inter', system-ui, sans-serif;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], footer { display:none !important; }
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 520px; padding: 0 0 48px 0 !important; background:#eef2fb;
}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] { gap: 0.8rem; }
.auth-head, .admin-head {
  background: linear-gradient(135deg,#2563eb 0%,#3b3fd8 100%);
  color:#fff; padding: 30px 24px 64px; position:relative; overflow:hidden;
}
.auth-head { border-radius:0 0 30px 30px; }
.auth-brand { display:flex; align-items:center; gap:13px; font-size:24px; font-weight:700; }
.auth-brand-icon { width:48px;height:48px;border-radius:15px;background:rgba(255,255,255,.18);
  display:flex;align-items:center;justify-content:center;font-size:25px; }
.auth-sub { margin-top:20px; font-size:14px; opacity:.85; }
.auth-greet { margin-top:5px; font-size:22px; font-weight:650; }
.st-key-loginbody {
  width:calc(100% - 36px) !important; margin:-34px 18px 0; background:#fff;
  border-radius:24px; padding:22px 20px 24px; position:relative; z-index:2;
  box-shadow:0 8px 28px rgba(37,52,140,.10);
}
.st-key-loginbody [data-testid="stRadio"] > label { font-weight:600; color:#27334a; }
.st-key-loginbody input, .st-key-loginbody [data-baseweb="select"] > div {
  border-radius:12px !important;
}
.st-key-loginbody button[kind="primary"] {
  background:linear-gradient(135deg,#2563eb,#3b3fd8) !important;
  border:0 !important; border-radius:14px !important; min-height:46px; font-weight:600;
}
.st-key-adminbody {
  width:calc(100% - 36px) !important; margin:-34px 18px 0; background:#fff;
  border-radius:24px; padding:22px 20px 26px; position:relative; z-index:2;
  box-shadow:0 8px 28px rgba(37,52,140,.10);
}
[data-testid="stMainBlockContainer"]:has(.st-key-adminbody), .block-container:has(.st-key-adminbody) {
  max-width:1100px !important; padding-bottom:40px !important;
}
.admin-head { border-radius:0 0 30px 30px; padding-bottom:58px; }
.admin-title { font-size:25px; font-weight:700; }
.admin-sub { margin-top:5px; font-size:14px; opacity:.82; }
.st-key-adminbody [data-testid="stButton"] button[kind="primary"],
.st-key-adminbody button[kind="primary"] {
  background:#2563eb !important; border:0 !important; border-radius:12px !important;
}
.st-key-adminbody [data-testid="stRadio"] > label { font-weight:600; }
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  border-radius:14px; overflow:hidden;
}
@media (max-width:600px) {
  [data-testid="stMainBlockContainer"]:has(.st-key-adminbody), .block-container:has(.st-key-adminbody) {
    max-width:100% !important;
  }
  .st-key-adminbody { padding:16px 12px 22px; width:calc(100% - 24px) !important; margin-left:12px; }
  .admin-head { padding:26px 18px 58px; }
  .auth-head { padding:28px 20px 60px; }
}

/* === Perbaikan kontras dan fokus: login + admin === */
.st-key-loginbody, .st-key-adminbody {
  color: #111827 !important;
  color-scheme: light !important;
}
.st-key-loginbody :is(p, label, span, small, .stMarkdown, [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"]),
.st-key-adminbody :is(p, label, span, small, .stMarkdown, [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"]) {
  color: #1f2937;
}
.st-key-loginbody h1, .st-key-loginbody h2, .st-key-loginbody h3,
.st-key-adminbody h1, .st-key-adminbody h2, .st-key-adminbody h3 {
  color: #111827 !important;
}
.st-key-loginbody input, .st-key-adminbody input,
.st-key-loginbody textarea, .st-key-adminbody textarea,
.st-key-loginbody [data-baseweb="select"] > div,
.st-key-adminbody [data-baseweb="select"] > div {
  background: #ffffff !important;
  color: #111827 !important;
  border-color: #d1d5db !important;
  border-radius: 12px !important;
  transition: border-color .16s ease, box-shadow .16s ease, transform .16s ease;
}
.st-key-loginbody input::placeholder, .st-key-adminbody input::placeholder,
.st-key-loginbody textarea::placeholder, .st-key-adminbody textarea::placeholder {
  color: #6b7280 !important;
  opacity: 1 !important;
}
.st-key-loginbody input:focus, .st-key-adminbody input:focus,
.st-key-loginbody textarea:focus, .st-key-adminbody textarea:focus,
.st-key-loginbody [data-baseweb="select"] > div:focus-within,
.st-key-adminbody [data-baseweb="select"] > div:focus-within {
  border-color: #60a5fa !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.22) !important;
  outline: none !important;
}
.st-key-loginbody [data-testid="stRadio"] label,
.st-key-adminbody [data-testid="stRadio"] label {
  color: #374151 !important;
  cursor: pointer;
}
.st-key-loginbody [data-testid="stRadio"] label > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label > div:first-child {
  transition: transform .16s ease, box-shadow .16s ease;
  border-radius: 50%;
}
.st-key-loginbody [data-testid="stRadio"] label:has(input:checked) > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label:has(input:checked) > div:first-child {
  transform: scale(1.10);
  box-shadow: 0 0 0 4px rgba(59,130,246,.18);
}
.st-key-loginbody [data-testid="stRadio"] input:checked + div,
.st-key-adminbody [data-testid="stRadio"] input:checked + div {
  border-color: #3b82f6 !important;
}
.st-key-loginbody [data-testid="stRadio"] input:focus-visible + div,
.st-key-adminbody [data-testid="stRadio"] input:focus-visible + div {
  box-shadow: 0 0 0 4px rgba(59,130,246,.22) !important;
}
.st-key-loginbody button[kind="primary"], .st-key-adminbody button[kind="primary"] {
  color: #ffffff !important;
}
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  color: #1f2937 !important;
}
.st-key-adminbody [data-testid="stDataFrame"] *, .st-key-adminbody [data-testid="stTable"] * {
  color: inherit;
}
/* Header tetap menggunakan teks putih di atas gradasi biru. */
.auth-head, .admin-head, .auth-head *, .admin-head * { color: #ffffff; }

</style>
"""



CSS_AUTH_ADMIN = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"] {
  background: #eef2fb !important; font-family: 'Inter', system-ui, sans-serif;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], footer { display:none !important; }
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 520px; padding: 0 0 48px 0 !important; background:#eef2fb;
}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] { gap: 0.8rem; }
.auth-head, .admin-head {
  background: linear-gradient(135deg,#2563eb 0%,#3b3fd8 100%);
  color:#fff; padding: 30px 24px 64px; position:relative; overflow:hidden;
}
.auth-head { border-radius:0 0 30px 30px; }
.auth-brand { display:flex; align-items:center; gap:13px; font-size:24px; font-weight:700; }
.auth-brand-icon { width:48px;height:48px;border-radius:15px;background:rgba(255,255,255,.18);
  display:flex;align-items:center;justify-content:center;font-size:25px; }
.auth-sub { margin-top:20px; font-size:14px; opacity:.85; }
.auth-greet { margin-top:5px; font-size:22px; font-weight:650; }
.st-key-loginbody {
  width:calc(100% - 36px) !important; margin:-34px 18px 0; background:#fff;
  border-radius:24px; padding:22px 20px 24px; position:relative; z-index:2;
  box-shadow:0 8px 28px rgba(37,52,140,.10);
}
.st-key-loginbody [data-testid="stRadio"] > label { font-weight:600; color:#27334a; }
.st-key-loginbody input, .st-key-loginbody [data-baseweb="select"] > div {
  border-radius:12px !important;
}
.st-key-loginbody button[kind="primary"] {
  background:linear-gradient(135deg,#2563eb,#3b3fd8) !important;
  border:0 !important; border-radius:14px !important; min-height:46px; font-weight:600;
}
.st-key-adminbody {
  width:calc(100% - 36px) !important; margin:-34px 18px 0; background:#fff;
  border-radius:24px; padding:22px 20px 26px; position:relative; z-index:2;
  box-shadow:0 8px 28px rgba(37,52,140,.10);
}
[data-testid="stMainBlockContainer"]:has(.st-key-adminbody), .block-container:has(.st-key-adminbody) {
  max-width:1100px !important; padding-bottom:40px !important;
}
.admin-head { border-radius:0 0 30px 30px; padding-bottom:58px; }
.admin-title { font-size:25px; font-weight:700; }
.admin-sub { margin-top:5px; font-size:14px; opacity:.82; }
.st-key-adminbody [data-testid="stButton"] button[kind="primary"],
.st-key-adminbody button[kind="primary"] {
  background:#2563eb !important; border:0 !important; border-radius:12px !important;
}
.st-key-adminbody [data-testid="stRadio"] > label { font-weight:600; }
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  border-radius:14px; overflow:hidden;
}
@media (max-width:600px) {
  [data-testid="stMainBlockContainer"]:has(.st-key-adminbody), .block-container:has(.st-key-adminbody) {
    max-width:100% !important;
  }
  .st-key-adminbody { padding:16px 12px 22px; width:calc(100% - 24px) !important; margin-left:12px; }
  .admin-head { padding:26px 18px 58px; }
  .auth-head { padding:28px 20px 60px; }
}

/* === Perbaikan kontras dan fokus: login + admin === */
.st-key-loginbody, .st-key-adminbody {
  color: #111827 !important;
  color-scheme: light !important;
}
.st-key-loginbody :is(p, label, span, small, .stMarkdown, [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"]),
.st-key-adminbody :is(p, label, span, small, .stMarkdown, [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"]) {
  color: #1f2937;
}
.st-key-loginbody h1, .st-key-loginbody h2, .st-key-loginbody h3,
.st-key-adminbody h1, .st-key-adminbody h2, .st-key-adminbody h3 {
  color: #111827 !important;
}
.st-key-loginbody input, .st-key-adminbody input,
.st-key-loginbody textarea, .st-key-adminbody textarea,
.st-key-loginbody [data-baseweb="select"] > div,
.st-key-adminbody [data-baseweb="select"] > div {
  background: #ffffff !important;
  color: #111827 !important;
  border-color: #d1d5db !important;
  border-radius: 12px !important;
  transition: border-color .16s ease, box-shadow .16s ease, transform .16s ease;
}
.st-key-loginbody input::placeholder, .st-key-adminbody input::placeholder,
.st-key-loginbody textarea::placeholder, .st-key-adminbody textarea::placeholder {
  color: #6b7280 !important;
  opacity: 1 !important;
}
.st-key-loginbody input:focus, .st-key-adminbody input:focus,
.st-key-loginbody textarea:focus, .st-key-adminbody textarea:focus,
.st-key-loginbody [data-baseweb="select"] > div:focus-within,
.st-key-adminbody [data-baseweb="select"] > div:focus-within {
  border-color: #60a5fa !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,.22) !important;
  outline: none !important;
}
.st-key-loginbody [data-testid="stRadio"] label,
.st-key-adminbody [data-testid="stRadio"] label {
  color: #374151 !important;
  cursor: pointer;
}
.st-key-loginbody [data-testid="stRadio"] label > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label > div:first-child {
  transition: transform .16s ease, box-shadow .16s ease;
  border-radius: 50%;
}
.st-key-loginbody [data-testid="stRadio"] label:has(input:checked) > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label:has(input:checked) > div:first-child {
  transform: scale(1.10);
  box-shadow: 0 0 0 4px rgba(59,130,246,.18);
}
.st-key-loginbody [data-testid="stRadio"] input:checked + div,
.st-key-adminbody [data-testid="stRadio"] input:checked + div {
  border-color: #3b82f6 !important;
}
.st-key-loginbody [data-testid="stRadio"] input:focus-visible + div,
.st-key-adminbody [data-testid="stRadio"] input:focus-visible + div {
  box-shadow: 0 0 0 4px rgba(59,130,246,.22) !important;
}
.st-key-loginbody button[kind="primary"], .st-key-adminbody button[kind="primary"] {
  color: #ffffff !important;
}
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  color: #1f2937 !important;
}
.st-key-adminbody [data-testid="stDataFrame"] *, .st-key-adminbody [data-testid="stTable"] * {
  color: inherit;
}
/* Header tetap menggunakan teks putih di atas gradasi biru. */
.auth-head, .admin-head, .auth-head *, .admin-head * { color: #ffffff; }


/* ===== Tema Presensi: biru, putih, pastel — tanpa teks hitam ===== */
:root {
  color-scheme: light;
  --pr-blue: #285bea;
  --pr-blue-deep: #3544d8;
  --pr-ink: #263858;
  --pr-muted: #64748b;
  --pr-border: #dce6f8;
  --pr-page: #eef3ff;
  --pr-card: #ffffff;
}
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"],
[data-testid="stMain"], [data-testid="stMainBlockContainer"], .block-container {
  background: var(--pr-page) !important;
  color: var(--pr-ink) !important;
  color-scheme: light !important;
  font-family: 'Inter', system-ui, sans-serif !important;
}
[data-testid="stMainBlockContainer"], .block-container {
  padding-top: 1rem !important;
}
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li, [data-testid="stWidgetLabel"],
[data-testid="stCaptionContainer"], label, .stText, .stMarkdown,
[data-testid="stMetricLabel"], [data-testid="stMetricValue"] {
  color: var(--pr-ink) !important;
}
h1, h2, h3, h4, h5, h6 {
  color: var(--pr-ink) !important;
  letter-spacing: -.02em;
}
.st-key-loginbody, .st-key-adminbody {
  background: var(--pr-card) !important;
  border: 1px solid rgba(211, 224, 250, .9) !important;
  border-radius: 26px !important;
  box-shadow: 0 12px 34px rgba(45, 77, 160, .10) !important;
  color: var(--pr-ink) !important;
}
.auth-head, .admin-head {
  background: linear-gradient(135deg, #2864ef 0%, #3d3dd7 100%) !important;
  color: #ffffff !important;
  box-shadow: 0 10px 26px rgba(45, 77, 200, .12);
}
.auth-head *, .admin-head *, .auth-brand, .auth-sub, .auth-greet,
.admin-title, .admin-sub {
  color: #ffffff !important;
}
.auth-brand-icon {
  background: rgba(255,255,255,.18) !important;
  border: 1px solid rgba(255,255,255,.18);
}
.st-key-loginbody [data-testid="stWidgetLabel"],
.st-key-adminbody [data-testid="stWidgetLabel"],
.st-key-loginbody label, .st-key-adminbody label {
  color: var(--pr-ink) !important;
  font-weight: 600;
}
/* Input putih, tulisan biru tua, fokus biru lembut */
.st-key-loginbody input, .st-key-adminbody input,
.st-key-loginbody textarea, .st-key-adminbody textarea,
.st-key-loginbody [data-baseweb="select"] > div,
.st-key-adminbody [data-baseweb="select"] > div,
.st-key-loginbody [data-baseweb="input"] > div,
.st-key-adminbody [data-baseweb="input"] > div {
  background: #ffffff !important;
  color: var(--pr-ink) !important;
  -webkit-text-fill-color: var(--pr-ink) !important;
  border: 1px solid var(--pr-border) !important;
  border-radius: 13px !important;
  box-shadow: 0 2px 7px rgba(52, 82, 160, .035);
  transition: border-color .16s ease, box-shadow .16s ease, transform .16s ease;
}
.st-key-loginbody input::placeholder, .st-key-adminbody input::placeholder,
.st-key-loginbody textarea::placeholder, .st-key-adminbody textarea::placeholder {
  color: #8493ad !important;
  -webkit-text-fill-color: #8493ad !important;
  opacity: 1 !important;
}
.st-key-loginbody input:focus, .st-key-adminbody input:focus,
.st-key-loginbody textarea:focus, .st-key-adminbody textarea:focus,
.st-key-loginbody [data-baseweb="select"] > div:focus-within,
.st-key-adminbody [data-baseweb="select"] > div:focus-within,
.st-key-loginbody [data-baseweb="input"] > div:focus-within,
.st-key-adminbody [data-baseweb="input"] > div:focus-within {
  border-color: #6da0ff !important;
  box-shadow: 0 0 0 3px rgba(67, 131, 255, .18) !important;
}
/* Radio: tidak hitam, pilihan mendapat aksen biru dan sedikit membesar */
.st-key-loginbody [data-testid="stRadio"] label,
.st-key-adminbody [data-testid="stRadio"] label {
  color: var(--pr-ink) !important;
}
.st-key-loginbody [data-testid="stRadio"] label > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label > div:first-child {
  border-radius: 50%;
  transition: transform .16s ease, box-shadow .16s ease;
}
.st-key-loginbody [data-testid="stRadio"] label:has(input:checked) > div:first-child,
.st-key-adminbody [data-testid="stRadio"] label:has(input:checked) > div:first-child {
  transform: scale(1.12);
  box-shadow: 0 0 0 4px rgba(59, 130, 246, .18);
}
.st-key-loginbody [data-testid="stRadio"] input:focus-visible + div,
.st-key-adminbody [data-testid="stRadio"] input:focus-visible + div {
  box-shadow: 0 0 0 4px rgba(59, 130, 246, .22) !important;
}
/* Tombol senada dengan header dan kartu pastel */
.st-key-loginbody button[kind="primary"],
.st-key-adminbody button[kind="primary"],
.st-key-loginbody [data-testid="stFormSubmitButton"] button,
.st-key-adminbody [data-testid="stFormSubmitButton"] button {
  background: linear-gradient(135deg, #3269f2, #4646dc) !important;
  color: #ffffff !important;
  border: 0 !important;
  border-radius: 14px !important;
  min-height: 44px;
  font-weight: 650 !important;
  box-shadow: 0 5px 13px rgba(49, 87, 218, .16);
  transition: transform .15s ease, box-shadow .15s ease;
}
.st-key-loginbody button[kind="primary"]:hover,
.st-key-adminbody button[kind="primary"]:hover {
  box-shadow: 0 8px 18px rgba(49, 87, 218, .22);
  transform: translateY(-1px);
}
/* Tombol sekunder juga memakai warna lembut, bukan hitam */
.st-key-loginbody button:not([kind="primary"]),
.st-key-adminbody button:not([kind="primary"]) {
  background: #edf3ff !important;
  color: #3158bd !important;
  border: 1px solid #dce7ff !important;
  border-radius: 12px !important;
}
.st-key-loginbody button:not([kind="primary"]) *,
.st-key-adminbody button:not([kind="primary"]) * {
  color: #3158bd !important;
}
/* Kartu metrik dan tabel admin */
.st-key-adminbody [data-testid="stMetric"],
.st-key-adminbody [data-testid="stVerticalBlockBorderWrapper"] {
  background: #ffffff !important;
  border: 1px solid #e1eaff !important;
  border-radius: 17px !important;
  box-shadow: 0 5px 16px rgba(45, 77, 160, .055);
}
.st-key-adminbody [data-testid="stDataFrame"],
.st-key-adminbody [data-testid="stTable"] {
  border: 1px solid #dce6f8 !important;
  border-radius: 14px !important;
  overflow: hidden;
  color: var(--pr-ink) !important;
}
.st-key-adminbody [data-testid="stDataFrame"] *,
.st-key-adminbody [data-testid="stTable"] * {
  color: var(--pr-ink) !important;
}
.st-key-adminbody hr, .st-key-loginbody hr {
  border-color: #e4ebfa !important;
}
[data-testid="stAlert"] {
  border-radius: 14px !important;
}
@media (max-width: 600px) {
  [data-testid="stMainBlockContainer"], .block-container {
    padding-left: 0 !important;
    padding-right: 0 !important;
  }
  .auth-head, .admin-head { padding-left: 20px !important; padding-right: 20px !important; }
  .st-key-loginbody { padding: 20px 16px 22px !important; }
  .st-key-adminbody { padding: 18px 14px 22px !important; }
}

</style>
"""


def inisial(nama):
    bagian = [b for b in nama.split() if b]
    return "".join(b[0] for b in bagian[:2]).upper() or "?"


def get_absen_hari(eng, nama, tanggal):
    """Absen hari ini milik peserta (baris berhasil diutamakan), atau None."""
    q = (
        select(absensi_t.c.jam, absensi_t.c.status, absensi_t.c.uang_makan, absensi_t.c.jarak)
        .where(and_(absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal))
        .order_by(absensi_t.c.jam)
    )
    with eng.connect() as c:
        rows = c.execute(q).mappings().all()
    for r in rows:
        if r["status"] == "berhasil":
            return r
    return rows[-1] if rows else None


def get_riwayat_peserta(eng, nama, limit=60):
    q = (
        select(absensi_t.c.tanggal, absensi_t.c.jam, absensi_t.c.status,
               absensi_t.c.jarak, absensi_t.c.uang_makan)
        .where(absensi_t.c.nama == nama)
        .order_by(absensi_t.c.tanggal.desc(), absensi_t.c.jam.desc())
        .limit(limit)
    )
    with eng.connect() as c:
        return c.execute(q).mappings().all()


def pindah_tab(tab):
    st.session_state.tab = tab


def tombol_overlay(key, html, tab=None, toast=None):
    """Tampilan custom (HTML) + tombol transparan di atasnya supaya bisa diklik."""
    with st.container(key=f"ov_{key}"):
        st.markdown(html, unsafe_allow_html=True)
        if st.button("‎", key=f"ovb_{key}", width="stretch"):
            if tab:
                st.session_state.tab = tab
                st.rerun()
            if toast:
                st.toast(toast, icon="ℹ️")


def navigasi_bawah(aktif):
    def item(ikon, label, tab):
        return f'<div class="nv {"on" if aktif == tab else ""}">{ic(ikon, 22)}{label}</div>'

    with st.container(key="bottomnav"):
        cols = st.columns([1, 1, 1, 1, 1])
        with cols[0]:
            tombol_overlay("nv_beranda", item("home", "Beranda", "beranda"), tab="beranda")
        with cols[1]:
            tombol_overlay("nv_riwayat", item("history", "Riwayat", "riwayat"), tab="riwayat")
        with cols[2]:
            tombol_overlay("fab", f'<div class="nv-fab">{ic("fingerprint", 30)}</div>', tab="absen")
        with cols[3]:
            tombol_overlay("nv_rekap", item("chart", "Rekap", "rekap"), tab="rekap")
        with cols[4]:
            tombol_overlay("nv_profil", item("user", "Profil", "profil"), tab="profil")


def header_halaman(judul, sub=""):
    sub_html = f"<small>{sub}</small>" if sub else ""
    st.markdown(f'<div class="pr-page-head">{judul}{sub_html}</div>', unsafe_allow_html=True)


def tab_beranda(eng, nama, peserta):
    now = datetime.now(WITA)
    absen = get_absen_hari(eng, nama, now.strftime("%Y-%m-%d"))
    berhasil = absen is not None and absen["status"] == "berhasil"

    jam_masuk = absen["jam"][:5] if berhasil else "--:--"
    uang = rupiah(absen["uang_makan"]) if berhasil else "-"
    batas = BATAS_JAM_MASUK.strftime("%H:%M")
    if berhasil:
        info_judul, info_sub = "Status presensi", f"Absen masuk berhasil pukul {absen['jam'][:5]} WITA"
    elif absen is not None:
        info_judul, info_sub = "Status presensi", "Absen terakhir gagal (di luar area). Coba lagi."
    else:
        info_judul, info_sub = "Status presensi", f"Belum absen. Waktu presensi sebelum {batas} WITA"

    st.markdown(
        f"""<div class="pr-head">
<div class="pr-top">
<div class="pr-logo">{ic("fingerprint", 28)}</div>
<div class="pr-title">Presensi</div>
<div class="pr-avatar">{inisial(nama)}</div>
<div class="pr-out">{ic("logout", 26)}</div>
</div>
<div class="pr-hello">{salam(now.hour)} 👋</div>
<div class="pr-name">{nama}</div>
<div class="pr-sub">{peserta["divisi"]} • Gedung {peserta["gedung"]}</div>
</div>
<div class="pr-card">
<div class="pr-date">{ic("calendar", 26)}<span>{tgl_indonesia(now.date())}</span><span class="pr-badge">Aktif</span></div>
<div class="pr-tiles">
<div class="pr-tile t-green">{ic("login", 26)}<div class="lb">Masuk</div><div class="vl">{jam_masuk}</div></div>
<div class="pr-tile t-yellow">{ic("file", 26)}<div class="lb">Uang Makan</div><div class="vl sm">{uang}</div></div>
<div class="pr-tile t-red">{ic("clock", 26)}<div class="lb">Batas Masuk</div><div class="vl">{batas}</div></div>
</div>
<div class="pr-info"><div class="h">{ic("clock", 22)}<span>{info_judul}</span></div><div class="s">{info_sub}</div></div>
</div>""",
        unsafe_allow_html=True,
    )
    # tombol keluar transparan di atas ikon logout header
    with st.container(key="ov_keluar"):
        if st.button("‎", key="ovb_keluar"):
            keluar()

    with st.container(key="menucard"):
        st.markdown('<div class="pr-menu-title">Menu</div>', unsafe_allow_html=True)
        menu = [
            ("izin", "calendar_days", "b-pink", "Izin"),
            ("approval", "inbox", "b-sky", "Approval Izin"),
            ("dinas", "plane", "b-mint", "Dinas Luar"),
            ("usulan", "file", "b-lav", "Usulan"),
        ]
        cols = st.columns(4)
        for col, (kunci, ikon, warna, label) in zip(cols, menu):
            with col:
                tombol_overlay(
                    f"mn_{kunci}",
                    f'<div class="mn"><div class="bx {warna}">{ic(ikon, 30)}</div><div class="tx">{label}</div></div>',
                    toast=f"Fitur {label} belum tersedia.",
                )


def tab_absen(eng, nama, peserta):
    gedung, divisi = peserta["gedung"], peserta["divisi"]
    header_halaman("Absen Masuk", f"{nama} • Gedung {gedung} • {divisi}")

    with st.container(key="pagebody"):
        lokasi = get_geolocation()
        coords = (lokasi or {}).get("coords")
        if coords:
            st.success(f"📡 Lokasi terdeteksi (akurasi ±{coords.get('accuracy', 0):.0f} m)")
        else:
            st.info("Menunggu lokasi... klik **Allow/Izinkan** saat browser meminta akses lokasi.")

        foto = st.camera_input("Foto selfie untuk absen")

        if st.button("Absen Sekarang", type="primary", width="stretch"):
            if not coords:
                st.error("Lokasi belum terdeteksi. Aktifkan GPS/izin lokasi lalu coba lagi.")
                return
            if foto is None:
                st.error("Foto belum diambil.")
                return

            now = datetime.now(WITA)
            tanggal, jam = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")

            if sudah_berhasil(eng, nama, tanggal, "masuk"):
                st.warning("Kamu sudah absen masuk hari ini.")
                return

            if now.time() > BATAS_JAM_MASUK:
                st.error(
                    f"❌ Absen ditolak: sudah lewat jam {BATAS_JAM_MASUK.strftime('%H:%M')} WITA "
                    f"(sekarang {now.strftime('%H:%M')})."
                )
                return

            lat, lon = coords["latitude"], coords["longitude"]
            cocok, jarak = cek_lokasi(lat, lon, gedung)
            status = "berhasil" if cocok else "gagal_lokasi"
            uang = UANG_MAKAN if cocok else 0

            simpan_absen(
                eng, nama=nama, tanggal=tanggal, jam=jam, jenis="masuk",
                lat=lat, lon=lon, akurasi=coords.get("accuracy"), jarak=jarak,
                status=status, uang_makan=uang, foto=kompres_foto(foto.getvalue()),
            )

            if cocok:
                st.success(f"✅ Absen masuk berhasil pukul {jam} WITA.")
                st.info(f"💰 Uang makan hari ini: {rupiah(uang)}")
            else:
                st.error(
                    f"❌ Absen GAGAL: kamu ~{jarak:.0f} m dari Gedung {gedung}, "
                    "di luar area yang diizinkan. Pastikan kamu sudah di gedung lalu coba lagi."
                )


def tab_riwayat(eng, nama):
    header_halaman("Riwayat Presensi", "60 absen terakhir")
    rows = get_riwayat_peserta(eng, nama)
    with st.container(key="pagebody"):
        if not rows:
            st.info("Belum ada riwayat absen.")
            return
        html = ""
        for r in rows:
            d = datetime.strptime(r["tanggal"], "%Y-%m-%d").date()
            ok = r["status"] == "berhasil"
            ket = rupiah(r["uang_makan"]) if ok else "Gagal lokasi"
            html += (
                f'<div class="pr-row"><div><div class="d">{tgl_indonesia(d)}</div>'
                f'<div class="m">Masuk {r["jam"][:5]} WITA • {r["jarak"] or 0:.0f} m</div></div>'
                f'<div class="r {"ok" if ok else "bad"}">{ket}</div></div>'
            )
        st.markdown(html, unsafe_allow_html=True)


def tab_rekap(eng, nama):
    now = datetime.now(WITA)
    header_halaman("Rekap Bulan Ini", f"{BULAN_ID[now.month - 1]} {now.year}")
    prefix = now.strftime("%Y-%m-")
    rows = [r for r in get_riwayat_peserta(eng, nama, limit=400) if r["tanggal"].startswith(prefix)]
    hadir = len({r["tanggal"] for r in rows if r["status"] == "berhasil"})
    total = sum(r["uang_makan"] for r in rows if r["status"] == "berhasil")
    gagal = sum(1 for r in rows if r["status"] != "berhasil")
    with st.container(key="pagebody"):
        st.markdown(
            f'<div class="pr-stat"><div><span>Hari hadir</span><b>{hadir}</b></div>'
            f'<div><span>Total uang makan</span><b>{rupiah(total)}</b></div></div>'
            f'<div class="pr-kv"><span>Percobaan gagal</span><span>{gagal}</span></div>',
            unsafe_allow_html=True,
        )


def tab_profil(nama, peserta):
    header_halaman("Profil", "Data peserta PKL")
    with st.container(key="pagebody"):
        st.markdown(
            f'<div class="pr-kv"><span>Nama</span><span>{nama}</span></div>'
            f'<div class="pr-kv"><span>Divisi</span><span>{peserta["divisi"]}</span></div>'
            f'<div class="pr-kv"><span>Gedung</span><span>Gedung {peserta["gedung"]}</span></div>'
            f'<div class="pr-kv"><span>Status</span><span class="ok">Aktif</span></div>',
            unsafe_allow_html=True,
        )
        if st.button("Keluar", key="keluar_profil", width="stretch"):
            keluar()


def halaman_peserta(eng, nama):
    """Tampilan peserta setelah login: Beranda (Presensi) + Absen / Riwayat / Rekap / Profil."""
    peserta = get_peserta(eng, nama)
    if peserta is None:
        st.error("Data kamu tidak ditemukan (mungkin sudah dihapus admin). Silakan keluar lalu hubungi admin.")
        if st.button("Keluar"):
            keluar()
        return

    st.markdown(CSS_PRESENSI, unsafe_allow_html=True)
    tab = st.session_state.get("tab", "beranda")
    if tab == "absen":
        tab_absen(eng, nama, peserta)
    elif tab == "riwayat":
        tab_riwayat(eng, nama)
    elif tab == "rekap":
        tab_rekap(eng, nama)
    elif tab == "profil":
        tab_profil(nama, peserta)
    else:
        tab = "beranda"
        tab_beranda(eng, nama, peserta)
    navigasi_bawah(tab)


# ------------------------------------------------------------------
# HALAMAN ADMIN: KELOLA PESERTA
# ------------------------------------------------------------------
def halaman_kelola_peserta(eng):
    st.header("👥 Kelola Peserta PKL")
    if not admin_login():
        return

    with st.form("tambah_peserta", clear_on_submit=True):
        nama = st.text_input("Nama peserta")
        divisi = st.selectbox("Divisi", list(DIVISI_GEDUNG.keys()))
        gedung = DIVISI_GEDUNG[divisi]
        pin = st.text_input("PIN (4-6 angka)", max_chars=6)
        kirim = st.form_submit_button("Tambah Peserta")

    if kirim:
        if not nama.strip():
            st.error("Nama wajib diisi.")
        elif not (pin.isdigit() and 4 <= len(pin) <= 6):
            st.error("PIN harus 4-6 angka.")
        else:
            try:
                add_peserta(eng, nama, gedung, divisi, pin)
                st.success(f"{nama.strip()} ditambahkan ke Gedung {gedung}. Kasih tahu PIN-nya ke yang bersangkutan.")
            except IntegrityError:
                st.error("Nama itu sudah terdaftar.")

    pesan = st.session_state.pop("pesan_hapus", None)
    if pesan:
        st.success(pesan)

    df = list_peserta(eng)
    st.dataframe(
        beri_nomor(df.rename(columns={"nama": "Nama", "gedung": "Gedung", "divisi": "Divisi"})),
        width="stretch",
    )

    if not df.empty:
        with st.expander("Edit Divisi / Reset PIN / Hapus peserta"):
            target = st.selectbox("Pilih peserta", df["nama"], key="target_kelola")

            opsi_divisi = list(DIVISI_GEDUNG.keys())
            divisi_lama = df.loc[df["nama"] == target, "divisi"].values[0]
            idx_default = opsi_divisi.index(divisi_lama) if divisi_lama in opsi_divisi else 0
            divisi_baru = st.selectbox("Divisi", opsi_divisi, index=idx_default, key="divisi_baru")
            st.caption(f"Ditempatkan di: Gedung {DIVISI_GEDUNG[divisi_baru]}")
            if st.button("Simpan Divisi"):
                set_divisi(eng, target, divisi_baru, DIVISI_GEDUNG[divisi_baru])
                st.success("Divisi & Gedung diperbarui.")
                st.rerun()

            pin_baru = st.text_input("PIN baru (4-6 angka)", max_chars=6, key="pin_baru")
            hapus_absen_juga = st.checkbox(
                "Saat menghapus peserta, hapus juga seluruh riwayat absensinya",
                key="hapus_absen_juga",
            )
            c1, c2 = st.columns(2)
            if c1.button("Reset PIN"):
                if pin_baru.isdigit() and 4 <= len(pin_baru) <= 6:
                    reset_pin(eng, target, pin_baru)
                    st.success("PIN diperbarui.")
                else:
                    st.error("PIN harus 4-6 angka.")
            if c2.button("Hapus peserta"):
                hapus_peserta(eng, target, dengan_absensi=hapus_absen_juga)
                if hapus_absen_juga:
                    st.session_state["pesan_hapus"] = f"{target} dihapus beserta seluruh riwayat absensinya."
                else:
                    st.session_state["pesan_hapus"] = (
                        f"{target} dihapus dari daftar peserta. Riwayat absensinya masih tersimpan "
                        "(hapus lewat bagian 'Hapus data absensi' di bawah kalau perlu)."
                    )
                st.rerun()

    with st.expander("🗑️ Hapus data absensi (tidak bisa dibatalkan)"):
        nama_absen = daftar_nama_absensi(eng)
        if not nama_absen:
            st.info("Belum ada data absensi.")
        else:
            semua = "(SEMUA peserta)"
            pilihan = st.selectbox("Hapus absensi milik", [semua] + nama_absen, key="hapus_absen_pilihan")
            konfirmasi = st.text_input("Ketik HAPUS (huruf besar) untuk konfirmasi", key="hapus_absen_konfirmasi")
            if st.button("Hapus data absensi", key="tombol_hapus_absen"):
                if konfirmasi.strip() != "HAPUS":
                    st.error("Konfirmasi belum benar. Ketik persis: HAPUS")
                else:
                    jumlah = hapus_absensi(eng, None if pilihan == semua else pilihan)
                    st.success(f"{jumlah} data absensi (beserta fotonya) dihapus.")


# ------------------------------------------------------------------
# HALAMAN ADMIN: REKAP, GRAFIK, EXPORT
# ------------------------------------------------------------------
def halaman_rekap(eng):
    st.header("📊 Rekap Absensi Masuk & Uang Makan")
    if not admin_login():
        return

    hari_ini = datetime.now(WITA).date()
    c1, c2 = st.columns(2)
    bulan = c1.selectbox("Bulan", list(range(1, 13)), index=hari_ini.month - 1)
    tahun = c2.number_input("Tahun", value=hari_ini.year, step=1)

    df = get_absensi_bulan(eng, int(tahun), int(bulan))
    if df.empty:
        st.info("Belum ada data absensi di bulan ini.")
        return

    berhasil = df[df["status"] == "berhasil"]
    rekap = (
        berhasil
        .groupby("nama")
        .agg(hari_hadir=("tanggal", "nunique"), total_uang_makan=("uang_makan", "sum"))
        .reset_index()
    )
    rekap.columns = ["Nama", "Hari Hadir", "Total Uang Makan"]

    m1, m2 = st.columns(2)
    m1.metric("Total uang makan bulan ini", f"Rp{int(rekap['Total Uang Makan'].sum()):,}".replace(",", "."))
    m2.metric("Percobaan absen gagal lokasi/telat", int((df["status"] != "berhasil").sum()))

    st.subheader("Total Uang Makan per Peserta")
    st.dataframe(beri_nomor(rekap), width="stretch")

    st.subheader("Grafik")
    if not rekap.empty:
        st.caption("Hari hadir per peserta")
        st.bar_chart(rekap.set_index("Nama")["Hari Hadir"])
    harian = berhasil.groupby("tanggal").size()
    if not harian.empty:
        st.caption("Jumlah peserta yang hadir per hari")
        st.bar_chart(harian)

    st.subheader("Detail Absensi")
    detail = df.rename(columns={
        "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis",
        "status": "Status", "jarak": "Jarak (m)", "uang_makan": "Uang Makan",
    })
    detail["Jarak (m)"] = detail["Jarak (m)"].round(0)
    st.dataframe(beri_nomor(detail), width="stretch")

    st.download_button(
        "⬇️ Download Excel", data=excel_bytes(rekap, detail),
        file_name=f"rekap_absensi_{int(tahun)}_{int(bulan):02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ------------------------------------------------------------------
# HALAMAN ADMIN: BUKTI FOTO
# ------------------------------------------------------------------
def halaman_foto(eng):
    st.header("🖼️ Bukti Foto Absensi")
    if not admin_login():
        return

    tgl = st.date_input("Tanggal", value=datetime.now(WITA).date())
    rows = get_absensi_hari_foto(eng, tgl.strftime("%Y-%m-%d"))
    if not rows:
        st.info("Tidak ada absensi di tanggal ini.")
        return

    kolom = st.columns(3)
    for i, r in enumerate(rows):
        with kolom[i % 3]:
            if r["foto"]:
                st.image(bytes(r["foto"]), width="stretch")
            ikon = "✅" if r["status"] == "berhasil" else "❌"
            st.caption(f"{ikon} {r['nama']}\n\n{r['jenis']} • {r['jam']} • {r['jarak']:.0f} m")


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def keluar():
    """Logout: hapus peran & data login, kembali ke layar login."""
    for k in ("peran", "nama_login", "admin_ok", "tab"):
        st.session_state.pop(k, None)
    st.rerun()


def halaman_login(eng):
    """Layar pertama: pilih peran, lalu login. Desain mengikuti tema Presensi peserta."""
    st.markdown(CSS_AUTH_ADMIN, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="auth-head">
          <div class="auth-brand"><div class="auth-brand-icon">◉</div><div>Presensi</div></div>
          <div class="auth-sub">BANK SULSELBAR • ABSENSI PKL</div>
          <div class="auth-greet">Selamat Datang 👋</div>
          <div style="font-size:14px;opacity:.82;margin-top:5px">Silakan masuk untuk melanjutkan</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.container(key="loginbody"):
        st.subheader("Masuk ke akun")
        peran = st.radio("Masuk sebagai", ["Peserta PKL", "Admin"], horizontal=True, key="login_peran")

        if peran == "Peserta PKL":
            df = list_peserta(eng)
            if df.empty:
                st.warning("Belum ada peserta terdaftar. Hubungi admin untuk didaftarkan.")
                return
            with st.form("login_peserta"):
                nama = st.selectbox("Nama peserta", df["nama"])
                pin = st.text_input("PIN", type="password", max_chars=6)
                kirim = st.form_submit_button("Masuk", type="primary", width="stretch")
            if kirim:
                if cek_pin(eng, nama, pin):
                    st.session_state.peran = "user"
                    st.session_state.nama_login = nama
                    st.rerun()
                else:
                    st.error("PIN salah. Silakan periksa kembali.")
        else:
            pw_benar = password_admin()
            if pw_benar is None:
                st.error("admin_password belum diatur di Secrets. Hubungi pengelola aplikasi.")
                return
            if not secret("admin_password") and not secret("database_url"):
                st.info("Mode lokal: password admin default = admin123")
            with st.form("login_admin"):
                pw = st.text_input("Password admin", type="password")
                kirim = st.form_submit_button("Masuk sebagai admin", type="primary", width="stretch")
            if kirim:
                if hmac.compare_digest(pw, pw_benar):
                    st.session_state.peran = "admin"
                    st.session_state.admin_ok = True
                    st.rerun()
                else:
                    st.error("Password salah. Silakan coba lagi.")


def halaman_admin(eng):
    """Area admin dengan gaya visual yang selaras dengan halaman peserta."""
    st.markdown(CSS_AUTH_ADMIN, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="admin-head">
          <div class="admin-title">Presensi · Panel Admin</div>
          <div class="admin-sub">BANK SULSELBAR • Pengelolaan Absensi PKL</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.container(key="adminbody"):
        kiri, kanan = st.columns([4, 1])
        kiri.caption("Masuk sebagai admin")
        if kanan.button("Keluar", key="keluar_admin", width="stretch"):
            keluar()

        menu = st.radio(
            "Menu admin",
            ["Kelola Peserta", "Rekap & Export", "Bukti Foto"],
            horizontal=True,
            key="menu_admin",
        )
        st.markdown("---")
        if menu == "Kelola Peserta":
            halaman_kelola_peserta(eng)
        elif menu == "Rekap & Export":
            halaman_rekap(eng)
        else:
            halaman_foto(eng)


def main():
    st.set_page_config(page_title="Presensi - Bank Sulselbar", page_icon="📍")
    eng = get_engine()

    peran = st.session_state.get("peran")
    nama = st.session_state.get("nama_login")
    if peran == "user" and nama:
        halaman_peserta(eng, nama)
    elif peran == "admin" and st.session_state.get("admin_ok"):
        halaman_admin(eng)
    else:
        halaman_login(eng)


if __name__ == "__main__":
    main()
