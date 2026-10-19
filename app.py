"""
Aplikasi Absensi Peserta PKL - Bank Sulselbar (versi siap deploy)
=================================================================
Fitur:
- Layar login di awal: pilih peran. Peserta (nama + PIN) -> halaman absen saja; Admin (password) -> area admin
- Absen MASUK saja (tidak ada absen pulang): foto selfie + lokasi GPS
- Login PIN per peserta (dibuat admin)
- Cek lokasi: harus di sekitar gedung yang terdaftar untuk peserta (A / B)
- Setiap peserta juga punya Divisi (dipilih dari daftar, Gedung ikut otomatis)
- Jam absen: masuk 07:00-09:00, pulang 16:30-18:00 (WITA); izin bisa kapan saja (dengan alasan)
- Uang makan Rp15.000 otomatis saat absen masuk berhasil
- Blokir absen dobel di hari yang sama
- Admin (pakai password): kelola peserta, rekap + grafik + export Excel, bukti foto
- Database: SQLite (lokal) atau Postgres/Supabase (cloud) - otomatis

Jalanin lokal:   streamlit run app.py
"""

import hashlib
import hmac
import io
import json
import os
import re
import shutil
import sqlite3
import tempfile
from functools import wraps
from copy import copy
from datetime import datetime, date, time as dtime, timedelta, timezone
from math import radians, sin, cos, sqrt, atan2
from html import escape as esc
from urllib.parse import quote

import math
import openpyxl
import requests
import streamlit.components.v1 as components
from openpyxl.styles import Alignment
from openpyxl.worksheet.properties import PageSetupProperties
import pandas as pd
import streamlit as st
from PIL import Image
from sqlalchemy import (
    Column, Integer, String, Float, LargeBinary, MetaData, Table,
    create_engine, select, insert, update, delete, func, and_, or_, inspect, text,
)
from sqlalchemy.engine import URL
from sqlalchemy.exc import IntegrityError, OperationalError, SQLAlchemyError
from streamlit_js_eval import get_geolocation

# ------------------------------------------------------------------
# KONFIGURASI - SESUAIKAN BAGIAN INI
# ------------------------------------------------------------------
UANG_MAKAN = 15000

# Koordinat asli Gedung A dan B Bank Sulselbar
GEDUNG_COORD = {
    "A": (-5.150070542826675, 119.41568236911743),  # (latitude, longitude)
    "B": (-5.150163347858429, 119.41329137592547),  # (latitude, longitude)
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

# Jendela waktu absen (WITA). Di luar jam ini absen ditolak. Izin bisa kapan saja.
JAM_MASUK_BUKA, JAM_MASUK_TUTUP = dtime(7, 0), dtime(15, 0)
JAM_PULANG_BUKA, JAM_PULANG_TUTUP = dtime(16, 50), dtime(18, 0)

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
# SUPABASE STORAGE (foto bukti absensi)
# ------------------------------------------------------------------
def storage_konfig():
    """(url, key, bucket) kalau Supabase Storage sudah diatur di Secrets, selain itu None."""
    url = str(secret("supabase_url") or "").strip().rstrip("/")
    key = str(secret("supabase_service_key") or "").strip()
    bucket = str(secret("supabase_bucket") or "foto-absensi").strip()
    return (url, key, bucket) if url and key else None


def _hdr(key, tambahan=None):
    h = {"apikey": key}
    if key.startswith("eyJ"):            # kunci lama (JWT) juga dikirim sebagai Bearer
        h["Authorization"] = f"Bearer {key}"
    h.update(tambahan or {})
    return h


@st.cache_resource(show_spinner=False)
def _pastikan_bucket(url, key, bucket):
    """Buat bucket privat kalau belum ada. Raise kalau gagal supaya tidak ikut ter-cache."""
    r = requests.post(
        f"{url}/storage/v1/bucket", headers=_hdr(key, {"Content-Type": "application/json"}),
        data=json.dumps({"id": bucket, "name": bucket, "public": False}), timeout=15,
    )
    if r.status_code in (200, 201, 409) or "already exists" in r.text.lower() or "duplicate" in r.text.lower():
        return True
    raise RuntimeError(f"Bucket gagal disiapkan ({r.status_code}): {r.text[:200]}")


def simpan_foto_storage(path, data):
    """Unggah foto ke Supabase Storage. True kalau berhasil; False kalau belum diatur / gagal."""
    cfg = storage_konfig()
    if not cfg:
        return False
    url, key, bucket = cfg
    try:
        _pastikan_bucket(url, key, bucket)
        r = requests.post(
            f"{url}/storage/v1/object/{bucket}/{quote(path)}",
            headers=_hdr(key, {"Content-Type": "image/jpeg", "x-upsert": "true"}),
            data=data, timeout=25,
        )
        return r.status_code in (200, 201)
    except Exception:
        return False


@st.cache_data(ttl=600, max_entries=200, show_spinner=False)
def _unduh_storage(path):
    cfg = storage_konfig()
    if not cfg:
        return None
    url, key, bucket = cfg
    try:
        r = requests.get(f"{url}/storage/v1/object/authenticated/{bucket}/{quote(path)}", headers=_hdr(key), timeout=20)
        return r.content if r.status_code == 200 else None
    except Exception:
        return None


def ambil_foto(foto_db, foto_path):
    """Foto dari Storage (kalau ada path-nya), kalau tidak ada dari kolom foto di database."""
    if foto_path:
        data = _unduh_storage(foto_path)
        if data:
            return data
    return bytes(foto_db) if foto_db else None


def hapus_foto_storage(daftar_path):
    cfg = storage_konfig()
    daftar_path = [p for p in daftar_path if p]
    if not cfg or not daftar_path:
        return
    url, key, bucket = cfg
    for i in range(0, len(daftar_path), 100):
        try:
            requests.delete(
                f"{url}/storage/v1/object/{bucket}",
                headers=_hdr(key, {"Content-Type": "application/json"}),
                data=json.dumps({"prefixes": daftar_path[i:i + 100]}), timeout=20,
            )
        except Exception:
            pass
    _unduh_storage.clear()



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
    Column("no_rekening", String(40)),
)

absensi_t = Table(
    "absensi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("nama", String(100), nullable=False, index=True),
    Column("tanggal", String(10), nullable=False, index=True),  # YYYY-MM-DD
    Column("jam", String(8), nullable=False),                   # HH:MM:SS
    Column("jenis", String(10), nullable=False, default="masuk"),  # masuk / pulang / izin
    Column("lat", Float),
    Column("lon", Float),
    Column("akurasi", Float),
    Column("jarak", Float),
    Column("status", String(20), nullable=False),               # berhasil / gagal_lokasi / izin
    Column("uang_makan", Integer, nullable=False, default=0),
    Column("foto", LargeBinary),                                # foto di DB (cadangan / data lama)
    Column("foto_path", String(200)),                           # lokasi foto di Supabase Storage
    Column("alasan", String(300)),                              # alasan izin
)

# Tabel kecil untuk Supabase Realtime: setiap absen baru menambah 1 baris,
# dashboard admin berlangganan perubahannya. Tidak berisi data sensitif.
notifikasi_t = Table(
    "notifikasi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("jenis", String(20)),
    Column("dibuat", String(19)),                               # YYYY-MM-DD HH:MM:SS (WITA)
)


def _bersihkan_url(u):
    """Buang sisa format tautan (mis. [teks](mailto:...)) yang ikut tersalin dari chat/web."""
    u = str(u or "").strip().strip("<>\"' ")
    return re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", u)


def url_database():
    """URL database dari Secrets: `database_url`, atau dirakit dari db_host/db_user/db_password/db_port/db_name."""
    url = _bersihkan_url(secret("database_url"))
    if url:
        url = url.replace("postgres://", "postgresql://", 1)
        # Paksa driver psycopg2 secara eksplisit, supaya SQLAlchemy tidak mencoba driver lain.
        if url.startswith("postgresql://"):
            url = url.replace("postgresql://", "postgresql+psycopg2://", 1)
        return url
    host = str(secret("db_host") or "").strip()
    if host:
        return URL.create(
            "postgresql+psycopg2",
            username=str(secret("db_user") or "postgres").strip(),
            password=str(secret("db_password") or ""),       # simbol apa pun aman, tidak perlu di-encode
            host=host,
            port=int(str(secret("db_port") or "6543").strip()),
            database=str(secret("db_name") or "postgres").strip(),
        )
    return None


def _sqlite_bisa_tulis(path):
    """True kalau file SQLite (dan foldernya) benar-benar bisa ditulis."""
    folder = os.path.dirname(os.path.abspath(path))
    if os.path.exists(path) and not os.access(path, os.W_OK):
        return False
    if not os.access(folder, os.W_OK):
        return False
    try:
        con = sqlite3.connect(path, timeout=5)
        v = con.execute("PRAGMA user_version").fetchone()[0]
        con.execute(f"PRAGMA user_version={int(v)}")  # tes tulis
        con.commit()
        con.close()
        return True
    except sqlite3.Error:
        return False


def _migrasi(eng):
    """Tambah kolom baru ke database lama (create_all tidak mengubah tabel yang sudah ada)."""
    perlu = {"peserta": [("no_rekening", "VARCHAR(40)")], "absensi": [("foto_path", "VARCHAR(200)"), ("alasan", "VARCHAR(300)")]}
    insp = inspect(eng)
    for tabel, daftar in perlu.items():
        ada = {k["name"] for k in insp.get_columns(tabel)}
        for kolom, tipe in daftar:
            if kolom not in ada:
                with eng.begin() as c:
                    c.execute(text(f"ALTER TABLE {tabel} ADD COLUMN {kolom} {tipe}"))


@st.cache_resource
def get_engine(versi_skema=6):
    url = url_database()
    sementara = False
    if url:
        eng = create_engine(url, pool_pre_ping=True, pool_recycle=300)
    else:
        # nama file baru (v2) supaya tidak bentrok dengan database versi lama
        path = "absensi_pkl_v2.db"
        if not _sqlite_bisa_tulis(path):
            # Folder aplikasi di hosting (mis. Streamlit Cloud) bisa bersifat read-only.
            # Salin ke folder sementara supaya tetap bisa menyimpan data. DATA SEMENTARA!
            tmp = os.path.join(tempfile.gettempdir(), "absensi_pkl_v2.db")
            if not os.path.exists(tmp) and os.path.exists(path):
                shutil.copyfile(path, tmp)
            path = tmp
            sementara = True
        eng = create_engine(
            f"sqlite:///{path}",
            connect_args={"check_same_thread": False, "timeout": 15},
        )
    metadata.create_all(eng)
    _migrasi(eng)
    eng.db_sementara = sementara
    return eng


def coba_ulang(fn):
    """Fungsi tulis ke DB: kalau koneksi terputus (OperationalError), reset koneksi lalu coba sekali lagi."""
    @wraps(fn)
    def bungkus(eng, *a, **k):
        try:
            return fn(eng, *a, **k)
        except OperationalError:
            eng.dispose()
            return fn(eng, *a, **k)
    return bungkus


def pesan_db(e):
    """Pesan error DB yang bisa dibaca (khusus admin; error asli Streamlit disembunyikan)."""
    asal = getattr(e, "orig", e)
    teks = f"{type(asal).__name__}: {asal}"
    if "readonly" in teks.lower():
        teks += " -> penyimpanan aplikasi bersifat read-only. Pakai database cloud (isi database_url di Secrets)."
    return f"Data gagal disimpan ke database. Detail: {teks}"


def hash_pin(nama, pin):
    garam = secret("salt", "pkl-sulselbar")
    return hashlib.sha256(f"{nama.strip().lower()}|{pin}|{garam}".encode()).hexdigest()


def list_peserta(eng):
    with eng.connect() as c:
        return pd.read_sql(
            select(
                peserta_t.c.nama, peserta_t.c.gedung, peserta_t.c.divisi, peserta_t.c.no_rekening
            ).order_by(peserta_t.c.nama), c
        )


def get_peserta(eng, nama):
    with eng.connect() as c:
        return c.execute(select(peserta_t).where(peserta_t.c.nama == nama)).mappings().first()


@coba_ulang
def add_peserta(eng, nama, gedung, divisi, pin, no_rekening=None):
    with eng.begin() as c:
        c.execute(insert(peserta_t).values(
            nama=nama.strip(), gedung=gedung, divisi=divisi.strip(), pin_hash=hash_pin(nama, pin),
            no_rekening=(no_rekening or None),
        ))


@coba_ulang
def set_rekening(eng, nama, no_rekening):
    with eng.begin() as c:
        c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(no_rekening=(no_rekening or None)))


def validasi_rekening(teks):
    """Kembalikan (nomor_rapi, pesan_error). Boleh angka dan tanda hubung, mis. 130-210-000554616-9."""
    rapi = re.sub(r"\s+", "", teks or "")
    if not rapi:
        return "", "Nomor rekening wajib diisi."
    if not re.fullmatch(r"[0-9\-]+", rapi) or not (8 <= len(re.sub(r"\D", "", rapi)) <= 25):
        return rapi, "Nomor rekening hanya boleh angka dan tanda hubung (-), minimal 8 angka."
    return rapi, ""


@coba_ulang
def reset_pin(eng, nama, pin_baru):
    with eng.begin() as c:
        c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(pin_hash=hash_pin(nama, pin_baru)))


@coba_ulang
def set_divisi(eng, nama, divisi, gedung):
    with eng.begin() as c:
        c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(divisi=divisi.strip(), gedung=gedung))


@coba_ulang
def hapus_peserta(eng, nama, dengan_absensi=False):
    """Hapus peserta. Kalau dengan_absensi=True, riwayat absensinya ikut dihapus."""
    paths = []
    with eng.begin() as c:
        c.execute(delete(peserta_t).where(peserta_t.c.nama == nama))
        if dengan_absensi:
            paths = [r[0] for r in c.execute(select(absensi_t.c.foto_path).where(absensi_t.c.nama == nama))]
            c.execute(delete(absensi_t).where(absensi_t.c.nama == nama))
    hapus_foto_storage(paths)


def daftar_nama_absensi(eng):
    """Semua nama yang punya riwayat absensi (termasuk peserta yang sudah dihapus)."""
    q = select(absensi_t.c.nama).distinct().order_by(absensi_t.c.nama)
    with eng.connect() as c:
        return [r[0] for r in c.execute(q).all()]


@coba_ulang
def hapus_absensi(eng, nama=None):
    """Hapus riwayat absensi (beserta fotonya). nama=None berarti SEMUA peserta.
    Mengembalikan jumlah baris yang terhapus."""
    stmt = delete(absensi_t)
    if nama is not None:
        stmt = stmt.where(absensi_t.c.nama == nama)
    sel = select(absensi_t.c.foto_path)
    if nama is not None:
        sel = sel.where(absensi_t.c.nama == nama)
    with eng.begin() as c:
        paths = [r[0] for r in c.execute(sel)]
        jumlah = c.execute(stmt).rowcount
    hapus_foto_storage(paths)
    return jumlah


def dalam_jendela(now, buka, tutup):
    """True kalau jam (setingkat menit) ada di antara buka dan tutup, keduanya termasuk."""
    t = now.time().replace(second=0, microsecond=0)
    return buka <= t <= tutup


def sudah_izin(eng, nama, tanggal):
    q = select(func.count()).select_from(absensi_t).where(and_(
        absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal, absensi_t.c.jenis == "izin",
    ))
    with eng.connect() as c:
        return c.execute(q).scalar() > 0


def sudah_berhasil(eng, nama, tanggal, jenis):
    q = select(func.count()).select_from(absensi_t).where(and_(
        absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal,
        absensi_t.c.jenis == jenis, absensi_t.c.status == "berhasil",
    ))
    with eng.connect() as c:
        return c.execute(q).scalar() > 0


@coba_ulang
def simpan_absen(eng, **nilai):
    sekarang = datetime.now(WITA)
    with eng.begin() as c:
        c.execute(insert(absensi_t).values(**nilai))
        # pemberitahuan untuk Supabase Realtime (dashboard admin), + bersihkan yang sudah > 7 hari
        c.execute(insert(notifikasi_t).values(jenis="absen", dibuat=sekarang.strftime("%Y-%m-%d %H:%M:%S")))
        c.execute(delete(notifikasi_t).where(
            notifikasi_t.c.dibuat < (sekarang - timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
        ))


def get_absensi_bulan(eng, tahun, bulan):
    prefix = f"{tahun:04d}-{bulan:02d}-"
    q = (
        select(
            absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jam, absensi_t.c.jenis,
            absensi_t.c.status, absensi_t.c.jarak, absensi_t.c.uang_makan, absensi_t.c.alasan,
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
            absensi_t.c.jarak, absensi_t.c.foto, absensi_t.c.foto_path,
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


# ------------------------------------------------------------------
# EXCEL FORMAT TEMPLATE (satu sheet per peserta)
# ------------------------------------------------------------------
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template_absen.xlsx")
# Penandatangan per divisi. Kalau divisi tidak ada di sini, dipakai tulisan bawaan di template_absen.xlsx.
# Contoh: {"Divisi Kepatuhan": ("Nama Pimpinan", "Pimdept. Kepatuhan")}
PENANDATANGAN = {}
BULAN_TTD = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
             "Agustus", "September", "Oktober", "November", "Desember"]


def _nama_sheet(nama, dipakai):
    """Nama sheet Excel: maks 31 karakter, tanpa karakter terlarang, unik."""
    dasar = re.sub(r"[\[\]:*?/\\]", " ", nama).strip()[:31] or "Peserta"
    hasil, i = dasar, 2
    while hasil.lower() in dipakai:
        akhiran = f" ({i})"
        hasil = dasar[: 31 - len(akhiran)] + akhiran
        i += 1
    dipakai.add(hasil.lower())
    return hasil


def tanggal_pertama_absen(eng, nama):
    with eng.connect() as c:
        return c.execute(select(func.min(absensi_t.c.tanggal)).where(absensi_t.c.nama == nama)).scalar()


def _salin_gaya(src, dst, nilai=None):
    dst._style = copy(src._style)
    if nilai is not None:
        dst.value = nilai


def _isi_sheet_peserta(ws, tpl, peserta, hari, pulang, izin, awal, akhir):
    """Isi satu sheet memakai gaya dari sheet 'Template'.
    hari = {tanggal_iso: (jam_masuk, uang)}, pulang = {tanggal_iso: jam_pulang}, izin = {tanggal_iso: alasan}."""
    isi = {
        "{nama}": peserta["nama"],
        "{unit_kerja}": (peserta["divisi"] or "-").upper(),
        "{no_rek}": peserta["no_rekening"] or "-",
        "{divisi}": peserta["divisi"] or "-",
        "{tanggal_ttd}": f"{akhir.day} {BULAN_TTD[akhir.month - 1]} {akhir.year}",
    }

    def ganti(v):
        if isinstance(v, str):
            for k, x in isi.items():
                v = v.replace(k, x)
        return v

    for k, dim in tpl.column_dimensions.items():
        nd = ws.column_dimensions[k]
        nd.width, nd.min, nd.max = dim.width, dim.min, dim.max

    def baris(sr, dr, nilai_ganti=True):
        for col in "BCDEF":
            sc = tpl[f"{col}{sr}"]
            _salin_gaya(sc, ws[f"{col}{dr}"], ganti(sc.value) if nilai_ganti else None)
        h = tpl.row_dimensions[sr].height
        if h:
            ws.row_dimensions[dr].height = h

    for r in range(3, 8):          # identitas + judul kolom
        baris(r, r)

    r = 8
    tgl = awal
    while tgl <= akhir:
        iso = tgl.isoformat()
        if iso in hari:                                   # hadir (masuk berhasil)
            jam, uang = hari[iso]
            baris(8, r, nilai_ganti=False)
            ws[f"B{r}"].value = datetime(tgl.year, tgl.month, tgl.day)
            ws[f"C{r}"].value = jam
            ws[f"D{r}"].value = pulang.get(iso, "-")      # "-" kalau belum/tidak absen pulang
            ws[f"E{r}"].value = "Hadir"
            ws[f"F{r}"].value = uang
        else:
            baris(9, r, nilai_ganti=False)
            ws[f"B{r}"].value = datetime(tgl.year, tgl.month, tgl.day)
            if iso in izin:                               # izin + alasan
                teks = f"Izin: {izin[iso]}"
                ws[f"C{r}"].value = teks
                ws[f"C{r}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[r].height = 15 * max(1, math.ceil(len(teks) / 42))
                ws[f"B{r}"].alignment = Alignment(horizontal="center", vertical="center")
            else:
                ws[f"C{r}"].value = "Libur " if tgl.weekday() >= 5 else "Tidak hadir"
            ws.merge_cells(f"C{r}:F{r}")
        r += 1
        tgl += timedelta(days=1)

    jumlah = r
    baris(10, jumlah)
    ws.merge_cells(f"B{jumlah}:E{jumlah}")
    ws[f"F{jumlah}"].value = f"=SUM(F8:F{jumlah - 1})" if jumlah > 8 else 0

    geser = jumlah - 10             # footer (tanda tangan) ikut turun
    penanda = PENANDATANGAN.get(peserta["divisi"])
    for sr in range(11, tpl.max_row + 1):
        baris(sr, sr + geser)
    for rng in tpl.merged_cells.ranges:
        if rng.min_row >= 11:
            ws.merge_cells(start_row=rng.min_row + geser, end_row=rng.max_row + geser,
                           start_column=rng.min_col, end_column=rng.max_col)
    if penanda:
        for sr in range(11, tpl.max_row + 1):
            if tpl[f"E{sr}"].value == "Faisal Hafiz ":
                ws[f"E{sr + geser}"].value = penanda[0]
                ws[f"E{sr + geser + 1}"].value = penanda[1]

    ws.page_margins = copy(tpl.page_margins)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def excel_template_bytes(eng, df_bulan, tahun, bulan, daftar_nama):
    """Excel format template: satu sheet per nama. Mengembalikan bytes (None kalau tidak ada data)."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    tpl = wb["Template"]
    awal_bulan = date(tahun, bulan, 1)
    akhir_bulan = (date(tahun + (bulan == 12), bulan % 12 + 1, 1) - timedelta(days=1))
    hari_ini = datetime.now(WITA).date()
    akhir = min(akhir_bulan, hari_ini) if (tahun, bulan) == (hari_ini.year, hari_ini.month) else akhir_bulan

    dipakai, jumlah_sheet = set(), 0
    for nama in daftar_nama:
        data = df_bulan[df_bulan["nama"] == nama]
        if data.empty:
            continue
        info = get_peserta(eng, nama)
        peserta = {
            "nama": nama,
            "divisi": (info["divisi"] if info else "") or "",
            "no_rekening": info["no_rekening"] if info else None,
        }
        hari, pulang, izin = {}, {}, {}
        for _, row in data.iterrows():
            iso = row["tanggal"]
            if row["jenis"] == "izin":
                izin.setdefault(iso, str(row["alasan"] or "-").strip())
            elif row["status"] == "berhasil":
                jam = datetime.strptime(row["jam"], "%H:%M:%S").time().replace(second=0)
                if row["jenis"] == "masuk":
                    hari.setdefault(iso, (jam, int(row["uang_makan"])))
                elif row["jenis"] == "pulang":
                    pulang[iso] = jam                      # kalau ada beberapa, ambil yang terakhir
        pertama = tanggal_pertama_absen(eng, nama) or data["tanggal"].min()
        awal = max(awal_bulan, date.fromisoformat(pertama))
        if awal > akhir:
            continue
        ws = wb.create_sheet(_nama_sheet(nama, dipakai))
        _isi_sheet_peserta(ws, tpl, peserta, hari, pulang, izin, awal, akhir)
        jumlah_sheet += 1

    if jumlah_sheet == 0:
        return None
    del wb["Template"]
    buf = io.BytesIO()
    wb.save(buf)
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
  color: #fff; padding: 30px 22px 76px; border-radius: 0 0 0 0;
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
.pr-card { background: #fff; border-radius: 22px; margin: -46px 22px 0; padding: 14px 14px 14px;
  box-shadow: 0 8px 28px rgba(37,52,140,.10); position: relative; z-index: 2; }
.pr-date { display: flex; align-items: center; gap: 9px; color: #2f3b6b; font-size: 14px; font-weight: 500; }
.pr-date svg { color: #4a5788; }
.pr-badge { margin-left: auto; background: #d9f7e8; color: #12a15f; font-weight: 600; font-size: 15px;
  padding: 4px 12px; border-radius: 999px; }
.pr-tiles { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-top: 12px; }
.pr-tile { border-radius: 14px; text-align: center; padding: 10px 4px 9px; }
.pr-tile .lb { font-size: 12px; font-weight: 500; margin-top: 3px; }
.pr-tile .vl { font-size: 19px; font-weight: 700; margin-top: 3px; line-height: 1.1; }
.pr-tile .vl.sm { font-size: 14px; margin-top: 5px; }
.t-green { background: #e6f9f0; color: #14a068; }
.t-yellow { background: #fff6dc; color: #c58a05; }
.t-red { background: #fdeaea; color: #e04a4a; }
.t-blue { background: #e8f0ff; color: #2f5fd0; }
.pr-info { margin-top: 10px; background: #eaf1ff; border-radius: 12px; padding: 9px 12px; color: #2f5fd0; }
.pr-info .h { display: flex; align-items: center; gap: 7px; font-size: 13px; font-weight: 600; }
.pr-info .s { font-size: 12.5px; margin-top: 2px; color: #3b6fdc; line-height: 1.3; }

/* kolom tidak boleh turun ke bawah di layar sempit */
[data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; }
[data-testid="stHorizontalBlock"] > [data-testid="stColumn"] { min-width: 0 !important; flex: 1 1 0 !important; width: auto !important; }

/* penyangga di akhir halaman: memberi ruang scroll di atas navigasi bawah */
.pr-spacer { height: 150px; }

/* ---------- tombol transparan di atas tampilan custom ---------- */
[class*="st-key-ov_"] { position: relative; }
[class*="st-key-ov_"] [data-testid="stElementContainer"]:has(button) { position: absolute; inset: 0; z-index: 5; }
[class*="st-key-ov_"] [data-testid="stButton"], [class*="st-key-ov_"] .stButton { height: 100%; width: 100% !important; }
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
.pr-row .d { font-weight: 600; color: #1e2a5a; font-size: 15px; }
.pr-row .m { font-size: 13px; color: #6b779e; margin-top: 2px; }
.pr-row .r { margin-left: auto; text-align: right; font-size: 13px; font-weight: 600; }
.ok { color: #14a068; } .bad { color: #e04a4a; } .izin { color: #c58a05; }
.pr-kv { display: flex; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid #eef0f5; font-size: 15px; }
.pr-kv:last-child { border-bottom: 0; }
.pr-kv span:first-child { color: #6b779e; } .pr-kv span:last-child { font-weight: 600; color: #1e2a5a; }
.pr-stat { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 8px; }
.pr-stat div { background: #eaf1ff; border-radius: 16px; padding: 14px; color: #2f5fd0; }
.pr-stat b { display: block; font-size: 22px; margin-top: 4px; }
.pr-stat span { font-size: 13px; }
.st-key-pagebody [data-testid="stButton"] button[kind="primary"] { background: #2563eb; border: 0; border-radius: 14px; }

</style>
"""



CSS_TEMA = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root {
  --biru: #2563eb; --ungu: #3b3fd8; --bg: #eef2fb; --tint: #eaf1ff;
  --navy: #1e2a5a; --navy2: #4a5788; --muted: #6b779e; --garis: #d5def3;
}
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"] {
  background: var(--bg) !important; color: var(--navy); font-family: 'Inter', system-ui, sans-serif;
  color-scheme: light;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"],
.stDeployButton, footer, [class*="viewerBadge"], [class*="StatusWidget"] { display: none !important; }
[data-testid="stMainBlockContainer"], .block-container {
  padding: 0 0 48px 0 !important; background: var(--bg); position: relative; min-height: 100vh;
}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] { gap: 0 !important; }

/* ---------- komponen form (login + admin) ---------- */
.st-key-loginbody, .st-key-adminbody { color: var(--navy); }
.st-key-loginbody :is(p, label, span, small, li, h1, h2, h3, h4),
.st-key-adminbody :is(p, label, span, small, li, h1, h2, h3, h4) { color: var(--navy); }
.st-key-loginbody [data-testid="stCaptionContainer"] *, .st-key-adminbody [data-testid="stCaptionContainer"] * { color: var(--muted); }
.st-key-loginbody h2, .st-key-adminbody h2 { font-size: 20px; font-weight: 700; padding: 4px 0 8px; }
.st-key-adminbody h3 { font-size: 17px; font-weight: 700; }
.st-key-loginbody hr, .st-key-adminbody hr { border: 0; border-top: 1px solid var(--garis); margin: 14px 0; }

.st-key-adminbody textarea,
.st-key-loginbody [data-baseweb="select"] > div, .st-key-adminbody [data-baseweb="select"] > div,
.st-key-adminbody [data-baseweb="base-input"], .st-key-loginbody [data-baseweb="base-input"] {
  background: #f6f8ff !important; color: var(--navy) !important; border-color: var(--garis) !important;
  border-radius: 14px !important;
}
.st-key-loginbody input, .st-key-adminbody input { background: transparent !important; color: var(--navy) !important; }
.st-key-loginbody input::placeholder, .st-key-adminbody input::placeholder { color: var(--muted) !important; opacity: 1; }
.st-key-loginbody [data-baseweb="select"] > div:focus-within, .st-key-adminbody [data-baseweb="select"] > div:focus-within,
.st-key-loginbody [data-baseweb="base-input"]:focus-within, .st-key-adminbody [data-baseweb="base-input"]:focus-within {
  border-color: var(--biru) !important; box-shadow: 0 0 0 3px rgba(37,99,235,.18) !important;
}

/* radio -> tombol pil */
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadio"] [role="radiogroup"] {
  background: var(--tint); border-radius: 16px; padding: 5px; gap: 4px; flex-wrap: nowrap; width: 100%;
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] {
  flex: 1 1 0; justify-content: center; border-radius: 12px; padding: 9px 8px; margin: 0 !important;
  cursor: pointer; transition: background .15s ease; min-width: 0;
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] > div { width: 100%; justify-content: center; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] > div > div:first-child { display: none; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] p { font-weight: 600; font-size: 15px; color: var(--biru); text-align: center; white-space: nowrap; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"][data-selected="true"] {
  background: linear-gradient(135deg, var(--biru), var(--ungu)); box-shadow: 0 4px 12px rgba(37,63,216,.3);
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"][data-selected="true"] p { color: #fff; }

/* tombol */
:is(.st-key-loginbody, .st-key-adminbody) button[kind="primary"],
:is(.st-key-loginbody, .st-key-adminbody) button[kind="primaryFormSubmit"] {
  background: linear-gradient(135deg, var(--biru), var(--ungu)) !important; border: 0 !important; color: #fff !important;
  border-radius: 14px !important; min-height: 48px; font-weight: 600; box-shadow: 0 6px 16px rgba(37,63,216,.3);
}
:is(.st-key-loginbody, .st-key-adminbody) button[kind="primary"] *, :is(.st-key-loginbody, .st-key-adminbody) button[kind="primaryFormSubmit"] * { color: #fff !important; }
:is(.st-key-loginbody, .st-key-adminbody) button[kind="secondary"],
:is(.st-key-loginbody, .st-key-adminbody) button[kind="secondaryFormSubmit"],
.st-key-adminbody [data-testid="stDownloadButton"] button {
  background: #fff !important; color: var(--biru) !important; border: 1.5px solid var(--biru) !important;
  border-radius: 14px !important; font-weight: 600; min-height: 44px;
}
:is(.st-key-loginbody, .st-key-adminbody) button[kind="secondary"] *, :is(.st-key-loginbody, .st-key-adminbody) button[kind="secondaryFormSubmit"] *, .st-key-adminbody [data-testid="stDownloadButton"] button * { color: var(--biru) !important; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadio"], :is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadio"] > div { width: 100%; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadio"] [role="radiogroup"] { display: flex; }

/* kartu statistik, tabel, expander, alert */
.st-key-adminbody [data-testid="stMetric"] { background: var(--tint); border-radius: 18px; padding: 14px 16px; }
.st-key-adminbody [data-testid="stMetricLabel"] * { color: #2f5fd0 !important; }
.st-key-adminbody [data-testid="stMetricValue"] * { color: var(--navy) !important; font-weight: 700; font-family: 'Inter', system-ui, sans-serif !important; }
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  border: 1px solid var(--garis); border-radius: 16px; overflow: hidden;
}
.st-key-adminbody [data-testid="stExpander"] details { background: #f6f8ff; border: 1px solid var(--garis); border-radius: 16px; }
.st-key-adminbody [data-testid="stExpander"] summary * { color: var(--navy); font-weight: 600; }
.st-key-adminbody [data-testid="stAlert"], .st-key-loginbody [data-testid="stAlert"] { border-radius: 14px; }
.st-key-adminbody [data-testid="stImage"] img { border-radius: 14px; }
.st-key-adminbody [data-testid="stForm"] {
  border: 1px solid var(--garis); border-radius: 18px; background: #fbfcff; padding: 16px;
}

/* ---------- komponen form di halaman peserta (radio pil, input, textarea) ---------- */
.st-key-pagebody [data-testid="stRadio"], .st-key-pagebody [data-testid="stRadio"] > div { width: 100%; }
.st-key-pagebody [data-testid="stRadio"] { margin-bottom: 8px; }
.st-key-pagebody [data-testid="stRadio"] [role="radiogroup"] {
  display: flex; background: var(--tint); border-radius: 16px; padding: 5px; gap: 4px; flex-wrap: nowrap; width: 100%;
}
.st-key-pagebody [data-testid="stRadioOption"] {
  flex: 1 1 0; justify-content: center; border-radius: 12px; padding: 9px 8px; margin: 0 !important;
  cursor: pointer; min-width: 0;
}
.st-key-pagebody [data-testid="stRadioOption"] > div { width: 100%; justify-content: center; }
.st-key-pagebody [data-testid="stRadioOption"] > div > div:first-child { display: none; }
.st-key-pagebody [data-testid="stRadioOption"] p { font-weight: 600; font-size: 15px; color: var(--biru); text-align: center; white-space: nowrap; }
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] {
  background: linear-gradient(135deg, var(--biru), var(--ungu)); box-shadow: 0 4px 12px rgba(37,63,216,.3);
}
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] p { color: #fff; }
.st-key-pagebody [data-baseweb="textarea"], .st-key-pagebody [data-baseweb="base-input"] {
  background: #f6f8ff !important; border-color: var(--garis) !important; border-radius: 14px !important;
}
.st-key-pagebody textarea, .st-key-pagebody input { background: transparent !important; color: var(--navy) !important; }
.st-key-pagebody textarea::placeholder { color: var(--muted) !important; opacity: 1; }
</style>
"""

CSS_AUTH_ADMIN = """
<style>
[data-testid="stMainBlockContainer"], .block-container { max-width: 460px !important; }
.auth-head {
  background: linear-gradient(135deg, #2563eb 0%, #3b3fd8 100%); color: #fff;
  padding: 36px 24px 92px; position: relative; overflow: hidden;
}
.auth-head::after {
  content: ""; position: absolute; left: -10%; right: -10%; bottom: -34px; height: 70px;
  background: #eef2fb; border-radius: 50% 50% 0 0 / 100% 100% 0 0;
}
.auth-brand { display: flex; align-items: center; gap: 14px; }
.auth-logo { width: 52px; height: 52px; border-radius: 16px; background: rgba(255,255,255,.2);
  display: flex; align-items: center; justify-content: center; color: #fff; }
.auth-title { font-size: 27px; font-weight: 700; letter-spacing: .2px; color: #fff; }
.auth-bank { font-size: 13px; color: rgba(255,255,255,.8); margin-top: 2px; }
.auth-greet { margin-top: 30px; font-size: 22px; font-weight: 600; color: #fff; }
.auth-desc { margin-top: 5px; font-size: 14px; color: rgba(255,255,255,.82); }
.st-key-loginbody {
  width: calc(100% - 44px) !important; margin: -50px 22px 0; background: #fff; border-radius: 26px;
  padding: 24px 18px 22px; position: relative; z-index: 2; box-shadow: 0 8px 28px rgba(37,52,140,.12);
}
.auth-foot { text-align: center; color: #6b779e; font-size: 13px; padding: 20px 0 0; }
</style>
"""

CSS_ADMIN = """
<style>
[data-testid="stMainBlockContainer"], .block-container { max-width: 1100px !important; }
.admin-head {
  background: linear-gradient(135deg, #2563eb 0%, #3b3fd8 100%); color: #fff;
  padding: 32px 28px 84px; position: relative; overflow: hidden;
}
.admin-head::after {
  content: ""; position: absolute; left: -10%; right: -10%; bottom: -34px; height: 70px;
  background: #eef2fb; border-radius: 50% 50% 0 0 / 100% 100% 0 0;
}
.admin-top { display: flex; align-items: center; gap: 14px; }
.admin-logo { width: 50px; height: 50px; border-radius: 15px; background: rgba(255,255,255,.2);
  display: flex; align-items: center; justify-content: center; color: #fff; }
.admin-title { font-size: 25px; font-weight: 700; color: #fff; flex: 1; }
.admin-sub { font-size: 13px; color: rgba(255,255,255,.82); margin-top: 2px; font-weight: 400; }
.admin-out { width: 50px; height: 50px; border-radius: 50%; background: rgba(255,255,255,.2);
  display: flex; align-items: center; justify-content: center; color: #fff; }
.st-key-adminlogout { position: absolute !important; top: 32px; right: 28px; width: 50px; height: 50px; z-index: 10; }
.st-key-adminlogout [data-testid="stElementContainer"], .st-key-adminlogout [data-testid="stButton"] { height: 50px; width: 50px !important; }
.st-key-adminlogout button { width: 50px !important; height: 50px !important; opacity: 0; cursor: pointer; padding: 0; }
.st-box { background: #f6f8ff; border: 1px solid #d5def3; border-radius: 14px; padding: 10px 14px; font-size: 13px; line-height: 1.7; color: #1e2a5a; margin-bottom: 8px; }
.st-key-rt_hidden { position: absolute; left: -9999px; width: 1px; height: 1px; overflow: hidden; }
.st-key-adminbody {
  width: calc(100% - 36px) !important; margin: -44px 18px 0; background: #fff; border-radius: 26px;
  padding: 22px 22px 26px; position: relative; z-index: 2; box-shadow: 0 8px 28px rgba(37,52,140,.12);
}
@media (max-width: 600px) {
  .admin-head { padding: 28px 18px 80px; }
  .st-key-adminlogout { right: 18px; top: 28px; }
  .st-key-adminbody { padding: 16px 12px 22px; width: calc(100% - 20px) !important; margin-left: 10px; }
}
</style>
"""


def inisial(nama):
    bagian = [b for b in nama.split() if b]
    return "".join(b[0] for b in bagian[:2]).upper() or "?"


def get_status_hari(eng, nama, tanggal):
    """Ringkasan absen satu hari: baris masuk/pulang yang berhasil, izin, dan apakah ada percobaan gagal."""
    q = (
        select(absensi_t.c.jenis, absensi_t.c.jam, absensi_t.c.status, absensi_t.c.uang_makan,
               absensi_t.c.jarak, absensi_t.c.alasan)
        .where(and_(absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal))
        .order_by(absensi_t.c.jam)
    )
    with eng.connect() as c:
        rows = c.execute(q).mappings().all()
    hasil = {"masuk": None, "pulang": None, "izin": None, "gagal_masuk": False, "gagal_pulang": False}
    for r in rows:
        if r["jenis"] == "izin":
            hasil["izin"] = r
        elif r["jenis"] in ("masuk", "pulang"):
            if r["status"] == "berhasil":
                if hasil[r["jenis"]] is None:
                    hasil[r["jenis"]] = r
            else:
                hasil["gagal_" + r["jenis"]] = True
    return hasil


def get_riwayat_peserta(eng, nama, limit=60, tanpa_gagal=False):
    """Catatan absen milik peserta. tanpa_gagal=True: sembunyikan percobaan gagal (hanya berhasil + izin)."""
    q = (
        select(absensi_t.c.tanggal, absensi_t.c.jam, absensi_t.c.jenis, absensi_t.c.status,
               absensi_t.c.jarak, absensi_t.c.uang_makan, absensi_t.c.alasan)
        .where(absensi_t.c.nama == nama)
        .order_by(absensi_t.c.tanggal.desc(), absensi_t.c.jam.desc())
        .limit(limit)
    )
    if tanpa_gagal:
        q = q.where(or_(absensi_t.c.status == "berhasil", absensi_t.c.jenis == "izin"))
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
                if tab == "absen":
                    st.session_state.pop("mode_absen", None)   # pilihan awal ikut kondisi hari ini
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


def teks_status_hari(now, sh):
    """(judul, keterangan) untuk kotak status di Beranda, sesuai kondisi & jam sekarang."""
    f = lambda t: t.strftime("%H:%M")
    jendela_masuk = f"{f(JAM_MASUK_BUKA)}-{f(JAM_MASUK_TUTUP)}"
    jendela_pulang = f"{f(JAM_PULANG_BUKA)}-{f(JAM_PULANG_TUTUP)}"
    t = now.time().replace(second=0, microsecond=0)
    if sh["masuk"] is None and sh["izin"] is not None:
        return "Izin hari ini", esc(str(sh["izin"]["alasan"] or "-"))
    if sh["masuk"] is None:
        if t < JAM_MASUK_BUKA:
            return "Status presensi", f"Absen masuk dibuka pukul {f(JAM_MASUK_BUKA)} WITA"
        if t <= JAM_MASUK_TUTUP:
            ekstra = " Percobaan terakhir gagal (di luar area)." if sh["gagal_masuk"] else ""
            return "Status presensi", f"Absen masuk dibuka sampai {f(JAM_MASUK_TUTUP)} WITA.{ekstra}"
        return "Status presensi", f"Waktu absen masuk ({jendela_masuk}) sudah lewat"
    if sh["pulang"] is None:
        if t < JAM_PULANG_BUKA:
            return "Status presensi", f"Absen pulang dibuka pukul {f(JAM_PULANG_BUKA)} WITA"
        if t <= JAM_PULANG_TUTUP:
            ekstra = " Percobaan terakhir gagal (di luar area)." if sh["gagal_pulang"] else ""
            return "Status presensi", f"Absen pulang dibuka sampai {f(JAM_PULANG_TUTUP)} WITA.{ekstra}"
        return "Status presensi", f"Waktu absen pulang ({jendela_pulang}) sudah lewat"
    return "Status presensi", f"Presensi hari ini lengkap: masuk {sh['masuk']['jam'][:5]}, pulang {sh['pulang']['jam'][:5]}"


def tab_beranda(eng, nama, peserta):
    now = datetime.now(WITA)
    sh = get_status_hari(eng, nama, now.strftime("%Y-%m-%d"))
    jam_masuk = sh["masuk"]["jam"][:5] if sh["masuk"] else "--:--"
    jam_pulang = sh["pulang"]["jam"][:5] if sh["pulang"] else "--:--"
    uang = rupiah(sh["masuk"]["uang_makan"]) if sh["masuk"] else "-"
    info_judul, info_sub = teks_status_hari(now, sh)

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
<div class="pr-date">{ic("calendar", 20)}<span>{tgl_indonesia(now.date())}</span><span class="pr-badge">Aktif</span></div>
<div class="pr-tiles">
<div class="pr-tile t-green">{ic("login", 20)}<div class="lb">Masuk</div><div class="vl">{jam_masuk}</div></div>
<div class="pr-tile t-yellow">{ic("logout", 20)}<div class="lb">Pulang</div><div class="vl">{jam_pulang}</div></div>
<div class="pr-tile t-blue">{ic("file", 20)}<div class="lb">Uang Makan</div><div class="vl sm">{uang}</div></div>
</div>
<div class="pr-info"><div class="h">{ic("clock", 16)}<span>{info_judul}</span></div><div class="s">{info_sub}</div></div>
</div>""",
        unsafe_allow_html=True,
    )
    # tombol keluar transparan di atas ikon logout header
    with st.container(key="ov_keluar"):
        if st.button("‎", key="ovb_keluar", width="stretch"):
            keluar()


def _form_absen(eng, nama, peserta, jenis, now, sh):
    """Form absen masuk / pulang: hanya di jendela waktunya, wajib di area gedung + selfie."""
    gedung = peserta["gedung"]
    buka, tutup = (JAM_MASUK_BUKA, JAM_MASUK_TUTUP) if jenis == "masuk" else (JAM_PULANG_BUKA, JAM_PULANG_TUTUP)
    rentang = f"{buka.strftime('%H:%M')}-{tutup.strftime('%H:%M')}"
    st.caption(f"Jam absen {jenis}: {rentang} WITA")

    if jenis == "pulang" and sh["masuk"] is None:
        st.warning("Kamu belum absen masuk hari ini, jadi belum bisa absen pulang.")
        return
    if sh[jenis] is not None:
        st.success(f"Kamu sudah absen {jenis} hari ini pukul {sh[jenis]['jam'][:5]} WITA.")
        return
    if not dalam_jendela(now, buka, tutup):
        st.warning(f"Absen {jenis} hanya dibuka pukul {rentang} WITA. Sekarang pukul {now.strftime('%H:%M')}.")
        return

    lokasi = get_geolocation()
    coords = (lokasi or {}).get("coords")
    if coords:
        st.success(f"📡 Lokasi terdeteksi (akurasi ±{coords.get('accuracy', 0):.0f} m)")
    else:
        st.info("Menunggu lokasi... klik **Allow/Izinkan** saat browser meminta akses lokasi.")

    foto = st.camera_input("Foto selfie untuk absen", key=f"foto_{jenis}")

    if st.button(f"Absen {jenis.capitalize()} Sekarang", type="primary", width="stretch", key=f"btn_{jenis}"):
        if not coords:
            st.error("Lokasi belum terdeteksi. Aktifkan GPS/izin lokasi lalu coba lagi.")
            return
        if foto is None:
            st.error("Foto belum diambil.")
            return

        now = datetime.now(WITA)                      # cek ulang: jam bisa saja bergeser saat foto diambil
        tanggal, jam = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")
        if sudah_berhasil(eng, nama, tanggal, jenis):
            st.warning(f"Kamu sudah absen {jenis} hari ini.")
            return
        if jenis == "pulang" and not sudah_berhasil(eng, nama, tanggal, "masuk"):
            st.warning("Kamu belum absen masuk hari ini, jadi belum bisa absen pulang.")
            return
        if not dalam_jendela(now, buka, tutup):
            st.error(f"❌ Absen ditolak: {jenis} hanya dibuka pukul {rentang} WITA (sekarang {now.strftime('%H:%M')}).")
            return

        lat, lon = coords["latitude"], coords["longitude"]
        cocok, jarak = cek_lokasi(lat, lon, gedung)
        status = "berhasil" if cocok else "gagal_lokasi"
        uang = UANG_MAKAN if (cocok and jenis == "masuk") else 0     # uang makan dihitung dari absen masuk

        foto_kecil = kompres_foto(foto.getvalue())
        nama_berkas = re.sub(r"[^A-Za-z0-9]+", "_", nama).strip("_") or "peserta"
        path_foto = f"{tanggal}/{nama_berkas}_{jenis}_{jam.replace(':', '')}.jpg"
        di_storage = simpan_foto_storage(path_foto, foto_kecil)   # False kalau Storage belum diatur / gagal
        try:
            simpan_absen(
                eng, nama=nama, tanggal=tanggal, jam=jam, jenis=jenis,
                lat=lat, lon=lon, akurasi=coords.get("accuracy"), jarak=jarak,
                status=status, uang_makan=uang,
                foto=None if di_storage else foto_kecil,
                foto_path=path_foto if di_storage else None,
            )
        except SQLAlchemyError:
            st.error("Absen gagal disimpan karena gangguan database. Coba lagi sebentar lagi atau hubungi admin.")
            return

        if cocok:
            st.success(f"✅ Absen {jenis} berhasil pukul {jam} WITA.")
            if jenis == "masuk":
                st.info(f"💰 Uang makan hari ini: {rupiah(uang)}")
        else:
            st.error(
                f"❌ Absen GAGAL: kamu ~{jarak:.0f} m dari Gedung {gedung}, "
                "di luar area yang diizinkan. Pastikan kamu sudah di gedung lalu coba lagi."
            )


def _form_izin(eng, nama, now):
    """Form izin: bisa kapan saja, tanpa lokasi/foto, wajib mengisi alasan."""
    st.caption("Izin bisa diajukan kapan saja. Alasan akan masuk ke laporan Excel.")
    tgl = st.date_input(
        "Tanggal izin", value=now.date(), min_value=now.date(), max_value=now.date() + timedelta(days=30),
        key="izin_tanggal",
    )
    alasan = st.text_area(
        "Alasan izin", max_chars=300, height=110, key="izin_alasan",
        placeholder="Contoh: sakit, ada keperluan keluarga, mengikuti kegiatan kampus...",
    )
    if st.button("Kirim Izin", type="primary", width="stretch", key="btn_izin"):
        teks = " ".join((alasan or "").split())
        tanggal = tgl.strftime("%Y-%m-%d")
        if len(teks) < 5:
            st.error("Alasan izin wajib diisi (minimal 5 karakter).")
            return
        if sudah_izin(eng, nama, tanggal):
            st.warning(f"Kamu sudah mengajukan izin untuk {tgl_indonesia(tgl)}.")
            return
        if sudah_berhasil(eng, nama, tanggal, "masuk"):
            st.warning("Kamu sudah absen masuk pada tanggal itu, jadi tidak bisa mengajukan izin.")
            return
        try:
            simpan_absen(
                eng, nama=nama, tanggal=tanggal, jam=datetime.now(WITA).strftime("%H:%M:%S"),
                jenis="izin", status="izin", uang_makan=0, alasan=teks,
            )
        except SQLAlchemyError:
            st.error("Izin gagal disimpan karena gangguan database. Coba lagi sebentar lagi atau hubungi admin.")
            return
        st.success(f"✅ Izin untuk {tgl_indonesia(tgl)} tersimpan.")


def tab_absen(eng, nama, peserta):
    header_halaman("Absen", f"{nama} • Gedung {peserta['gedung']} • {peserta['divisi']}")
    now = datetime.now(WITA)
    sh = get_status_hari(eng, nama, now.strftime("%Y-%m-%d"))

    if "mode_absen" not in st.session_state:         # pilihan awal mengikuti kondisi hari ini
        st.session_state.mode_absen = "Pulang" if (sh["masuk"] is not None and sh["pulang"] is None) else "Masuk"

    with st.container(key="pagebody"):
        mode = st.radio("Jenis absen", ["Masuk", "Pulang", "Izin"], horizontal=True,
                        key="mode_absen", label_visibility="collapsed")
        if mode == "Izin":
            _form_izin(eng, nama, now)
        else:
            _form_absen(eng, nama, peserta, mode.lower(), now, sh)


def tab_riwayat(eng, nama):
    header_halaman("Riwayat Presensi", "60 catatan terakhir")
    rows = get_riwayat_peserta(eng, nama, tanpa_gagal=True)      # percobaan gagal tidak ditampilkan
    with st.container(key="pagebody"):
        if not rows:
            st.info("Belum ada riwayat absen.")
            return
        html = ""
        for r in rows:
            d = datetime.strptime(r["tanggal"], "%Y-%m-%d").date()
            if r["jenis"] == "izin":
                info = f"Izin • {esc(str(r['alasan'] or '-'))}"
                ket, kelas = "Izin", "izin"
            else:
                label = "Masuk" if r["jenis"] == "masuk" else "Pulang"
                info = f"{label} {r['jam'][:5]} WITA • {r['jarak'] or 0:.0f} m"
                ok = r["status"] == "berhasil"
                if not ok:
                    ket, kelas = "Gagal lokasi", "bad"
                else:
                    ket, kelas = (rupiah(r["uang_makan"]) if r["jenis"] == "masuk" else "Berhasil"), "ok"
            html += (
                f'<div class="pr-row"><div><div class="d">{tgl_indonesia(d)}</div>'
                f'<div class="m">{info}</div></div>'
                f'<div class="r {kelas}">{ket}</div></div>'
            )
        st.markdown(html, unsafe_allow_html=True)


def tab_rekap(eng, nama):
    now = datetime.now(WITA)
    header_halaman("Rekap Bulan Ini", f"{BULAN_ID[now.month - 1]} {now.year}")
    prefix = now.strftime("%Y-%m-")
    rows = [r for r in get_riwayat_peserta(eng, nama, limit=400) if r["tanggal"].startswith(prefix)]
    masuk_ok = [r for r in rows if r["jenis"] == "masuk" and r["status"] == "berhasil"]
    hadir = len({r["tanggal"] for r in masuk_ok})
    total = sum(r["uang_makan"] for r in masuk_ok)
    izin = len({r["tanggal"] for r in rows if r["jenis"] == "izin"})
    gagal = sum(1 for r in rows if r["jenis"] != "izin" and r["status"] != "berhasil")
    with st.container(key="pagebody"):
        st.markdown(
            f'<div class="pr-stat"><div><span>Hari hadir</span><b>{hadir}</b></div>'
            f'<div><span>Total uang makan</span><b>{rupiah(total)}</b></div></div>'
            f'<div class="pr-kv"><span>Izin</span><span>{izin}</span></div>'
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
    st.markdown('<div class="pr-spacer"></div>', unsafe_allow_html=True)
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
        no_rek = st.text_input("Nomor rekening", placeholder="contoh: 130-210-000554616-9", max_chars=30)
        pin = st.text_input("PIN (4-6 angka)", max_chars=6)
        kirim = st.form_submit_button("Tambah Peserta")

    if kirim:
        if (no_rek or "").strip():
            rek_rapi, rek_err = validasi_rekening(no_rek)        # kalau diisi, formatnya harus benar
        else:
            rek_rapi, rek_err = "", ""                           # kosong = boleh, bisa diisi nanti lewat Edit
        if not nama.strip():
            st.error("Nama wajib diisi.")
        elif rek_err:
            st.error(rek_err)
        elif not (pin.isdigit() and 4 <= len(pin) <= 6):
            st.error("PIN harus 4-6 angka.")
        else:
            try:
                add_peserta(eng, nama, gedung, divisi, pin, rek_rapi)
                catatan = "" if rek_rapi else " Nomor rekening belum diisi, bisa ditambahkan nanti lewat panel Edit di bawah."
                st.success(f"{nama.strip()} ditambahkan ke Gedung {gedung}. Kasih tahu PIN-nya ke yang bersangkutan.{catatan}")
            except IntegrityError:
                st.error("Nama itu sudah terdaftar.")
            except SQLAlchemyError as e:
                st.error(pesan_db(e))

    pesan = st.session_state.pop("pesan_hapus", None)
    if pesan:
        st.success(pesan)

    df = list_peserta(eng)
    st.dataframe(
        beri_nomor(df.rename(columns={
            "nama": "Nama", "gedung": "Gedung", "divisi": "Divisi", "no_rekening": "No. Rekening",
        }).fillna({"No. Rekening": "-"})),
        width="stretch",
    )
    belum = df[df["no_rekening"].isna() | (df["no_rekening"].astype(str).str.strip() == "")]["nama"].tolist()
    if belum:
        st.caption(
            f"⚠️ {len(belum)} peserta belum punya nomor rekening: {', '.join(belum)}. "
            "Isi lewat panel Edit di bawah (di Excel akan tertulis “-” sampai diisi)."
        )

    if not df.empty:
        with st.expander("Edit Divisi / No. Rekening / Reset PIN / Hapus peserta"):
            target = st.selectbox("Pilih peserta", df["nama"], key="target_kelola")

            opsi_divisi = list(DIVISI_GEDUNG.keys())
            divisi_lama = df.loc[df["nama"] == target, "divisi"].values[0]
            idx_default = opsi_divisi.index(divisi_lama) if divisi_lama in opsi_divisi else 0
            divisi_baru = st.selectbox("Divisi", opsi_divisi, index=idx_default, key="divisi_baru")
            st.caption(f"Ditempatkan di: Gedung {DIVISI_GEDUNG[divisi_baru]}")
            if st.button("Simpan Divisi"):
                try:
                    set_divisi(eng, target, divisi_baru, DIVISI_GEDUNG[divisi_baru])
                except SQLAlchemyError as e:
                    st.error(pesan_db(e))
                else:
                    st.success("Divisi & Gedung diperbarui.")
                    st.rerun()

            rek_lama = df.loc[df["nama"] == target, "no_rekening"].values[0]
            rek_baru = st.text_input(
                "Nomor rekening", value=rek_lama if isinstance(rek_lama, str) else "",
                max_chars=30, key=f"rek_baru_{target}",
            )
            if st.button("Simpan No. Rekening"):
                rek_rapi, rek_err = validasi_rekening(rek_baru)
                if rek_err:
                    st.error(rek_err)
                else:
                    try:
                        set_rekening(eng, target, rek_rapi)
                    except SQLAlchemyError as e:
                        st.error(pesan_db(e))
                    else:
                        st.success("Nomor rekening diperbarui.")
                        st.rerun()

            pin_baru = st.text_input("PIN baru (4-6 angka)", max_chars=6, key="pin_baru")
            hapus_absen_juga = st.checkbox(
                "Saat menghapus peserta, hapus juga seluruh riwayat absensinya",
                key="hapus_absen_juga",
            )
            c1, c2 = st.columns(2)
            if c1.button("Reset PIN"):
                if pin_baru.isdigit() and 4 <= len(pin_baru) <= 6:
                    try:
                        reset_pin(eng, target, pin_baru)
                        st.success("PIN diperbarui.")
                    except SQLAlchemyError as e:
                        st.error(pesan_db(e))
                else:
                    st.error("PIN harus 4-6 angka.")
            if c2.button("Hapus peserta"):
                try:
                    hapus_peserta(eng, target, dengan_absensi=hapus_absen_juga)
                except SQLAlchemyError as e:
                    st.error(pesan_db(e))
                else:
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
                    try:
                        jumlah = hapus_absensi(eng, None if pilihan == semua else pilihan)
                        st.success(f"{jumlah} data absensi (beserta fotonya) dihapus.")
                    except SQLAlchemyError as e:
                        st.error(pesan_db(e))


# ------------------------------------------------------------------
# HALAMAN ADMIN: REKAP, GRAFIK, EXPORT
# ------------------------------------------------------------------
def halaman_rekap(eng):
    st.header("📊 Rekap Absensi & Uang Makan")
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

    berhasil = df[(df["status"] == "berhasil") & (df["jenis"] == "masuk")]       # hari hadir = absen masuk berhasil
    pulang_ok = df[(df["status"] == "berhasil") & (df["jenis"] == "pulang")]
    izin_df = df[df["jenis"] == "izin"]
    rekap = pd.DataFrame({"Nama": sorted(df["nama"].unique())})
    rekap["Hari Hadir"] = rekap["Nama"].map(berhasil.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Absen Pulang"] = rekap["Nama"].map(pulang_ok.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Izin"] = rekap["Nama"].map(izin_df.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Total Uang Makan"] = rekap["Nama"].map(berhasil.groupby("nama")["uang_makan"].sum()).fillna(0).astype(int)

    m1, m2, m3 = st.columns(3)
    m1.metric("Total uang makan bulan ini", f"Rp{int(rekap['Total Uang Makan'].sum()):,}".replace(",", "."))
    m2.metric("Izin bulan ini", int(izin_df.groupby(["nama", "tanggal"]).ngroups))
    m3.metric("Percobaan gagal lokasi", int(((df["status"] != "berhasil") & (df["jenis"] != "izin")).sum()))

    st.subheader("Total Uang Makan per Peserta")
    st.dataframe(beri_nomor(rekap), width="stretch")

    st.subheader("Grafik")
    if not rekap.empty:
        st.caption("Hari hadir per peserta")
        st.bar_chart(rekap.set_index("Nama")["Hari Hadir"], color="#2563eb")
    harian = berhasil.groupby("tanggal").size()
    if not harian.empty:
        st.caption("Jumlah peserta yang hadir per hari")
        st.bar_chart(harian, color="#2563eb")

    st.subheader("Detail Absensi")
    detail = df.rename(columns={
        "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis",
        "status": "Status", "jarak": "Jarak (m)", "uang_makan": "Uang Makan", "alasan": "Alasan",
    })
    detail["Alasan"] = detail["Alasan"].fillna("")
    detail["Jarak (m)"] = detail["Jarak (m)"].round(0)

    daftar_nama = sorted(detail["Nama"].unique())
    pilih = st.selectbox("Pilih nama peserta", daftar_nama)
    detail_pilih = detail[detail["Nama"] == pilih].reset_index(drop=True)
    ok = detail_pilih[(detail_pilih["Status"] == "berhasil") & (detail_pilih["Jenis"] == "masuk")]
    n_izin = detail_pilih[detail_pilih["Jenis"] == "izin"]["Tanggal"].nunique()
    st.caption(
        f"{len(detail_pilih)} catatan • {ok['Tanggal'].nunique()} hari hadir • {n_izin} izin • "
        f"Uang makan Rp{int(ok['Uang Makan'].sum()):,}".replace(",", ".")
    )
    tampil = detail_pilih.copy()
    tampil["Jarak (m)"] = tampil["Jarak (m)"].apply(lambda v: "" if pd.isna(v) else str(int(round(v))))  # izin: tanpa jarak
    st.dataframe(beri_nomor(tampil), width="stretch")

    st.subheader("Download Excel")
    st.caption(
        "Format template absensi magang (Nama, Unit Kerja, No. Rekening, tabel kehadiran, tanda tangan) "
        "terisi otomatis dari data peserta."
    )
    tahun_i, bulan_i = int(tahun), int(bulan)
    aman = lambda t: re.sub(r"[^A-Za-z0-9]+", "_", t).strip("_")
    mime_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    try:
        data_satu = excel_template_bytes(eng, df, tahun_i, bulan_i, [pilih])
        data_semua = excel_template_bytes(eng, df, tahun_i, bulan_i, daftar_nama)
    except FileNotFoundError:
        data_satu = data_semua = None
        st.warning(
            "File **template_absen.xlsx** tidak ditemukan di folder aplikasi. "
            "Unggah file itu ke folder yang sama dengan app.py (GitHub), lalu muat ulang."
        )
    d1, d2 = st.columns(2)
    if data_satu:
        d1.download_button(
            f"⬇️ Excel {pilih}", data=data_satu, width="stretch",
            file_name=f"ABSEN_MAGANG_{aman(pilih)}_{tahun_i}_{bulan_i:02d}.xlsx", mime=mime_xlsx,
        )
    if data_semua:
        d2.download_button(
            "⬇️ Excel semua peserta", data=data_semua, width="stretch",
            file_name=f"ABSEN_MAGANG_SEMUA_{tahun_i}_{bulan_i:02d}.xlsx", mime=mime_xlsx,
        )
    st.download_button(
        "⬇️ Rekap ringkas (tabel)", data=excel_bytes(rekap, detail),
        file_name=f"rekap_absensi_{tahun_i}_{bulan_i:02d}.xlsx", mime=mime_xlsx,
    )


# ------------------------------------------------------------------
# HALAMAN ADMIN: BUKTI FOTO
# ------------------------------------------------------------------
def halaman_foto(eng):
    st.header("🖼️ Bukti Foto Absensi")
    if not admin_login():
        return

    tgl = st.date_input("Tanggal", value=datetime.now(WITA).date())
    rows = [r for r in get_absensi_hari_foto(eng, tgl.strftime("%Y-%m-%d")) if r["jenis"] != "izin"]
    if not rows:
        st.info("Tidak ada absen masuk/pulang di tanggal ini.")
        return

    daftar_nama = sorted({r["nama"] for r in rows})
    pilih = st.selectbox("Pilih nama peserta", daftar_nama)
    milik = [r for r in rows if r["nama"] == pilih]
    st.caption(f"{len(milik)} catatan absen pada tanggal ini. Klik baris untuk melihat fotonya.")

    for r in milik:
        ikon = "✅" if r["status"] == "berhasil" else "❌"
        judul = f"{ikon} {r['jam']} • {r['jenis']} • {(r['jarak'] or 0):.0f} m"
        with st.expander(judul):
            isi = ambil_foto(r["foto"], r["foto_path"])
            if isi:
                st.image(isi, width=170)
                st.caption("Disimpan di Supabase Storage" if r["foto_path"] else "Disimpan di database")
            else:
                st.caption("Foto tidak tersedia.")


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
        f"""<div class="auth-head">
<div class="auth-brand"><div class="auth-logo">{ic("fingerprint", 30)}</div>
<div><div class="auth-title">Presensi</div><div class="auth-bank">Bank Sulselbar • Absensi PKL</div></div></div>
<div class="auth-greet">Selamat Datang 👋</div>
<div class="auth-desc">Silakan masuk untuk melanjutkan</div>
</div>""",
        unsafe_allow_html=True,
    )
    with st.container(key="loginbody"):
        peran = st.radio(
            "Masuk sebagai", ["Peserta PKL", "Admin"], horizontal=True, key="login_peran",
            label_visibility="collapsed",
        )

        if peran == "Peserta PKL":
            df = list_peserta(eng)
            if df.empty:
                st.warning("Belum ada peserta terdaftar. Hubungi admin untuk didaftarkan.")
            else:
                with st.form("login_peserta", border=False):
                    nama = st.selectbox("Nama peserta", df["nama"])
                    pin = st.text_input("PIN", type="password", max_chars=6, placeholder="Masukkan PIN kamu")
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
            else:
                if not secret("admin_password") and not secret("database_url"):
                    st.info("Mode lokal: password admin default = admin123")
                with st.form("login_admin", border=False):
                    pw = st.text_input("Password admin", type="password", placeholder="Masukkan password admin")
                    kirim = st.form_submit_button("Masuk sebagai admin", type="primary", width="stretch")
                if kirim:
                    if hmac.compare_digest(pw, pw_benar):
                        st.session_state.peran = "admin"
                        st.session_state.admin_ok = True
                        st.rerun()
                    else:
                        st.error("Password salah. Silakan coba lagi.")

    st.markdown('<div class="auth-foot">© Bank Sulselbar • Absensi Peserta PKL</div>', unsafe_allow_html=True)


def secret_nyasar(kunci):
    """Cari kunci Secrets yang salah tempat (terselip di dalam bagian [xxx])."""
    try:
        for nama, isi in st.secrets.items():
            if hasattr(isi, "keys") and kunci in isi.keys():
                return nama
    except Exception:
        pass
    return None


def panel_status(eng):
    """Kotak kecil di panel admin: database & storage yang sedang dipakai."""
    baris = []
    if eng.dialect.name == "postgresql":
        host = eng.url.host or eng.url.query.get("host") or "lokal"
        baris.append(("ok", f"Database: PostgreSQL / Supabase ({host})"))
    else:
        petunjuk = "database_url belum terbaca dari Secrets, data TIDAK masuk ke Supabase."
        nyasar = secret_nyasar("database_url")
        if nyasar:
            petunjuk = f"database_url terselip di dalam bagian [{nyasar}] pada Secrets. Pindahkan ke paling atas."
        baris.append(("warn", f"Database: SQLite lokal. {petunjuk}"))
    cfg = storage_konfig()
    if cfg:
        baris.append(("ok", f"Foto: Supabase Storage (bucket {cfg[2]})"))
    else:
        baris.append(("warn", "Foto: disimpan di database (supabase_url / supabase_service_key belum terbaca)"))
    html = "".join(
        f'<div class="st-{k}">{"🟢" if k == "ok" else "🟠"} {t}</div>' for k, t in baris
    )
    st.markdown(f'<div class="st-box">{html}</div>', unsafe_allow_html=True)


def realtime_admin():
    """Berlangganan Supabase Realtime di browser admin. Ada absen baru -> dashboard otomatis dimuat ulang."""
    url = str(secret("supabase_url") or "").strip().rstrip("/")
    anon = str(secret("supabase_anon_key") or "").strip()
    if not (url and anon):
        st.caption("🔕 Realtime belum aktif (isi `supabase_url` dan `supabase_anon_key` di Secrets).")
        return
    with st.container(key="rt_hidden"):
        if st.button("segarkan", key="rt_refresh"):
            st.toast("Ada data absensi baru. Dashboard diperbarui.", icon="🔔")
    st.caption("")
    components.html(
        """<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.min.js"></script>
<script>
let timer = null;
function segarkan() {
  clearTimeout(timer);
  timer = setTimeout(function () {
    try {
      var b = window.parent.document.querySelector('.st-key-rt_refresh button');
      if (b) b.click();
    } catch (e) { console.error(e); }
  }, 800);
}
window.__segarkan = segarkan;
try {
  var sb = supabase.createClient(%s, %s);
  sb.channel('dashboard-absensi')
    .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'notifikasi' }, segarkan)
    .subscribe();
} catch (e) { console.error(e); }
</script>""" % (json.dumps(url), json.dumps(anon)),
        height=0,
    )


def halaman_admin(eng):
    """Area admin dengan gaya visual yang selaras dengan halaman peserta."""
    st.markdown(CSS_ADMIN, unsafe_allow_html=True)
    st.markdown(
        f"""<div class="admin-head"><div class="admin-top">
<div class="admin-logo">{ic("fingerprint", 28)}</div>
<div><div class="admin-title">Panel Admin</div><div class="admin-sub">Bank Sulselbar • Pengelolaan Absensi PKL</div></div>
<div style="flex:1"></div>
<div class="admin-out">{ic("logout", 24)}</div>
</div></div>""",
        unsafe_allow_html=True,
    )
    with st.container(key="adminlogout"):
        if st.button("‎", key="keluar_admin", width="stretch"):
            keluar()

    with st.container(key="adminbody"):
        if getattr(eng, "db_sementara", False):
            st.warning(
                "Database masih bersifat SEMENTARA: penyimpanan aplikasi read-only, jadi data disimpan di folder "
                "sementara dan akan HILANG saat aplikasi restart/reboot. Isi `database_url` (Supabase/Postgres) "
                "di Secrets agar data permanen."
            )
        panel_status(eng)
        realtime_admin()
        menu = st.radio(
            "Menu admin",
            ["Kelola Peserta", "Rekap & Export", "Bukti Foto"],
            horizontal=True,
            key="menu_admin",
            label_visibility="collapsed",
        )
        st.markdown("---")
        if menu == "Kelola Peserta":
            halaman_kelola_peserta(eng)
        elif menu == "Rekap & Export":
            halaman_rekap(eng)
        else:
            halaman_foto(eng)


def main():
    st.set_page_config(page_title="Presensi - Bank Sulselbar", page_icon="📍", layout="centered")
    st.markdown(CSS_TEMA, unsafe_allow_html=True)
    try:
        eng = get_engine()
    except SQLAlchemyError as e:
        asal = getattr(e, "orig", e)
        st.error(
            "Tidak bisa terhubung ke database. Periksa pengaturan database di Secrets "
            "(`database_url`, atau `db_host`, `db_user`, `db_password`). Pakai alamat Transaction pooler; "
            "nama pengguna harus `postgres.KODE_PROJECT`.\n\n"
            f"Detail: {type(asal).__name__}: {str(asal).strip()[:400]}"
        )
        st.stop()

    if not st.session_state.get("skema_ok"):
        try:
            _migrasi(eng)       # tambah kolom yang belum ada (aman dijalankan berulang)
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
            st.stop()
        st.session_state.skema_ok = True

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
