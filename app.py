"""
Aplikasi Absensi Peserta PKL - Bank Sulselbar (versi siap deploy)
=================================================================
Fitur:
- Layar login di awal: pilih peran. Peserta (nama + PIN) -> halaman absen saja; Admin (password) -> area admin
- Absen MASUK saja (tidak ada absen pulang): foto selfie + lokasi GPS
- Login PIN per peserta (dibuat admin)
- Cek lokasi: harus di sekitar gedung yang terdaftar untuk peserta (A / B)
- Setiap peserta juga punya Divisi (dipilih dari daftar, Gedung ikut otomatis)
- Jam absen: masuk 06:00-12:30 (sampai 08:00 tidak terlambat, lewat itu terlambat), pulang 16:50-18:00 (WITA).
  Uang makan didapat saat absen pulang, termasuk yang masuknya terlambat.
  Izin (sebentar, tetap bisa absen) dan cuti (sehari penuh, tidak bisa absen masuk): kapan saja, dengan alasan
- Uang makan Rp15.000 otomatis saat absen masuk berhasil
- Blokir absen dobel di hari yang sama
- Admin (pakai password): kelola peserta, rekap + export Excel, analitik EDA (grafik), model ML (deteksi absen mencurigakan, prediksi terlambat, monitoring), bukti foto
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
from PIL import Image, ImageOps
from sqlalchemy import (
    Column, ForeignKey, Index, Integer, String, Float, LargeBinary, MetaData, Table, Text,
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
    "A": (-5.1635049454681825, 119.42393534081793),  # (latitude, longitude)
    "B": (-5.1635049454681825, 119.42393534081793),  # (latitude, longitude)
}

RADIUS_METER = 200  # toleransi jarak dari titik gedung (meter)

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
JAM_MASUK_BUKA, JAM_MASUK_TUTUP = dtime(21, 0), dtime(23, 30)
JAM_PULANG_BUKA, JAM_PULANG_TUTUP = dtime(16, 50), dtime(23, 45)
# Hari libur PKL (0=Senin ... 5=Sabtu, 6=Minggu). Pada hari ini TIDAK bisa absen masuk, pulang, izin, maupun cuti.
HARI_LIBUR = (5, 6)
# Masuk sampai jam ini = tidak terlambat; setelahnya (sampai JAM_MASUK_TUTUP) = terlambat, tetap bisa absen pulang & dapat uang makan.
BATAS_TEPAT_WAKTU = dtime(8, 0)
KET_TEPAT_WAKTU, KET_TERLAMBAT = "Tidak terlambat", "Terlambat"      # tulisan di aplikasi & kolom Keterangan Excel

WITA = timezone(timedelta(hours=8))  # server cloud pakai UTC, jadi jam dipaksa WITA


# ------------------------------------------------------------------
# HELPER: SECRETS
# ------------------------------------------------------------------
def secret(key, default=None):
    """Nilai dari Streamlit Secrets; kalau tidak ada, dari environment variable (nama huruf besar, mis. DB_HOST).
    Aman kalau file secrets belum ada."""
    try:
        v = st.secrets.get(key, None)
    except Exception:
        v = None
    if v is None:
        v = os.environ.get(key.upper())
    return default if v is None else v


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
    Column("peserta_id", Integer, ForeignKey("peserta.id", ondelete="SET NULL")),  # penghubung resmi ke peserta
    Column("nama", String(100), nullable=False, index=True),
    Column("tanggal", String(10), nullable=False, index=True),  # YYYY-MM-DD
    Column("jam", String(8), nullable=False),                   # HH:MM:SS
    Column("jenis", String(10), nullable=False, default="masuk"),  # masuk / pulang / izin / cuti
    Column("lat", Float),
    Column("lon", Float),
    Column("akurasi", Float),
    Column("jarak", Float),
    Column("status", String(20), nullable=False),               # berhasil / gagal_lokasi / izin / cuti
    Column("uang_makan", Integer, nullable=False, default=0),
    Column("foto", LargeBinary),                                # foto di DB (cadangan / data lama)
    Column("foto_path", String(200)),                           # lokasi foto di Supabase Storage
    Column("alasan", String(300)),                              # alasan izin / cuti
)

# Anti absen ganda DI TINGKAT DATABASE: satu catatan per (nama, tanggal, jenis) untuk masuk/pulang yang berhasil
# dan izin/cuti. Percobaan "gagal_lokasi" boleh berkali-kali. Ini menutup celah balapan (dua klik/dua perangkat
# bersamaan) yang tidak bisa dicegah oleh pengecekan sudah_berhasil() saja.
NAMA_INDEKS_UNIK = "uq_absensi_sekali"
NAMA_INDEKS_GABUNGAN = "ix_absensi_nama_tanggal"
Index(NAMA_INDEKS_UNIK, absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jenis, unique=True,
      sqlite_where=text("status <> 'gagal_lokasi'"), postgresql_where=text("status <> 'gagal_lokasi'"))
Index(NAMA_INDEKS_GABUNGAN, absensi_t.c.nama, absensi_t.c.tanggal)

# Tabel kecil untuk Supabase Realtime: setiap absen baru menambah 1 baris,
# dashboard admin berlangganan perubahannya. Tidak berisi data sensitif.
notifikasi_t = Table(
    "notifikasi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("jenis", String(20)),
    Column("dibuat", String(19)),                               # YYYY-MM-DD HH:MM:SS (WITA)
)


# Versi model ML (prediksi risiko terlambat). Artefak = model terlatih (joblib); metrik & baseline = JSON.
# Hanya satu versi per jenis yang aktif. Model anomali tidak disimpan: tanpa label, dilatih ulang tiap analisis.
model_versi_t = Table(
    "model_versi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("jenis", String(30), nullable=False, index=True),
    Column("dilatih", String(19), nullable=False),               # YYYY-MM-DD HH:MM:SS (WITA)
    Column("n_data", Integer, nullable=False, default=0),
    Column("tanggal_akhir", String(10)),                         # tanggal data terakhir yang ikut dilatih
    Column("metrik", Text),
    Column("baseline", Text),
    Column("artefak", LargeBinary),
    Column("aktif", Integer, nullable=False, default=0),
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


def _pasang_indeks(eng):
    """Pasang indeks gabungan dan indeks unik anti-absen-ganda pada database yang sudah ada (create_all tidak
    menambah indeks ke tabel lama). Mengembalikan True kalau indeks unik terpasang; False kalau gagal
    (biasanya karena sudah ada data ganda lama; lihat cari_duplikat / rapikan_duplikat)."""
    ada = {i["name"] for i in inspect(eng).get_indexes("absensi")}
    if NAMA_INDEKS_GABUNGAN not in ada:
        try:
            with eng.begin() as c:
                c.execute(text(f"CREATE INDEX IF NOT EXISTS {NAMA_INDEKS_GABUNGAN} ON absensi (nama, tanggal)"))
        except SQLAlchemyError:
            pass
    if NAMA_INDEKS_UNIK in ada:
        return True
    try:
        with eng.begin() as c:
            c.execute(text(
                f"CREATE UNIQUE INDEX IF NOT EXISTS {NAMA_INDEKS_UNIK} "
                "ON absensi (nama, tanggal, jenis) WHERE status <> 'gagal_lokasi'"
            ))
        return True
    except SQLAlchemyError:
        return False


def _migrasi(eng):
    """Tambah kolom/indeks baru ke database lama (create_all tidak mengubah tabel yang sudah ada). Aman diulang."""
    perlu = {
        "peserta": [("no_rekening", "VARCHAR(40)")],
        "absensi": [
            ("foto_path", "VARCHAR(200)"), ("alasan", "VARCHAR(300)"),
            ("peserta_id", "INTEGER REFERENCES peserta(id) ON DELETE SET NULL"),
        ],
    }
    insp = inspect(eng)
    for tabel, daftar in perlu.items():
        ada = {k["name"] for k in insp.get_columns(tabel)}
        for kolom, tipe in daftar:
            if kolom not in ada:
                with eng.begin() as c:
                    c.execute(text(f"ALTER TABLE {tabel} ADD COLUMN {kolom} {tipe}"))
    with eng.begin() as c:    # isi peserta_id untuk riwayat lama (hanya yang namanya masih terdaftar)
        c.execute(text(
            "UPDATE absensi SET peserta_id = (SELECT p.id FROM peserta p WHERE p.nama = absensi.nama) "
            "WHERE peserta_id IS NULL AND EXISTS (SELECT 1 FROM peserta p WHERE p.nama = absensi.nama)"
        ))
    eng.indeks_unik_ok = _pasang_indeks(eng)


@st.cache_resource
def get_engine(versi_skema=7):
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


def cari_duplikat(eng):
    """Kelompok absensi ganda (nama, tanggal, jenis, jumlah) yang menghalangi pemasangan indeks unik."""
    q = (
        select(absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jenis, func.count().label("jumlah"))
        .where(absensi_t.c.status != "gagal_lokasi")
        .group_by(absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jenis)
        .having(func.count() > 1)
        .order_by(absensi_t.c.tanggal, absensi_t.c.nama)
    )
    with eng.connect() as c:
        return [tuple(r) for r in c.execute(q).all()]


@coba_ulang
def rapikan_duplikat(eng):
    """Hapus absensi ganda, sisakan catatan PALING AWAL di tiap kelompok (jadi uang makan tidak terhitung dobel),
    lalu pasang indeks unik. Mengembalikan jumlah baris yang dihapus."""
    paths_hapus, paths_simpan, hapus = [], set(), 0
    with eng.begin() as c:
        grup = c.execute(
            select(absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jenis, func.min(absensi_t.c.id).label("simpan"))
            .where(absensi_t.c.status != "gagal_lokasi")
            .group_by(absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jenis)
            .having(func.count() > 1)
        ).all()
        for nama, tanggal, jenis, simpan in grup:
            kembar = and_(
                absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal, absensi_t.c.jenis == jenis,
                absensi_t.c.status != "gagal_lokasi",
            )
            paths_simpan.update(r[0] for r in c.execute(select(absensi_t.c.foto_path).where(absensi_t.c.id == simpan)))
            paths_hapus += [r[0] for r in c.execute(select(absensi_t.c.foto_path).where(and_(kembar, absensi_t.c.id != simpan)))]
            hapus += c.execute(delete(absensi_t).where(and_(kembar, absensi_t.c.id != simpan))).rowcount
    hapus_foto_storage([p for p in paths_hapus if p and p not in paths_simpan])   # jangan hapus foto milik catatan yang dipertahankan
    eng.indeks_unik_ok = _pasang_indeks(eng)
    return hapus


def foto_dipakai(eng, path):
    """True kalau ada baris absensi yang memakai path foto itu."""
    with eng.connect() as c:
        return c.execute(select(func.count()).select_from(absensi_t).where(absensi_t.c.foto_path == path)).scalar() > 0


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
        # riwayat yang dipertahankan dilepas dari peserta (SQLite tidak menjalankan ON DELETE SET NULL)
        c.execute(update(absensi_t).where(absensi_t.c.nama == nama).values(peserta_id=None))
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


def hari_libur(tgl):
    """True kalau tanggal (date/datetime) jatuh di hari libur PKL (Sabtu/Minggu)."""
    return tgl.weekday() in HARI_LIBUR


def pesan_libur(tgl):
    return (f"Hari {HARI_ID[tgl.weekday()]} adalah hari libur PKL. Absen masuk, pulang, izin, dan cuti "
            "tidak bisa dilakukan pada hari Sabtu dan Minggu.")


def terlambat(jam):
    """True kalau jam masuk (str 'HH:MM[:SS]' atau time) lewat BATAS_TEPAT_WAKTU. Dihitung per menit: 08:00 masih tepat waktu."""
    if isinstance(jam, str):
        jam = datetime.strptime(jam[:5], "%H:%M").time()
    return jam.replace(second=0, microsecond=0) > BATAS_TEPAT_WAKTU


def ketepatan(jam):
    return KET_TERLAMBAT if terlambat(jam) else KET_TEPAT_WAKTU


def ada_catatan(eng, nama, tanggal, jenis):
    """True kalau peserta sudah punya catatan izin/cuti (jenis) pada tanggal itu."""
    q = select(func.count()).select_from(absensi_t).where(and_(
        absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal, absensi_t.c.jenis == jenis,
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
        if nilai.get("peserta_id") is None:
            nilai["peserta_id"] = c.execute(select(peserta_t.c.id).where(peserta_t.c.nama == nilai.get("nama"))).scalar()
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
# Penandatangan di bagian bawah Excel, tergantung divisi peserta: (nama, jabatan).
# Divisi yang masih ("", "") akan dikosongkan di Excel (bukan memakai nama divisi lain),
# dan admin diberi peringatan. Kalau datanya sudah ada, isi di sini, contoh:
#   "Divisi Kepatuhan": ("Nama Pimpinan", "Pimdept. Kepatuhan"),
PENANDATANGAN = {
    "Divisi Manajemen Risiko": ("Faisal Hafiz", "Pimdept. Risiko Operasional, TI & Lainnya"),
    # --- Gedung A ---
    "Divisi Korporasi dan Komersial": ("", ""),
    "Divisi Ritel dan Konsumer": ("", ""),
    "Divisi Corporate Secretary (CORSEC)": ("", ""),
    "Divisi Human Capital": ("", ""),
    "Divisi SIMO": ("", ""),
    "Divisi Keuangan dan Akuntansi": ("", ""),
    "Divisi Treasury": ("", ""),
    "Divisi Teknologi dan Informasi": ("", ""),
    "Divisi International Banking": ("", ""),
    "Divisi Digitalisasi dan Layanan": ("", ""),
    "Divisi Umum": ("", ""),
    # --- Gedung B ---
    "Divisi Kepatuhan": ("", ""),
    "Divisi Dana dan Jasa": ("", ""),
    "Divisi Syariah": ("", ""),
    "Divisi SKAI": ("", ""),
}


def cari_penandatangan(divisi):
    """(nama, jabatan) penandatangan untuk divisi itu; ("", "") kalau belum diisi."""
    kunci = (divisi or "").strip().lower()
    for k, v in PENANDATANGAN.items():
        if k.strip().lower() == kunci:
            nama, jabatan = (v[0] or "").strip(), (v[1] or "").strip()
            if nama or jabatan:                      # lewati entri kosong kalau ada entri lain yang berisi
                return nama, jabatan
    return "", ""


def divisi_tanpa_penandatangan(eng, daftar_nama):
    """Daftar divisi (dari peserta yang akan diekspor) yang penandatangannya belum diisi."""
    hasil = []
    for nama in daftar_nama:
        info = get_peserta(eng, nama)
        divisi = (info["divisi"] if info else "") or "(tanpa divisi)"
        if not any(cari_penandatangan(divisi)) and divisi not in hasil:
            hasil.append(divisi)
    return hasil


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


def _isi_sheet_peserta(ws, tpl, peserta, hari, pulang, izin, cuti, awal, akhir):
    """Isi satu sheet memakai gaya dari sheet 'Template' (kolom B..G).
    hari = {tanggal_iso: jam_masuk}, pulang = {tanggal_iso: (jam_pulang, uang_makan)},
    izin / cuti = {tanggal_iso: alasan}."""
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
    ws.column_dimensions["E"].width = max(ws.column_dimensions["E"].width or 0, 16)   # muat "Tidak terlambat"

    def baris(sr, dr, nilai_ganti=True):
        for col in "BCDEFG":
            sc = tpl[f"{col}{sr}"]
            _salin_gaya(sc, ws[f"{col}{dr}"], ganti(sc.value) if nilai_ganti else None)
        h = tpl.row_dimensions[sr].height
        if h:
            ws.row_dimensions[dr].height = h

    def tinggi(r, teks, lebar=28):
        ws.row_dimensions[r].height = 15 * max(1, math.ceil(len(teks) / lebar))

    for r in range(3, 8):          # identitas + judul kolom
        baris(r, r)

    r = 8
    tgl = awal
    while tgl <= akhir:
        iso = tgl.isoformat()
        tanggal_dt = datetime(tgl.year, tgl.month, tgl.day)
        if iso in hari:                                   # hadir (masuk berhasil)
            baris(8, r, nilai_ganti=False)
            ws[f"B{r}"].value = tanggal_dt
            ws[f"C{r}"].value = hari[iso]
            jam_p, uang = pulang.get(iso, (None, 0))      # uang makan hanya kalau sudah absen pulang
            ws[f"D{r}"].value = jam_p if jam_p else "-"
            ws[f"E{r}"].value = ketepatan(hari[iso])          # "Tidak terlambat" / "Terlambat"
            if iso in izin:                               # kolom Izin: alasan izin hari itu, selain itu "-"
                ws[f"F{r}"].value = izin[iso]
                ws[f"F{r}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                tinggi(r, izin[iso])
            else:
                ws[f"F{r}"].value = "-"
            ws[f"G{r}"].value = uang
        else:                                             # satu baris blok: cuti / tidak hadir / libur
            baris(9, r, nilai_ganti=False)
            ws[f"B{r}"].value = tanggal_dt
            if iso in cuti:
                teks = f"Cuti: {cuti[iso]}"
            elif iso in izin and tgl.weekday() < 5:
                teks = f"Tidak hadir (izin: {izin[iso]})"
            else:
                teks = "Libur " if tgl.weekday() >= 5 else "Tidak hadir"
            if teks.startswith(("Cuti", "Tidak hadir (")):
                ws[f"C{r}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws[f"B{r}"].alignment = Alignment(horizontal="center", vertical="center")
                tinggi(r, teks, lebar=62)
            ws[f"C{r}"].value = teks
            ws.merge_cells(f"C{r}:G{r}")
        r += 1
        tgl += timedelta(days=1)

    jumlah = r
    baris(10, jumlah)
    ws.merge_cells(f"B{jumlah}:F{jumlah}")
    ws[f"G{jumlah}"].value = f"=SUM(G8:G{jumlah - 1})" if jumlah > 8 else 0

    geser = jumlah - 10             # footer (tanda tangan) ikut turun
    penanda_nama, penanda_jabatan = cari_penandatangan(peserta["divisi"])
    for sr in range(11, tpl.max_row + 1):
        baris(sr, sr + geser)
    for rng in tpl.merged_cells.ranges:
        if rng.min_row >= 11:
            ws.merge_cells(start_row=rng.min_row + geser, end_row=rng.max_row + geser,
                           start_column=rng.min_col, end_column=rng.max_col)
    for sr in range(11, tpl.max_row + 1):             # tulisan "Faisal Hafiz" di template selalu diganti sesuai divisi
        if tpl[f"F{sr}"].value == "Faisal Hafiz ":
            ws[f"F{sr + geser}"].value = penanda_nama or None
            ws[f"F{sr + geser + 1}"].value = penanda_jabatan or None

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
        hari, pulang, izin, cuti = {}, {}, {}, {}
        for _, row in data.iterrows():
            iso = row["tanggal"]
            if row["jenis"] == "izin":
                izin.setdefault(iso, str(row["alasan"] or "-").strip())
            elif row["jenis"] == "cuti":
                cuti.setdefault(iso, str(row["alasan"] or "-").strip())
            elif row["status"] == "berhasil":
                jam = datetime.strptime(row["jam"], "%H:%M:%S").time().replace(second=0)
                if row["jenis"] == "masuk":
                    hari.setdefault(iso, jam)
                elif row["jenis"] == "pulang":
                    pulang[iso] = (jam, int(row["uang_makan"]))   # jam pulang + uang makan (didapat saat pulang)
        pertama = tanggal_pertama_absen(eng, nama) or data["tanggal"].min()
        awal = max(awal_bulan, date.fromisoformat(pertama))
        if awal > akhir:
            continue
        ws = wb.create_sheet(_nama_sheet(nama, dipakai))
        _isi_sheet_peserta(ws, tpl, peserta, hari, pulang, izin, cuti, awal, akhir)
        jumlah_sheet += 1

    if jumlah_sheet == 0:
        return None
    del wb["Template"]
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()



def kompres_foto(data: bytes) -> bytes:
    """Kecilkan foto (maks 640px, JPEG q70) supaya database tidak cepat penuh."""
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail((640, 640))
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=70, optimize=True)
    return out.getvalue()

# ------------------------------------------------------------------
# VERIFIKASI WAJAH (absen masuk / pulang): wajah harus ada dan pas di dalam bingkai oval di layar kamera
# ------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def _muat_kaskade():
    """Muat pendeteksi wajah Haar bawaan OpenCV 4.x. Raise kalau gagal (supaya kegagalan tidak ikut ter-cache)."""
    import cv2
    if not hasattr(cv2, "CascadeClassifier"):
        raise RuntimeError("tanpa CascadeClassifier")
    hasil = []
    for nama in ("haarcascade_frontalface_default.xml", "haarcascade_frontalface_alt2.xml"):
        c = cv2.CascadeClassifier(os.path.join(cv2.data.haarcascades, nama))
        if not c.empty():
            hasil.append(c)
    if not hasil:
        raise RuntimeError("file model tidak ditemukan")
    return hasil


def kaskade_wajah():
    try:
        return _muat_kaskade()
    except Exception:
        return None


def status_wajah():
    """(aktif, keterangan) untuk panel admin."""
    try:
        import cv2
    except Exception as e:
        return False, f"OpenCV belum terpasang ({type(e).__name__}). Cek requirements.txt dan Reboot app."
    if not hasattr(cv2, "CascadeClassifier"):
        return False, (f"OpenCV {cv2.__version__} tidak punya pendeteksi wajah Haar (dihapus di OpenCV 5). "
                       "Pakai opencv-python-headless versi 4.x di requirements.txt, lalu Reboot app.")
    if not kaskade_wajah():
        return False, f"OpenCV {cv2.__version__}: file model wajah tidak ditemukan."
    return True, f"OpenCV {cv2.__version__}"


def ukuran_oval(lebar, tinggi):
    """Ukuran bingkai oval yang digambar di layar kamera (sama dengan CSS): tinggi maks 78% frame,
    lebar maks 62% frame, rasio lebar:tinggi 0.78."""
    h = min(0.78 * tinggi, 0.62 * lebar / 0.78)
    return 0.78 * h, h


def _iou(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    irisan = ix * iy
    gabungan = aw * ah + bw * bh - irisan
    return irisan / gabungan if gabungan else 0


def periksa_wajah(data: bytes):
    """(ok, pesan). Foto valid kalau ada wajah menghadap depan yang pas di dalam bingkai oval.
    Wajah dihitung hanya jika dua pendeteksi sepakat (mengurangi deteksi palsu); kemiringan sampai ~24 derajat dicoba."""
    kaskade = kaskade_wajah()
    if not kaskade:
        return False, "Verifikasi wajah belum aktif di server. Hubungi admin."
    import cv2
    import numpy as np
    try:
        img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("L")
    except Exception:
        return False, "Foto tidak bisa dibaca. Ambil foto lagi."
    if img.width > 640:
        img = img.resize((640, max(1, int(img.height * 640 / img.width))))
    gray0 = cv2.equalizeHist(np.array(img))
    tinggi, lebar = gray0.shape
    ow, oh = ukuran_oval(lebar, tinggi)
    min_sisi = max(20, int(min(lebar, tinggi) * 0.15))

    terdekat = None                                  # (jarak_ke_pusat, rasio_lebar) wajah paling dekat bingkai
    for sudut in (0, -12, 12, -24, 24):
        if sudut:
            M = cv2.getRotationMatrix2D((lebar / 2, tinggi / 2), sudut, 1.0)
            gray = cv2.warpAffine(gray0, M, (lebar, tinggi), borderMode=cv2.BORDER_REPLICATE)
            Mi = cv2.invertAffineTransform(M)
        else:
            gray = gray0
        per_pendeteksi = [
            [tuple(int(v) for v in f) for f in c.detectMultiScale(gray, 1.1, 5, minSize=(min_sisi, min_sisi))]
            for c in kaskade
        ]
        if len(per_pendeteksi) > 1:
            kandidat = [f for f in per_pendeteksi[0] if any(_iou(f, g) > 0.4 for g in per_pendeteksi[1])]
        else:
            kandidat = per_pendeteksi[0]
        for x, y, w, h in kandidat:
            cx, cy = x + w / 2, y + h / 2
            if sudut:
                cx, cy = Mi @ np.array([cx, cy, 1.0])
            dx, dy = (cx - lebar / 2) / (ow / 2), (cy - tinggi / 2) / (oh / 2)
            jarak, rasio = dx * dx + dy * dy, w / ow
            if jarak <= 1.0 and 0.40 <= rasio <= 1.30:
                return True, "ok"                    # ada wajah yang pas di dalam bingkai
            if terdekat is None or jarak < terdekat[0]:
                terdekat = (jarak, rasio)

    if terdekat is None:
        return False, "Wajah tidak terdeteksi. Pastikan wajahmu terlihat jelas menghadap kamera dan cahaya cukup."
    if terdekat[0] > 1.0:
        return False, "Posisikan wajahmu tepat di dalam bingkai oval."
    if terdekat[1] < 0.40:
        return False, "Wajah terlalu jauh. Dekatkan wajah ke kamera sampai memenuhi bingkai oval."
    return False, "Wajah terlalu dekat. Mundurkan sedikit sampai wajah pas di dalam bingkai oval."




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
.ok { color: #14a068; } .bad { color: #e04a4a; } .izin { color: #c58a05; } .cuti { color: #2f5fd0; }
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

/* ---------- petunjuk bawaan input: hilangkan "Press Enter to submit form", sisakan penghitung (0/6) ---------- */
[data-testid="InputInstructions"] > span:first-child:not(:last-child),
[data-testid="stForm"] [data-testid="InputInstructions"] > span:only-child { display: none; }
[data-testid="InputInstructions"] > span::before { content: none !important; }   /* tanda titik pemisah sebelum 0/6 */

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
.st-key-pagebody [data-testid="stRadioOption"] p { font-weight: 600; font-size: 14px; color: var(--biru); text-align: center; white-space: nowrap; }
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] {
  background: linear-gradient(135deg, var(--biru), var(--ungu)); box-shadow: 0 4px 12px rgba(37,63,216,.3);
}
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] p { color: #fff; }
.st-key-pagebody [data-baseweb="textarea"], .st-key-pagebody [data-baseweb="base-input"] {
  background: #f6f8ff !important; border-color: var(--garis) !important; border-radius: 14px !important;
}
.st-key-pagebody textarea, .st-key-pagebody input { background: transparent !important; color: var(--navy) !important; }
.st-key-pagebody textarea::placeholder { color: var(--muted) !important; opacity: 1; }

/* ---------- bingkai wajah di layar kamera (sama dengan area yang dicek server) ---------- */
.st-key-pagebody [data-testid="stCameraInputWebcamStyledBox"] { position: relative; }
.st-key-pagebody [data-testid="stCameraInputWebcamStyledBox"]::after {
  content: ""; position: absolute; inset: 11% 19%; pointer-events: none; z-index: 2;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 78 100'%3E%3Cellipse cx='39' cy='50' rx='37' ry='48' fill='none' stroke='black' stroke-opacity='0.35' stroke-width='3.6'/%3E%3Cellipse cx='39' cy='50' rx='37' ry='48' fill='none' stroke='white' stroke-width='1.8' stroke-dasharray='6 4'/%3E%3Ccircle cx='28' cy='42' r='1.8' fill='white' fill-opacity='0.8'/%3E%3Ccircle cx='50' cy='42' r='1.8' fill='white' fill-opacity='0.8'/%3E%3Cpath d='M39 46 L37 56 L41 56' fill='none' stroke='white' stroke-opacity='0.7' stroke-width='1.2'/%3E%3Cpath d='M30 66 Q39 72 48 66' fill='none' stroke='white' stroke-opacity='0.8' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E") center / contain no-repeat;
}
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
    hasil = {"masuk": None, "pulang": None, "izin": None, "cuti": None,
             "gagal_masuk": False, "gagal_pulang": False}
    for r in rows:
        if r["jenis"] in ("izin", "cuti"):
            hasil[r["jenis"]] = r
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
        q = q.where(or_(absensi_t.c.status == "berhasil", absensi_t.c.jenis.in_(["izin", "cuti"])))
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


def teks_status_hari(now, sh, html=True):
    """(judul, keterangan) untuk kotak status di Beranda, sesuai kondisi & jam sekarang.
    html=False: teks polos (untuk API / aplikasi native), alasan tidak di-escape."""
    esc = globals()["esc"] if html else (lambda x: x)
    f = lambda t: t.strftime("%H:%M")
    jendela_masuk = f"{f(JAM_MASUK_BUKA)}-{f(JAM_MASUK_TUTUP)}"
    jendela_pulang = f"{f(JAM_PULANG_BUKA)}-{f(JAM_PULANG_TUTUP)}"
    t = now.time().replace(second=0, microsecond=0)
    if hari_libur(now):
        return "Hari libur", f"{HARI_ID[now.weekday()]} libur PKL. Absen masuk, pulang, izin, dan cuti tidak tersedia."
    if sh["masuk"] is None and sh["cuti"] is not None:
        return "Cuti hari ini", esc(str(sh["cuti"]["alasan"] or "-"))
    if sh["masuk"] is None:
        if t < JAM_MASUK_BUKA:
            judul, sub = "Status presensi", f"Absen masuk dibuka pukul {f(JAM_MASUK_BUKA)} WITA"
        elif t <= JAM_MASUK_TUTUP:
            ekstra = " Percobaan terakhir gagal (di luar area)." if sh["gagal_masuk"] else ""
            if t <= BATAS_TEPAT_WAKTU:
                info = f"tidak terlambat sampai {f(BATAS_TEPAT_WAKTU)}, batas absen {f(JAM_MASUK_TUTUP)}"
            else:
                info = f"sudah lewat {f(BATAS_TEPAT_WAKTU)}, dihitung terlambat. Batas absen {f(JAM_MASUK_TUTUP)}"
            judul, sub = "Status presensi", f"Absen masuk dibuka ({info}) WITA.{ekstra}"
        else:
            judul, sub = "Status presensi", f"Waktu absen masuk ({jendela_masuk}) sudah lewat"
    elif sh["pulang"] is None:
        if t < JAM_PULANG_BUKA:
            judul, sub = "Status presensi", f"Absen pulang dibuka pukul {f(JAM_PULANG_BUKA)} WITA"
        elif t <= JAM_PULANG_TUTUP:
            ekstra = " Percobaan terakhir gagal (di luar area)." if sh["gagal_pulang"] else ""
            judul, sub = "Status presensi", f"Absen pulang dibuka sampai {f(JAM_PULANG_TUTUP)} WITA.{ekstra}"
        else:
            judul, sub = "Status presensi", f"Waktu absen pulang ({jendela_pulang}) sudah lewat"
    else:
        judul, sub = "Status presensi", (f"Presensi hari ini lengkap: masuk {sh['masuk']['jam'][:5]} "
                                         f"({ketepatan(sh['masuk']['jam']).lower()}), pulang {sh['pulang']['jam'][:5]}")
    if sh["izin"] is not None:
        sub += f" • Izin: {esc(str(sh['izin']['alasan'] or '-'))}"
    return judul, sub


def tab_beranda(eng, nama, peserta):
    now = datetime.now(WITA)
    sh = get_status_hari(eng, nama, now.strftime("%Y-%m-%d"))
    jam_masuk = sh["masuk"]["jam"][:5] if sh["masuk"] else "--:--"
    jam_pulang = sh["pulang"]["jam"][:5] if sh["pulang"] else "--:--"
    uang = rupiah(sh["pulang"]["uang_makan"]) if sh["pulang"] else "-"      # uang makan didapat saat absen pulang
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


def _hasil(level, pesan, **ekstra):
    return {"level": level, "pesan": pesan, "ok": level == "success", **ekstra}


def proses_absen(eng, nama, peserta, jenis, foto_bytes, lat, lon, akurasi=None, now=None):
    """Satu-satunya tempat aturan absen masuk/pulang (dipakai website Streamlit DAN API aplikasi native).
    Mengembalikan dict: level (success/warning/error), pesan, ok, dan info tambahan."""
    gedung = peserta["gedung"]
    buka, tutup = (JAM_MASUK_BUKA, JAM_MASUK_TUTUP) if jenis == "masuk" else (JAM_PULANG_BUKA, JAM_PULANG_TUTUP)
    rentang = f"{buka.strftime('%H:%M')}-{tutup.strftime('%H:%M')}"
    now = now or datetime.now(WITA)
    tanggal, jam = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")

    if hari_libur(now):
        return _hasil("warning", pesan_libur(now), alasan="libur")
    if ada_catatan(eng, nama, tanggal, "cuti") and not sudah_berhasil(eng, nama, tanggal, "masuk"):
        return _hasil("warning", f"Kamu cuti hari ini, jadi tidak bisa absen {jenis}.")
    if sudah_berhasil(eng, nama, tanggal, jenis):
        return _hasil("warning", f"Kamu sudah absen {jenis} hari ini.")
    if jenis == "pulang" and not sudah_berhasil(eng, nama, tanggal, "masuk"):
        return _hasil("warning", "Kamu belum absen masuk hari ini, jadi belum bisa absen pulang.")
    if not dalam_jendela(now, buka, tutup):
        return _hasil("error", f"❌ Absen ditolak: {jenis} hanya dibuka pukul {rentang} WITA (sekarang {now.strftime('%H:%M')}).")

    wajah_ok, pesan_wajah = periksa_wajah(foto_bytes)        # bukan wajah / posisi salah = tidak dicatat sama sekali
    if not wajah_ok:
        return _hasil("error", f"📷 {pesan_wajah} Ambil foto lagi.", alasan="wajah")

    cocok, jarak = cek_lokasi(lat, lon, gedung)
    status = "berhasil" if cocok else "gagal_lokasi"
    uang = UANG_MAKAN if (cocok and jenis == "pulang") else 0    # uang makan didapat saat absen PULANG

    foto_simpan, path_foto, di_storage = None, None, False
    if cocok:                                                    # foto HANYA disimpan kalau absen berhasil
        foto_kecil = kompres_foto(foto_bytes)
        nama_berkas = re.sub(r"[^A-Za-z0-9]+", "_", nama).strip("_") or "peserta"
        path_foto = f"{tanggal}/{nama_berkas}_{jenis}_{jam.replace(':', '')}.jpg"
        di_storage = simpan_foto_storage(path_foto, foto_kecil)   # False kalau Storage belum diatur / gagal
        foto_simpan = None if di_storage else foto_kecil
    try:
        simpan_absen(
            eng, nama=nama, tanggal=tanggal, jam=jam, jenis=jenis,
            lat=lat, lon=lon, akurasi=akurasi, jarak=jarak,
            status=status, uang_makan=uang,
            foto=foto_simpan,
            foto_path=path_foto if di_storage else None,
        )
    except IntegrityError:                                       # kalah balapan: absen yang sama sudah tersimpan
        if di_storage and path_foto and not foto_dipakai(eng, path_foto):
            hapus_foto_storage([path_foto])                      # buang foto yatim (kecuali dipakai catatan pemenang)
        return _hasil("warning", f"Kamu sudah absen {jenis} hari ini.")
    except SQLAlchemyError:
        return _hasil("error", "Absen gagal disimpan karena gangguan database. Coba lagi sebentar lagi atau hubungi admin.")

    if not cocok:
        return _hasil(
            "error",
            f"❌ Absen GAGAL: kamu ~{jarak:.0f} m dari Gedung {gedung}, "
            "di luar area yang diizinkan. Pastikan kamu sudah di gedung lalu coba lagi.",
            status=status, jarak=round(jarak),
        )
    ket = ketepatan(jam) if jenis == "masuk" else None
    tambahan = f" ({ket.lower()})" if ket else ""
    if jenis == "pulang":
        info, gaya = f"💰 Uang makan hari ini: {rupiah(uang)}", "info"
    else:
        info, gaya = "Uang makan didapat setelah absen pulang.", "caption"
    return _hasil("success", f"✅ Absen {jenis} berhasil pukul {jam} WITA{tambahan}.",
                  status=status, jarak=round(jarak), jam=jam, uang=uang, ketepatan=ket, info=info, info_gaya=gaya)


def proses_tidak_hadir(eng, nama, jenis, tgl, alasan, now=None):
    """Izin / cuti (tgl = date). Aturan bersama untuk website dan aplikasi native."""
    label = jenis.capitalize()
    now = now or datetime.now(WITA)
    teks = " ".join((alasan or "").split())
    tanggal = tgl.strftime("%Y-%m-%d")
    if hari_libur(now):
        return _hasil("warning", pesan_libur(now), alasan="libur")
    if hari_libur(tgl):
        return _hasil("error", f"{tgl_indonesia(tgl)} jatuh pada hari libur (Sabtu/Minggu), jadi tidak perlu "
                               f"mengajukan {jenis}. Pilih hari kerja (Senin sampai Jumat).")
    if len(teks) < 5:
        return _hasil("error", f"Alasan {jenis} wajib diisi (minimal 5 karakter).")
    if tgl < now.date() or tgl > now.date() + timedelta(days=30):
        return _hasil("error", f"Tanggal {jenis} harus antara hari ini dan 30 hari ke depan.")
    if ada_catatan(eng, nama, tanggal, jenis):
        return _hasil("warning", f"Kamu sudah mengajukan {jenis} untuk {tgl_indonesia(tgl)}.")
    if jenis == "cuti" and sudah_berhasil(eng, nama, tanggal, "masuk"):
        return _hasil("warning", "Kamu sudah absen masuk pada tanggal itu, jadi tidak bisa mengajukan cuti.")
    if jenis == "izin" and ada_catatan(eng, nama, tanggal, "cuti"):
        return _hasil("warning", "Kamu cuti pada tanggal itu, jadi tidak perlu mengajukan izin.")
    try:
        simpan_absen(
            eng, nama=nama, tanggal=tanggal, jam=now.strftime("%H:%M:%S"),
            jenis=jenis, status=jenis, uang_makan=0, alasan=teks,
        )
    except IntegrityError:
        return _hasil("warning", f"Kamu sudah mengajukan {jenis} untuk {tgl_indonesia(tgl)}.")
    except SQLAlchemyError:
        return _hasil("error", f"{label} gagal disimpan karena gangguan database. Coba lagi sebentar lagi atau hubungi admin.")
    return _hasil("success", f"✅ {label} untuk {tgl_indonesia(tgl)} tersimpan.")


def _tampilkan_hasil(hasil):
    {"success": st.success, "warning": st.warning, "error": st.error}[hasil["level"]](hasil["pesan"])
    if hasil.get("info"):
        (st.info if hasil.get("info_gaya") == "info" else st.caption)(hasil["info"])


def _form_absen(eng, nama, peserta, jenis, now, sh):
    """Form absen masuk / pulang: hanya di jendela waktunya, wajib di area gedung + selfie."""
    gedung = peserta["gedung"]
    buka, tutup = (JAM_MASUK_BUKA, JAM_MASUK_TUTUP) if jenis == "masuk" else (JAM_PULANG_BUKA, JAM_PULANG_TUTUP)
    rentang = f"{buka.strftime('%H:%M')}-{tutup.strftime('%H:%M')}"
    if jenis == "masuk":
        st.caption(f"Jam absen masuk: {rentang} WITA (tidak terlambat sampai {BATAS_TEPAT_WAKTU.strftime('%H:%M')}, "
                   "lewat itu dihitung terlambat)")
    else:
        st.caption(f"Jam absen pulang: {rentang} WITA")

    if sh["cuti"] is not None and sh["masuk"] is None:
        st.warning(f"Kamu cuti hari ini ({sh['cuti']['alasan']}), jadi tidak bisa absen {jenis}.")
        return
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

    st.caption("Posisikan wajahmu di dalam bingkai oval, lalu ambil foto. Foto tanpa wajah tidak diterima.")
    foto = st.camera_input("Foto selfie untuk absen", key=f"foto_{jenis}")

    if st.button(f"Absen {jenis.capitalize()} Sekarang", type="primary", width="stretch", key=f"btn_{jenis}"):
        if not coords:
            st.error("Lokasi belum terdeteksi. Aktifkan GPS/izin lokasi lalu coba lagi.")
            return
        if foto is None:
            st.error("Foto belum diambil.")
            return

        lat, lon = coords["latitude"], coords["longitude"]
        hasil = proses_absen(eng, nama, peserta, jenis, foto.getvalue(), lat, lon, coords.get("accuracy"))
        _tampilkan_hasil(hasil)


def _form_tidak_hadir(eng, nama, now, jenis):
    """Form izin / cuti: bisa kapan saja, tanpa lokasi/foto, wajib mengisi alasan.
    Izin  = keluar sebentar, TETAP bisa absen masuk/pulang.
    Cuti  = sehari penuh, TIDAK bisa absen masuk pada tanggal itu."""
    label = jenis.capitalize()
    if jenis == "izin":
        st.caption("Izin untuk keperluan sebentar (misalnya ke kampus beberapa jam). Kamu tetap bisa absen masuk dan pulang. "
                   "Alasan akan masuk ke laporan Excel.")
    else:
        st.caption("Cuti sehari penuh. Pada tanggal cuti kamu tidak bisa absen masuk. Alasan akan masuk ke laporan Excel.")
    tgl = st.date_input(
        f"Tanggal {jenis}", value=now.date(), min_value=now.date(), max_value=now.date() + timedelta(days=30),
        key=f"{jenis}_tanggal",
    )
    alasan = st.text_area(
        f"Alasan {jenis}", max_chars=300, height=110, key=f"{jenis}_alasan",
        placeholder="Contoh: ke kampus untuk bimbingan, sakit, keperluan keluarga..." if jenis == "izin"
        else "Contoh: sakit, acara keluarga, ujian kampus...",
    )
    if st.button(f"Kirim {label}", type="primary", width="stretch", key=f"btn_{jenis}"):
        _tampilkan_hasil(proses_tidak_hadir(eng, nama, jenis, tgl, alasan))


def tab_absen(eng, nama, peserta):
    header_halaman("Absen", f"{nama} • Gedung {peserta['gedung']} • {peserta['divisi']}")
    now = datetime.now(WITA)
    sh = get_status_hari(eng, nama, now.strftime("%Y-%m-%d"))

    if hari_libur(now):
        with st.container(key="pagebody"):
            st.info(f"🗓️ {pesan_libur(now)}")
        return

    if "mode_absen" not in st.session_state:         # pilihan awal mengikuti kondisi hari ini
        st.session_state.mode_absen = "Pulang" if (sh["masuk"] is not None and sh["pulang"] is None) else "Masuk"

    with st.container(key="pagebody"):
        mode = st.radio("Jenis absen", ["Masuk", "Pulang", "Izin", "Cuti"], horizontal=True,
                        key="mode_absen", label_visibility="collapsed")
        if mode in ("Izin", "Cuti"):
            _form_tidak_hadir(eng, nama, now, mode.lower())
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
            if r["jenis"] in ("izin", "cuti"):
                label = r["jenis"].capitalize()
                info = f"{label} • {esc(str(r['alasan'] or '-'))}"
                ket, kelas = label, r["jenis"]
            else:
                label = "Masuk" if r["jenis"] == "masuk" else "Pulang"
                info = f"{label} {r['jam'][:5]} WITA • {r['jarak'] or 0:.0f} m"
                if r["jenis"] == "pulang" and r["uang_makan"]:
                    ket, kelas = rupiah(r["uang_makan"]), "ok"          # uang makan muncul di baris pulang
                elif r["jenis"] == "masuk":
                    ket = ketepatan(r["jam"])
                    kelas = "izin" if terlambat(r["jam"]) else "ok"       # terlambat = oranye, tidak terlambat = hijau
                else:
                    ket, kelas = "Berhasil", "ok"
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
    pulang_ok = [r for r in rows if r["jenis"] == "pulang" and r["status"] == "berhasil"]
    hadir = len({r["tanggal"] for r in masuk_ok})
    total = sum(r["uang_makan"] for r in pulang_ok)              # uang makan dihitung dari absen pulang
    izin = len({r["tanggal"] for r in rows if r["jenis"] == "izin"})
    cuti = len({r["tanggal"] for r in rows if r["jenis"] == "cuti"})
    n_terlambat = len({r["tanggal"] for r in masuk_ok if terlambat(r["jam"])})
    with st.container(key="pagebody"):
        st.markdown(
            f'<div class="pr-stat"><div><span>Hari hadir</span><b>{hadir}</b></div>'
            f'<div><span>Total uang makan</span><b>{rupiah(total)}</b></div></div>'
            f'<div class="pr-kv"><span>Terlambat</span><span>{n_terlambat}</span></div>'
            f'<div class="pr-kv"><span>Izin</span><span>{izin}</span></div>'
            f'<div class="pr-kv"><span>Cuti</span><span>{cuti}</span></div>',
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


# Streamlit Cloud menaruh dua ikon melayang di pojok kanan bawah ("Manage app" dan "Hosted with Streamlit"). Ikon itu
# dipasang platform di halaman induk, DI ATAS bingkai aplikasi, jadi tidak bisa disembunyikan atau ditimpa dari dalam
# aplikasi. Solusinya: bilah navigasi diperpanjang ke bawah (jalur putih kosong setinggi NAV_NAIK_PX) supaya ikon
# duduk di jalur itu dan tidak menutupi tombol. Otomatis aktif di alamat .streamlit.app; di aplikasi Android
# "Presensi PKL" tidak dipakai karena aplikasi itu menghilangkan ikonnya sendiri.
# Paksa dengan ?angkat=1, matikan dengan ?angkat=0.
NAV_NAIK_PX = 56
CSS_NAV_NAIK = f"""
<style>
.st-key-bottomnav {{ padding-bottom: {12 + NAV_NAIK_PX}px !important; }}
.pr-spacer {{ height: {150 + NAV_NAIK_PX}px !important; }}
</style>
"""


def nav_perlu_naik():
    try:
        pilihan = str(st.query_params.get("angkat", "")).strip()
    except Exception:
        pilihan = ""
    if pilihan in ("0", "1"):
        return pilihan == "1"
    try:
        if "PresensiPKLApp" in (st.context.headers.get("user-agent") or ""):
            return False                                   # aplikasi Android: ikon sudah dihilangkan oleh aplikasinya
        host = (st.context.headers.get("host") or "").lower().split(":")[0]
        return host.endswith(".streamlit.app")
    except Exception:
        return False


def halaman_peserta(eng, nama):
    """Tampilan peserta setelah login: Beranda (Presensi) + Absen / Riwayat / Rekap / Profil."""
    peserta = get_peserta(eng, nama)
    if peserta is None:
        st.error("Data kamu tidak ditemukan (mungkin sudah dihapus admin). Silakan keluar lalu hubungi admin.")
        if st.button("Keluar"):
            keluar()
        return

    st.markdown(CSS_PRESENSI, unsafe_allow_html=True)
    if nav_perlu_naik():
        st.markdown(CSS_NAV_NAIK, unsafe_allow_html=True)
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

    duplikat = [] if getattr(eng, "indeks_unik_ok", True) else cari_duplikat(eng)
    if duplikat:
        with st.expander(f"🧹 Rapikan absensi ganda ({len(duplikat)} kelompok)"):
            st.caption(
                "Ada catatan ganda (nama + tanggal + jenis sama) dari sebelum aturan anti-ganda dipasang. "
                "Yang dipertahankan adalah catatan paling awal; sisanya dihapus beserta fotonya. "
                "Setelah itu indeks anti-ganda dipasang otomatis."
            )
            st.dataframe(beri_nomor(pd.DataFrame(duplikat, columns=["Nama", "Tanggal", "Jenis", "Jumlah"])), width="stretch")
            ok_rapikan = st.text_input("Ketik RAPIKAN (huruf besar) untuk konfirmasi", key="rapikan_konfirmasi")
            if st.button("Rapikan absensi ganda", key="tombol_rapikan"):
                if ok_rapikan.strip() != "RAPIKAN":
                    st.error("Konfirmasi belum benar. Ketik persis: RAPIKAN")
                else:
                    try:
                        n = rapikan_duplikat(eng)
                    except SQLAlchemyError as e:
                        st.error(pesan_db(e))
                    else:
                        st.success(f"{n} catatan ganda dihapus. Indeks anti-ganda "
                                   + ("terpasang." if eng.indeks_unik_ok else "BELUM terpasang, cek lagi data."))
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
    pulang_ok = df[(df["status"] == "berhasil") & (df["jenis"] == "pulang")]     # uang makan dihitung dari absen pulang
    izin_df = df[df["jenis"] == "izin"]
    cuti_df = df[df["jenis"] == "cuti"]
    rekap = pd.DataFrame({"Nama": sorted(df["nama"].unique())})
    rekap["Hari Hadir"] = rekap["Nama"].map(berhasil.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Absen Pulang"] = rekap["Nama"].map(pulang_ok.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    terlambat_df = berhasil[berhasil["jam"].apply(terlambat)]
    rekap["Terlambat"] = rekap["Nama"].map(terlambat_df.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Izin"] = rekap["Nama"].map(izin_df.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Cuti"] = rekap["Nama"].map(cuti_df.groupby("nama")["tanggal"].nunique()).fillna(0).astype(int)
    rekap["Total Uang Makan"] = rekap["Nama"].map(pulang_ok.groupby("nama")["uang_makan"].sum()).fillna(0).astype(int)

    m1, m2, m3 = st.columns(3)
    m1.metric("Total uang makan bulan ini", f"Rp{int(rekap['Total Uang Makan'].sum()):,}".replace(",", "."))
    m2.metric("Izin / Cuti bulan ini", f"{int(rekap['Izin'].sum())} / {int(rekap['Cuti'].sum())}")
    m3.metric("Percobaan gagal lokasi", int(((df["status"] != "berhasil") & (~df["jenis"].isin(["izin", "cuti"]))).sum()))

    st.subheader("Total Uang Makan per Peserta")
    st.dataframe(beri_nomor(rekap), width="stretch")

    st.subheader("Detail Absensi")
    detail = df.rename(columns={
        "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis",
        "status": "Status", "jarak": "Jarak (m)", "uang_makan": "Uang Makan", "alasan": "Alasan",
    })
    detail["Alasan"] = detail["Alasan"].fillna("")
    detail["Keterangan"] = detail.apply(
        lambda r: ketepatan(r["Jam"]) if (r["Jenis"] == "masuk" and r["Status"] == "berhasil") else "", axis=1)
    detail["Jarak (m)"] = detail["Jarak (m)"].round(0)

    daftar_nama = sorted(detail["Nama"].unique())
    pilih = st.selectbox("Pilih nama peserta", daftar_nama)
    detail_pilih = detail[detail["Nama"] == pilih].reset_index(drop=True)
    ok = detail_pilih[(detail_pilih["Status"] == "berhasil") & (detail_pilih["Jenis"] == "masuk")]
    uang_ok = detail_pilih[(detail_pilih["Status"] == "berhasil") & (detail_pilih["Jenis"] == "pulang")]
    n_izin = detail_pilih[detail_pilih["Jenis"] == "izin"]["Tanggal"].nunique()
    n_cuti = detail_pilih[detail_pilih["Jenis"] == "cuti"]["Tanggal"].nunique()
    st.caption(
        f"{len(detail_pilih)} catatan • {ok['Tanggal'].nunique()} hari hadir • {n_izin} izin • {n_cuti} cuti • "
        f"Uang makan Rp{int(uang_ok['Uang Makan'].sum()):,}".replace(",", ".")
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
    kurang = divisi_tanpa_penandatangan(eng, daftar_nama)
    if kurang:
        st.warning(
            "Penandatangan belum diisi untuk: " + ", ".join(kurang) +
            ". Bagian tanda tangan di Excel untuk divisi itu dikosongkan. "
            "Isi nama dan jabatannya di `PENANDATANGAN` pada app.py."
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
# HALAMAN ADMIN: ANALITIK (EDA - Exploratory Data Analysis)
# ------------------------------------------------------------------
def get_absensi_rentang(eng, awal, akhir):
    """Semua catatan absensi antara dua tanggal (date), untuk analitik."""
    q = (
        select(
            absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jam, absensi_t.c.jenis, absensi_t.c.status,
            absensi_t.c.jarak, absensi_t.c.akurasi, absensi_t.c.uang_makan,
        )
        .where(and_(absensi_t.c.tanggal >= awal.isoformat(), absensi_t.c.tanggal <= akhir.isoformat()))
        .order_by(absensi_t.c.tanggal, absensi_t.c.jam)
    )
    with eng.connect() as c:
        return pd.read_sql(q, c)


def menit_hari(jam):
    """'HH:MM[:SS]' -> menit sejak tengah malam."""
    return int(jam[0:2]) * 60 + int(jam[3:5])


def label_menit(m):
    m = max(0, int(round(m)))
    return f"{m // 60:02d}:{m % 60:02d}"


def label_slot(menit, lebar):
    return label_menit((menit // lebar) * lebar)


def daftar_slot(lebar):
    """Semua blok waktu dari jam buka sampai tutup absen masuk (supaya blok kosong tetap tampil sebagai 0)."""
    awal, akhir = JAM_MASUK_BUKA.hour * 60 + JAM_MASUK_BUKA.minute, JAM_MASUK_TUTUP.hour * 60 + JAM_MASUK_TUTUP.minute
    return [label_menit(m) for m in range((awal // lebar) * lebar, akhir + 1, lebar)]


def batas_iqr(seri):
    """(q1, q3, batas_bawah, batas_atas) aturan 1,5 x IQR; None kalau data terlalu sedikit (< 4)."""
    x = pd.Series(seri).dropna().astype(float)
    if len(x) < 4:
        return None
    q1, q3 = float(x.quantile(0.25)), float(x.quantile(0.75))
    iqr = q3 - q1
    return q1, q3, q1 - 1.5 * iqr, q3 + 1.5 * iqr


def siapkan_eda(df, peserta_df):
    """Ubah catatan absensi mentah jadi tabel-tabel siap grafik + ringkasan angka. Fungsi murni (tanpa Streamlit/DB).
    df: kolom nama, tanggal, jam, jenis, status, jarak. peserta_df: kolom nama, gedung, divisi."""
    d = df.copy()
    ref = peserta_df.drop_duplicates("nama").set_index("nama") if len(peserta_df) else pd.DataFrame(columns=["gedung", "divisi"])
    d["gedung"] = d["nama"].map(ref["gedung"]).fillna("?") if len(ref) else "?"
    d["divisi"] = d["nama"].map(ref["divisi"]).fillna("(tidak terdaftar)") if len(ref) else "(tidak terdaftar)"
    d["hari_idx"] = pd.to_datetime(d["tanggal"]).dt.weekday
    d["hari"] = d["hari_idx"].map(lambda i: HARI_ID[i])

    masuk = d[(d["jenis"] == "masuk") & (d["status"] == "berhasil")].copy()
    masuk["menit"] = masuk["jam"].map(menit_hari)
    masuk["terlambat"] = masuk["jam"].map(terlambat).astype(bool)
    masuk["ketepatan"] = masuk["terlambat"].map({True: KET_TERLAMBAT, False: KET_TEPAT_WAKTU})
    masuk["slot15"] = masuk["menit"].map(lambda m: label_slot(m, 15))
    masuk["slot30"] = masuk["menit"].map(lambda m: label_slot(m, 30))

    # tren harian
    tanggal = sorted(d["tanggal"].unique())
    harian = pd.DataFrame({"tanggal": tanggal})
    def per_tgl(sub):
        return harian["tanggal"].map(sub.groupby("tanggal")["nama"].nunique()).fillna(0).astype(int)
    harian["Hadir"] = per_tgl(masuk)
    harian["Terlambat"] = per_tgl(masuk[masuk["terlambat"]])
    harian["Izin"] = per_tgl(d[d["jenis"] == "izin"])
    harian["Cuti"] = per_tgl(d[d["jenis"] == "cuti"])

    # histogram jam masuk (grid lengkap: blok x ketepatan, isi 0 kalau kosong)
    slot15 = daftar_slot(15)
    grid = pd.MultiIndex.from_product([slot15, [KET_TEPAT_WAKTU, KET_TERLAMBAT]], names=["slot", "ketepatan"])
    hist = (masuk.groupby(["slot15", "ketepatan"]).size().rename_axis(["slot", "ketepatan"])
            .reindex(grid, fill_value=0).reset_index(name="jumlah"))
    # slot di luar jam buka (data lama/koreksi) tidak boleh hilang diam-diam
    luar = masuk[~masuk["slot15"].isin(slot15)]
    if len(luar):
        tambahan = luar.groupby(["slot15", "ketepatan"]).size().rename_axis(["slot", "ketepatan"]).reset_index(name="jumlah")
        hist = pd.concat([hist, tambahan], ignore_index=True).sort_values("slot", kind="stable").reset_index(drop=True)

    # heatmap hari x blok 30 menit
    heat = masuk.groupby(["hari", "hari_idx", "slot30"]).size().reset_index(name="jumlah")

    # % terlambat per hari dalam seminggu
    per_hari = (masuk.groupby(["hari_idx", "hari"]).agg(total=("nama", "size"), terlambat=("terlambat", "sum"))
                .reset_index().sort_values("hari_idx"))
    per_hari["terlambat"] = per_hari["terlambat"].astype(int)
    per_hari["persen"] = (per_hari["terlambat"] / per_hari["total"] * 100) if len(per_hari) else []

    per_peserta = (masuk.groupby("nama").agg(hari_hadir=("tanggal", "nunique"), terlambat=("terlambat", "sum"))
                   .reset_index().sort_values(["terlambat", "hari_hadir"], ascending=[False, False]).reset_index(drop=True))
    per_peserta["terlambat"] = per_peserta["terlambat"].astype(int)

    # jarak ke titik gedung (semua percobaan masuk/pulang, termasuk yang gagal lokasi)
    jarak = d[d["jenis"].isin(["masuk", "pulang"]) & d["jarak"].notna()][
        ["nama", "tanggal", "jam", "jenis", "status", "jarak", "gedung"]].copy()
    jarak["hasil"] = jarak["status"].map({"berhasil": "Berhasil", "gagal_lokasi": "Gagal lokasi"}).fillna(jarak["status"])
    jarak["jarak"] = jarak["jarak"].astype(float).round(1)

    hari_aktif = harian[harian["Hadir"] > 0]
    return {
        "kosong": d.empty, "harian": harian, "masuk": masuk, "hist": hist, "heat": heat, "slot30": daftar_slot(30),
        "per_hari": per_hari, "per_peserta": per_peserta, "jarak": jarak,
        "rata_hadir": float(hari_aktif["Hadir"].mean()) if len(hari_aktif) else 0.0,
        "pct_telat": float(masuk["terlambat"].mean() * 100) if len(masuk) else 0.0,
        "median_menit": float(masuk["menit"].median()) if len(masuk) else None,
        "n_masuk": int(len(masuk)),
        "n_gagal": int((jarak["status"] == "gagal_lokasi").sum()),
        "n_percobaan": int(len(jarak)),
    }


def buat_insight(eda):
    """Temuan otomatis dalam kalimat, dari hasil siapkan_eda."""
    out, m = [], eda["masuk"]
    if len(m):
        out.append(f"Rata-rata {eda['rata_hadir']:.1f} peserta hadir per hari; {eda['pct_telat']:.0f}% absen masuk terlambat "
                   f"(batas {BATAS_TEPAT_WAKTU.strftime('%H:%M')}).")
        ph = eda["per_hari"][eda["per_hari"]["total"] >= 5]                       # hindari kesimpulan dari data terlalu sedikit
        if len(ph):
            r = ph.sort_values("persen", ascending=False).iloc[0]
            if r["persen"] > 0:
                out.append(f"Hari dengan keterlambatan tertinggi: {r['hari']} ({r['persen']:.0f}% dari {int(r['total'])} absen masuk).")
        iqr = batas_iqr(m["menit"])
        txt = f"Median jam masuk {label_menit(eda['median_menit'])}."
        if iqr:
            q1, q3, lo, hi = iqr
            n_out = int(((m["menit"] < lo) | (m["menit"] > hi)).sum())
            txt += f" Rentang tengah (IQR) {label_menit(q1)} sampai {label_menit(q3)}; {n_out} absen termasuk outlier."
        out.append(txt)
        pp = eda["per_peserta"]
        if len(pp) and pp.iloc[0]["terlambat"] > 0:
            r = pp.iloc[0]
            out.append(f"Peserta paling sering terlambat: {r['nama']} ({int(r['terlambat'])} dari {int(r['hari_hadir'])} hari hadir).")
        sibuk = m.groupby("slot15").size()
        out.append(f"Blok tersibuk absen masuk: {sibuk.idxmax()} (+15 menit), {int(sibuk.max())} absen.")
    if eda["n_percobaan"]:
        gedung_gagal = eda["jarak"][eda["jarak"]["status"] == "gagal_lokasi"]["gedung"].value_counts()
        txt = f"{eda['n_gagal']} dari {eda['n_percobaan']} percobaan absen ({eda['n_gagal'] / eda['n_percobaan'] * 100:.1f}%) gagal lokasi"
        if len(gedung_gagal):
            txt += f"; terbanyak di Gedung {gedung_gagal.index[0]} ({int(gedung_gagal.iloc[0])})"
        out.append(txt + ".")
        ok = eda["jarak"][eda["jarak"]["status"] == "berhasil"]["jarak"]        # outlier hanya bermakna di antara absen yang berhasil
        iqr = batas_iqr(ok)
        if iqr:
            n_out = int(((ok > iqr[3]) | (ok < iqr[2])).sum())
            txt = f"Jarak absen berhasil: median {ok.median():.0f} m; {n_out} absen berjarak ekstrem (outlier)"
            if n_out:
                txt += f", mis. absen di pinggir batas {RADIUS_METER} m"
            out.append(txt + ".")
    return out


def grafik_eda(eda):
    """Kumpulan grafik Altair dari hasil siapkan_eda."""
    import altair as alt
    biru, oranye = "#2563eb", "#f59e0b"
    skala_ket = alt.Scale(domain=[KET_TEPAT_WAKTU, KET_TERLAMBAT], range=[biru, oranye])
    g = {}

    panjang = eda["harian"].melt(id_vars="tanggal", value_vars=["Hadir", "Terlambat", "Izin", "Cuti"],
                                 var_name="Kategori", value_name="Jumlah")
    g["tren"] = alt.Chart(panjang).mark_line(point=True).encode(
        x=alt.X("tanggal:T", title="Tanggal"),
        y=alt.Y("Jumlah:Q", title="Jumlah peserta", axis=alt.Axis(tickMinStep=1)),
        color=alt.Color("Kategori:N", legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["tanggal:T", "Kategori:N", "Jumlah:Q"],
    ).properties(height=260)

    urut = list(dict.fromkeys(eda["hist"]["slot"]))
    g["hist"] = alt.Chart(eda["hist"]).mark_bar().encode(
        x=alt.X("slot:N", sort=urut, title="Jam masuk (blok 15 menit)", axis=alt.Axis(labelAngle=-60)),
        y=alt.Y("jumlah:Q", title="Jumlah absen masuk", axis=alt.Axis(tickMinStep=1)),
        color=alt.Color("ketepatan:N", scale=skala_ket, legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["slot:N", "ketepatan:N", "jumlah:Q"],
    ).properties(height=260)

    if len(eda["heat"]):
        g["heat"] = alt.Chart(eda["heat"]).mark_rect(cornerRadius=3).encode(
            x=alt.X("slot30:N", sort=eda["slot30"], title="Jam masuk (blok 30 menit)", axis=alt.Axis(labelAngle=-45)),
            y=alt.Y("hari:N", sort=HARI_ID[:5], title=None),
            color=alt.Color("jumlah:Q", scale=alt.Scale(scheme="blues"), title="Absen"),
            tooltip=["hari:N", "slot30:N", "jumlah:Q"],
        ).properties(height=200)

    dasar = alt.Chart(eda["per_hari"]).encode(
        x=alt.X("hari:N", sort=HARI_ID[:5], title=None),
        y=alt.Y("persen:Q", title="% absen masuk terlambat", scale=alt.Scale(domain=[0, 100])),
    )
    g["per_hari"] = (
        dasar.mark_bar(color=oranye, cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            tooltip=["hari:N", "total:Q", "terlambat:Q", alt.Tooltip("persen:Q", format=".1f")])
        + dasar.mark_text(dy=-8, color="#1e2a5a").encode(text=alt.Text("persen:Q", format=".0f"))
    ).properties(height=240)

    if len(eda["jarak"]):
        g["jarak"] = alt.Chart(eda["jarak"]).mark_boxplot(extent=1.5, size=40).encode(
            x=alt.X("gedung:N", title="Gedung"),
            y=alt.Y("jarak:Q", title="Jarak ke titik gedung (m)"),
            color=alt.Color("gedung:N", legend=None),
            column=alt.Column("hasil:N", title=None),
        ).properties(width=140, height=260)
    return g


def halaman_analitik(eng):
    st.header("📈 Analitik Absensi (EDA)")
    if not admin_login():
        return

    hari_ini = datetime.now(WITA).date()
    c1, c2 = st.columns(2)
    awal = c1.date_input("Dari tanggal", value=hari_ini - timedelta(days=29), key="eda_awal")
    akhir = c2.date_input("Sampai tanggal", value=hari_ini, key="eda_akhir")
    if awal > akhir:
        st.error("Tanggal awal tidak boleh setelah tanggal akhir.")
        return
    peserta_df = list_peserta(eng)
    opsi = ["(Semua divisi)"] + sorted(d for d in peserta_df["divisi"].dropna().unique() if d)
    divisi = st.selectbox("Divisi", opsi, key="eda_divisi")

    df = get_absensi_rentang(eng, awal, akhir)
    if divisi != "(Semua divisi)":
        df = df[df["nama"].isin(peserta_df.loc[peserta_df["divisi"] == divisi, "nama"])]
    if df.empty:
        st.info("Belum ada data absensi pada filter ini.")
        return

    eda = siapkan_eda(df, peserta_df)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Absen masuk berhasil", eda["n_masuk"])
    m2.metric("Rata-rata hadir/hari", f"{eda['rata_hadir']:.1f}")
    m3.metric("% terlambat", f"{eda['pct_telat']:.0f}%")
    m4.metric("Median jam masuk", label_menit(eda["median_menit"]) if eda["median_menit"] is not None else "-")

    st.subheader("Temuan otomatis")
    insight = buat_insight(eda)
    if insight:
        st.markdown("\n".join(f"- {t}" for t in insight))
    else:
        st.caption("Belum cukup data untuk membuat temuan.")

    grafik = grafik_eda(eda)
    st.subheader("Tren kehadiran harian")
    st.caption("Jumlah peserta per hari: hadir, terlambat, izin, dan cuti.")
    st.altair_chart(grafik["tren"], width="stretch")

    st.subheader("Distribusi jam masuk")
    st.caption("Histogram absen masuk berhasil per blok 15 menit. Biru = tidak terlambat, oranye = terlambat.")
    st.altair_chart(grafik["hist"], width="stretch")

    if "heat" in grafik:
        st.subheader("Peta kepadatan absen masuk")
        st.caption("Heatmap hari x jam masuk. Warna lebih gelap = lebih banyak absen di blok itu.")
        st.altair_chart(grafik["heat"], width="stretch")

    st.subheader("Keterlambatan per hari")
    st.caption("Persentase absen masuk yang terlambat pada tiap hari dalam seminggu.")
    st.altair_chart(grafik["per_hari"], width="stretch")

    if "jarak" in grafik:
        st.subheader("Sebaran jarak absen ke titik gedung")
        st.caption("Box plot (kotak = 50% data tengah, garis = median, titik = outlier). "
                   "Percobaan 'Gagal lokasi' selalu berjarak lebih dari radius yang diizinkan.")
        st.altair_chart(grafik["jarak"], width="stretch")
        j = eda["jarak"][eda["jarak"]["status"] == "berhasil"]
        iqr = batas_iqr(j["jarak"])
        if iqr:
            ekstrem = j[(j["jarak"] > iqr[3]) | (j["jarak"] < iqr[2])].sort_values("jarak", ascending=False).head(10)
            if len(ekstrem):
                st.caption("Absen berhasil dengan jarak ekstrem (10 teratas), layak dicek bila berulang pada peserta yang sama:")
                st.dataframe(beri_nomor(ekstrem[["nama", "tanggal", "jam", "jenis", "hasil", "jarak", "gedung"]].rename(columns={
                    "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis",
                    "hasil": "Hasil", "jarak": "Jarak (m)", "gedung": "Gedung"})), width="stretch")

    st.subheader("Ringkasan per peserta")
    st.dataframe(beri_nomor(eda["per_peserta"].rename(columns={
        "nama": "Nama", "hari_hadir": "Hari hadir", "terlambat": "Terlambat"})), width="stretch")

    st.subheader("Dataset untuk analisis lanjutan")
    st.caption("Unduh data mentah pada filter ini (CSV) untuk dianalisis di Jupyter/Google Colab.")
    st.download_button(
        "⬇️ Dataset absensi (CSV)", data=df.to_csv(index=False).encode("utf-8"),
        file_name=f"dataset_absensi_{awal.isoformat()}_{akhir.isoformat()}.csv", mime="text/csv",
    )


# ------------------------------------------------------------------
# MODEL ML: penyimpanan versi, pelatihan, prediksi, monitoring (logikanya ada di ml_absensi.py)
# ------------------------------------------------------------------
JENIS_MODEL = "terlambat"
KEBIJAKAN_AKTIF = ("jika_lebih_baik", "selalu", "tidak")


def get_absensi_lengkap(eng, awal, akhir):
    """Catatan absensi lengkap dengan koordinat (untuk deteksi anomali)."""
    q = (
        select(
            absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jam, absensi_t.c.jenis, absensi_t.c.status,
            absensi_t.c.lat, absensi_t.c.lon, absensi_t.c.akurasi, absensi_t.c.jarak,
        )
        .where(and_(absensi_t.c.tanggal >= awal.isoformat(), absensi_t.c.tanggal <= akhir.isoformat()))
        .order_by(absensi_t.c.tanggal, absensi_t.c.jam)
    )
    with eng.connect() as c:
        return pd.read_sql(q, c)


def data_masuk_semua(eng):
    """Semua absen masuk yang berhasil (bahan dataset model keterlambatan)."""
    q = (
        select(absensi_t.c.nama, absensi_t.c.tanggal, absensi_t.c.jam)
        .where(and_(absensi_t.c.jenis == "masuk", absensi_t.c.status == "berhasil"))
        .order_by(absensi_t.c.tanggal, absensi_t.c.jam)
    )
    with eng.connect() as c:
        return pd.read_sql(q, c)


@coba_ulang
def simpan_model(eng, hasil, aktif=False, jenis=JENIS_MODEL):
    """Simpan hasil ml.latih_terlambat sebagai versi baru. aktif=True menonaktifkan versi lain. Mengembalikan id."""
    with eng.begin() as c:
        if aktif:
            c.execute(update(model_versi_t).where(model_versi_t.c.jenis == jenis).values(aktif=0))
        r = c.execute(insert(model_versi_t).values(
            jenis=jenis, dilatih=datetime.now(WITA).strftime("%Y-%m-%d %H:%M:%S"), n_data=hasil["n_data"],
            tanggal_akhir=hasil["tanggal_akhir"], metrik=json.dumps(hasil["metrik"]),
            baseline=json.dumps(hasil["baseline"]), artefak=hasil["artefak"], aktif=1 if aktif else 0,
        ))
        return int(r.inserted_primary_key[0])


def daftar_model(eng, jenis=JENIS_MODEL):
    """Riwayat versi model (tanpa artefak), terbaru dulu."""
    q = (
        select(model_versi_t.c.id, model_versi_t.c.dilatih, model_versi_t.c.n_data, model_versi_t.c.tanggal_akhir,
               model_versi_t.c.metrik, model_versi_t.c.aktif)
        .where(model_versi_t.c.jenis == jenis).order_by(model_versi_t.c.id.desc())
    )
    with eng.connect() as c:
        return [dict(r, metrik=json.loads(r["metrik"] or "{}")) for r in c.execute(q).mappings().all()]


def ambil_model(eng, jenis=JENIS_MODEL, model_id=None):
    """Versi model aktif (atau model_id tertentu) lengkap dengan artefak; None kalau belum ada."""
    q = select(model_versi_t).where(model_versi_t.c.jenis == jenis)
    q = q.where(model_versi_t.c.id == model_id) if model_id is not None else q.where(model_versi_t.c.aktif == 1)
    with eng.connect() as c:
        r = c.execute(q.order_by(model_versi_t.c.id.desc()).limit(1)).mappings().first()
    if r is None:
        return None
    return dict(r, metrik=json.loads(r["metrik"] or "{}"), baseline=json.loads(r["baseline"] or "{}"))


@coba_ulang
def aktifkan_model(eng, model_id, jenis=JENIS_MODEL):
    with eng.begin() as c:
        if c.execute(select(model_versi_t.c.id).where(and_(model_versi_t.c.id == model_id, model_versi_t.c.jenis == jenis))).first() is None:
            return False
        c.execute(update(model_versi_t).where(model_versi_t.c.jenis == jenis).values(aktif=0))
        c.execute(update(model_versi_t).where(model_versi_t.c.id == model_id).values(aktif=1))
        return True


def latih_dan_simpan(eng, aktifkan="jika_lebih_baik"):
    """Latih model baru dari seluruh data, simpan sebagai versi baru.
    aktifkan: 'jika_lebih_baik' (aktif bila F1 uji tidak lebih buruk dari model aktif, toleransi 0,02), 'selalu', 'tidak'.
    Dipakai bersama oleh halaman admin, API, dan skrip jalankan_latih.py. Raise ml_absensi.DataKurang kalau data belum cukup."""
    import ml_absensi as ml
    if aktifkan not in KEBIJAKAN_AKTIF:
        raise ValueError(f"aktifkan harus salah satu dari {KEBIJAKAN_AKTIF}")
    hasil = ml.latih_terlambat(ml.bangun_dataset(data_masuk_semua(eng), list_peserta(eng)))
    lama = ambil_model(eng)
    if aktifkan == "selalu" or lama is None:
        aktif = True
    elif aktifkan == "jika_lebih_baik":
        aktif = hasil["metrik"]["uji"]["f1"] >= (lama["metrik"].get("uji") or {}).get("f1", 0.0) - 0.02
    else:
        aktif = False
    mid = simpan_model(eng, hasil, aktif)
    return {"id": mid, "aktif": bool(aktif), "metrik": hasil["metrik"], "ada_model_sebelumnya": lama is not None}


def prediksi_semua(eng, tanggal, hanya_nama=None):
    """Probabilitas terlambat pada `tanggal` (date) untuk semua peserta (atau satu nama), dari riwayat SEBELUM tanggal itu.
    Mengembalikan (DataFrame, baris_model). Raise ml_absensi.DataKurang kalau belum ada model aktif."""
    import ml_absensi as ml
    m = ambil_model(eng)
    if m is None:
        raise ml.DataKurang("Belum ada model aktif. Latih model terlebih dulu di menu Model.")
    paket = ml.dari_bytes(m["artefak"])
    peserta = list_peserta(eng)
    if hanya_nama is not None:
        peserta = peserta[peserta["nama"] == hanya_nama]
    masuk = data_masuk_semua(eng)
    baris = []
    for _, p in peserta.iterrows():
        f = ml.fitur_prediksi(masuk[masuk["nama"] == p["nama"]][["tanggal", "jam"]], tanggal, p["gedung"])
        baris.append({"nama": p["nama"], "divisi": p["divisi"], **f})
    kolom = ["nama", "divisi", *ml.FITUR_TERLAMBAT, "probabilitas", "risiko", "faktor"]
    if not baris:
        return pd.DataFrame(columns=kolom), m
    df = pd.DataFrame(baris)
    df["probabilitas"] = ml.prediksi_proba(paket, df)
    df["risiko"] = df["probabilitas"].map(ml.tingkat_risiko)
    df["faktor"] = ["; ".join(ml.faktor_risiko(r)) for r in df[ml.FITUR_TERLAMBAT].to_dict("records")]
    return df[kolom].sort_values("probabilitas", ascending=False).reset_index(drop=True), m


def ringkasan_pantau(eng, jendela_hari=14):
    """Kesehatan model aktif: drift fitur (PSI), kinerja pada data SESUDAH tanggal latih, dan rekomendasi retrain."""
    import ml_absensi as ml
    m = ambil_model(eng)
    if m is None:
        return None
    paket = ml.dari_bytes(m["artefak"])
    ds = ml.bangun_dataset(data_masuk_semua(eng), list_peserta(eng))
    baru = ds[ds["tanggal"] > (m["tanggal_akhir"] or "")]           # data yang belum pernah dilihat model
    if len(ds):
        batas = (date.fromisoformat(ds["tanggal"].max()) - timedelta(days=jendela_hari)).isoformat()
        terbaru = ds[ds["tanggal"] > batas]
    else:
        terbaru = ds
    drift = ml.pantau_drift(m["baseline"], terbaru)
    ev = ml.evaluasi(paket, baru)
    umur = (datetime.now(WITA).replace(tzinfo=None) - datetime.strptime(m["dilatih"], "%Y-%m-%d %H:%M:%S")).days
    perlu, alasan = ml.rekomendasi_retrain(m["metrik"], drift, ev, umur, len(baru), m["n_data"])
    info = {k: v for k, v in m.items() if k not in ("artefak", "baseline")}
    return {"model": info, "drift": drift, "eval_baru": ev, "n_baru": int(len(baru)), "n_terbaru": int(len(terbaru)),
            "umur_hari": int(umur), "perlu_retrain": perlu, "alasan": alasan}


# ------------------------------------------------------------------
# HALAMAN ADMIN: MODEL (deteksi absen mencurigakan, prediksi terlambat, monitoring)
# ------------------------------------------------------------------
def _hari_kerja_berikut(d):
    while hari_libur(d):
        d += timedelta(days=1)
    return d


def _ringkas_model(m):
    """Kartu ringkas satu versi model: metrik uji, pembanding baseline, confusion matrix, fitur penting."""
    k = m["metrik"]
    uji = k.get("uji", {})
    st.caption(f"Versi {m['id']} • {k.get('model', '-')} • dilatih {m['dilatih']} WITA • {m['n_data']} data "
               f"({k.get('tanggal_awal', '?')} sampai {k.get('tanggal_akhir', '?')})")
    c = st.columns(5)
    c[0].metric("Accuracy", f"{uji.get('accuracy', 0):.2f}")
    c[1].metric("Precision", f"{uji.get('precision', 0):.2f}")
    c[2].metric("Recall", f"{uji.get('recall', 0):.2f}")
    c[3].metric("F1", f"{uji.get('f1', 0):.2f}")
    c[4].metric("ROC-AUC", "-" if uji.get("roc_auc") is None else f"{uji['roc_auc']:.2f}")
    base = k.get("baseline", {})
    st.caption(
        f"Diuji pada {uji.get('n', 0)} data terakhir (mulai {uji.get('tanggal_mulai', '?')}), yang tidak dipakai saat tuning. "
        f"Baseline ({base.get('strategi', '-')}): accuracy {base.get('accuracy', 0):.2f}, F1 {base.get('f1', 0):.2f}. "
        + ("Model LEBIH BAIK dari baseline." if k.get("lebih_baik_dari_baseline") else "Model BELUM lebih baik dari baseline: jangan diandalkan.")
    )
    cm = uji.get("confusion_matrix")
    if cm:
        st.dataframe(pd.DataFrame(cm, index=["Aktual: tidak terlambat", "Aktual: terlambat"],
                                  columns=["Prediksi: tidak terlambat", "Prediksi: terlambat"]), width="stretch")
    import ml_absensi as ml
    penting = pd.DataFrame({"Kepentingan": k.get("kepentingan", {})})
    if len(penting):
        penting.index = [ml.LABEL_FITUR.get(i, i) for i in penting.index]
        st.caption("Fitur yang paling berpengaruh pada model (dinormalkan, total = 1):")
        st.bar_chart(penting, horizontal=True)


def _tab_anomali(eng, ml):
    hari_ini = datetime.now(WITA).date()
    c1, c2 = st.columns(2)
    awal = c1.date_input("Dari tanggal", value=hari_ini - timedelta(days=29), key="an_awal")
    akhir = c2.date_input("Sampai tanggal", value=hari_ini, key="an_akhir")
    persen = st.slider("Sensitivitas: persentase absen yang ditandai", 1, 15, 5, key="an_persen",
                       help="Isolation Forest selalu menandai sebagian kecil data yang paling tidak lazim. "
                            "Naikkan untuk menjaring lebih luas, turunkan agar hanya yang paling janggal.")
    if awal > akhir:
        st.error("Tanggal awal tidak boleh setelah tanggal akhir.")
        return
    try:
        hasil = ml.deteksi_anomali(get_absensi_lengkap(eng, awal, akhir), list_peserta(eng), GEDUNG_COORD, persen=persen)
    except ml.DataKurang as e:
        st.info(str(e))
        return
    tanda = hasil[hasil["ditandai"]]
    m1, m2, m3 = st.columns(3)
    m1.metric("Catatan dianalisis", len(hasil))
    m2.metric("Ditandai", len(tanda))
    m3.metric("Peserta terkait", tanda["nama"].nunique())
    st.caption("Ini sinyal untuk ditinjau manusia, bukan bukti kecurangan. Koordinat yang sama bisa muncul alami "
               "(mis. posisi dari Wi-Fi kantor), dan absen janggal bisa punya alasan sah. Cek bukti foto sebelum menyimpulkan.")
    if tanda.empty:
        st.success("Tidak ada absen yang ditandai pada rentang ini.")
    else:
        tampil = tanda.assign(hasil=tanda["status"].map({"berhasil": "Berhasil", "gagal_lokasi": "Gagal lokasi"}))
        tampil = tampil[["nama", "tanggal", "jam", "jenis", "hasil", "jarak", "akurasi", "skor", "alasan"]].rename(columns={
            "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis", "hasil": "Hasil",
            "jarak": "Jarak (m)", "akurasi": "Akurasi (m)", "skor": "Skor", "alasan": "Alasan"})
        tampil[["Jarak (m)", "Akurasi (m)", "Skor"]] = tampil[["Jarak (m)", "Akurasi (m)", "Skor"]].round(0)
        st.dataframe(beri_nomor(tampil), width="stretch")

    import altair as alt
    import numpy as np
    st.subheader("Peta posisi absen terhadap gedung")
    st.caption(f"Titik (0, 0) = titik gedung peserta. Lingkaran putus-putus = batas {RADIUS_METER} m. Merah = ditandai.")
    ang = np.linspace(0, 2 * np.pi, 121)
    ring = pd.DataFrame({"dx": RADIUS_METER * np.cos(ang), "dy": RADIUS_METER * np.sin(ang), "u": range(121)})
    titik = hasil.assign(Penanda=np.where(hasil["ditandai"], "Ditandai", "Normal"))
    batas = float(max(titik["dx"].abs().max(), titik["dy"].abs().max(), RADIUS_METER) * 1.1)
    skala = alt.Scale(domain=[-batas, batas])
    garis = alt.Chart(ring).mark_line(strokeDash=[5, 4], color="#6b779e").encode(
        x=alt.X("dx:Q", scale=skala, title="Timur-barat (m)"), y=alt.Y("dy:Q", scale=skala, title="Utara-selatan (m)"), order="u:Q")
    pts = alt.Chart(titik).mark_circle(size=55).encode(
        x=alt.X("dx:Q", scale=skala), y=alt.Y("dy:Q", scale=skala),
        color=alt.Color("Penanda:N", scale=alt.Scale(domain=["Normal", "Ditandai"], range=["#9db8f2", "#e04a4a"]),
                        legend=alt.Legend(orient="bottom", title=None)),
        opacity=alt.condition(alt.datum.Penanda == "Ditandai", alt.value(0.95), alt.value(0.4)),
        tooltip=["nama:N", "tanggal:N", "jam:N", alt.Tooltip("skor:Q", format=".0f"), "alasan:N"])
    st.altair_chart((garis + pts).properties(width=380, height=380), width="content")   # persegi: lingkaran tetap bulat


def _tab_prediksi(eng, ml):
    pesan = st.session_state.pop("pesan_model", None)
    if pesan:
        {"success": st.success, "warning": st.warning}[pesan[0]](pesan[1])
    aktif = ambil_model(eng)
    st.subheader("Model aktif")
    if aktif is None:
        st.info("Belum ada model. Latih model pertama pada bagian di bawah (butuh data absensi beberapa minggu).")
    else:
        _ringkas_model(aktif)

    with st.expander("Latih model baru"):
        st.caption("Dua kandidat (Logistic Regression dan Random Forest) di-tuning dengan GridSearchCV + cross-validation, "
                   "diuji pada hari-hari terakhir, lalu dibandingkan dengan baseline. Versi baru disimpan, versi lama tidak dihapus.")
        pilihan = {"jika_lebih_baik": "Aktifkan otomatis jika tidak lebih buruk dari model aktif",
                   "selalu": "Selalu aktifkan versi baru", "tidak": "Simpan saja (aktifkan manual)"}
        pol = st.selectbox("Kebijakan aktivasi", list(pilihan), format_func=pilihan.get, key="model_kebijakan")
        if st.button("Latih sekarang", key="model_latih", type="primary"):
            try:
                with st.spinner("Melatih model, mohon tunggu..."):
                    r = latih_dan_simpan(eng, pol)
            except ml.DataKurang as e:
                st.warning(str(e))
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                f1 = r["metrik"]["uji"]["f1"]
                st.session_state["pesan_model"] = ("success", f"Versi {r['id']} tersimpan (F1 uji {f1:.2f}). "
                                                   + ("Versi ini sekarang AKTIF." if r["aktif"] else "Belum diaktifkan: tidak lebih baik dari model aktif."))
                st.rerun()

    versi = daftar_model(eng)
    if versi:
        st.subheader("Riwayat versi")
        tabel = pd.DataFrame([{
            "ID": v["id"], "Dilatih": v["dilatih"], "Data": v["n_data"], "Model": v["metrik"].get("model", "-"),
            "F1 uji": round((v["metrik"].get("uji") or {}).get("f1", 0), 2),
            "ROC-AUC": (v["metrik"].get("uji") or {}).get("roc_auc"), "Aktif": "✅" if v["aktif"] else "",
        } for v in versi])
        st.dataframe(tabel, width="stretch", hide_index=True)
        c1, c2 = st.columns([2, 1])
        pilih = c1.selectbox("Aktifkan versi", [v["id"] for v in versi], format_func=lambda i: f"Versi {i}", key="model_pilih")
        if c2.button("Aktifkan", key="model_aktifkan"):
            try:
                aktifkan_model(eng, pilih)
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                st.rerun()

    if aktif is None:
        return
    st.subheader("Prediksi risiko terlambat")
    sekarang = datetime.now(WITA)
    tgl = st.date_input("Prediksi untuk tanggal", value=_hari_kerja_berikut(sekarang.date()), key="pr_tgl")
    if hari_libur(tgl):
        st.info(pesan_libur(tgl))
        return
    df, _ = prediksi_semua(eng, tgl)
    if df.empty:
        st.info("Belum ada peserta terdaftar.")
        return
    st.caption("Probabilitas dihitung dari riwayat kehadiran SEBELUM tanggal itu. Gunakan untuk pengingat dan pendampingan, "
               "bukan untuk menghukum.")
    tinggi = int((df["risiko"] == "Tinggi").sum())
    st.metric("Peserta berisiko tinggi", tinggi)
    st.dataframe(beri_nomor(pd.DataFrame({
        "Nama": df["nama"], "Divisi": df["divisi"], "Probabilitas terlambat (%)": (df["probabilitas"] * 100).round(0).astype(int),
        "Risiko": df["risiko"], "Faktor": df["faktor"],
    })), width="stretch")


def _tab_monitoring(eng, ml):
    r = ringkasan_pantau(eng)
    if r is None:
        st.info("Belum ada model aktif untuk dipantau. Latih model di tab Prediksi.")
        return
    m = r["model"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Umur model", f"{r['umur_hari']} hari")
    c2.metric("Data baru sejak dilatih", r["n_baru"])
    ev = r["eval_baru"]
    c3.metric("F1 pada data baru", "-" if ev is None else f"{ev['f1']:.2f}",
              delta=None if ev is None else f"{ev['f1'] - (m['metrik'].get('uji') or {}).get('f1', 0):+.2f} vs uji")
    if r["perlu_retrain"]:
        st.warning("Disarankan melatih ulang model:\n\n" + "\n".join(f"- {a}" for a in r["alasan"]))
    else:
        st.success("Model masih sehat: tidak ada tanda drift atau penurunan kinerja.")

    st.subheader(f"Drift fitur (PSI), {r['n_terbaru']} data terbaru vs data latih")
    ikon = {"Stabil": "🟢", "Waspada": "🟠", "Drift": "🔴", "-": "⚪"}
    d = r["drift"]
    st.dataframe(pd.DataFrame({
        "Fitur": d["label"], "PSI": d["psi"].round(3), "Status": [f"{ikon[s]} {s}" for s in d["status"]],
    }), width="stretch", hide_index=True)
    st.caption("PSI < 0,1 stabil, 0,1 sampai 0,25 waspada, >= 0,25 drift (pola data sudah berbeda dari saat model dilatih).")

    st.subheader("Kinerja pada data baru")
    if ev is None:
        st.info("Belum cukup data sesudah tanggal latih (butuh minimal 10 baris yang mencakup hari terlambat dan tidak terlambat).")
    else:
        e = st.columns(5)
        for kolom, nama in zip(e, ["accuracy", "precision", "recall", "f1", "roc_auc"]):
            kolom.metric(nama.upper() if nama == "roc_auc" else nama.capitalize(), "-" if ev[nama] is None else f"{ev[nama]:.2f}")

    st.subheader("Jadwal pelatihan ulang")
    st.caption("Aplikasi ini tidak menjalankan jadwal sendiri. Latih ulang berkala dengan `python jalankan_latih.py` "
               "(cron/GitHub Actions) atau panggil endpoint `POST /v1/model/latih` pada API. "
               "Kebijakan 'aktifkan jika tidak lebih buruk' mencegah model yang lebih jelek menggantikan yang sehat.")


def halaman_model(eng):
    st.header("🤖 Model Prediksi & Deteksi")
    if not admin_login():
        return
    try:
        import ml_absensi as ml
    except ImportError as e:
        st.error("Modul ML belum bisa dimuat. Pastikan `scikit-learn` ada di requirements.txt dan file `ml_absensi.py` "
                 f"berada satu folder dengan app.py, lalu Reboot app. Detail: {e}")
        return
    sub = st.radio("Bagian", ["Mencurigakan", "Prediksi", "Monitoring"], horizontal=True,
                   key="menu_model", label_visibility="collapsed")
    if sub == "Mencurigakan":
        _tab_anomali(eng, ml)
    elif sub == "Prediksi":
        _tab_prediksi(eng, ml)
    else:
        _tab_monitoring(eng, ml)


# ------------------------------------------------------------------
# HALAMAN ADMIN: BUKTI FOTO
# ------------------------------------------------------------------
def halaman_foto(eng):
    st.header("🖼️ Bukti Foto Absensi")
    if not admin_login():
        return

    tgl = st.date_input("Tanggal", value=datetime.now(WITA).date())
    rows = [r for r in get_absensi_hari_foto(eng, tgl.strftime("%Y-%m-%d")) if r["jenis"] not in ("izin", "cuti")]
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
                st.caption("Absen gagal lokasi: foto tidak disimpan." if r["status"] != "berhasil" else "Foto tidak tersedia.")


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def keluar():
    """Logout: hapus peran & data login, kembali ke layar login."""
    for k in ("peran", "nama_login", "admin_ok", "tab"):
        st.session_state.pop(k, None)
    st.rerun()


def mode_pkl():
    """True kalau dibuka lewat aplikasi PKL (alamat berakhiran ?mode=pkl): hanya login peserta, tanpa admin."""
    try:
        return str(st.query_params.get("mode", "")).strip().lower() in ("pkl", "app")
    except Exception:
        return False


def halaman_login(eng, hanya_peserta=False):
    """Layar pertama: pilih peran, lalu login. Desain mengikuti tema Presensi peserta.
    hanya_peserta=True (aplikasi PKL): langsung login peserta, tanpa pilihan Admin."""
    st.markdown(CSS_AUTH_ADMIN, unsafe_allow_html=True)
    st.markdown(
        f"""<div class="auth-head">
<div class="auth-brand"><div class="auth-logo">{ic("fingerprint", 30)}</div>
<div><div class="auth-title">{"Presensi PKL" if hanya_peserta else "Presensi"}</div><div class="auth-bank">Bank Sulselbar • Absensi PKL</div></div></div>
<div class="auth-greet">Selamat Datang 👋</div>
<div class="auth-desc">Silakan masuk untuk melanjutkan</div>
</div>""",
        unsafe_allow_html=True,
    )
    with st.container(key="loginbody"):
        if hanya_peserta:
            peran = "Peserta PKL"
        else:
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
    if getattr(eng, "indeks_unik_ok", False):
        baris.append(("ok", "Integritas: absen ganda dicegah langsung oleh database (unique index)"))
    else:
        n = len(cari_duplikat(eng))
        if n:
            baris.append(("warn", f"Integritas: indeks anti-absen-ganda belum terpasang karena {n} kelompok data ganda lama. "
                                  "Buka Kelola Peserta, lalu bagian Rapikan absensi ganda."))
        else:
            baris.append(("warn", "Integritas: indeks anti-absen-ganda gagal dipasang (cek hak akses database)."))
    aktif, ket = status_wajah()
    if aktif:
        baris.append(("ok", f"Verifikasi wajah: aktif ({ket})"))
    else:
        baris.append(("warn", f"Verifikasi wajah: NONAKTIF, absen masuk/pulang akan ditolak. {ket}"))
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
    st.caption("🔔 Realtime aktif: dashboard diperbarui otomatis saat ada absen/izin baru.")
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
            ["Peserta", "Rekap", "Analitik", "Model", "Foto"],
            horizontal=True,
            key="menu_admin",
            label_visibility="collapsed",
        )
        st.markdown("---")
        if menu == "Peserta":
            halaman_kelola_peserta(eng)
        elif menu == "Rekap":
            halaman_rekap(eng)
        elif menu == "Analitik":
            halaman_analitik(eng)
        elif menu == "Model":
            halaman_model(eng)
        else:
            halaman_foto(eng)


def main():
    pkl = mode_pkl()
    st.set_page_config(page_title="Presensi PKL" if pkl else "Presensi - Bank Sulselbar", page_icon="📍", layout="centered")
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
    elif peran == "admin" and st.session_state.get("admin_ok") and not pkl:
        halaman_admin(eng)
    else:
        halaman_login(eng, hanya_peserta=pkl)


if __name__ == "__main__":
    main()
