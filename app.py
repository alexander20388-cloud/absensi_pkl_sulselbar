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
import logging
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
    Column, ForeignKey, Index, Integer, String, Float, LargeBinary, MetaData, Table, Text, UniqueConstraint,
    create_engine, select, insert, update, delete, func, and_, or_, inspect, text,
)
from sqlalchemy.engine import URL
from sqlalchemy.exc import IntegrityError, OperationalError, SQLAlchemyError
from streamlit_js_eval import get_geolocation

import proyek

# ------------------------------------------------------------------
# KONFIGURASI - SESUAIKAN BAGIAN INI
# ------------------------------------------------------------------
UANG_MAKAN = 15000

# Koordinat asli Gedung A dan B Bank Sulselbar
GEDUNG_COORD = {
    "A": (-5.150070542826675, 119.41568236911743),  # (latitude, longitude)
    "B": (-5.150163347858429, 119.41329137592547),  # (latitude, longitude)
}

RADIUS_METER = 200  # toleransi jarak dari titik gedung (meter)
MAKS_AKURASI_METER = 100  # absen ditolak bila ketelitian GPS perangkat lebih buruk dari ini (BUKAN toleransi jarak; itu RADIUS_METER)

# Pembatasan percobaan login: setelah MAKS_GAGAL_LOGIN kali salah berturut-turut (dalam jendela KUNCI_LOGIN_MENIT menit),
# akun/peran itu dikunci selama KUNCI_LOGIN_MENIT menit.
MAKS_GAGAL_LOGIN = 5
KUNCI_LOGIN_MENIT = 15

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
JAM_MASUK_BUKA, JAM_MASUK_TUTUP = dtime(6, 0), dtime(12, 30)
JAM_PULANG_BUKA, JAM_PULANG_TUTUP = dtime(16, 50), dtime(18, 0)
# Hari libur PKL (0=Senin ... 5=Sabtu, 6=Minggu). Pada hari ini TIDAK bisa absen masuk, pulang, izin, maupun cuti.
HARI_LIBUR = (5, 6)
HARI_LIBUR_KHUSUS = ()       # tanggal 'YYYY-MM-DD' libur nasional/cuti bersama; hanya dipakai rekap agar tidak dihitung 'tidak hadir'
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
    Column("pin_hash", String(200), nullable=False),            # scrypt (baru) atau SHA-256 64 hex (lama; ditingkatkan saat login berhasil)
    Column("no_rekening", String(40)),
    Column("setuju_versi", String(20)),                         # versi pemberitahuan privasi yang disetujui
    Column("setuju_pada", String(19)),                          # YYYY-MM-DD HH:MM:SS (WITA)
)

absensi_t = Table(
    "absensi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("peserta_id", Integer, ForeignKey("peserta.id", ondelete="SET NULL")),  # penghubung resmi ke peserta
    Column("divisi", String(100)),                              # divisi peserta SAAT absen (snapshot; anggaran per divisi tetap benar bila peserta pindah)
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


# Umpan balik peserta (hanya dikumpulkan lewat aplikasi Android). Rating terikat ke peserta supaya hanya bisa diberikan
# SEKALI; ulasan/saran sengaja TIDAK memuat nama atau peserta_id (hanya divisi dan tanggal) dan tidak masuk audit,
# supaya tidak bisa ditelusuri ke orangnya. Tanggal saja (bukan jam) agar tidak bisa dicocokkan lewat waktu kirim.
rating_aplikasi_t = Table(
    "rating_aplikasi", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("peserta_id", Integer, ForeignKey("peserta.id", ondelete="SET NULL"), unique=True),
    Column("bintang", Integer, nullable=False),                 # 1-5
    Column("dibuat", String(10), nullable=False),               # YYYY-MM-DD
)

ulasan_t = Table(
    "ulasan", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("divisi", String(100)),
    Column("kategori", String(20), nullable=False),             # absen / logbook / tampilan / lainnya
    Column("isi", String(500), nullable=False),
    Column("dibuat", String(10), nullable=False, index=True),   # YYYY-MM-DD
    Column("status", String(12), nullable=False, default="baru"),   # baru / diproses / selesai (ditindaklanjuti admin)
    Column("catatan_admin", String(300)),
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


# Jejak perubahan (change log) berantai hash: tiap catatan memuat hash catatan sebelumnya, jadi mengubah atau menghapus
# catatan di tengah memutus rantai (lihat proyek.verifikasi_rantai). hash_prev unik = rantai tidak bisa bercabang.
audit_log_t = Table(
    "audit_log", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("waktu", String(19), nullable=False, index=True),    # YYYY-MM-DD HH:MM:SS (WITA)
    Column("aktor", String(100), nullable=False),
    Column("peran", String(20), nullable=False),                # admin / pembimbing / peserta / sistem
    Column("aksi", String(40), nullable=False, index=True),     # mis. peserta.hapus, logbook.setuju
    Column("objek", String(200), nullable=False, default=""),
    Column("detail", Text, nullable=False, default="{}"),       # JSON
    Column("hash_prev", String(64), nullable=False, unique=True),
    Column("hash", String(64), nullable=False),
)

# Percobaan login gagal per akun/peran (kunci: "peserta:<nama>", "pembimbing:<nama>", "admin").
login_gagal_t = Table(
    "login_gagal", metadata,
    Column("kunci", String(150), primary_key=True),
    Column("gagal", Integer, nullable=False, default=0),
    Column("kunci_sampai", String(19)),                         # YYYY-MM-DD HH:MM:SS (WITA); kosong = tidak terkunci
    Column("terakhir", String(19), nullable=False),
)

# Pembimbing divisi (stakeholder yang menyetujui logbook dan mengelola paket kerja divisinya).
pembimbing_t = Table(
    "pembimbing", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("nama", String(100), nullable=False, unique=True),
    Column("divisi", String(100), nullable=False),
    Column("pin_hash", String(200), nullable=False),            # scrypt dengan garam acak (proyek.hash_pin_kuat)
    Column("dibuat", String(19)),
)

# Paket kerja (WBS) per divisi: dasar jadwal dan progres logbook.
paket_kerja_t = Table(
    "paket_kerja", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("divisi", String(100), nullable=False, index=True),
    Column("kode", String(20), nullable=False),                 # kode WBS: 1, 1.1, 2.3
    Column("nama", String(120), nullable=False),
    Column("mulai", String(10), nullable=False),                # YYYY-MM-DD
    Column("selesai", String(10), nullable=False),
    Column("estimasi_jam", Float, nullable=False),
    Column("ditutup", Integer, nullable=False, default=0),      # 1 = ditutup manual (dianggap selesai 100%)
    Column("dibuat_oleh", String(100)),
    UniqueConstraint("divisi", "kode", name="uq_paket_divisi_kode"),
)

# Logbook kegiatan harian peserta. divisi = snapshot saat dibuat (menentukan pembimbing yang berwenang).
logbook_t = Table(
    "logbook", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("peserta_id", Integer, ForeignKey("peserta.id", ondelete="SET NULL")),
    Column("nama", String(100), nullable=False, index=True),
    Column("divisi", String(100), nullable=False, default="", index=True),
    Column("tanggal", String(10), nullable=False, index=True),
    Column("uraian", Text, nullable=False),
    Column("durasi_jam", Float, nullable=False),
    Column("paket_id", Integer, ForeignKey("paket_kerja.id", ondelete="SET NULL")),
    Column("status", String(12), nullable=False, default="diajukan", index=True),   # diajukan / disetujui / ditolak
    Column("catatan_peninjau", Text),
    Column("ditinjau_oleh", String(100)),
    Column("ditinjau_pada", String(19)),
    Column("dibuat", String(19), nullable=False),
    Column("diubah", String(19)),
)

# Anggaran uang makan per divisi per bulan (Rupiah).
anggaran_t = Table(
    "anggaran", metadata,
    Column("id", Integer, primary_key=True, autoincrement=True),
    Column("tahun", Integer, nullable=False),
    Column("bulan", Integer, nullable=False),
    Column("divisi", String(100), nullable=False),
    Column("jumlah", Integer, nullable=False),
    Column("catatan", String(300)),
    Column("diubah_oleh", String(100)),
    Column("diubah_pada", String(19)),
    UniqueConstraint("tahun", "bulan", "divisi", name="uq_anggaran"),
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
    """Buat tabel yang belum ada, lalu tambah kolom/indeks baru ke database lama (create_all tidak mengubah tabel yang
    sudah ada). Aman diulang. create_all ikut dipanggil di sini karena get_engine di-cache: tanpa ini, tabel baru dari
    pembaruan kode tidak terbentuk selama proses aplikasi belum restart."""
    metadata.create_all(eng)
    perlu = {
        "peserta": [("no_rekening", "VARCHAR(40)"), ("setuju_versi", "VARCHAR(20)"), ("setuju_pada", "VARCHAR(19)")],
        "absensi": [
            ("foto_path", "VARCHAR(200)"), ("alasan", "VARCHAR(300)"),
            ("peserta_id", "INTEGER REFERENCES peserta(id) ON DELETE SET NULL"),
            ("divisi", "VARCHAR(100)"),
        ],
    }
    insp = inspect(eng)
    if eng.dialect.name == "postgresql":      # hash scrypt lebih panjang dari 64 karakter (SQLite tidak membatasi panjang)
        for k in insp.get_columns("peserta"):
            if k["name"] == "pin_hash" and (getattr(k["type"], "length", None) or 0) < 200:
                with eng.begin() as c:
                    c.execute(text("ALTER TABLE peserta ALTER COLUMN pin_hash TYPE VARCHAR(200)"))
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
    with eng.begin() as c:    # snapshot divisi untuk riwayat lama: divisi peserta SAAT INI (satu-satunya data yang tersedia)
        c.execute(text(
            "UPDATE absensi SET divisi = (SELECT p.divisi FROM peserta p WHERE p.nama = absensi.nama) "
            "WHERE divisi IS NULL AND EXISTS (SELECT 1 FROM peserta p WHERE p.nama = absensi.nama)"
        ))
    eng.indeks_unik_ok = _pasang_indeks(eng)


@st.cache_resource
def get_engine(versi_skema=9):
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
def rapikan_duplikat(eng, aktor="admin", peran="admin"):
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
        if hapus:
            _audit(c, aktor, peran, "absensi.rapikan_duplikat", "absensi", {"dihapus": hapus})
    hapus_foto_storage([p for p in paths_hapus if p and p not in paths_simpan])   # jangan hapus foto milik catatan yang dipertahankan
    eng.indeks_unik_ok = _pasang_indeks(eng)
    return hapus


def foto_dipakai(eng, path):
    """True kalau ada baris absensi yang memakai path foto itu."""
    with eng.connect() as c:
        return c.execute(select(func.count()).select_from(absensi_t).where(absensi_t.c.foto_path == path)).scalar() > 0


_KUNCI_ADVISORY_AUDIT = 727001


def _audit(c, aktor, peran, aksi, objek, detail=None):
    """Tambah satu catatan ke rantai audit DI DALAM transaksi yang sedang berjalan (c), jadi perubahan data dan
    jejaknya berhasil atau gagal bersama-sama. Di Postgres penulisan audit diserialkan lewat advisory lock supaya
    rantai tidak bercabang; di SQLite penguncian tulis bawaannya sudah menyerialkan. Pemanggil sebaiknya
    memanggil ini SETELAH perintah perubahan datanya."""
    if c.dialect.name == "postgresql":
        c.execute(text("SELECT pg_advisory_xact_lock(:k)"), {"k": _KUNCI_ADVISORY_AUDIT})
    prev = c.execute(select(audit_log_t.c.hash).order_by(audit_log_t.c.id.desc()).limit(1)).scalar() or proyek.GENESIS
    waktu = datetime.now(WITA).strftime("%Y-%m-%d %H:%M:%S")
    aktor, peran, aksi, objek = str(aktor)[:100], str(peran)[:20], str(aksi)[:40], str(objek)[:200]
    detail_s = proyek.detail_json(detail)
    c.execute(insert(audit_log_t).values(
        waktu=waktu, aktor=aktor, peran=peran, aksi=aksi, objek=objek, detail=detail_s, hash_prev=prev,
        hash=proyek.hitung_hash(prev, waktu, aktor, peran, aksi, objek, detail_s),
    ))


def hash_pin(nama, pin):
    """Hash LAMA (SHA-256 bergaram tetap). Hanya untuk memverifikasi PIN lama; PIN baru memakai proyek.hash_pin_kuat."""
    garam = secret("salt", "pkl-sulselbar")
    return hashlib.sha256(f"{nama.strip().lower()}|{pin}|{garam}".encode()).hexdigest()


# ---------- pembatasan percobaan login ----------
def _fmt_waktu(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def sisa_kunci_login(eng, kunci, now=None):
    """Sisa detik akun/peran itu terkunci (0 = tidak terkunci)."""
    now = now or datetime.now(WITA)
    with eng.connect() as c:
        v = c.execute(select(login_gagal_t.c.kunci_sampai).where(login_gagal_t.c.kunci == kunci)).scalar()
    if not v:
        return 0
    sampai = datetime.strptime(v, "%Y-%m-%d %H:%M:%S").replace(tzinfo=WITA)
    return max(0, math.ceil((sampai - now).total_seconds()))


@coba_ulang
def catat_login_gagal(eng, kunci, now=None):
    """Catat satu percobaan gagal. Mengembalikan sisa detik terkunci (0 bila belum mencapai batas)."""
    now = now or datetime.now(WITA)
    s = _fmt_waktu(now)
    jendela_awal = _fmt_waktu(now - timedelta(minutes=KUNCI_LOGIN_MENIT))
    try:
        with eng.begin() as c:
            row = c.execute(select(login_gagal_t.c.gagal, login_gagal_t.c.kunci_sampai, login_gagal_t.c.terakhir)
                            .where(login_gagal_t.c.kunci == kunci)).first()
            gagal = 0
            if row is not None and row[2] >= jendela_awal and not (row[1] and row[1] <= s):
                gagal = int(row[0])                       # hitungan lama dibuang bila sudah di luar jendela atau masa kunci habis
            gagal += 1
            sampai = _fmt_waktu(now + timedelta(minutes=KUNCI_LOGIN_MENIT)) if gagal >= MAKS_GAGAL_LOGIN else None
            if row is None:
                c.execute(insert(login_gagal_t).values(kunci=kunci, gagal=gagal, kunci_sampai=sampai, terakhir=s))
            else:
                c.execute(update(login_gagal_t).where(login_gagal_t.c.kunci == kunci)
                          .values(gagal=gagal, kunci_sampai=sampai, terakhir=s))
            if sampai:
                _audit(c, "sistem", "sistem", "login.terkunci", kunci, {"gagal": gagal, "menit": KUNCI_LOGIN_MENIT})
    except IntegrityError:                                # dua percobaan bersamaan menyisip baris yang sama: satu cukup
        pass
    return sisa_kunci_login(eng, kunci, now)


@coba_ulang
def reset_login_gagal(eng, kunci):
    with eng.begin() as c:
        c.execute(delete(login_gagal_t).where(login_gagal_t.c.kunci == kunci))


def proses_login(eng, kunci, benar_fn, now=None):
    """Satu pintu untuk semua login. benar_fn() dipanggil HANYA bila akun tidak sedang terkunci.
    Mengembalikan (status, menit): status 'ok' | 'salah' | 'terkunci'; menit = sisa menit kunci (dibulatkan ke atas)."""
    sisa = sisa_kunci_login(eng, kunci, now)
    if sisa > 0:
        return "terkunci", math.ceil(sisa / 60)
    if benar_fn():
        reset_login_gagal(eng, kunci)
        return "ok", 0
    sisa = catat_login_gagal(eng, kunci, now)
    return ("terkunci", math.ceil(sisa / 60)) if sisa > 0 else ("salah", 0)


def pesan_terkunci(menit):
    return f"Terlalu banyak percobaan gagal. Akun dikunci sementara, coba lagi dalam {menit} menit."


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
def add_peserta(eng, nama, gedung, divisi, pin, no_rekening=None, aktor="admin", peran="admin"):
    with eng.begin() as c:
        c.execute(insert(peserta_t).values(
            nama=nama.strip(), gedung=gedung, divisi=divisi.strip(), pin_hash=proyek.hash_pin_kuat(pin),
            no_rekening=(no_rekening or None),
        ))
        _audit(c, aktor, peran, "peserta.tambah", f"peserta:{nama.strip()}",
               {"divisi": divisi.strip(), "gedung": gedung, "rekening_diisi": bool(no_rekening)})


def _akhir_rekening(nomor):
    """Hanya 4 angka terakhir yang masuk jejak audit (nomor rekening adalah data pribadi)."""
    angka = re.sub(r"\D", "", nomor or "")
    return angka[-4:] if angka else None


@coba_ulang
def set_rekening(eng, nama, no_rekening, aktor="admin", peran="admin"):
    with eng.begin() as c:
        lama = c.execute(select(peserta_t.c.no_rekening).where(peserta_t.c.nama == nama)).first()
        r = c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(no_rekening=(no_rekening or None)))
        if r.rowcount:
            _audit(c, aktor, peran, "peserta.ubah_rekening", f"peserta:{nama}",
                   {"dari_akhir": _akhir_rekening(lama[0]), "ke_akhir": _akhir_rekening(no_rekening)})


def validasi_rekening(teks):
    """Kembalikan (nomor_rapi, pesan_error). Boleh angka dan tanda hubung, mis. 130-210-000554616-9."""
    rapi = re.sub(r"\s+", "", teks or "")
    if not rapi:
        return "", "Nomor rekening wajib diisi."
    if not re.fullmatch(r"[0-9\-]+", rapi) or not (8 <= len(re.sub(r"\D", "", rapi)) <= 25):
        return rapi, "Nomor rekening hanya boleh angka dan tanda hubung (-), minimal 8 angka."
    return rapi, ""


@coba_ulang
def reset_pin(eng, nama, pin_baru, aktor="admin", peran="admin"):
    with eng.begin() as c:
        r = c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(pin_hash=proyek.hash_pin_kuat(pin_baru)))
        if r.rowcount:
            _audit(c, aktor, peran, "peserta.reset_pin", f"peserta:{nama}")           # PIN-nya sendiri tidak pernah dicatat


@coba_ulang
def set_divisi(eng, nama, divisi, gedung, aktor="admin", peran="admin"):
    with eng.begin() as c:
        lama = c.execute(select(peserta_t.c.divisi, peserta_t.c.gedung).where(peserta_t.c.nama == nama)).first()
        r = c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(divisi=divisi.strip(), gedung=gedung))
        if r.rowcount:
            _audit(c, aktor, peran, "peserta.ubah_divisi", f"peserta:{nama}",
                   {"dari": lama[0], "ke": divisi.strip(), "gedung_dari": lama[1], "gedung_ke": gedung})


@coba_ulang
def hapus_peserta(eng, nama, dengan_absensi=False, aktor="admin", peran="admin"):
    """Hapus peserta. Kalau dengan_absensi=True, riwayat absensi dan logbooknya ikut dihapus."""
    paths, n_absen, n_log = [], 0, 0
    with eng.begin() as c:
        # riwayat yang dipertahankan dilepas dari peserta (SQLite tidak menjalankan ON DELETE SET NULL)
        c.execute(update(absensi_t).where(absensi_t.c.nama == nama).values(peserta_id=None))
        c.execute(update(logbook_t).where(logbook_t.c.nama == nama).values(peserta_id=None))
        pid = c.execute(select(peserta_t.c.id).where(peserta_t.c.nama == nama)).scalar()
        if pid is not None:                                     # rating tetap sebagai statistik, tetapi dilepas dari peserta
            c.execute(update(rating_aplikasi_t).where(rating_aplikasi_t.c.peserta_id == pid).values(peserta_id=None))
        ada = c.execute(delete(peserta_t).where(peserta_t.c.nama == nama)).rowcount
        if dengan_absensi:
            paths = [r[0] for r in c.execute(select(absensi_t.c.foto_path).where(absensi_t.c.nama == nama))]
            n_absen = c.execute(delete(absensi_t).where(absensi_t.c.nama == nama)).rowcount
            n_log = c.execute(delete(logbook_t).where(logbook_t.c.nama == nama)).rowcount
        if ada or n_absen or n_log:
            _audit(c, aktor, peran, "peserta.hapus", f"peserta:{nama}",
                   {"dengan_riwayat": bool(dengan_absensi), "absensi_dihapus": n_absen, "logbook_dihapus": n_log})
    hapus_foto_storage(paths)


def daftar_nama_absensi(eng):
    """Semua nama yang punya riwayat absensi (termasuk peserta yang sudah dihapus)."""
    q = select(absensi_t.c.nama).distinct().order_by(absensi_t.c.nama)
    with eng.connect() as c:
        return [r[0] for r in c.execute(q).all()]


@coba_ulang
def hapus_absensi(eng, nama=None, aktor="admin", peran="admin"):
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
        if jumlah:
            _audit(c, aktor, peran, "absensi.hapus", f"absensi:{nama or 'SEMUA'}", {"jumlah": jumlah})
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
        if nilai.get("peserta_id") is None or nilai.get("divisi") is None:
            p = c.execute(select(peserta_t.c.id, peserta_t.c.divisi).where(peserta_t.c.nama == nilai.get("nama"))).first()
            if p is not None:
                if nilai.get("peserta_id") is None:
                    nilai["peserta_id"] = p[0]
                if nilai.get("divisi") is None:
                    nilai["divisi"] = p[1]                  # snapshot: anggaran per divisi tetap benar bila peserta pindah
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
    # Angka dihitung langsung (bukan rumus =SUM): file dari openpyxl tidak menyimpan hasil rumus, sehingga di pratinjau
    # HP / Google Drive / Protected View / viewer tanpa kalkulasi, total tampil kosong.
    ws[f"G{jumlah}"].value = sum(int(ws[f"G{rr}"].value) for rr in range(8, jumlah)
                                 if isinstance(ws[f"G{rr}"].value, (int, float)))

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
    if url_database() is not None:      # database cloud apa pun (database_url ATAU db_host dst.): jangan pernah pakai password bawaan
        return None
    return "admin123"  # mode lokal (tanpa database cloud)


def cek_pin(eng, nama, pin):
    """PIN peserta. Hash scrypt (baru) dan SHA-256 (lama) sama-sama diterima; yang lama otomatis ditingkatkan ke scrypt
    saat login berhasil, jadi tidak ada peserta yang perlu direset."""
    peserta = get_peserta(eng, nama)
    if peserta is None:
        proyek.cek_pin_kuat(pin, _hash_dummy())            # waktu proses sama, supaya nama yang ada/tidak ada tak bisa dibedakan
        return False
    tersimpan = peserta["pin_hash"] or ""
    if tersimpan.startswith("scrypt$"):
        return proyek.cek_pin_kuat(pin, tersimpan)
    if not hmac.compare_digest(tersimpan.encode(), hash_pin(nama, pin).encode()):
        return False
    try:                                                     # tingkatkan ke scrypt; gagal menyimpan tidak boleh menggagalkan login
        with eng.begin() as c:
            c.execute(update(peserta_t).where(and_(peserta_t.c.nama == nama, peserta_t.c.pin_hash == tersimpan))
                      .values(pin_hash=proyek.hash_pin_kuat(pin)))
    except SQLAlchemyError:
        pass
    return True


def _sama(a, b):
    """Perbandingan waktu-konstan untuk teks apa pun (compare_digest pada str melempar TypeError bila ada huruf non-ASCII)."""
    return hmac.compare_digest(str(a).encode(), str(b).encode())


def admin_login():
    """Penjaga halaman admin. Login normal lewat layar login di awal; ini cadangan
    kalau halaman admin terbuka tanpa lewat login."""
    if st.session_state.get("admin_ok"):
        return True

    pw_benar = password_admin()
    if pw_benar is None:
        st.error("admin_password belum diatur di Secrets. Hubungi pengelola aplikasi.")
        return False
    if not secret("admin_password") and url_database() is None:
        st.info("Mode lokal: password admin default = admin123")

    pw = st.text_input("Password admin", type="password")
    if st.button("Masuk sebagai admin"):
        status, menit = proses_login(get_engine(), "admin", lambda: _sama(pw, str(pw_benar)))
        if status == "ok":
            st.session_state.admin_ok = True
            st.rerun()
        elif status == "terkunci":
            st.error(pesan_terkunci(menit))
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
    "menu": '<line x1="4" x2="20" y1="12" y2="12"/><line x1="4" x2="20" y1="6" y2="6"/><line x1="4" x2="20" y1="18" y2="18"/>',
    "x": '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
}


# Logo Bank Sulselbar (PNG transparan, base64) untuk layar login.
LOGO_BANK_B64 = "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"


def ic(nama, size=24, stroke=2):
    return _svg(ICON[nama], size, stroke)


CSS_PRESENSI = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"] {
  background: #edf3f8 !important; font-family: 'Inter', system-ui, sans-serif;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], footer { display: none !important; }
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 430px !important; padding: 0 0 130px 0 !important; position: relative;
  background: #edf3f8; min-height: 100vh;
}
[data-testid="stMainBlockContainer"] > [data-testid="stVerticalBlock"] { gap: 0 !important; }

/* ---------- header biru ---------- */
.pr-head {
  background: linear-gradient(135deg, #0a5a8f 0%, #063f6b 100%);
  color: #fff; padding: 30px 22px 76px; border-radius: 0 0 0 0;
  position: relative; overflow: hidden;
}
.pr-head::after {
  content: ""; position: absolute; left: -10%; right: -10%; bottom: -34px; height: 70px;
  background: #edf3f8; border-radius: 50% 50% 0 0 / 100% 100% 0 0;
}
.pr-top { display: flex; align-items: center; gap: 14px; }
.pr-logo { width: 46px; height: 46px; border-radius: 14px; background: rgba(255,255,255,.18);
  display: flex; align-items: center; justify-content: center; }
.pr-title { font-size: 25px; font-weight: 600; letter-spacing: .2px; flex: 1; }
.pr-avatar { width: 46px; height: 46px; border-radius: 12px; background: rgba(255,255,255,.9);
  color: #0a5a8f; font-weight: 700; font-size: 18px; display: flex; align-items: center; justify-content: center; }
.pr-out { width: 52px; height: 52px; border-radius: 50%; background: rgba(255,255,255,.18);
  display: flex; align-items: center; justify-content: center; }
.pr-hello { margin-top: 26px; font-size: 16px; color: rgba(255,255,255,.82); }
.pr-name { font-size: 20px; font-weight: 600; margin-top: 4px; line-height: 1.25; }
.pr-sub { font-size: 13px; color: rgba(255,255,255,.75); margin-top: 3px; }

/* ---------- kartu presensi ---------- */
.pr-card { background: #fff; border-radius: 22px; margin: -46px 22px 0; padding: 14px 14px 14px;
  box-shadow: 0 8px 28px rgba(6,63,107,.10); position: relative; z-index: 2; }
.pr-date { display: flex; align-items: center; gap: 9px; color: #21496b; font-size: 14px; font-weight: 500; }
.pr-date svg { color: #3f6785; }
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
.t-blue { background: #e4f0f8; color: #0a5a8f; }
.pr-info { margin-top: 10px; background: #e4f0f8; border-radius: 12px; padding: 9px 12px; color: #0a5a8f; }
.pr-info .h { display: flex; align-items: center; gap: 7px; font-size: 13px; font-weight: 600; }
.pr-info .s { font-size: 12.5px; margin-top: 2px; color: #1f78b4; line-height: 1.3; }

/* ---------- tema Bank Sulselbar: lengkung hijau lime di header ---------- */
.pr-head, .pr-page-head { position: relative; isolation: isolate; overflow: hidden; }
.pr-head::before, .pr-page-head::before {
  content: ""; position: absolute; right: 0; bottom: 0; width: 74%; height: 100%; z-index: -1; pointer-events: none;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 300 160' preserveAspectRatio='none'%3E%3Cpath d='M0 160C90 160 205 118 300 8V160Z' fill='%238cc63f'/%3E%3Cpath d='M90 160C170 150 240 110 300 56V160Z' fill='%23b5dd6a' fill-opacity='.6'/%3E%3C/svg%3E") right bottom / 100% 100% no-repeat;
}

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
  box-shadow: 0 -6px 24px rgba(6,63,107,.10); }
.st-key-bottomnav [data-testid="stHorizontalBlock"] { gap: 0 !important; align-items: flex-end; }
.nv { text-align: center; color: #8da3b5; font-size: 12px; padding-top: 6px; min-height: 52px; }
.nv svg { display: block; margin: 0 auto 3px; }
.nv.on { color: var(--hijau-tua, #3f8f1e); font-weight: 600; background: #e9f5d6; border-radius: 14px; }
.nv-fab { width: 64px; height: 64px; border-radius: 50%; background: linear-gradient(135deg,#0a5a8f,#063f6b);
  color: #fff; margin: -34px auto 0; display: flex; align-items: center; justify-content: center;
  box-shadow: 0 8px 20px rgba(6,63,107,.45); border: 5px solid #fff; box-sizing: content-box; }
.st-key-ov_fab { height: 64px; }

/* ---------- halaman dalam ---------- */
.pr-page-head { background: linear-gradient(135deg, #0a5a8f 0%, #063f6b 100%); color: #fff;
  padding: 30px 22px 56px; font-size: 22px; font-weight: 600; border-radius: 0 0 28px 28px; }
.pr-page-head small { display: block; font-size: 13px; font-weight: 400; opacity: .8; margin-top: 4px; }
.st-key-pagebody { width: calc(100% - 36px) !important; margin: -34px 18px 0; background: #fff; border-radius: 24px; padding: 18px 16px;
  box-shadow: 0 8px 28px rgba(6,63,107,.10); position: relative; z-index: 2; }
.pr-row { display: flex; align-items: center; gap: 12px; padding: 12px 0; border-bottom: 1px solid #e6eef4; }
.pr-row:last-child { border-bottom: 0; }
.pr-row .d { font-weight: 600; color: #0b3a5e; font-size: 15px; }
.pr-row .m { font-size: 13px; color: #5f7b92; margin-top: 2px; }
.pr-row .r { margin-left: auto; text-align: right; font-size: 13px; font-weight: 600; }
.ok { color: #14a068; } .bad { color: #e04a4a; } .izin { color: #c58a05; } .cuti { color: #0a5a8f; }
.pr-kv { display: flex; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid #e6eef4; font-size: 15px; }
.pr-kv:last-child { border-bottom: 0; }
.pr-kv span:first-child { color: #5f7b92; } .pr-kv span:last-child { font-weight: 600; color: #0b3a5e; }
.pr-stat { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 8px; }
.pr-stat div { background: #e4f0f8; border-radius: 16px; padding: 14px; color: #0a5a8f; }
.pr-stat b { display: block; font-size: 22px; margin-top: 4px; }
.pr-stat span { font-size: 13px; }
.st-key-pagebody [data-testid="stButton"] button[kind="primary"] { background: #0a5a8f; border: 0; border-radius: 14px; }

</style>
"""



CSS_TEMA = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root {
  --biru: #0a5a8f; --ungu: #063f6b; --bg: #edf3f8; --tint: #e4f0f8;
  --navy: #0b3a5e; --navy2: #3f6785; --muted: #5f7b92; --garis: #d3e2ed;
  --hijau: #8cc63f; --hijau2: #6bb02f; --hijau-tua: #3f8f1e; --hijau-muda: #e9f5d6;
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
  background: #f3f8fb !important; color: var(--navy) !important; border-color: var(--garis) !important;
  border-radius: 14px !important;
}
.st-key-loginbody input, .st-key-adminbody input { background: transparent !important; color: var(--navy) !important; }
.st-key-loginbody input::placeholder, .st-key-adminbody input::placeholder { color: var(--muted) !important; opacity: 1; }
.st-key-loginbody [data-baseweb="select"] > div:focus-within, .st-key-adminbody [data-baseweb="select"] > div:focus-within,
.st-key-loginbody [data-baseweb="base-input"]:focus-within, .st-key-adminbody [data-baseweb="base-input"]:focus-within {
  border-color: var(--biru) !important; box-shadow: 0 0 0 3px rgba(10,90,143,.18) !important;
}

/* radio -> tombol pil */
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stRadio"] [role="radiogroup"] > div { flex: 1 1 0; min-width: 0; }
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stRadioOption"] { width: 100%; box-sizing: border-box; }
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
  background: linear-gradient(135deg, var(--hijau), var(--hijau2)); box-shadow: 0 4px 12px rgba(107,176,47,.35);
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"][data-selected="true"] p { color: #fff; }

/* tombol */
:is(.st-key-loginbody, .st-key-adminbody) button[kind="primary"],
:is(.st-key-loginbody, .st-key-adminbody) button[kind="primaryFormSubmit"] {
  background: linear-gradient(135deg, var(--biru), var(--ungu)) !important; border: 0 !important; color: #fff !important;
  border-radius: 14px !important; min-height: 48px; font-weight: 600; box-shadow: 0 6px 16px rgba(6,63,107,.3);
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
.st-key-adminbody [data-testid="stMetricLabel"] * { color: #0a5a8f !important; }
.st-key-adminbody [data-testid="stMetricValue"] * { color: var(--navy) !important; font-weight: 700; font-family: 'Inter', system-ui, sans-serif !important; }
.st-key-adminbody [data-testid="stDataFrame"], .st-key-adminbody [data-testid="stTable"] {
  border: 1px solid var(--garis); border-radius: 16px; overflow: hidden;
}
.st-key-adminbody [data-testid="stExpander"] details { background: #f3f8fb; border: 1px solid var(--garis); border-radius: 16px; }
.st-key-adminbody [data-testid="stExpander"] summary * { color: var(--navy); font-weight: 600; }
.st-key-adminbody [data-testid="stAlert"], .st-key-loginbody [data-testid="stAlert"] { border-radius: 14px; }
.st-key-adminbody [data-testid="stImage"] img { border-radius: 14px; }
.st-key-adminbody [data-testid="stForm"] {
  border: 1px solid var(--garis); border-radius: 18px; background: #f8fbfd; padding: 16px;
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
  background: linear-gradient(135deg, var(--hijau), var(--hijau2)); box-shadow: 0 4px 12px rgba(107,176,47,.35);
}
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] p { color: #fff; }
.st-key-pagebody [data-baseweb="textarea"], .st-key-pagebody [data-baseweb="base-input"] {
  background: #f3f8fb !important; border-color: var(--garis) !important; border-radius: 14px !important;
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
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 460px !important; background: linear-gradient(180deg, #0a5a8f 0%, #063f6b 100%) !important; overflow: hidden;
}
[data-testid="stMainBlockContainer"]::after {
  content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 200px; z-index: 0; pointer-events: none;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 200' preserveAspectRatio='none'%3E%3Cpath d='M0 108C110 136 270 68 400 26V60C270 100 110 158 0 146Z' fill='%238cc63f'/%3E%3Cpath d='M0 150C110 162 270 104 400 64' fill='none' stroke='white' stroke-opacity='.85' stroke-width='1.6'/%3E%3C/svg%3E") bottom / 100% 100% no-repeat;
}
.auth-head { background: #fff; padding: 28px 24px 52px; position: relative; z-index: 1; }
.auth-logo-img { display: block; width: 150px; max-width: 60%; height: auto; margin: 0 auto 24px; }
.auth-brand { display: flex; align-items: center; gap: 12px; }
.auth-logo { width: 50px; height: 50px; border-radius: 15px; background: var(--tint); color: var(--biru);
  display: flex; align-items: center; justify-content: center; flex: 0 0 50px; }
.auth-title { font-size: 25px; font-weight: 700; letter-spacing: .2px; color: var(--navy); line-height: 1.1; }
.auth-bank { font-size: 11.5px; color: var(--muted); margin-top: 3px; }
.auth-greet { margin-top: 22px; font-size: 22px; font-weight: 700; color: var(--navy); }
.auth-desc { margin-top: 4px; font-size: 13px; color: var(--muted); }
.st-key-loginbody {
  width: calc(100% - 28px) !important; margin: -34px 14px 0; background: #fff; border-radius: 24px;
  padding: 18px 18px 22px; position: relative; z-index: 2; box-shadow: 0 12px 30px rgba(3,30,52,.28);
}
.st-key-loginbody [data-testid="stElementContainer"]:has([data-testid="stRadio"]) { width: 100% !important; }
.auth-foot { position: fixed; left: 0; right: 0; bottom: 14px; text-align: center; color: rgba(255,255,255,.92); font-size: 12px; z-index: 1; }
</style>
"""

CSS_ADMIN = """
<style>
[data-testid="stMainBlockContainer"], .block-container { max-width: 1100px !important; }
.admin-head {
  background: linear-gradient(135deg, #0a5a8f 0%, #063f6b 100%); color: #fff;
  padding: 32px 28px 84px; position: relative; overflow: hidden; isolation: isolate;
}
.admin-head::before {
  content: ""; position: absolute; right: 0; bottom: 0; width: 60%; height: 100%; z-index: -1; pointer-events: none;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 300 160' preserveAspectRatio='none'%3E%3Cpath d='M0 160C90 160 205 118 300 8V160Z' fill='%238cc63f'/%3E%3Cpath d='M90 160C170 150 240 110 300 56V160Z' fill='%23b5dd6a' fill-opacity='.6'/%3E%3C/svg%3E") right bottom / 100% 100% no-repeat;
}
.admin-head::after {
  content: ""; position: absolute; left: -10%; right: -10%; bottom: -34px; height: 70px;
  background: #edf3f8; border-radius: 50% 50% 0 0 / 100% 100% 0 0;
}
.admin-top { display: flex; align-items: center; gap: 14px; position: relative; z-index: 1; }
.admin-logo { width: 50px; height: 50px; border-radius: 15px; background: rgba(255,255,255,.2);
  display: flex; align-items: center; justify-content: center; color: #fff; }
.admin-title { font-size: 25px; font-weight: 700; color: #fff; flex: 1; }
.admin-sub { font-size: 13px; color: rgba(255,255,255,.82); margin-top: 2px; font-weight: 400; }
.admin-out { width: 50px; height: 50px; border-radius: 50%; background: rgba(255,255,255,.2);
  display: flex; align-items: center; justify-content: center; color: #fff; }
.st-key-adminlogout { position: absolute !important; top: 32px; right: 28px; width: 50px; height: 50px; z-index: 10; }
.st-key-adminmenu { position: absolute !important; top: 32px; left: 28px; width: 50px; height: 50px; z-index: 10; }
.st-key-adminmenu [data-testid="stElementContainer"], .st-key-adminmenu [data-testid="stButton"] { height: 50px; width: 50px !important; }
.st-key-adminmenu button { width: 50px !important; height: 50px !important; opacity: 0; cursor: pointer; padding: 0; }

/* ---------- menu hamburger (drawer) ---------- */
.st-key-drawerbg { position: fixed !important; inset: 0; z-index: 9998; background: rgba(4,32,58,.55); animation: drFade .18s ease-out; }
.st-key-drawerbg [data-testid="stElementContainer"] { position: absolute; inset: 0; }
.st-key-drawerbg [data-testid="stButton"], .st-key-drawerbg .stButton { height: 100%; width: 100% !important; }
.st-key-drawerbg button { width: 100% !important; height: 100% !important; opacity: 0; cursor: pointer; min-height: 0; }
.st-key-drawer {
  position: fixed !important; top: 0; left: 0; bottom: 0; width: min(290px, 84vw) !important; z-index: 9999; background: #fff;
  overflow-y: auto; gap: 4px !important; padding: 0 0 18px; box-shadow: 8px 0 30px rgba(3,30,52,.35); animation: drIn .2s ease-out;
}
.st-key-drawer [data-testid="stElementContainer"]:has(button) { padding: 0 12px; }
.st-key-drawer [data-testid="stElementContainer"]:has(.dr-head) { padding: 0; }
.dr-head { background: linear-gradient(160deg, #0a5a8f 0%, #063f6b 100%); color: #fff; padding: 20px 20px 18px; position: relative; overflow: hidden; isolation: isolate; margin-bottom: 8px; }
.dr-head::before {
  content: ""; position: absolute; right: 0; bottom: 0; width: 80%; height: 70%; z-index: -1; pointer-events: none;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 300 160' preserveAspectRatio='none'%3E%3Cpath d='M0 160C90 160 205 118 300 8V160Z' fill='%238cc63f'/%3E%3Cpath d='M90 160C170 150 240 110 300 56V160Z' fill='%23b5dd6a' fill-opacity='.6'/%3E%3C/svg%3E") right bottom / 100% 100% no-repeat;
}
.dr-top { display: flex; align-items: center; justify-content: space-between; height: 26px; color: #fff; }
.dr-t { font-size: 19px; font-weight: 700; margin-top: 18px; color: #fff; }
.dr-s { font-size: 12px; color: rgba(255,255,255,.85); margin-top: 3px; }
.dr-m { font-size: 10.5px; font-weight: 700; letter-spacing: .8px; color: rgba(255,255,255,.9); margin-top: 18px; }
.st-key-drawerx { position: absolute !important; top: 16px; right: 16px; width: 36px; height: 36px; z-index: 5; }
.st-key-drawerx [data-testid="stElementContainer"] { padding: 0 !important; height: 36px; width: 36px; }
.st-key-drawerx [data-testid="stButton"] { height: 36px; width: 36px !important; }
.st-key-drawerx button { width: 36px !important; height: 36px !important; opacity: 0; cursor: pointer; padding: 0; min-height: 0; }
.st-key-drawer [class*="st-key-drw_"] button {
  position: relative; width: 100%; justify-content: flex-start !important; text-align: left; padding: 10px 14px; min-height: 46px;
  background: #fff; border: 0; border-radius: 12px; box-shadow: none; color: var(--navy);
}
.st-key-drawer [class*="st-key-drw_"] button:hover { background: #f3f8fb; }
.st-key-drawer [class*="st-key-drw_"] button > div { width: 100%; padding-right: 22px; }
.st-key-drawer [class*="st-key-drw_"] button > div > span { display: flex; align-items: center; justify-content: flex-start; gap: 12px; width: 100%; }
.st-key-drawer [class*="st-key-drw_"] button > div > span > span:last-child { flex: 1; text-align: left; }
.st-key-drawer [class*="st-key-drw_"] button [data-testid="stMarkdownContainer"] { text-align: left; }
.st-key-drawer [class*="st-key-drw_"] button p { text-align: left; font-size: 15px; font-weight: 600; color: inherit; }
.st-key-drawer [class*="st-key-drw_"] button::after { content: "›"; position: absolute; right: 14px; top: 50%; transform: translateY(-52%); font-size: 20px; color: #9bb3c4; }
.st-key-drawer .st-key-drw_keluar button { color: #d64545; }
.st-key-drawer .st-key-drw_keluar button::after { content: none; }
.dr-garis { border-top: 1px solid var(--garis); margin: 8px 20px 4px; }
.dr-foot { text-align: center; font-size: 11px; color: var(--muted); padding: 10px 0 0; }
@keyframes drIn { from { transform: translateX(-100%); } to { transform: none; } }
@keyframes drFade { from { opacity: 0; } to { opacity: 1; } }
.st-key-adminlogout [data-testid="stElementContainer"], .st-key-adminlogout [data-testid="stButton"] { height: 50px; width: 50px !important; }
.st-key-adminlogout button { width: 50px !important; height: 50px !important; opacity: 0; cursor: pointer; padding: 0; }
.st-box { background: #f3f8fb; border: 1px solid #d3e2ed; border-radius: 14px; padding: 10px 14px; font-size: 13px; line-height: 1.7; color: #0b3a5e; margin-bottom: 8px; }
.st-key-rt_hidden { position: absolute; left: -9999px; width: 1px; height: 1px; overflow: hidden; }
.st-key-adminbody {
  width: calc(100% - 36px) !important; margin: -44px 18px 0; background: #fff; border-radius: 26px;
  padding: 22px 22px 26px; position: relative; z-index: 2; box-shadow: 0 8px 28px rgba(6,63,107,.12);
}
@media (max-width: 600px) {
  .admin-head { padding: 28px 18px 80px; }
  .st-key-adminlogout { right: 18px; top: 28px; }
  .st-key-adminmenu { left: 18px; top: 28px; }
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
<div class="pr-avatar">{esc(inisial(nama))}</div>
<div class="pr-out">{ic("logout", 26)}</div>
</div>
<div class="pr-hello">{salam(now.hour)} 👋</div>
<div class="pr-name">{esc(nama)}</div>
<div class="pr-sub">{esc(peserta["divisi"])} • Gedung {esc(peserta["gedung"])}</div>
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
    _kartu_logbook_beranda(eng, nama)


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

    if akurasi is not None and akurasi > MAKS_AKURASI_METER:   # ketelitian GPS buruk (mis. lokasi dari IP/menara); BUKAN soal jarak ke gedung
        return _hasil("error", f"📡 Akurasi lokasi terlalu rendah (±{akurasi:.0f} m; maksimal ±{MAKS_AKURASI_METER} m). "
                               "Aktifkan GPS presisi tinggi atau pindah ke tempat terbuka, tunggu sebentar, lalu coba lagi.",
                      alasan="akurasi")
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
    if coords and (coords.get("accuracy") or 0) > MAKS_AKURASI_METER:
        st.warning(f"📡 Akurasi lokasi masih rendah (±{coords['accuracy']:.0f} m, perlu ±{MAKS_AKURASI_METER} m atau lebih baik). "
                   "Tunggu GPS menyesuaikan atau pindah ke tempat terbuka.")
    elif coords:
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


def tab_profil(eng, nama, peserta):
    header_halaman("Profil", "Data peserta PKL")
    with st.container(key="pagebody"):
        st.markdown(
            f'<div class="pr-kv"><span>Nama</span><span>{esc(nama)}</span></div>'
            f'<div class="pr-kv"><span>Divisi</span><span>{esc(peserta["divisi"])}</span></div>'
            f'<div class="pr-kv"><span>Gedung</span><span>Gedung {esc(peserta["gedung"])}</span></div>'
            f'<div class="pr-kv"><span>Status</span><span class="ok">Aktif</span></div>'
            f'<div class="pr-kv"><span>Persetujuan privasi</span><span>{esc(str(peserta["setuju_pada"] or "-")[:10])}</span></div>',
            unsafe_allow_html=True,
        )
        with st.expander("Pemberitahuan privasi data"):
            st.markdown(TEKS_PRIVASI.format(kontak=esc(kontak_privasi())))
        if di_aplikasi():
            _profil_umpan_balik(eng, nama)
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
    st.markdown(CSS_PROYEK, unsafe_allow_html=True)
    if perlu_persetujuan(peserta):
        halaman_persetujuan(eng, nama)
        return
    if di_aplikasi():
        if _perlu_umpan_balik_sesi(eng, nama) and not st.session_state.get(f"umpan_lewati:{nama}"):
            halaman_minta_umpan_balik(eng, nama)
            return
        pesan_umpan = st.session_state.pop("pesan_umpan", None)
        if pesan_umpan:
            st.toast(pesan_umpan)
    if nav_perlu_naik():
        st.markdown(CSS_NAV_NAIK, unsafe_allow_html=True)
    tab = st.session_state.get("tab", "beranda")
    if tab == "absen":
        tab_absen(eng, nama, peserta)
    elif tab == "logbook":
        tab_logbook(eng, nama, peserta)
    elif tab == "riwayat":
        tab_riwayat(eng, nama)
    elif tab == "rekap":
        tab_rekap(eng, nama)
    elif tab == "profil":
        tab_profil(eng, nama, peserta)
    else:
        tab = "beranda"
        tab_beranda(eng, nama, peserta)
    st.markdown('<div class="pr-spacer"></div>', unsafe_allow_html=True)
    navigasi_bawah(tab)


# ------------------------------------------------------------------
# ------------------------------------------------------------------
# TAMPILAN PESERTA: PRIVASI DAN LOGBOOK
# ------------------------------------------------------------------
CSS_PROYEK = """
<style>
.pr-link { display: flex; align-items: center; gap: 12px; background: #fff; border-radius: 18px; margin: 12px 22px 0;
  padding: 13px 14px; box-shadow: 0 4px 18px rgba(6,63,107,.08); position: relative; z-index: 2; }
.pr-link .ic { width: 42px; height: 42px; border-radius: 12px; background: #e4f0f8; color: #0a5a8f; display: flex;
  align-items: center; justify-content: center; flex: none; }
.pr-link .t { font-weight: 600; color: #0b3a5e; font-size: 15px; }
.pr-link .s { font-size: 12.5px; color: #5f7b92; margin-top: 2px; }
.pr-link.warn .s { color: #c2410c; font-weight: 600; }
.pr-link .go { margin-left: auto; color: #9aa6c8; font-size: 24px; line-height: 1; }
.lb-chip { display: inline-block; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 999px; }
.lb-diajukan { background: #fff4d6; color: #b45309; }
.lb-disetujui { background: #d9f7e8; color: #12a15f; }
.lb-ditolak { background: #fde2e2; color: #c0392b; }
.lb-bar { height: 8px; background: #e6edfb; border-radius: 999px; overflow: hidden; margin-top: 7px; }
.lb-bar > i { display: block; height: 100%; background: #0a5a8f; border-radius: 999px; }
.lb-bar.rencana { margin-top: 3px; height: 4px; background: #e6eef4; }
.lb-bar.rencana > i { background: #8fc0de; }
</style>
"""

# CATATAN UNTUK PENGELOLA: ini draf pemberitahuan. Sesuaikan dengan kebijakan pelindungan data pribadi instansi dan
# mintalah tinjauan unit hukum/Human Capital sebelum dipakai. Naikkan VERSI_PRIVASI bila isinya diubah supaya semua
# peserta diminta menyetujui ulang. Setiap kalimat di bawah menggambarkan perilaku aplikasi yang sebenarnya.
TEKS_PRIVASI = """
**Data apa yang dikumpulkan**
- **Lokasi GPS** saat kamu menekan absen masuk/pulang (dan percobaan absen yang gagal karena di luar area).
- **Foto selfie** saat absen masuk/pulang.
- **Jam absen**, izin/cuti beserta alasannya, dan **logbook kegiatan** yang kamu tulis.
- Nama, divisi, gedung, dan nomor rekening yang didaftarkan admin.

**Untuk apa**
Memverifikasi kehadiran PKL, menghitung uang makan, dan mengevaluasi kegiatan PKL. Data tidak dipakai untuk tujuan lain.

**Siapa yang dapat melihat**
- **Admin/Human Capital**: absensi, lokasi, dan foto.
- **Pembimbing divisi**: hanya logbook peserta di divisinya. Pembimbing **tidak** dapat melihat foto maupun lokasi.
- Perubahan data oleh admin dan pembimbing tercatat pada jejak audit yang tidak dapat diubah diam-diam.

**Berapa lama**
Selama masa PKL dan jangka waktu administrasi yang ditetapkan instansi. Admin dapat menghapus data sesuai kebijakan tersebut.

**Hakmu**
Kamu dapat menanyakan data apa saja yang tersimpan tentangmu, meminta koreksi, atau menarik persetujuan dengan menghubungi {kontak}.
Tanpa persetujuan, absen lewat aplikasi ini tidak dapat digunakan; hubungi pihak yang sama untuk cara absen alternatif.
"""


def kontak_privasi():
    return (secret("kontak_privasi") or "admin atau unit Human Capital Bank Sulselbar").strip()


def halaman_persetujuan(eng, nama):
    """Layar wajib sebelum peserta memakai aplikasi (dan setiap VERSI_PRIVASI berubah)."""
    header_halaman("Privasi Data", "Mohon dibaca sebelum menggunakan aplikasi")
    with st.container(key="pagebody"):
        st.markdown(TEKS_PRIVASI.format(kontak=esc(kontak_privasi())))
        setuju = st.checkbox("Saya sudah membaca dan menyetujui pemberitahuan ini.", key="privasi_cek")
        if st.button("Setuju dan lanjutkan", key="privasi_ok", type="primary", disabled=not setuju, width="stretch"):
            try:
                catat_persetujuan(eng, nama)
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                st.rerun()
        if st.button("Keluar", key="privasi_keluar", width="stretch"):
            keluar()


def _chip(status):
    return f'<span class="lb-chip lb-{status}">{status.capitalize()}</span>'


def _kartu_logbook_beranda(eng, nama):
    """Kartu tautan ke Logbook di Beranda, dengan ringkasan yang perlu perhatian peserta."""
    df = daftar_logbook(eng, nama=nama, limit=100)
    n_tolak = int((df["status"] == "ditolak").sum())
    n_ajuan = int((df["status"] == "diajukan").sum())
    hari_ini = datetime.now(WITA)
    sudah_hari_ini = bool((df["tanggal"] == hari_ini.strftime("%Y-%m-%d")).any())
    kelas = ""
    if n_tolak:
        sub, kelas = f"{n_tolak} logbook ditolak, perlu kamu perbaiki", "warn"
    elif hari_libur(hari_ini):
        sub = "Tulis kegiatan harianmu"
    elif not sudah_hari_ini:
        sub = "Belum ada kegiatan tercatat hari ini"
    elif n_ajuan:
        sub = f"{n_ajuan} menunggu persetujuan pembimbing"
    else:
        sub = "Kegiatan hari ini sudah tercatat"
    html = (f'<div class="pr-link {kelas}"><div class="ic">{ic("file", 22)}</div>'
            f'<div><div class="t">Logbook kegiatan</div><div class="s">{sub}</div></div><div class="go">›</div></div>')
    tombol_overlay("logbook", html, tab="logbook")


def _hari_kerja_sebelum(d):
    while hari_libur(d):
        d -= timedelta(days=1)
    return d


def _opsi_paket(eng, divisi):
    paket = daftar_paket(eng, divisi)
    opsi = {None: "(Tidak terkait paket kerja)"}
    for r in paket.itertuples():
        opsi[int(r.id)] = f"{r.kode}  {r.nama}"
    return opsi


def _lb_tulis(eng, nama, peserta, now):
    # Isian dikosongkan HANYA setelah berhasil. clear_on_submit akan menghapus teks juga saat validasi gagal,
    # padahal peserta baru saja mengetiknya. Mengubah session_state sebelum widget dibuat itu diizinkan.
    if st.session_state.pop("lb_reset", False):
        st.session_state["lb_uraian"], st.session_state["lb_durasi"], st.session_state["lb_paket"] = "", 4.0, None
    pesan = st.session_state.pop("pesan_logbook", None)
    if pesan:
        st.success(pesan[0])
        if pesan[1]:
            st.info(pesan[1])
    opsi = _opsi_paket(eng, peserta["divisi"])
    with st.form("form_logbook"):
        tgl = st.date_input("Tanggal kegiatan", value=_hari_kerja_sebelum(now.date()), key="lb_tgl",
                            min_value=now.date() - timedelta(days=14), max_value=now.date())
        paket_id = st.selectbox("Paket kerja (opsional)", list(opsi), format_func=opsi.get, key="lb_paket")
        uraian = st.text_area("Apa yang kamu kerjakan?", max_chars=1000, height=130, key="lb_uraian",
                              placeholder="Contoh: Menyusun rekap data nasabah bulan Oktober dan memeriksa kesesuaian angkanya")
        durasi = st.number_input("Durasi (jam)", min_value=0.5, max_value=10.0, value=4.0, step=0.5, key="lb_durasi")
        kirim = st.form_submit_button("Ajukan ke pembimbing", type="primary", width="stretch", key="lb_kirim")
    if kirim:
        try:
            tambah_logbook(eng, nama, tgl, uraian, durasi, paket_id, now=now)
        except ValueError as e:
            st.error(str(e))
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
        else:
            hadir = tgl.isoformat() in {r["tanggal"] for r in get_riwayat_peserta(eng, nama, limit=60)
                                        if r["jenis"] == "masuk" and r["status"] == "berhasil"}
            st.session_state["lb_reset"] = True
            st.session_state["pesan_logbook"] = (
                "Logbook diajukan. Pembimbing akan meninjaunya.",
                None if hadir else "Catatan: tidak ada absen masuk berhasil pada tanggal itu. Pembimbing akan melihat penandanya.")
            st.rerun()


def _lb_daftar(eng, nama, peserta, now):
    df = daftar_logbook(eng, nama=nama, limit=60)
    if df.empty:
        st.info("Belum ada logbook. Buka bagian Tulis untuk menambahkan.")
        return
    c = st.columns(3)
    c[0].metric("Menunggu", int((df["status"] == "diajukan").sum()))
    c[1].metric("Disetujui", int((df["status"] == "disetujui").sum()))
    c[2].metric("Ditolak", int((df["status"] == "ditolak").sum()))
    html = ""
    for r in df.itertuples():
        paket_txt = f"{esc(r.kode_paket)} • " if isinstance(r.kode_paket, str) else ""
        uraian = esc(r.uraian[:90]) + ("…" if len(r.uraian) > 90 else "")
        extra = ""
        if r.status == "ditolak" and isinstance(r.catatan_peninjau, str):
            extra = f'<div class="m" style="color:#c0392b">Catatan pembimbing: {esc(r.catatan_peninjau)}</div>'
        html += (f'<div class="pr-row"><div><div class="d">{tgl_indonesia(date.fromisoformat(r.tanggal))}</div>'
                 f'<div class="m">{paket_txt}{r.durasi_jam:g} jam</div><div class="m">{uraian}</div>{extra}</div>'
                 f'<div class="r">{_chip(r.status)}</div></div>')
    st.markdown(html, unsafe_allow_html=True)

    bisa = df[df["status"].isin(["diajukan", "ditolak"])]
    if bisa.empty:
        return
    st.markdown("**Perbaiki atau hapus**")
    ringkas = {int(r.id): f"{r.tanggal} • {r.uraian[:38]}" for r in bisa.itertuples()}
    pilih = st.selectbox("Pilih kegiatan", list(ringkas), format_func=ringkas.get, key="lb_pilih")
    row = bisa[bisa["id"] == pilih].iloc[0]
    opsi = _opsi_paket(eng, peserta["divisi"])
    idx = list(opsi).index(int(row["paket_id"])) if pd.notna(row["paket_id"]) and int(row["paket_id"]) in opsi else 0
    uraian_baru = st.text_area("Uraian", value=row["uraian"], key=f"lb_u_{pilih}", max_chars=1000, height=120)
    durasi_baru = st.number_input("Durasi (jam)", min_value=0.5, max_value=10.0, value=float(row["durasi_jam"]), step=0.5, key=f"lb_d_{pilih}")
    paket_baru = st.selectbox("Paket kerja", list(opsi), index=idx, format_func=opsi.get, key=f"lb_p_{pilih}")
    k1, k2 = st.columns(2)
    if k1.button("Simpan dan ajukan ulang", key=f"lb_simpan_{pilih}", type="primary", width="stretch"):
        try:
            ubah_logbook(eng, int(pilih), nama, uraian_baru, durasi_baru, paket_baru, now=now)
        except ValueError as e:
            st.error(str(e))
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
        else:
            st.rerun()
    if k2.button("Hapus", key=f"lb_hapus_{pilih}", width="stretch"):
        try:
            hapus_logbook(eng, int(pilih), nama)
        except ValueError as e:
            st.error(str(e))
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
        else:
            st.rerun()


def _baris_progres_html(r):
    spi = "-" if r.spi is None or pd.isna(r.spi) else f"{r.spi:.2f}"
    return (f'<div class="pr-row"><div style="flex:1"><div class="d">{esc(r.kode)}  {esc(r.nama)}</div>'
            f'<div class="m">{r.jam_disetujui:g} dari {r.estimasi_jam:g} jam • rencana {r.rencana * 100:.0f}% • SPI {spi} • {esc(r.status)}</div>'
            f'<div class="lb-bar"><i style="width:{r.progres * 100:.0f}%"></i></div>'
            f'<div class="lb-bar rencana"><i style="width:{r.rencana * 100:.0f}%"></i></div></div>'
            f'<div class="r">{r.progres * 100:.0f}%</div></div>')


def _lb_progres(eng, nama, peserta, now):
    saya = daftar_logbook(eng, nama=nama, status="disetujui", limit=1000)
    st.metric("Jam logbook yang sudah disetujui (kamu)", f"{saya['durasi_jam'].sum():g}")
    hasil = progres_divisi(eng, peserta["divisi"], now.date())
    if hasil.empty:
        st.info("Pembimbing belum menyusun paket kerja untuk divisimu.")
        return
    st.caption("Progres paket kerja divisi (gabungan semua peserta). Batang biru = progres, abu-abu = rencana jadwal.")
    st.markdown("".join(_baris_progres_html(r) for r in hasil.itertuples()), unsafe_allow_html=True)


def tab_logbook(eng, nama, peserta):
    now = datetime.now(WITA)
    header_halaman("Logbook", peserta["divisi"])
    with st.container(key="pagebody"):
        if st.button("← Beranda", key="lb_kembali"):
            st.session_state.tab = "beranda"
            st.rerun()
        sub = st.radio("Bagian", ["Tulis", "Daftar", "Progres"], horizontal=True, key="lb_sub", label_visibility="collapsed")
        if sub == "Tulis":
            _lb_tulis(eng, nama, peserta, now)
        elif sub == "Daftar":
            _lb_daftar(eng, nama, peserta, now)
        else:
            _lb_progres(eng, nama, peserta, now)


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
                "Saat menghapus peserta, hapus juga seluruh riwayat absensi dan logbooknya",
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
CSS_REKAP = """
<style>
.st-key-rekap_tabel, .st-key-rekap_tabel > [data-testid="stVerticalBlock"] { overflow: auto !important; gap: 0 !important; row-gap: 0 !important; }
.st-key-rekap_tabel { border: 1px solid var(--garis); border-radius: 12px; padding: 0 4px; }
.st-key-rekap_tabel > * { border-bottom: 1px solid var(--garis); }
.st-key-rekap_tabel > :has(.rk-h) { position: sticky; top: 0; z-index: 3; background: #eef2ff; }
.st-key-rekap_tabel > :not(:has(.rk-h)):hover { background: #f3f8fb; }
.st-key-rekap_tabel [data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; min-width: 300px; gap: 3px !important; align-items: center; padding: 2px 6px; margin: 0; }
.st-key-rekap_tabel [data-testid="stElementContainer"] { margin: 0 !important; }
.st-key-rekap_tabel [data-testid="stMarkdownContainer"] { min-width: 0; overflow: hidden; margin: 0 !important; }
.st-key-rekap_tabel [data-testid="stMarkdownContainer"] p { margin: 0 !important; }
.st-key-rekap_tabel :is([data-testid="stColumn"], [data-testid="column"]) { width: auto !important; min-width: 0 !important; flex: 1 1 0 !important; }
.st-key-rekap_tabel :is([data-testid="stColumn"], [data-testid="column"]):first-child { flex: 3 1 0 !important; }
.st-key-rekap_tabel :is([data-testid="stColumn"], [data-testid="column"]):nth-child(7) { flex: 2.4 1 0 !important; }
.st-key-rekap_tabel :is([data-testid="stColumn"], [data-testid="column"]):last-child { flex: 0 0 28px !important; }
/* menimpa aturan tombol global panel admin (min-height 44px, border biru, !important) */
.st-key-adminbody .st-key-rekap_tabel button[kind="secondary"] {
  min-height: 26px !important; height: 26px !important; padding: 0 !important; width: 100% !important;
  border-radius: 8px !important; border: 1px solid var(--garis) !important; background: var(--tint) !important; box-shadow: none !important;
}
.st-key-adminbody .st-key-rekap_tabel button[kind="secondary"] p { font-size: 16px !important; font-weight: 700; line-height: 1; margin: 0 !important; }
.rk-h { font-size: 10.5px; font-weight: 700; color: var(--muted); text-align: center; white-space: nowrap; }
.rk-h.kiri { text-align: left; }
.rk-n { font-size: 12.5px; font-weight: 600; color: var(--navy); line-height: 1.2; white-space: normal; overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; word-break: break-word; }
.rk-v { font-size: 12.5px; color: var(--navy); text-align: center; white-space: nowrap; }
.rk-sub { font-size: 12px; color: var(--muted); }
.rk-n .rk-sub { font-size: 10px; font-weight: 400; }
@media (max-width: 600px) { .db-k-n { font-size: 19px; white-space: nowrap; } }
.rk-kartu { display: flex; align-items: center; gap: 14px; background: var(--tint); border-radius: 18px; padding: 14px 16px; margin-bottom: 12px; }
.rk-avatar { width: 54px; height: 54px; border-radius: 50%; background: #fff; color: var(--biru); font-size: 24px; font-weight: 700;
  display: flex; align-items: center; justify-content: center; flex: 0 0 54px; }
.rk-nama { font-size: 17px; font-weight: 700; color: var(--navy); }
.rk-chip { display: inline-block; margin-top: 6px; font-size: 12px; font-weight: 600; padding: 3px 12px; border-radius: 999px; background: #d9f7e8; color: #12a15f; }
.rk-chip.abu { background: #eaf0f5; color: #3f6785; }
.rk-periode { background: #f3f8fb; border: 1px solid var(--garis); border-radius: 16px; padding: 12px 14px; margin-bottom: 4px; }
.rk-periode-t { font-size: 14px; font-weight: 700; color: var(--navy); margin-bottom: 8px; }
.rk-periode-g { display: grid; grid-template-columns: 2fr 1fr; gap: 8px; }
.rk-periode-g div { background: var(--tint); border-radius: 12px; padding: 9px 12px; font-size: 14px; color: var(--navy); }
</style>
"""

HARI_LIBUR_KHUSUS_NOTE = "Hari libur nasional/cuti bersama dihitung sebagai hari kerja kecuali diisi di HARI_LIBUR_KHUSUS (app.py)."
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _rp(n):
    return "Rp" + f"{int(n):,}".replace(",", ".")


def _susun_rekap_detail(df):
    """Tabel 'rekap' dan 'detail' bentuk lama (dipakai ekspor Excel ringkas): hanya nama yang punya data di bulan itu."""
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
    detail = df.rename(columns={
        "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis",
        "status": "Status", "jarak": "Jarak (m)", "uang_makan": "Uang Makan", "alasan": "Alasan",
    })
    detail["Alasan"] = detail["Alasan"].fillna("")
    detail["Keterangan"] = detail.apply(
        lambda r: ketepatan(r["Jam"]) if (r["Jenis"] == "masuk" and r["Status"] == "berhasil") else "", axis=1)
    detail["Jarak (m)"] = detail["Jarak (m)"].round(0)
    return rekap, detail


def hari_kerja_periode(tahun, bulan, sekarang=None):
    """Hari kerja (bukan Sabtu/Minggu/HARI_LIBUR_KHUSUS) dalam satu bulan yang SUDAH LEWAT. Hari ini baru dihitung
    setelah jendela absen masuk tutup, karena sebelum itu peserta masih boleh absen."""
    sekarang = sekarang or datetime.now(WITA)
    awal = date(tahun, bulan, 1)
    akhir_bulan = date(tahun + (bulan == 12), bulan % 12 + 1, 1) - timedelta(days=1)
    tutup = sekarang.time().replace(second=0, microsecond=0) > JAM_MASUK_TUTUP
    batas = sekarang.date() if tutup else sekarang.date() - timedelta(days=1)
    akhir = min(akhir_bulan, batas)
    if akhir < awal:
        return []
    semua = (awal + timedelta(days=i) for i in range((akhir - awal).days + 1))
    return [d for d in semua if not hari_libur(d) and d.isoformat() not in HARI_LIBUR_KHUSUS]


def absen_pertama(eng):
    """{nama: tanggal absen pertama} (percobaan gagal lokasi tidak dihitung); patokan awal hitungan 'tidak hadir'."""
    q = select(absensi_t.c.nama, func.min(absensi_t.c.tanggal)).where(absensi_t.c.status != "gagal_lokasi").group_by(absensi_t.c.nama)
    with eng.connect() as c:
        return {n: t for n, t in c.execute(q).all()}


def rekap_peserta_bulan(eng, df, tahun, bulan, sekarang=None):
    """Satu baris per peserta (terdaftar, ditambah nama yang hanya tersisa di riwayat). 'Tidak Hadir' = hari kerja tanpa
    absen masuk, izin, atau cuti, dihitung sejak absen pertama peserta (bukan sejak didaftarkan, karena tanggal daftar
    tidak disimpan); peserta yang belum pernah absen sama sekali tidak dihitung tidak hadir."""
    terdaftar = set(list_peserta(eng)["nama"])
    nama_semua = sorted(terdaftar | set(df["nama"].unique()))
    masuk = df[(df["status"] == "berhasil") & (df["jenis"] == "masuk")]
    masuk = masuk.assign(telat=masuk["jam"].map(terlambat).astype(bool))
    pulang = df[(df["status"] == "berhasil") & (df["jenis"] == "pulang")]
    izin, cuti = df[df["jenis"] == "izin"], df[df["jenis"] == "cuti"]
    tgl = lambda d, n: set(d.loc[d["nama"] == n, "tanggal"])
    hari_kerja = [d.isoformat() for d in hari_kerja_periode(tahun, bulan, sekarang)]
    pertama = absen_pertama(eng)
    baris = []
    for nama in nama_semua:
        t_hadir, t_izin, t_cuti = tgl(masuk, nama), tgl(izin, nama), tgl(cuti, nama)
        diabsen = t_hadir | t_izin | t_cuti
        mulai = pertama.get(nama)
        baris.append({
            "Nama": nama, "Hari Hadir": len(t_hadir), "Terlambat": len(tgl(masuk[masuk["telat"]], nama)),
            "Izin": len(t_izin), "Cuti": len(t_cuti),
            "Tidak Hadir": sum(1 for d in hari_kerja if mulai and d >= mulai and d not in diabsen),
            "Absen Pulang": len(tgl(pulang, nama)), "Total Uang Makan": int(pulang.loc[pulang["nama"] == nama, "uang_makan"].sum()),
            "Terdaftar": nama in terdaftar,
        })
    return pd.DataFrame(baris)


def detail_absensi_peserta(df, nama):
    """(tabel per tanggal, ringkasan) untuk halaman detail satu peserta. Hari yang hanya berisi percobaan gagal lokasi
    tidak ditampilkan. Bila ada absen masuk dan izin di hari yang sama, absen masuk yang menentukan."""
    d = df[df["nama"] == nama]
    baris, hadir, telat, izin, cuti, uang = [], 0, 0, 0, 0, 0
    for tgl_iso, g in d.groupby("tanggal", sort=True):
        masuk = g[(g["jenis"] == "masuk") & (g["status"] == "berhasil")]
        pulang = g[(g["jenis"] == "pulang") & (g["status"] == "berhasil")]
        izin_g, cuti_g = g[g["jenis"] == "izin"], g[g["jenis"] == "cuti"]
        if not masuk.empty:
            jam = masuk["jam"].iloc[0]
            status = "Terlambat" if terlambat(jam) else "Hadir"
            hadir += 1
            telat += status == "Terlambat"
            jarak = masuk["jarak"].iloc[0]
            ket = "-"
        elif not izin_g.empty or not cuti_g.empty:
            ada_izin = not izin_g.empty
            status = "Izin" if ada_izin else "Cuti"
            izin += ada_izin
            cuti += not ada_izin
            jam, jarak = "", None
            alasan = (izin_g if ada_izin else cuti_g)["alasan"].iloc[0]
            ket = alasan if isinstance(alasan, str) and alasan.strip() else "-"
        else:
            continue
        uang_hari = int(pulang["uang_makan"].sum()) if not pulang.empty else 0
        uang += uang_hari
        baris.append({
            "_tgl": tgl_iso, "Tanggal": date.fromisoformat(tgl_iso).strftime("%d/%m/%Y"),
            "Absen masuk": jam[:5] if jam else "-", "Status": status,
            "Jarak (m)": "-" if jarak is None or pd.isna(jarak) else str(int(round(jarak))), "Keterangan": ket,
            "Absen pulang": pulang["jam"].iloc[0][:5] if not pulang.empty else "-", "Uang makan": _rp(uang_hari),
        })
    kolom = ["_tgl", "Tanggal", "Absen masuk", "Status", "Jarak (m)", "Keterangan", "Absen pulang", "Uang makan"]
    return pd.DataFrame(baris, columns=kolom), {"hadir": hadir, "terlambat": telat, "izin": izin, "cuti": cuti, "uang": uang}


def foto_absen_peserta(eng, nama, tanggal):
    """{'masuk': (jam, bytes|None), 'pulang': (...)} untuk satu peserta di satu tanggal (hanya absen yang berhasil).
    Foto diambil saat dibutuhkan, bukan ikut dimuat di tabel."""
    q = select(absensi_t.c.jenis, absensi_t.c.jam, absensi_t.c.foto, absensi_t.c.foto_path).where(and_(
        absensi_t.c.nama == nama, absensi_t.c.tanggal == tanggal, absensi_t.c.status == "berhasil",
        absensi_t.c.jenis.in_(("masuk", "pulang"))))
    with eng.connect() as c:
        baris = c.execute(q).mappings().all()
    return {r["jenis"]: (r["jam"], ambil_foto(r["foto"], r["foto_path"])) for r in baris}


def _html_kartu_rekap(aktif, r, label_periode):
    kartu = [
        ("db-hadir", aktif, "Peserta Aktif"), ("db-izin", int(r["Hari Hadir"].sum()), "Hari Hadir"),
        ("db-telat", int(r["Terlambat"].sum()), "Terlambat"), ("db-izin", int(r["Izin"].sum()), "Izin"),
        ("db-cuti", int(r["Cuti"].sum()), "Cuti"), ("db-belum", int(r["Tidak Hadir"].sum()), "Tidak Hadir"),
    ]
    return (
        f'<div class="db-judul"><div class="db-judul-t">Rekap Total Peserta</div><div class="db-judul-s">{esc(label_periode)}</div></div>'
        '<div class="db-grid">' + "".join(
            f'<div class="db-k {kelas}"><div class="db-k-n">{n}</div><div class="db-k-t">{esc(label)}</div></div>'
            for kelas, n, label in kartu) + "</div>"
    )


def _tabel_rekap_peserta(rekap, tahun, bulan):
    """Tabel ringkas peserta dengan kolom terakhir tombol ›. Klik membuka halaman detail peserta itu."""
    lebar = [3, 1, 1, 1, 1, 1, 2, 1]
    judul = ["Nama", "Hadir", "Telat", "Izin", "Cuti", "Absen", "Uang", ""]
    # Tabel punya scroll sendiri (atas-bawah dan kiri-kanan) dengan tinggi dibatasi; halaman tidak ikut memanjang.
    with st.container(key="rekap_tabel", height=min(340, 40 + 35 * len(rekap))):
        for i, (kol, t) in enumerate(zip(st.columns(lebar), judul)):
            kelas = "rk-h kiri" if i == 0 else "rk-h"
            bantu = {"Absen": ' title="Tidak hadir"', "Telat": ' title="Terlambat"', "Uang": ' title="Uang makan"'}.get(t, "")
            kol.markdown(f'<div class="{kelas}"{bantu}>{t}</div>', unsafe_allow_html=True)
        for i, r in enumerate(rekap.to_dict("records"), 1):
            kol = st.columns(lebar, vertical_alignment="center")
            sisa = "" if r["Terdaftar"] else ' <span class="rk-sub">(tidak terdaftar)</span>'
            kol[0].markdown(f'<div class="rk-n" title="{esc(r["Nama"])}">{esc(r["Nama"])}{sisa}</div>', unsafe_allow_html=True)
            for k, nilai in zip(kol[1:6], (r["Hari Hadir"], r["Terlambat"], r["Izin"], r["Cuti"], r["Tidak Hadir"])):
                k.markdown(f'<div class="rk-v">{nilai}</div>', unsafe_allow_html=True)
            kol[6].markdown(f'<div class="rk-v">{_rp(r["Total Uang Makan"])}</div>', unsafe_allow_html=True)
            if kol[7].button("›", key=f"rk_buka_{i}", help=f"Detail absensi {r['Nama']}"):
                st.session_state["rekap_detail"] = {"nama": r["Nama"], "tahun": tahun, "bulan": bulan}
                st.rerun()


def _halaman_detail_rekap(eng, df, nama, tahun, bulan):
    """Halaman detail absensi satu peserta untuk satu bulan: ringkasan, tabel per tanggal, bukti foto, dan Excel peserta."""
    if st.button("‹ Kembali ke rekap", key="rk_kembali"):
        st.session_state.pop("rekap_detail", None)
        st.rerun()
    info = get_peserta(eng, nama)
    divisi = (info["divisi"] if info else "") or "-"
    gedung = f"Gedung {info['gedung']}" if info else "-"
    chip = '<span class="rk-chip">Aktif</span>' if info else '<span class="rk-chip abu">Tidak terdaftar</span>'
    st.markdown(
        f'<div class="rk-kartu"><div class="rk-avatar">{esc(inisial(nama))}</div><div>'
        f'<div class="rk-nama">{esc(nama)}</div><div class="rk-sub">Divisi {esc(divisi)} • {esc(gedung)}</div>{chip}</div></div>'
        f'<div class="rk-periode"><div class="rk-periode-t">Periode Rekap</div>'
        f'<div class="rk-periode-g"><div>{BULAN_ID[bulan - 1]}</div><div>{tahun}</div></div></div>',
        unsafe_allow_html=True,
    )
    tabel, ring = detail_absensi_peserta(df, nama)
    st.markdown(
        '<div class="db-judul"><div class="db-judul-t">Tabel Detail Absensi Peserta</div></div>'
        '<div class="db-grid">'
        f'<div class="db-k db-hadir"><div class="db-k-n">{ring["hadir"]}</div><div class="db-k-t">Hari Hadir</div></div>'
        f'<div class="db-k db-telat"><div class="db-k-n">{ring["terlambat"]}</div><div class="db-k-t">Terlambat</div></div>'
        f'<div class="db-k db-izin"><div class="db-k-n">{_rp(ring["uang"])}</div><div class="db-k-t">Uang Makan</div></div></div>'
        f'<div class="db-catatan">Izin {ring["izin"]} • Cuti {ring["cuti"]}</div>',
        unsafe_allow_html=True,
    )
    if tabel.empty:
        st.info("Tidak ada absen yang berhasil, izin, atau cuti untuk peserta ini di bulan tersebut.")
    else:
        st.dataframe(beri_nomor(tabel.drop(columns="_tgl")), width="stretch")
        tgl_foto = list(tabel.loc[tabel["Status"].isin(["Hadir", "Terlambat"]), "_tgl"])
        pilih = st.selectbox("Lihat bukti foto: pilih tanggal", tgl_foto, index=None, placeholder="Pilih tanggal",
                             format_func=lambda t: date.fromisoformat(t).strftime("%d/%m/%Y"), key=f"rk_foto_{nama}_{tahun}_{bulan}")
        if pilih:
            foto = foto_absen_peserta(eng, nama, pilih)
            kol = st.columns(2)
            for k, jenis, judul in ((kol[0], "masuk", "Foto absen masuk"), (kol[1], "pulang", "Foto absen pulang")):
                jam, isi = foto.get(jenis, ("", None))
                k.caption(f"{judul} {jam[:5]}".strip())
                if isi:
                    k.image(isi, width=170)
                else:
                    k.caption("Belum absen pulang." if jenis == "pulang" and not jam else "Foto tidak tersedia.")
    aman = re.sub(r"[^A-Za-z0-9]+", "_", nama).strip("_")
    try:
        data = excel_template_bytes(eng, df, tahun, bulan, [nama])
    except FileNotFoundError:
        data = None
        st.warning("File **template_absen.xlsx** tidak ditemukan di folder aplikasi. Unggah file itu ke folder yang sama dengan app.py (GitHub), lalu muat ulang.")
    if data:
        st.download_button("⬇️ Download Excel Peserta", data=data, width="stretch", type="primary", mime=MIME_XLSX,
                           file_name=f"ABSEN_MAGANG_{aman}_{tahun}_{bulan:02d}.xlsx")
    if info and divisi_tanpa_penandatangan(eng, [nama]):
        st.warning(f"Penandatangan untuk Divisi {divisi} belum diisi, jadi bagian tanda tangan di Excel dikosongkan. Isi di `PENANDATANGAN` pada app.py.")


def halaman_rekap(eng):
    if st.session_state.get("rekap_detail"):
        st.header("🧾 Detail Absensi")
    else:
        st.header("📊 Rekap Absensi & Uang Makan")
    if not admin_login():
        return
    st.markdown(CSS_DASHBOARD + CSS_REKAP, unsafe_allow_html=True)
    try:
        det = st.session_state.get("rekap_detail")
        if det:
            _halaman_detail_rekap(eng, get_absensi_bulan(eng, det["tahun"], det["bulan"]), det["nama"], det["tahun"], det["bulan"])
            return

        hari_ini = datetime.now(WITA).date()
        c1, c2 = st.columns(2)
        bulan = c1.selectbox("Bulan", list(range(1, 13)), index=hari_ini.month - 1, format_func=lambda m: BULAN_ID[m - 1], key="rekap_bulan")
        tahun = c2.number_input("Tahun", value=hari_ini.year, step=1, key="rekap_tahun")
        tahun_i, bulan_i = int(tahun), int(bulan)
        df = get_absensi_bulan(eng, tahun_i, bulan_i)
        if df.empty:
            st.info("Belum ada data absensi di bulan ini.")
            return
        rekap_p = rekap_peserta_bulan(eng, df, tahun_i, bulan_i)
    except SQLAlchemyError as e:
        st.error(pesan_db(e))
        return

    aktif = int(rekap_p["Terdaftar"].sum())
    st.markdown(_html_kartu_rekap(aktif, rekap_p, f"{BULAN_ID[bulan_i - 1]} {tahun_i}"), unsafe_allow_html=True)
    st.markdown('<div class="db-judul"><div class="db-judul-t">Rekap Absensi Peserta</div></div>', unsafe_allow_html=True)
    st.caption("Klik › di ujung baris untuk membuka detail absensi peserta itu.")
    _tabel_rekap_peserta(rekap_p, tahun_i, bulan_i)

    daftar_nama = sorted(df["nama"].unique())
    # Tombol langsung tampil; file Excel baru dibuat saat tombol diklik (data berupa fungsi), jadi halaman tidak tertahan.
    if os.path.exists(TEMPLATE_PATH):
        st.download_button("⬇️ Download Semua Excel Peserta", data=lambda: excel_template_bytes(eng, df, tahun_i, bulan_i, daftar_nama),
                           width="stretch", type="primary", mime=MIME_XLSX,
                           file_name=f"ABSEN_MAGANG_SEMUA_{tahun_i}_{bulan_i:02d}.xlsx", key="rk_dl_semua")
    else:
        st.download_button("⬇️ Download Semua Excel Peserta", data=b"", width="stretch", type="primary", disabled=True,
                           key="rk_dl_semua_off", mime=MIME_XLSX)
        st.warning("File **template_absen.xlsx** tidak ditemukan di folder aplikasi. Unggah file itu ke folder yang sama dengan app.py (GitHub), lalu muat ulang.")
    st.download_button("⬇️ Rekap ringkas (tabel)", data=lambda: excel_bytes(*_susun_rekap_detail(df)), width="stretch", mime=MIME_XLSX,
                       file_name=f"rekap_absensi_{tahun_i}_{bulan_i:02d}.xlsx", key="rk_dl_ringkas")
    kurang = divisi_tanpa_penandatangan(eng, daftar_nama)
    if kurang:
        st.warning("Penandatangan belum diisi untuk: " + ", ".join(kurang) +
                   ". Bagian tanda tangan di Excel untuk divisi itu dikosongkan. Isi nama dan jabatannya di `PENANDATANGAN` pada app.py.")
    st.caption("Excel per peserta ada di halaman detail (klik › pada barisnya).")


# ------------------------------------------------------------------
# PROYEK: audit, persetujuan privasi, pembimbing, paket kerja, logbook, anggaran (logika murni di proyek.py)
# ------------------------------------------------------------------
WAKTU_FMT = "%Y-%m-%d %H:%M:%S"
STATUS_LOGBOOK = ("diajukan", "disetujui", "ditolak")


def _sekarang_str():
    return datetime.now(WITA).strftime(WAKTU_FMT)


def _cek_wewenang(peran, divisi_scope, divisi_objek):
    """Penjagaan di SISI SERVER (bukan hanya di tampilan): pembimbing hanya boleh menyentuh data divisinya."""
    if peran == "pembimbing" and divisi_scope is None:
        raise PermissionError("Pembimbing harus terikat pada satu divisi.")
    if divisi_scope is not None and divisi_objek != divisi_scope:
        raise PermissionError("Tidak berwenang atas data divisi lain.")


# ---------- audit: baca & verifikasi ----------
def daftar_audit(eng, limit=200, aksi=None, aktor=None):
    """Catatan audit terbaru dulu (tanpa kolom hash). Filter opsional aksi/aktor (cocok persis)."""
    q = select(audit_log_t.c.id, audit_log_t.c.waktu, audit_log_t.c.aktor, audit_log_t.c.peran, audit_log_t.c.aksi,
               audit_log_t.c.objek, audit_log_t.c.detail).order_by(audit_log_t.c.id.desc()).limit(limit)
    if aksi:
        q = q.where(audit_log_t.c.aksi == aksi)
    if aktor:
        q = q.where(audit_log_t.c.aktor == aktor)
    with eng.connect() as c:
        return pd.read_sql(q, c)


def verifikasi_audit(eng):
    """Periksa seluruh rantai hash. Lihat proyek.verifikasi_rantai untuk apa yang bisa dan tidak bisa dideteksi."""
    with eng.connect() as c:
        rows = [dict(r) for r in c.execute(select(audit_log_t).order_by(audit_log_t.c.id)).mappings().all()]
    return proyek.verifikasi_rantai(rows)


def hash_audit_ada(eng, hash_catatan):
    """True bila hash itu masih menjadi bagian rantai audit. Dipakai sebagai 'jangkar': hash puncak yang dicatat di luar
    sistem pada waktu lampau HARUS tetap ada; bila hilang, ada catatan yang dipotong atau ditulis ulang."""
    with eng.connect() as c:
        return c.execute(select(func.count()).select_from(audit_log_t).where(audit_log_t.c.hash == hash_catatan)).scalar() > 0


# ---------- persetujuan privasi ----------
VERSI_PRIVASI = "2026-11"          # naikkan bila isi pemberitahuan berubah: semua peserta diminta menyetujui ulang


def persetujuan_wajib():
    """Bisa dimatikan lewat secret/env privasi_wajib=0 (mis. pengembangan lokal). Bawaan: wajib."""
    return str(secret("privasi_wajib", "1")).strip().lower() not in ("0", "false", "tidak", "no")


def perlu_persetujuan(peserta):
    return persetujuan_wajib() and peserta["setuju_versi"] != VERSI_PRIVASI


@coba_ulang
def catat_persetujuan(eng, nama):
    with eng.begin() as c:
        r = c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(setuju_versi=VERSI_PRIVASI, setuju_pada=_sekarang_str()))
        if r.rowcount:
            _audit(c, nama, "peserta", "privasi.setuju", f"peserta:{nama}", {"versi": VERSI_PRIVASI})


# ---------- pembimbing ----------
_hash_dummy_cache = []


def _hash_dummy():
    """Hash tiruan agar login dengan nama yang tidak ada memakan waktu sama dengan nama yang ada (cegah pendataan nama)."""
    if not _hash_dummy_cache:
        _hash_dummy_cache.append(proyek.hash_pin_kuat("000000"))
    return _hash_dummy_cache[0]


@coba_ulang
def add_pembimbing(eng, nama, divisi, pin, aktor="admin", peran="admin"):
    nama = (nama or "").strip()
    if len(nama) < 3:
        raise ValueError("Nama pembimbing minimal 3 karakter.")
    if divisi not in DIVISI_GEDUNG:
        raise ValueError("Divisi tidak dikenal.")
    err = proyek.validasi_pin_pembimbing(pin)
    if err:
        raise ValueError(err)
    with eng.begin() as c:
        c.execute(insert(pembimbing_t).values(nama=nama, divisi=divisi, pin_hash=proyek.hash_pin_kuat(pin), dibuat=_sekarang_str()))
        _audit(c, aktor, peran, "pembimbing.tambah", f"pembimbing:{nama}", {"divisi": divisi})


def list_pembimbing(eng):
    with eng.connect() as c:
        return pd.read_sql(select(pembimbing_t.c.nama, pembimbing_t.c.divisi, pembimbing_t.c.dibuat).order_by(pembimbing_t.c.nama), c)


def get_pembimbing(eng, nama):
    with eng.connect() as c:
        return c.execute(select(pembimbing_t).where(pembimbing_t.c.nama == nama)).mappings().first()


def cek_pin_pembimbing(eng, nama, pin):
    b = get_pembimbing(eng, nama)
    if b is None:
        proyek.cek_pin_kuat(pin, _hash_dummy())
        return False
    return proyek.cek_pin_kuat(pin, b["pin_hash"])


@coba_ulang
def reset_pin_pembimbing(eng, nama, pin_baru, aktor="admin", peran="admin"):
    err = proyek.validasi_pin_pembimbing(pin_baru)
    if err:
        raise ValueError(err)
    with eng.begin() as c:
        r = c.execute(update(pembimbing_t).where(pembimbing_t.c.nama == nama).values(pin_hash=proyek.hash_pin_kuat(pin_baru)))
        if r.rowcount:
            _audit(c, aktor, peran, "pembimbing.reset_pin", f"pembimbing:{nama}")


@coba_ulang
def hapus_pembimbing(eng, nama, aktor="admin", peran="admin"):
    with eng.begin() as c:
        r = c.execute(delete(pembimbing_t).where(pembimbing_t.c.nama == nama))
        if r.rowcount:
            _audit(c, aktor, peran, "pembimbing.hapus", f"pembimbing:{nama}")


# ---------- paket kerja (WBS) ----------
def _tgl(teks):
    return date.fromisoformat(teks)


def daftar_paket(eng, divisi=None):
    q = select(paket_kerja_t).order_by(paket_kerja_t.c.divisi, paket_kerja_t.c.id)
    if divisi is not None:
        q = q.where(paket_kerja_t.c.divisi == divisi)
    with eng.connect() as c:
        df = pd.read_sql(q, c)
    if len(df):
        df["mulai"], df["selesai"] = df["mulai"].map(_tgl), df["selesai"].map(_tgl)
    return df


def _paket_atau_error(c, paket_id):
    r = c.execute(select(paket_kerja_t).where(paket_kerja_t.c.id == paket_id)).mappings().first()
    if r is None:
        raise ValueError("Paket kerja tidak ditemukan.")
    return r


@coba_ulang
def tambah_paket(eng, divisi, kode, nama, mulai, selesai, estimasi_jam, aktor="admin", peran="admin", divisi_scope=None):
    _cek_wewenang(peran, divisi_scope, divisi)
    if divisi not in DIVISI_GEDUNG:
        raise ValueError("Divisi tidak dikenal.")
    errs = proyek.validasi_paket(kode, nama, mulai, selesai, estimasi_jam)
    if errs:
        raise ValueError(" ".join(errs))
    try:
        with eng.begin() as c:
            r = c.execute(insert(paket_kerja_t).values(
                divisi=divisi, kode=kode.strip(), nama=nama.strip(), mulai=mulai.isoformat(), selesai=selesai.isoformat(),
                estimasi_jam=float(estimasi_jam), ditutup=0, dibuat_oleh=aktor))
            pid = int(r.inserted_primary_key[0])
            _audit(c, aktor, peran, "paket.tambah", f"paket:{divisi}:{kode.strip()}",
                   {"id": pid, "nama": nama.strip(), "mulai": mulai.isoformat(), "selesai": selesai.isoformat(), "estimasi_jam": float(estimasi_jam)})
    except IntegrityError:
        raise ValueError(f"Kode {kode.strip()} sudah dipakai di divisi ini.") from None
    return pid


@coba_ulang
def ubah_paket(eng, paket_id, kode, nama, mulai, selesai, estimasi_jam, aktor="admin", peran="admin", divisi_scope=None):
    errs = proyek.validasi_paket(kode, nama, mulai, selesai, estimasi_jam)
    if errs:
        raise ValueError(" ".join(errs))
    try:
        with eng.begin() as c:
            lama = _paket_atau_error(c, paket_id)
            _cek_wewenang(peran, divisi_scope, lama["divisi"])
            c.execute(update(paket_kerja_t).where(paket_kerja_t.c.id == paket_id).values(
                kode=kode.strip(), nama=nama.strip(), mulai=mulai.isoformat(), selesai=selesai.isoformat(), estimasi_jam=float(estimasi_jam)))
            _audit(c, aktor, peran, "paket.ubah", f"paket:{lama['divisi']}:{kode.strip()}", {
                "id": paket_id, "dari": {"kode": lama["kode"], "nama": lama["nama"], "mulai": lama["mulai"], "selesai": lama["selesai"],
                                         "estimasi_jam": lama["estimasi_jam"]},
                "ke": {"kode": kode.strip(), "nama": nama.strip(), "mulai": mulai.isoformat(), "selesai": selesai.isoformat(),
                       "estimasi_jam": float(estimasi_jam)}})
    except IntegrityError:
        raise ValueError(f"Kode {kode.strip()} sudah dipakai di divisi ini.") from None


@coba_ulang
def tutup_paket(eng, paket_id, ditutup, aktor="admin", peran="admin", divisi_scope=None):
    with eng.begin() as c:
        lama = _paket_atau_error(c, paket_id)
        _cek_wewenang(peran, divisi_scope, lama["divisi"])
        c.execute(update(paket_kerja_t).where(paket_kerja_t.c.id == paket_id).values(ditutup=1 if ditutup else 0))
        _audit(c, aktor, peran, "paket.tutup" if ditutup else "paket.buka_kembali", f"paket:{lama['divisi']}:{lama['kode']}", {"id": paket_id})


@coba_ulang
def hapus_paket(eng, paket_id, aktor="admin", peran="admin", divisi_scope=None):
    """Hapus paket kerja. Logbook yang merujuknya tidak ikut terhapus; hanya dilepas dari paket."""
    with eng.begin() as c:
        lama = _paket_atau_error(c, paket_id)
        _cek_wewenang(peran, divisi_scope, lama["divisi"])
        n = c.execute(update(logbook_t).where(logbook_t.c.paket_id == paket_id).values(paket_id=None)).rowcount
        c.execute(delete(paket_kerja_t).where(paket_kerja_t.c.id == paket_id))
        _audit(c, aktor, peran, "paket.hapus", f"paket:{lama['divisi']}:{lama['kode']}", {"id": paket_id, "logbook_dilepas": n})


def jam_disetujui_per_paket(eng, divisi=None):
    """{paket_id: jam logbook yang SUDAH disetujui}; dasar perhitungan progres."""
    q = (select(logbook_t.c.paket_id, func.sum(logbook_t.c.durasi_jam))
         .where(and_(logbook_t.c.status == "disetujui", logbook_t.c.paket_id.is_not(None))).group_by(logbook_t.c.paket_id))
    if divisi is not None:
        q = q.where(logbook_t.c.divisi == divisi)
    with eng.connect() as c:
        return {int(pid): float(jam) for pid, jam in c.execute(q).all()}


def progres_divisi(eng, divisi, hari_ini=None):
    """Tabel progres paket kerja satu divisi (lihat proyek.progres_paket)."""
    paket = daftar_paket(eng, divisi)
    if paket.empty:
        return paket
    return proyek.progres_paket(paket, jam_disetujui_per_paket(eng, divisi), hari_ini or datetime.now(WITA).date(), HARI_LIBUR)


# ---------- logbook ----------
def _jam_logbook_hari(c, nama, tanggal_iso, kecuali_id=None):
    q = select(func.coalesce(func.sum(logbook_t.c.durasi_jam), 0.0)).where(and_(
        logbook_t.c.nama == nama, logbook_t.c.tanggal == tanggal_iso, logbook_t.c.status != "ditolak"))
    if kecuali_id is not None:
        q = q.where(logbook_t.c.id != kecuali_id)
    return float(c.execute(q).scalar())


def _cek_paket_divisi(c, paket_id, divisi):
    if paket_id is None:
        return None
    r = c.execute(select(paket_kerja_t.c.divisi).where(paket_kerja_t.c.id == paket_id)).first()
    return None if (r is not None and r[0] == divisi) else "Paket kerja itu tidak termasuk divisi kamu."


@coba_ulang
def tambah_logbook(eng, nama, tanggal, uraian, durasi, paket_id=None, now=None):
    """Peserta mengajukan kegiatan harian. ValueError berisi pesan validasi yang siap ditampilkan."""
    now = now or datetime.now(WITA)
    with eng.begin() as c:
        pes = c.execute(select(peserta_t.c.id, peserta_t.c.divisi).where(peserta_t.c.nama == nama)).first()
        if pes is None:
            raise ValueError("Peserta tidak ditemukan.")
        errs = proyek.validasi_logbook(tanggal, now.date(), uraian, float(durasi), _jam_logbook_hari(c, nama, tanggal.isoformat()), HARI_LIBUR)
        salah_paket = _cek_paket_divisi(c, paket_id, pes[1])
        if salah_paket:
            errs.append(salah_paket)
        if errs:
            raise ValueError("\n".join(errs))
        r = c.execute(insert(logbook_t).values(
            peserta_id=pes[0], nama=nama, divisi=pes[1], tanggal=tanggal.isoformat(), uraian=uraian.strip(), durasi_jam=float(durasi),
            paket_id=paket_id, status="diajukan", dibuat=now.strftime(WAKTU_FMT)))
        lid = int(r.inserted_primary_key[0])
        _audit(c, nama, "peserta", "logbook.ajukan", f"logbook:{lid}", {"tanggal": tanggal.isoformat(), "durasi": float(durasi), "paket_id": paket_id})
    return lid


def _logbook_milik(c, logbook_id, nama):
    r = c.execute(select(logbook_t).where(logbook_t.c.id == logbook_id)).mappings().first()
    if r is None or r["nama"] != nama:
        raise ValueError("Logbook tidak ditemukan.")                     # sengaja sama untuk 'bukan milikmu': tidak membocorkan keberadaan
    if r["status"] == "disetujui":
        raise ValueError("Logbook yang sudah disetujui tidak bisa diubah atau dihapus.")
    return r


@coba_ulang
def ubah_logbook(eng, logbook_id, nama, uraian, durasi, paket_id=None, now=None):
    """Ubah logbook milik sendiri yang belum disetujui. Mengajukannya kembali (status jadi 'diajukan')."""
    now = now or datetime.now(WITA)
    with eng.begin() as c:
        lama = _logbook_milik(c, logbook_id, nama)
        errs = proyek.validasi_logbook(_tgl(lama["tanggal"]), now.date(), uraian, float(durasi),
                                       _jam_logbook_hari(c, nama, lama["tanggal"], kecuali_id=logbook_id), HARI_LIBUR, maks_mundur=10_000)
        salah_paket = _cek_paket_divisi(c, paket_id, lama["divisi"])
        if salah_paket:
            errs.append(salah_paket)
        if errs:
            raise ValueError("\n".join(errs))
        c.execute(update(logbook_t).where(logbook_t.c.id == logbook_id).values(
            uraian=uraian.strip(), durasi_jam=float(durasi), paket_id=paket_id, status="diajukan", diubah=now.strftime(WAKTU_FMT),
            catatan_peninjau=None, ditinjau_oleh=None, ditinjau_pada=None))
        _audit(c, nama, "peserta", "logbook.ubah", f"logbook:{logbook_id}",
               {"status_sebelum": lama["status"], "durasi_dari": lama["durasi_jam"], "durasi_ke": float(durasi)})


@coba_ulang
def hapus_logbook(eng, logbook_id, nama):
    with eng.begin() as c:
        lama = _logbook_milik(c, logbook_id, nama)
        c.execute(delete(logbook_t).where(logbook_t.c.id == logbook_id))
        _audit(c, nama, "peserta", "logbook.hapus", f"logbook:{logbook_id}", {"tanggal": lama["tanggal"], "status_sebelum": lama["status"]})


@coba_ulang
def tinjau_logbook(eng, logbook_id, keputusan, catatan, aktor, peran, divisi_scope=None):
    """Pembimbing/admin menyetujui atau menolak logbook yang berstatus 'diajukan'. Menolak wajib disertai alasan."""
    if keputusan not in ("disetujui", "ditolak"):
        raise ValueError("Keputusan harus 'disetujui' atau 'ditolak'.")
    catatan = (catatan or "").strip()
    if keputusan == "ditolak" and len(catatan) < 5:
        raise ValueError("Alasan penolakan wajib diisi (minimal 5 karakter) agar peserta tahu apa yang harus diperbaiki.")
    with eng.begin() as c:
        lama = c.execute(select(logbook_t).where(logbook_t.c.id == logbook_id)).mappings().first()
        if lama is None:
            raise ValueError("Logbook tidak ditemukan.")
        _cek_wewenang(peran, divisi_scope, lama["divisi"])
        if lama["status"] != "diajukan":
            raise ValueError(f"Logbook ini sudah {lama['status']}; hanya yang berstatus 'diajukan' yang bisa ditinjau.")
        c.execute(update(logbook_t).where(logbook_t.c.id == logbook_id).values(
            status=keputusan, catatan_peninjau=catatan or None, ditinjau_oleh=aktor, ditinjau_pada=_sekarang_str()))
        _audit(c, aktor, peran, "logbook.setuju" if keputusan == "disetujui" else "logbook.tolak", f"logbook:{logbook_id}",
               {"peserta": lama["nama"], "tanggal": lama["tanggal"], "durasi": lama["durasi_jam"], "catatan": catatan or None})


def daftar_logbook(eng, nama=None, divisi=None, status=None, awal=None, akhir=None, limit=300):
    """Logbook terbaru dulu, dengan kolom tanpa_absensi (True = tidak ada absen masuk berhasil pada tanggal itu)."""
    j = logbook_t.outerjoin(paket_kerja_t, logbook_t.c.paket_id == paket_kerja_t.c.id)
    q = (select(
        logbook_t.c.id, logbook_t.c.nama, logbook_t.c.divisi, logbook_t.c.tanggal, logbook_t.c.uraian, logbook_t.c.durasi_jam,
        logbook_t.c.paket_id, paket_kerja_t.c.kode.label("kode_paket"), paket_kerja_t.c.nama.label("nama_paket"),
        logbook_t.c.status, logbook_t.c.catatan_peninjau, logbook_t.c.ditinjau_oleh, logbook_t.c.ditinjau_pada,
        logbook_t.c.dibuat, logbook_t.c.diubah,
    ).select_from(j).order_by(logbook_t.c.tanggal.desc(), logbook_t.c.id.desc()).limit(limit))
    for kolom, nilai in ((logbook_t.c.nama, nama), (logbook_t.c.divisi, divisi), (logbook_t.c.status, status)):
        if nilai is not None:
            q = q.where(kolom == nilai)
    if awal is not None:
        q = q.where(logbook_t.c.tanggal >= awal.isoformat())
    if akhir is not None:
        q = q.where(logbook_t.c.tanggal <= akhir.isoformat())
    with eng.connect() as c:
        df = pd.read_sql(q, c)
        if df.empty:
            df["tanpa_absensi"] = pd.Series(dtype=bool)
            return df
        hadir = set(map(tuple, c.execute(select(absensi_t.c.nama, absensi_t.c.tanggal).where(and_(
            absensi_t.c.jenis == "masuk", absensi_t.c.status == "berhasil", absensi_t.c.nama.in_(list(df["nama"].unique())),
            absensi_t.c.tanggal >= df["tanggal"].min(), absensi_t.c.tanggal <= df["tanggal"].max()))).all()))
    df["tanpa_absensi"] = [(n, t) not in hadir for n, t in zip(df["nama"], df["tanggal"])]
    return df


# ---------- anggaran uang makan ----------
def _rentang_bulan(tahun, bulan):
    awal = date(tahun, bulan, 1)
    akhir = date(tahun + (bulan == 12), bulan % 12 + 1, 1) - timedelta(days=1)
    return awal, akhir


def get_anggaran(eng, tahun, bulan):
    with eng.connect() as c:
        return {d: int(j) for d, j in c.execute(select(anggaran_t.c.divisi, anggaran_t.c.jumlah).where(
            and_(anggaran_t.c.tahun == tahun, anggaran_t.c.bulan == bulan))).all()}


@coba_ulang
def set_anggaran(eng, tahun, bulan, divisi, jumlah, catatan=None, aktor="admin", peran="admin"):
    """Tetapkan/ubah anggaran satu divisi untuk satu bulan. Mengembalikan True bila ada perubahan (dan tercatat di audit)."""
    if not (2020 <= int(tahun) <= 2100 and 1 <= int(bulan) <= 12):
        raise ValueError("Tahun atau bulan tidak valid.")
    if divisi not in DIVISI_GEDUNG:
        raise ValueError("Divisi tidak dikenal.")
    jumlah = int(jumlah)
    if not (0 <= jumlah <= 10_000_000_000):
        raise ValueError("Jumlah anggaran harus antara Rp0 dan Rp10 miliar.")
    catatan = (catatan or "").strip()[:300] or None
    kunci = and_(anggaran_t.c.tahun == tahun, anggaran_t.c.bulan == bulan, anggaran_t.c.divisi == divisi)
    with eng.begin() as c:
        ada = c.execute(select(anggaran_t.c.jumlah, anggaran_t.c.catatan).where(kunci)).first()
        if ada is not None and int(ada[0]) == jumlah and (ada[1] or None) == catatan:
            return False
        if ada is None:
            c.execute(insert(anggaran_t).values(tahun=tahun, bulan=bulan, divisi=divisi, jumlah=jumlah, catatan=catatan,
                                                diubah_oleh=aktor, diubah_pada=_sekarang_str()))
        else:
            c.execute(update(anggaran_t).where(kunci).values(jumlah=jumlah, catatan=catatan, diubah_oleh=aktor, diubah_pada=_sekarang_str()))
        _audit(c, aktor, peran, "anggaran.set", f"anggaran:{tahun}-{int(bulan):02d}:{divisi}",
               {"dari": None if ada is None else int(ada[0]), "ke": jumlah, "catatan": catatan})
    return True


@coba_ulang
def hapus_anggaran(eng, tahun, bulan, divisi, aktor="admin", peran="admin"):
    kunci = and_(anggaran_t.c.tahun == tahun, anggaran_t.c.bulan == bulan, anggaran_t.c.divisi == divisi)
    with eng.begin() as c:
        ada = c.execute(select(anggaran_t.c.jumlah).where(kunci)).first()
        if ada is None:
            return False
        c.execute(delete(anggaran_t).where(kunci))
        _audit(c, aktor, peran, "anggaran.hapus", f"anggaran:{tahun}-{int(bulan):02d}:{divisi}", {"dari": int(ada[0])})
    return True


def realisasi_harian(eng, tahun, bulan):
    """Uang makan yang terealisasi per (divisi, tanggal): dicatat saat absen pulang berhasil."""
    awal, akhir = _rentang_bulan(tahun, bulan)
    div = func.coalesce(absensi_t.c.divisi, "(tidak terdaftar)").label("divisi")
    q = (select(div, absensi_t.c.tanggal, func.sum(absensi_t.c.uang_makan).label("jumlah"))
         .where(and_(absensi_t.c.jenis == "pulang", absensi_t.c.status == "berhasil",
                     absensi_t.c.tanggal >= awal.isoformat(), absensi_t.c.tanggal <= akhir.isoformat()))
         .group_by(div, absensi_t.c.tanggal))
    with eng.connect() as c:
        df = pd.read_sql(q, c)
    df["tanggal"] = df["tanggal"].map(_tgl)
    return df


def jumlah_peserta_per_divisi(eng):
    with eng.connect() as c:
        return {d: int(n) for d, n in c.execute(select(peserta_t.c.divisi, func.count()).group_by(peserta_t.c.divisi)).all()}


def tingkat_kehadiran_historis(eng, tahun, bulan, n_bulan=3):
    """Rata-rata tingkat kehadiran per divisi pada n bulan SEBELUM (tahun, bulan):
    jumlah absen pulang berhasil / (peserta yang aktif x hari kerja). Divisi tanpa riwayat tidak muncul."""
    per_divisi = {}
    for k in range(1, n_bulan + 1):
        t, b = proyek.bulan_sebelumnya(tahun, bulan, k)
        awal, akhir = _rentang_bulan(t, b)
        hk = len(proyek.hari_kerja_bulan(t, b, HARI_LIBUR))
        q = select(absensi_t.c.divisi, absensi_t.c.nama, absensi_t.c.jenis).where(and_(
            absensi_t.c.status == "berhasil", absensi_t.c.jenis.in_(["masuk", "pulang"]),
            absensi_t.c.tanggal >= awal.isoformat(), absensi_t.c.tanggal <= akhir.isoformat()))
        with eng.connect() as c:
            df = pd.read_sql(q, c)
        if df.empty:
            continue
        df["divisi"] = df["divisi"].fillna("(tidak terdaftar)")
        for dv, g in df.groupby("divisi"):
            per_divisi.setdefault(dv, []).append(int((g["jenis"] == "pulang").sum()) / (g["nama"].nunique() * hk))
    return {dv: sum(v) / len(v) for dv, v in per_divisi.items()}


def saran_anggaran_bulan(eng, tahun, bulan, cadangan_pct=10.0):
    peserta_n, tingkat = jumlah_peserta_per_divisi(eng), tingkat_kehadiran_historis(eng, tahun, bulan)
    stats = {d: {"n_peserta": n, "tingkat_kehadiran": tingkat.get(d)} for d, n in peserta_n.items()}
    return proyek.saran_anggaran(stats, len(proyek.hari_kerja_bulan(tahun, bulan, HARI_LIBUR)), UANG_MAKAN, cadangan_pct)


def terapkan_saran_anggaran(eng, tahun, bulan, aktor="admin", peran="admin", timpa=False, cadangan_pct=10.0):
    """Tulis saran sebagai anggaran. Divisi yang sudah punya anggaran tidak diubah kecuali timpa=True. Mengembalikan jumlah ditulis."""
    ada = get_anggaran(eng, tahun, bulan)
    n = 0
    for r in saran_anggaran_bulan(eng, tahun, bulan, cadangan_pct).to_dict("records"):
        if r["divisi"] in ada and not timpa:
            continue
        n += int(set_anggaran(eng, tahun, bulan, r["divisi"], r["saran"], f"Saran sistem (cadangan {cadangan_pct:g}%)", aktor, peran))
    return n


def ringkasan_anggaran(eng, tahun, bulan, now=None):
    """Anggaran vs realisasi vs proyeksi per divisi. Mengembalikan (DataFrame, meta)."""
    now = now or datetime.now(WITA)
    hk = proyek.hari_kerja_bulan(tahun, bulan, HARI_LIBUR)
    selesai_hari_ini = now.time() >= JAM_PULANG_TUTUP
    real = realisasi_harian(eng, tahun, bulan)
    per_divisi = {d: dict(zip(g["tanggal"], g["jumlah"].astype(int))) for d, g in real.groupby("divisi")} if len(real) else {}
    n_peserta, anggaran = jumlah_peserta_per_divisi(eng), get_anggaran(eng, tahun, bulan)
    tingkat = tingkat_kehadiran_historis(eng, tahun, bulan)
    baris = []
    for dv in sorted(set(n_peserta) | set(anggaran) | set(per_divisi)):
        n = n_peserta.get(dv, 0)
        rate_awal = n * UANG_MAKAN * (proyek.ASUMSI_KEHADIRAN if tingkat.get(dv) is None else min(1.0, max(0.5, tingkat[dv])))
        pr = proyek.proyeksi_anggaran(per_divisi.get(dv, {}), hk, now.date(), selesai_hari_ini, n, UANG_MAKAN, rate_awal)
        ang = anggaran.get(dv, 0)
        baris.append({
            "divisi": dv, "peserta": n, "anggaran": ang, "realisasi": pr["realisasi"],
            "serapan_pct": (pr["realisasi"] / ang * 100) if ang else None, "sisa_anggaran": ang - pr["realisasi"],
            "proyeksi": pr["proyeksi"], "proyeksi_maks": pr["proyeksi_maks"], "selisih": ang - pr["proyeksi"],
            "status": proyek.status_anggaran(ang, pr["proyeksi"]), "sumber": pr["sumber"],
        })
    kolom = ["divisi", "peserta", "anggaran", "realisasi", "serapan_pct", "sisa_anggaran", "proyeksi", "proyeksi_maks",
             "selisih", "status", "sumber"]
    kosong = proyek.proyeksi_anggaran({}, hk, now.date(), selesai_hari_ini, 0, 0)
    meta = {"hari_total": kosong["hari_total"], "hari_selesai": kosong["hari_selesai"], "hari_sisa": kosong["hari_sisa"]}
    return pd.DataFrame(baris, columns=kolom), meta


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
    d["hari_idx"] = pd.to_datetime(d["tanggal"]).dt.day_of_week
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
    biru, oranye = "#0a5a8f", "#f59e0b"
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
        + dasar.mark_text(dy=-8, color="#0b3a5e").encode(text=alt.Text("persen:Q", format=".0f"))
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
def simpan_model(eng, hasil, aktif=False, jenis=JENIS_MODEL, aktor="admin", peran="admin"):
    """Simpan hasil ml.latih_terlambat sebagai versi baru. aktif=True menonaktifkan versi lain. Mengembalikan id."""
    with eng.begin() as c:
        if aktif:
            c.execute(update(model_versi_t).where(model_versi_t.c.jenis == jenis).values(aktif=0))
        r = c.execute(insert(model_versi_t).values(
            jenis=jenis, dilatih=datetime.now(WITA).strftime("%Y-%m-%d %H:%M:%S"), n_data=hasil["n_data"],
            tanggal_akhir=hasil["tanggal_akhir"], metrik=json.dumps(hasil["metrik"]),
            baseline=json.dumps(hasil["baseline"]), artefak=hasil["artefak"], aktif=1 if aktif else 0,
        ))
        mid = int(r.inserted_primary_key[0])
        _audit(c, aktor, peran, "model.simpan", f"model:{jenis}:{mid}",
               {"aktif": bool(aktif), "n_data": hasil["n_data"], "f1_uji": hasil["metrik"]["uji"]["f1"]})
        return mid


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
def aktifkan_model(eng, model_id, jenis=JENIS_MODEL, aktor="admin", peran="admin"):
    with eng.begin() as c:
        if c.execute(select(model_versi_t.c.id).where(and_(model_versi_t.c.id == model_id, model_versi_t.c.jenis == jenis))).first() is None:
            return False
        c.execute(update(model_versi_t).where(model_versi_t.c.jenis == jenis).values(aktif=0))
        c.execute(update(model_versi_t).where(model_versi_t.c.id == model_id).values(aktif=1))
        _audit(c, aktor, peran, "model.aktifkan", f"model:{jenis}:{model_id}")
        return True


def latih_dan_simpan(eng, aktifkan="jika_lebih_baik", aktor="admin", peran="admin"):
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
    mid = simpan_model(eng, hasil, aktif, aktor=aktor, peran=peran)
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
    garis = alt.Chart(ring).mark_line(strokeDash=[5, 4], color="#5f7b92").encode(
        x=alt.X("dx:Q", scale=skala, title="Timur-barat (m)"), y=alt.Y("dy:Q", scale=skala, title="Utara-selatan (m)"), order="u:Q")
    pts = alt.Chart(titik).mark_circle(size=55).encode(
        x=alt.X("dx:Q", scale=skala), y=alt.Y("dy:Q", scale=skala),
        color=alt.Color("Penanda:N", scale=alt.Scale(domain=["Normal", "Ditandai"], range=["#8fc0de", "#e04a4a"]),
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
# HALAMAN PEMBIMBING DAN ADMIN: PROYEK (logbook, paket kerja, anggaran, pembimbing, audit)
# ------------------------------------------------------------------
def _pilih_divisi(divisi_scope, key, semua=True):
    """Admin memilih divisi; pembimbing terkunci pada divisinya. Mengembalikan nama divisi atau None (= semua)."""
    if divisi_scope is not None:
        st.caption(f"Divisi: {divisi_scope}")
        return divisi_scope
    opsi = (["(Semua divisi)"] if semua else []) + sorted(DIVISI_GEDUNG)
    pilih = st.selectbox("Divisi", opsi, key=key)
    return None if pilih == "(Semua divisi)" else pilih


def ui_tinjau_logbook(eng, divisi_scope, aktor, peran):
    divisi = _pilih_divisi(divisi_scope, "tj_divisi")
    status = st.radio("Status", list(STATUS_LOGBOOK), horizontal=True, key="tj_status", label_visibility="collapsed")
    df = daftar_logbook(eng, divisi=divisi, status=status, limit=100)
    if df.empty:
        st.info(f"Tidak ada logbook berstatus '{status}'.")
        return
    st.caption(f"{len(df)} logbook (maksimal 100 terbaru).")
    for r in df.itertuples():
        with st.container(border=True):
            paket = f" • paket {r.kode_paket}" if isinstance(r.kode_paket, str) else ""
            st.markdown(f"**{r.nama}** • {r.divisi} • {tgl_indonesia(date.fromisoformat(r.tanggal))} • {r.durasi_jam:g} jam{paket}")
            st.write(r.uraian)
            if r.tanpa_absensi:
                st.warning("Tidak ada absen masuk berhasil pada tanggal ini.")
            if r.status == "diajukan":
                catatan = st.text_input("Catatan (wajib bila menolak)", key=f"tj_cat_{r.id}")
                k1, k2 = st.columns(2)
                for kolom, keputusan, label, tipe in ((k1, "disetujui", "Setujui", "primary"), (k2, "ditolak", "Tolak", "secondary")):
                    if kolom.button(label, key=f"tj_{keputusan}_{r.id}", type=tipe, width="stretch"):
                        try:
                            tinjau_logbook(eng, int(r.id), keputusan, catatan, aktor, peran, divisi_scope)
                        except (ValueError, PermissionError) as e:
                            st.error(str(e))
                        except SQLAlchemyError as e:
                            st.error(pesan_db(e))
                        else:
                            st.rerun()
            else:
                st.caption(f"{r.status.capitalize()} oleh {r.ditinjau_oleh} pada {r.ditinjau_pada}"
                           + (f" • catatan: {r.catatan_peninjau}" if isinstance(r.catatan_peninjau, str) else ""))


def ui_paket_kerja(eng, divisi_scope, aktor, peran):
    if st.session_state.pop("pk_reset", False):                  # kosongkan isian hanya setelah berhasil
        st.session_state["pk_kode"], st.session_state["pk_nama"], st.session_state["pk_est"] = "", "", 40.0
    divisi = _pilih_divisi(divisi_scope, "pk_divisi", semua=False)
    hari_ini = datetime.now(WITA).date()
    hasil = progres_divisi(eng, divisi, hari_ini)
    if hasil.empty:
        st.info("Belum ada paket kerja untuk divisi ini. Tambahkan di bawah.")
    else:
        st.caption("Progres = jam logbook yang DISETUJUI dibanding estimasi jam (proksi berbasis usaha, bukan ukuran keluaran). "
                   "SPI < 0,9 terlambat dari jadwal, > 1,1 lebih cepat.")
        st.altair_chart(proyek.grafik_gantt(hasil, hari_ini), width="stretch")
        tabel = hasil.assign(mulai=hasil["mulai"].map(str), selesai=hasil["selesai"].map(str),
                             rencana=(hasil["rencana"] * 100).round(0), progres=(hasil["progres"] * 100).round(0),
                             spi=hasil["spi"].round(2))[["kode", "nama", "mulai", "selesai", "estimasi_jam", "jam_disetujui",
                                                         "rencana", "progres", "spi", "status"]]
        st.dataframe(tabel.rename(columns={"kode": "Kode", "nama": "Paket", "mulai": "Mulai", "selesai": "Selesai",
                                           "estimasi_jam": "Estimasi jam", "jam_disetujui": "Jam disetujui",
                                           "rencana": "Rencana %", "progres": "Progres %", "spi": "SPI", "status": "Status"}),
                     width="stretch", hide_index=True)

    with st.expander("➕ Tambah paket kerja"):
        with st.form("form_paket_baru"):
            c1, c2 = st.columns([1, 3])
            kode = c1.text_input("Kode WBS", placeholder="1.1", key="pk_kode")
            nama = c2.text_input("Nama paket", key="pk_nama")
            d1, d2 = st.columns(2)
            mulai = d1.date_input("Mulai", value=hari_ini, key="pk_mulai")
            selesai = d2.date_input("Selesai", value=hari_ini + timedelta(days=13), key="pk_selesai")
            est = st.number_input("Estimasi jam kerja", min_value=1.0, max_value=2000.0, value=40.0, step=1.0, key="pk_est")
            kirim = st.form_submit_button("Simpan paket", type="primary", key="pk_simpan_baru")
        if kirim:
            try:
                tambah_paket(eng, divisi, kode, nama, mulai, selesai, est, aktor, peran, divisi_scope)
            except (ValueError, PermissionError) as e:
                st.error(str(e))
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                st.session_state["pk_reset"] = True
                st.rerun()

    paket = daftar_paket(eng, divisi)
    if paket.empty:
        return
    with st.expander("✏️ Ubah, tutup, atau hapus paket"):
        label = {int(r.id): f"{r.kode}  {r.nama}" for r in paket.itertuples()}
        pid = st.selectbox("Pilih paket", list(label), format_func=label.get, key="pk_pilih")
        cur = paket[paket["id"] == pid].iloc[0]
        c1, c2 = st.columns([1, 3])
        kode_b = c1.text_input("Kode WBS", value=cur["kode"], key=f"pk_k_{pid}")
        nama_b = c2.text_input("Nama paket", value=cur["nama"], key=f"pk_n_{pid}")
        d1, d2 = st.columns(2)
        mulai_b = d1.date_input("Mulai", value=cur["mulai"], key=f"pk_m_{pid}")
        selesai_b = d2.date_input("Selesai", value=cur["selesai"], key=f"pk_s_{pid}")
        est_b = st.number_input("Estimasi jam kerja", min_value=1.0, max_value=2000.0, value=float(cur["estimasi_jam"]), step=1.0, key=f"pk_e_{pid}")
        b1, b2, b3 = st.columns(3)
        aksi = None
        if b1.button("Simpan perubahan", key=f"pk_ubah_{pid}", type="primary", width="stretch"):
            aksi = lambda: ubah_paket(eng, int(pid), kode_b, nama_b, mulai_b, selesai_b, est_b, aktor, peran, divisi_scope)  # noqa: E731
        if b2.button("Buka kembali" if cur["ditutup"] else "Tutup (selesai)", key=f"pk_tutup_{pid}", width="stretch"):
            aksi = lambda: tutup_paket(eng, int(pid), not bool(cur["ditutup"]), aktor, peran, divisi_scope)  # noqa: E731
        if b3.button("Hapus", key=f"pk_hapus_{pid}", width="stretch"):
            aksi = lambda: hapus_paket(eng, int(pid), aktor, peran, divisi_scope)  # noqa: E731
        if aksi is not None:
            try:
                aksi()
            except (ValueError, PermissionError) as e:
                st.error(str(e))
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                st.rerun()


def _grafik_anggaran(df):
    import altair as alt
    d = df[(df["anggaran"] > 0) | (df["realisasi"] > 0)]
    if d.empty:
        return None
    panjang = d.melt(id_vars="divisi", value_vars=["anggaran", "realisasi", "proyeksi"], var_name="Jenis", value_name="Rp")
    panjang["Jenis"] = panjang["Jenis"].map({"anggaran": "Anggaran", "realisasi": "Realisasi", "proyeksi": "Proyeksi"})
    return alt.Chart(panjang).mark_bar().encode(
        x=alt.X("divisi:N", title=None, axis=alt.Axis(labelAngle=-35, labelLimit=150)), xOffset="Jenis:N",
        y=alt.Y("Rp:Q", title="Rupiah"),
        color=alt.Color("Jenis:N", scale=alt.Scale(domain=["Anggaran", "Realisasi", "Proyeksi"], range=["#8fc0de", "#0a5a8f", "#f59e0b"]),
                        legend=alt.Legend(orient="bottom", title=None)),
        tooltip=["divisi:N", "Jenis:N", alt.Tooltip("Rp:Q", format=",.0f")],
    ).properties(height=300)


def ui_anggaran(eng, aktor):
    now = datetime.now(WITA)
    c1, c2 = st.columns(2)
    tahun = int(c1.number_input("Tahun", min_value=2025, max_value=2100, value=now.year, step=1, key="ag_tahun"))
    bulan = BULAN_ID.index(c2.selectbox("Bulan", BULAN_ID, index=now.month - 1, key="ag_bulan")) + 1
    df, meta = ringkasan_anggaran(eng, tahun, bulan, now)
    st.caption(f"Hari kerja bulan ini: {meta['hari_total']} (selesai {meta['hari_selesai']}, sisa {meta['hari_sisa']}). "
               f"Tarif uang makan {rupiah(UANG_MAKAN)} per hari hadir, tercatat saat absen pulang.")
    m = st.columns(3)
    m[0].metric("Total anggaran", rupiah(int(df["anggaran"].sum())))
    m[1].metric("Realisasi", rupiah(int(df["realisasi"].sum())))
    m[2].metric("Proyeksi akhir bulan", rupiah(int(df["proyeksi"].sum())))
    tampil = df.assign(
        anggaran=df["anggaran"].map(rupiah), realisasi=df["realisasi"].map(rupiah), sisa_anggaran=df["sisa_anggaran"].map(rupiah),
        proyeksi=df["proyeksi"].map(rupiah), proyeksi_maks=df["proyeksi_maks"].map(rupiah), selisih=df["selisih"].map(rupiah),
        serapan_pct=df["serapan_pct"].map(lambda v: "-" if pd.isna(v) else f"{v:.0f}%"),
    ).rename(columns={"divisi": "Divisi", "peserta": "Peserta", "anggaran": "Anggaran", "realisasi": "Realisasi", "serapan_pct": "Serapan",
                      "sisa_anggaran": "Sisa", "proyeksi": "Proyeksi", "proyeksi_maks": "Batas atas", "selisih": "Selisih",
                      "status": "Status", "sumber": "Dasar proyeksi"})
    st.dataframe(tampil, width="stretch", hide_index=True)
    st.caption("Proyeksi = realisasi hari-hari yang sudah selesai + rata-rata harian x hari kerja tersisa. Batas atas = semua peserta hadir "
               "tiap hari tersisa. Selisih = anggaran dikurangi proyeksi (negatif = diperkirakan melebihi). Realisasi per divisi "
               "memakai divisi peserta SAAT absen; riwayat sebelum fitur ini memakai divisi peserta saat ini.")
    grafik = _grafik_anggaran(df)
    if grafik is not None:
        st.altair_chart(grafik, width="stretch")
    st.download_button("⬇️ Ringkasan anggaran (CSV)", df.to_csv(index=False).encode("utf-8"),
                       file_name=f"anggaran_{tahun}-{bulan:02d}.csv", mime="text/csv", key="ag_csv")

    st.subheader("Tetapkan anggaran")
    ver = st.session_state.get("ag_ver", 0)
    ada = get_anggaran(eng, tahun, bulan)
    basis = pd.DataFrame({"Divisi": sorted(DIVISI_GEDUNG), "Anggaran (Rp)": [ada.get(d, 0) for d in sorted(DIVISI_GEDUNG)]})
    edit = st.data_editor(basis, hide_index=True, width="stretch", disabled=["Divisi"], key=f"ag_editor_{tahun}_{bulan}_{ver}",
                          column_config={"Anggaran (Rp)": st.column_config.NumberColumn(min_value=0, max_value=10_000_000_000, step=10_000, format="%d")})
    if st.button("Simpan anggaran", key="ag_simpan", type="primary"):
        diubah, galat = 0, None
        for lama, baru in zip(basis.itertuples(), edit.itertuples()):
            try:
                if int(baru[2]) == int(lama[2]):
                    continue
                if int(baru[2]) == 0:
                    diubah += int(hapus_anggaran(eng, tahun, bulan, lama.Divisi, aktor))
                else:
                    diubah += int(set_anggaran(eng, tahun, bulan, lama.Divisi, int(baru[2]), None, aktor))
            except (ValueError, SQLAlchemyError) as e:
                galat = str(e) if isinstance(e, ValueError) else pesan_db(e)
                break
        if galat:
            st.error(galat)
        else:
            st.session_state["ag_ver"] = ver + 1
            st.session_state["pesan_proyek"] = f"{diubah} anggaran diperbarui." if diubah else "Tidak ada perubahan."
            st.rerun()

    with st.expander("💡 Saran anggaran (estimasi + cadangan)"):
        cad = st.slider("Cadangan (contingency) %", 0, 30, 10, key="ag_cad")
        saran = saran_anggaran_bulan(eng, tahun, bulan, float(cad))
        if saran.empty:
            st.info("Belum ada peserta terdaftar.")
        else:
            st.caption("Estimasi dasar = peserta x hari kerja x tarif x tingkat kehadiran (rata-rata 3 bulan terakhir; bila belum ada riwayat "
                       "dipakai asumsi 90%), lalu ditambah cadangan dan dibulatkan ke atas Rp10.000.")
            st.dataframe(saran.assign(
                tingkat_kehadiran=(saran["tingkat_kehadiran"] * 100).round(0).astype(int).astype(str) + "%",
                estimasi_dasar=saran["estimasi_dasar"].map(rupiah), cadangan=saran["cadangan"].map(rupiah), saran=saran["saran"].map(rupiah),
                asumsi_kehadiran=saran["asumsi_kehadiran"].map({True: "asumsi", False: "riwayat"}),
            ).rename(columns={"divisi": "Divisi", "peserta": "Peserta", "hari_kerja": "Hari kerja", "tingkat_kehadiran": "Kehadiran",
                              "estimasi_dasar": "Estimasi dasar", "cadangan": "Cadangan", "saran": "Saran anggaran",
                              "asumsi_kehadiran": "Dasar kehadiran"}), width="stretch", hide_index=True)
            timpa = st.checkbox("Timpa anggaran yang sudah ditetapkan", key="ag_timpa")
            if st.button("Terapkan saran", key="ag_terapkan"):
                try:
                    n = terapkan_saran_anggaran(eng, tahun, bulan, aktor, "admin", timpa, float(cad))
                except (ValueError, SQLAlchemyError) as e:
                    st.error(str(e) if isinstance(e, ValueError) else pesan_db(e))
                else:
                    st.session_state["ag_ver"] = ver + 1
                    st.session_state["pesan_proyek"] = f"{n} anggaran ditulis dari saran."
                    st.rerun()


def ui_pembimbing(eng, aktor):
    if st.session_state.pop("pb_reset", False):                  # kosongkan isian hanya setelah berhasil
        st.session_state["pb_nama"], st.session_state["pb_pin"] = "", ""
    st.caption("Pembimbing menyetujui logbook dan menyusun paket kerja untuk divisinya. Akun dibuat di sini; PIN 6-8 angka.")
    with st.form("form_pembimbing"):
        nama = st.text_input("Nama pembimbing", key="pb_nama")
        divisi = st.selectbox("Divisi", sorted(DIVISI_GEDUNG), key="pb_divisi")
        pin = st.text_input("PIN (6-8 angka)", max_chars=8, type="password", key="pb_pin")
        kirim = st.form_submit_button("Tambah pembimbing", key="pb_tambah")
    if kirim:
        try:
            add_pembimbing(eng, nama, divisi, pin, aktor)
        except ValueError as e:
            st.error(str(e))
        except IntegrityError:
            st.error("Nama pembimbing itu sudah terdaftar.")
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
        else:
            st.session_state["pb_reset"] = True
            st.session_state["pesan_proyek"] = f"{nama.strip()} ditambahkan. Sampaikan PIN-nya langsung ke yang bersangkutan."
            st.rerun()
    df = list_pembimbing(eng)
    if df.empty:
        st.info("Belum ada pembimbing.")
        return
    st.dataframe(df.rename(columns={"nama": "Nama", "divisi": "Divisi", "dibuat": "Dibuat"}), width="stretch", hide_index=True)
    with st.expander("Reset PIN atau hapus pembimbing"):
        target = st.selectbox("Pilih pembimbing", df["nama"], key="pb_target")
        pin_baru = st.text_input("PIN baru (6-8 angka)", max_chars=8, type="password", key="pb_pin_baru")
        c1, c2 = st.columns(2)
        if c1.button("Reset PIN", key="pb_reset"):
            try:
                reset_pin_pembimbing(eng, target, pin_baru, aktor)
            except ValueError as e:
                st.error(str(e))
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                st.success("PIN diperbarui.")
        if c2.button("Hapus pembimbing", key="pb_hapus"):
            try:
                hapus_pembimbing(eng, target, aktor)
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                st.rerun()


def ui_audit(eng):
    st.caption("Jejak perubahan data oleh admin, pembimbing, peserta, dan sistem. Tiap catatan dikaitkan dengan hash catatan sebelumnya, "
               "jadi perubahan atau penghapusan di tengah rantai terdeteksi.")
    if st.button("Verifikasi integritas rantai", key="au_verif", type="primary"):
        v = verifikasi_audit(eng)
        if v["ok"]:
            st.success(f"Rantai utuh: {v['jumlah']} catatan terverifikasi.")
        else:
            st.error(f"RANTAI RUSAK pada catatan #{v['rusak_id']}: {v['alasan']}.")
        if v["hash_terakhir"]:
            st.code(v["hash_terakhir"], language=None)
            st.caption("Hash puncak di atas. Simpan di luar sistem (mis. dicetak atau dicatat berkala): menghapus catatan TERAKHIR "
                       "tidak memutus rantai, tetapi mengubah hash puncak ini.")
    semua = daftar_audit(eng, limit=1000)
    c1, c2, c3 = st.columns([2, 2, 1])
    aksi = c1.selectbox("Aksi", ["(semua)"] + sorted(semua["aksi"].unique()), key="au_aksi")
    aktor = c2.text_input("Aktor (nama persis)", key="au_aktor")
    batas = c3.number_input("Jumlah", min_value=10, max_value=1000, value=100, step=10, key="au_batas")
    df = daftar_audit(eng, int(batas), None if aksi == "(semua)" else aksi, aktor.strip() or None)
    if df.empty:
        st.info("Belum ada catatan audit pada filter ini.")
        return
    st.dataframe(df.rename(columns={"id": "ID", "waktu": "Waktu (WITA)", "aktor": "Aktor", "peran": "Peran", "aksi": "Aksi",
                                    "objek": "Objek", "detail": "Detail"}), width="stretch", hide_index=True)
    st.download_button("⬇️ Audit (CSV)", df.to_csv(index=False).encode("utf-8"), file_name="audit_log.csv", mime="text/csv", key="au_csv")


def halaman_proyek(eng):
    st.header("📋 Proyek PKL")
    if not admin_login():
        return
    pesan = st.session_state.pop("pesan_proyek", None)
    if pesan:
        st.success(pesan)
    sub = st.radio("Bagian", ["Logbook", "Paket", "Anggaran", "Pembimbing", "Audit"], horizontal=True,
                   key="menu_proyek", label_visibility="collapsed")
    if sub == "Logbook":
        ui_tinjau_logbook(eng, None, "admin", "admin")
    elif sub == "Paket":
        ui_paket_kerja(eng, None, "admin", "admin")
    elif sub == "Anggaran":
        ui_anggaran(eng, "admin")
    elif sub == "Pembimbing":
        ui_pembimbing(eng, "admin")
    else:
        ui_audit(eng)


def halaman_pembimbing(eng, nama):
    """Area pembimbing: tinjau logbook dan kelola paket kerja, terbatas pada divisinya sendiri."""
    b = get_pembimbing(eng, nama)
    if b is None:
        st.error("Akun pembimbing tidak ditemukan (mungkin sudah dihapus admin). Silakan keluar.")
        if st.button("Keluar", key="pb_keluar_hilang"):
            keluar()
        return
    divisi = b["divisi"]
    st.markdown(CSS_ADMIN, unsafe_allow_html=True)
    st.markdown(
        f"""<div class="admin-head"><div class="admin-top">
<div class="admin-logo">{ic("fingerprint", 28)}</div>
<div><div class="admin-title">Panel Pembimbing</div><div class="admin-sub">{esc(nama)} • {esc(divisi)}</div></div>
<div style="flex:1"></div>
<div class="admin-out">{ic("logout", 24)}</div>
</div></div>""",
        unsafe_allow_html=True,
    )
    with st.container(key="adminlogout"):
        if st.button("‎", key="keluar_pembimbing", width="stretch"):
            keluar()
    with st.container(key="adminbody"):
        menu = st.radio("Menu", ["Tinjau logbook", "Paket kerja"], horizontal=True, key="menu_pembimbing", label_visibility="collapsed")
        st.markdown("---")
        if menu == "Tinjau logbook":
            ui_tinjau_logbook(eng, divisi, nama, "pembimbing")
        else:
            ui_paket_kerja(eng, divisi, nama, "pembimbing")


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
# UMPAN BALIK PESERTA: rating, ulasan, dan saran (hanya di aplikasi Android)
# ------------------------------------------------------------------
HARI_MINTA_UMPAN_BALIK = 7              # peserta diarahkan ke form setelah sekian hari sejak absen pertamanya
KATEGORI_ULASAN = {"absen": "Absen", "logbook": "Logbook", "tampilan": "Tampilan", "lainnya": "Lainnya"}
STATUS_ULASAN = ("baru", "diproses", "selesai")
ULASAN_MIN, ULASAN_MAKS = 10, 500
ULASAN_MAKS_PER_SESI = 3                # penahan spam ringan; ulasan tanpa nama jadi tidak bisa dibatasi per peserta di DB


def di_aplikasi():
    """True kalau halaman dibuka lewat aplikasi Android Presensi PKL. Aplikasi itu selalu menambahkan ?mode=pkl dan
    menempelkan 'PresensiPKLApp' pada user-agent; salah satunya cukup. Ini penentu tampilan, BUKAN pengaman: yang
    sengaja memalsukannya di browser hanya akan melihat form umpan balik."""
    if mode_pkl():
        return True
    try:
        return "PresensiPKLApp" in (st.context.headers.get("user-agent") or "")
    except Exception:
        return False


def hari_pemakaian(eng, nama, hari_ini=None):
    """Jumlah hari sejak absen pertama peserta (percobaan yang gagal lokasi tidak dihitung); None kalau belum pernah absen."""
    with eng.connect() as c:
        awal = c.execute(select(func.min(absensi_t.c.tanggal)).where(
            and_(absensi_t.c.nama == nama, absensi_t.c.status != "gagal_lokasi"))).scalar()
    if not awal:
        return None
    return ((hari_ini or datetime.now(WITA).date()) - date.fromisoformat(awal)).days


def rating_peserta(eng, nama):
    """Jumlah bintang (1-5) yang sudah diberikan peserta, atau None kalau belum pernah menilai."""
    with eng.connect() as c:
        return c.execute(
            select(rating_aplikasi_t.c.bintang)
            .select_from(rating_aplikasi_t.join(peserta_t, peserta_t.c.id == rating_aplikasi_t.c.peserta_id))
            .where(peserta_t.c.nama == nama)
        ).scalar()


def perlu_diminta_umpan_balik(eng, nama, hari_ini=None):
    """True kalau peserta belum pernah menilai DAN sudah memakai aplikasi minimal HARI_MINTA_UMPAN_BALIK hari."""
    if rating_peserta(eng, nama) is not None:
        return False
    hari = hari_pemakaian(eng, nama, hari_ini)
    return hari is not None and hari >= HARI_MINTA_UMPAN_BALIK


def validasi_ulasan(isi, kategori, wajib=True):
    """Daftar pesan kesalahan (kosong = valid). wajib=False: isi boleh kosong (peserta hanya memberi rating)."""
    teks = (isi or "").strip()
    if not teks and not wajib:
        return []
    salah = []
    if kategori not in KATEGORI_ULASAN:
        salah.append("Pilih kategori ulasan.")
    if len(teks) < ULASAN_MIN:
        salah.append(f"Ulasan atau saran minimal {ULASAN_MIN} karakter.")
    elif len(teks) > ULASAN_MAKS:
        salah.append(f"Ulasan atau saran maksimal {ULASAN_MAKS} karakter.")
    return salah


@coba_ulang
def simpan_rating(eng, nama, bintang, isi="", kategori="lainnya", now=None):
    """Rating SEKALI seumur akun (tidak bisa diubah). Ulasan opsional ikut disimpan, tanpa nama. Salah -> ValueError."""
    if isinstance(bintang, bool) or not isinstance(bintang, int) or not 1 <= bintang <= 5:
        raise ValueError("Penilaian harus 1 sampai 5 bintang.")
    salah = validasi_ulasan(isi, kategori, wajib=False)
    if salah:
        raise ValueError(" ".join(salah))
    tanggal = (now or datetime.now(WITA)).strftime("%Y-%m-%d")
    try:
        with eng.begin() as c:
            pes = c.execute(select(peserta_t.c.id, peserta_t.c.divisi).where(peserta_t.c.nama == nama)).first()
            if pes is None:
                raise ValueError("Data kamu tidak ditemukan.")
            c.execute(insert(rating_aplikasi_t).values(peserta_id=pes[0], bintang=bintang, dibuat=tanggal))
            if (isi or "").strip():
                c.execute(insert(ulasan_t).values(divisi=pes[1], kategori=kategori, isi=isi.strip(), dibuat=tanggal, status="baru"))
    except IntegrityError:
        raise ValueError("Kamu sudah memberi penilaian. Terima kasih!") from None


@coba_ulang
def simpan_ulasan(eng, nama, isi, kategori, now=None):
    """Ulasan atau saran dari Profil (boleh berkali-kali). Hanya divisi dan tanggal yang disimpan, tanpa nama."""
    salah = validasi_ulasan(isi, kategori)
    if salah:
        raise ValueError(" ".join(salah))
    tanggal = (now or datetime.now(WITA)).strftime("%Y-%m-%d")
    with eng.begin() as c:
        pes = c.execute(select(peserta_t.c.divisi).where(peserta_t.c.nama == nama)).first()
        if pes is None:
            raise ValueError("Data kamu tidak ditemukan.")
        c.execute(insert(ulasan_t).values(divisi=pes[0], kategori=kategori, isi=isi.strip(), dibuat=tanggal, status="baru"))


def ringkasan_rating(eng):
    """Angka gabungan untuk admin: jumlah penilai, rata-rata, sebaran per bintang, dan jumlah peserta."""
    with eng.connect() as c:
        baris = c.execute(select(rating_aplikasi_t.c.bintang, func.count()).group_by(rating_aplikasi_t.c.bintang)).all()
        n_peserta = c.execute(select(func.count()).select_from(peserta_t)).scalar() or 0
    sebaran = {b: 0 for b in range(1, 6)}
    for bintang, jumlah in baris:
        sebaran[int(bintang)] = int(jumlah)
    n = sum(sebaran.values())
    rata = sum(b * k for b, k in sebaran.items()) / n if n else None
    return {"n": n, "rata": rata, "sebaran": sebaran, "n_peserta": int(n_peserta)}


def daftar_ulasan(eng, status=None, limit=500):
    q = select(ulasan_t).order_by(ulasan_t.c.id.desc()).limit(limit)
    if status:
        q = q.where(ulasan_t.c.status == status)
    with eng.connect() as c:
        return pd.read_sql(q, c)


@coba_ulang
def ubah_status_ulasan(eng, ulasan_id, status, catatan="", aktor="admin"):
    """Admin menandai tindak lanjut ulasan (baru, diproses, selesai). Tercatat di audit; ulasan itu sendiri tanpa nama."""
    if status not in STATUS_ULASAN:
        raise ValueError("Status tidak dikenal.")
    catatan = (catatan or "").strip()[:300]
    with eng.begin() as c:
        lama = c.execute(select(ulasan_t.c.status).where(ulasan_t.c.id == ulasan_id)).first()
        if lama is None:
            raise ValueError("Ulasan tidak ditemukan.")
        c.execute(update(ulasan_t).where(ulasan_t.c.id == ulasan_id).values(status=status, catatan_admin=catatan or None))
        _audit(c, aktor, "admin", "ulasan.status", f"ulasan:{ulasan_id}", {"dari": lama[0], "ke": status})


def _perlu_umpan_balik_sesi(eng, nama):
    """Dihitung sekali per sesi (bukan tiap interaksi). Kalau database bermasalah, fitur opsional ini dilewati supaya
    peserta tetap bisa absen."""
    kunci = f"umpan_perlu:{nama}"
    if kunci not in st.session_state:
        try:
            st.session_state[kunci] = perlu_diminta_umpan_balik(eng, nama)
        except SQLAlchemyError as e:
            logging.getLogger("absensi").warning("Cek umpan balik dilewati: %s", type(getattr(e, "orig", e)).__name__)
            return False
    return bool(st.session_state[kunci])


def _bintang(n):
    return "★" * n + "☆" * (5 - n)


def _form_umpan_balik(eng, nama, kunci, minta_rating, boleh_lewati):
    """Form rating + ulasan/saran. minta_rating=False: peserta sudah menilai, jadi hanya ulasan. True bila ada yang tersimpan."""
    k_bintang, k_kat, k_isi = f"ub_bintang_{kunci}", f"ub_kategori_{kunci}", f"ub_isi_{kunci}"
    if st.session_state.pop(f"ub_reset_{kunci}", False):          # kosongkan isian hanya setelah berhasil terkirim
        st.session_state.pop(k_isi, None)
        st.session_state.pop(k_bintang, None)
    bintang = None
    if minta_rating:
        bintang = st.radio("Seberapa puas kamu dengan aplikasi ini?", [1, 2, 3, 4, 5], index=None, horizontal=True,
                           format_func=lambda n: f"{n} ★", key=k_bintang)
        st.caption("1 = sangat kurang, 5 = sangat baik. Penilaian hanya diminta satu kali.")
    kategori = st.selectbox("Tentang apa?", list(KATEGORI_ULASAN), format_func=KATEGORI_ULASAN.get, key=k_kat)
    isi = st.text_area("Ulasan dan saran" + (" (boleh dikosongkan)" if minta_rating else ""), max_chars=ULASAN_MAKS,
                       height=120, key=k_isi, placeholder="Ceritakan yang sudah enak dipakai atau yang masih menyulitkan.")
    st.caption("Ulasan disimpan tanpa namamu (hanya divisi dan tanggal).")
    if boleh_lewati:
        kol_kirim, kol_lewati = st.columns(2)
    else:
        kol_kirim, kol_lewati = st.container(), None
    kirim = kol_kirim.button("Kirim", key=f"ub_kirim_{kunci}", type="primary", width="stretch")
    if kol_lewati is not None and kol_lewati.button("Lewati", key=f"ub_lewati_{kunci}", width="stretch"):
        st.session_state[f"umpan_lewati:{nama}"] = True           # hanya untuk sesi ini; dibuka lagi muncul lagi
        st.rerun()
    if not kirim:
        return False
    if minta_rating and bintang is None:
        st.error("Pilih jumlah bintang dulu.")
        return False
    ada_isi = bool((isi or "").strip())
    kunci_sesi = f"ulasan_terkirim:{nama}"
    if ada_isi and st.session_state.get(kunci_sesi, 0) >= ULASAN_MAKS_PER_SESI:
        st.error("Kamu sudah mengirim beberapa ulasan. Coba lagi nanti.")
        return False
    try:
        if minta_rating:
            simpan_rating(eng, nama, int(bintang), isi, kategori)
        else:
            simpan_ulasan(eng, nama, isi, kategori)
    except ValueError as e:
        st.error(str(e))
        return False
    except SQLAlchemyError as e:
        st.error(pesan_db(e))
        return False
    if ada_isi:
        st.session_state[kunci_sesi] = st.session_state.get(kunci_sesi, 0) + 1
    st.session_state[f"ub_reset_{kunci}"] = True
    return True


def halaman_minta_umpan_balik(eng, nama):
    """Layar yang muncul setelah seminggu pemakaian sampai peserta menilai. 'Lewati' hanya berlaku selama sesi ini."""
    header_halaman("Ulasan dan saran", "Bantu kami memperbaiki aplikasi")
    with st.container(key="pagebody"):
        st.write("Kamu sudah lebih dari seminggu memakai aplikasi ini. Boleh beri penilaian dan saran? "
                 "Hanya sebentar, dan hanya diminta satu kali. Setelah ini, ulasan dan saran bisa kamu tulis kapan saja di menu Profil.")
        if _form_umpan_balik(eng, nama, "awal", minta_rating=True, boleh_lewati=True):
            st.session_state[f"umpan_perlu:{nama}"] = False
            st.session_state["pesan_umpan"] = "Terima kasih atas penilaian dan sarannya!"
            st.rerun()


def _profil_umpan_balik(eng, nama):
    """Bagian 'Ulasan dan saran' di Profil (hanya di aplikasi). Peserta yang belum menilai juga bisa memberi bintang di sini."""
    with st.container(border=True):
        st.markdown("**Ulasan dan saran**")
        try:
            bintang = rating_peserta(eng, nama)
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
            return
        if bintang is not None:
            st.caption(f"Penilaianmu: {_bintang(bintang)}. Terima kasih! Penilaian hanya diminta satu kali.")
        if _form_umpan_balik(eng, nama, "profil", minta_rating=bintang is None, boleh_lewati=False):
            st.session_state[f"umpan_perlu:{nama}"] = False
            st.session_state["pesan_umpan"] = "Terima kasih! Masukanmu sudah kami terima."
            st.rerun()


def halaman_umpan_balik(eng):
    """Admin: ringkasan rating dan daftar ulasan beserta status tindak lanjutnya (bukti iterasi desain)."""
    st.header("💬 Umpan Balik Peserta")
    if not admin_login():
        return
    st.caption("Dikumpulkan lewat aplikasi Android. Ulasan tersimpan tanpa nama (hanya divisi dan tanggal). "
               "Rating hanya ditampilkan sebagai angka gabungan.")
    try:
        ring = ringkasan_rating(eng)
        semua = daftar_ulasan(eng)
    except SQLAlchemyError as e:
        st.error(pesan_db(e))
        return
    k1, k2, k3 = st.columns(3)
    k1.metric("Rata-rata penilaian", "-" if ring["rata"] is None else f"{ring['rata']:.1f} / 5")
    k2.metric("Sudah menilai", f"{ring['n']} dari {ring['n_peserta']}")
    k3.metric("Ulasan masuk", len(semua))
    if ring["n"]:
        urut = (5, 4, 3, 2, 1)
        st.dataframe(pd.DataFrame({"Penilaian": [_bintang(b) for b in urut], "Jumlah": [ring["sebaran"][b] for b in urut]}),
                     width="stretch", hide_index=True)
    if semua.empty:
        st.info("Belum ada ulasan.")
        return
    st.download_button("⬇️ Ulasan (CSV)", semua.to_csv(index=False).encode("utf-8"), file_name="ulasan_peserta.csv",
                       mime="text/csv", key="ub_csv")
    pilih = st.radio("Status", ["(semua)", *STATUS_ULASAN], horizontal=True, key="ub_filter", label_visibility="collapsed")
    df = semua if pilih == "(semua)" else semua[semua["status"] == pilih]
    if df.empty:
        st.info("Tidak ada ulasan dengan status itu.")
        return
    for u in df.itertuples():
        with st.container(border=True):
            st.markdown(f"**{esc(KATEGORI_ULASAN.get(u.kategori, u.kategori))}** · {esc(str(u.divisi or '-'))} · {esc(u.dibuat)}")
            st.write(u.isi)
            kol_status, kol_catatan = st.columns([1, 2])
            status_baru = kol_status.selectbox("Status", STATUS_ULASAN, index=STATUS_ULASAN.index(u.status), key=f"ub_st_{u.id}")
            catatan_lama = u.catatan_admin if isinstance(u.catatan_admin, str) else ""
            catatan = kol_catatan.text_input("Catatan tindak lanjut", value=catatan_lama, max_chars=300, key=f"ub_cat_{u.id}")
            if st.button("Simpan", key=f"ub_simpan_{u.id}"):
                try:
                    ubah_status_ulasan(eng, int(u.id), status_baru, catatan)
                except ValueError as e:
                    st.error(str(e))
                except SQLAlchemyError as e:
                    st.error(pesan_db(e))
                else:
                    st.rerun()


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def keluar():
    """Logout: hapus peran & data login, kembali ke layar login."""
    for k in ("peran", "nama_login", "admin_ok", "tab", "rekap_detail", "drawer_admin", "menu_admin"):
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
{'<img class="auth-logo-img" alt="Bank Sulselbar" src="data:image/png;base64,' + LOGO_BANK_B64 + '">' if LOGO_BANK_B64 else ""}
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
                "Masuk sebagai", ["Peserta PKL", "Pembimbing", "Admin"], horizontal=True, key="login_peran",
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
                    status, menit = proses_login(eng, f"peserta:{nama}", lambda: cek_pin(eng, nama, pin))
                    if status == "ok":
                        st.session_state.peran = "user"
                        st.session_state.nama_login = nama
                        st.rerun()
                    elif status == "terkunci":
                        st.error(pesan_terkunci(menit))
                    else:
                        st.error("PIN salah. Silakan periksa kembali.")
        elif peran == "Pembimbing":
            db = list_pembimbing(eng)
            if db.empty:
                st.warning("Belum ada pembimbing terdaftar. Hubungi admin untuk dibuatkan akun.")
            else:
                with st.form("login_pembimbing", border=False):
                    nama_pb = st.selectbox("Nama pembimbing", db["nama"])
                    pin_pb = st.text_input("PIN", type="password", max_chars=8, placeholder="Masukkan PIN kamu")
                    kirim = st.form_submit_button("Masuk", type="primary", width="stretch")
                if kirim:
                    status, menit = proses_login(eng, f"pembimbing:{nama_pb}", lambda: cek_pin_pembimbing(eng, nama_pb, pin_pb))
                    if status == "ok":
                        st.session_state.peran = "pembimbing"
                        st.session_state.nama_login = nama_pb
                        st.rerun()
                    elif status == "terkunci":
                        st.error(pesan_terkunci(menit))
                    else:
                        st.error("PIN salah. Silakan periksa kembali.")
        else:
            pw_benar = password_admin()
            if pw_benar is None:
                st.error("admin_password belum diatur di Secrets. Hubungi pengelola aplikasi.")
            else:
                if not secret("admin_password") and url_database() is None:
                    st.info("Mode lokal: password admin default = admin123")
                with st.form("login_admin", border=False):
                    pw = st.text_input("Password admin", type="password", placeholder="Masukkan password admin")
                    kirim = st.form_submit_button("Masuk sebagai admin", type="primary", width="stretch")
                if kirim:
                    status, menit = proses_login(eng, "admin", lambda: _sama(pw, pw_benar))
                    if status == "ok":
                        st.session_state.peran = "admin"
                        st.session_state.admin_ok = True
                        st.rerun()
                    elif status == "terkunci":
                        st.error(pesan_terkunci(menit))
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


# ------------------------------------------------------------------
# BERANDA ADMIN: ringkasan peserta aktif dan kehadiran hari ini
# ------------------------------------------------------------------
CSS_DASHBOARD = """
<style>
.db-hero { background: var(--tint); border-radius: 18px; padding: 16px 18px; display: flex; align-items: center; gap: 18px; }
.db-hero-n { font-size: 40px; font-weight: 700; color: var(--biru); line-height: 1; min-width: 56px; }
.db-hero-t { font-size: 17px; font-weight: 700; color: var(--navy); }
.db-hero-s { font-size: 12px; color: var(--muted); margin-top: 2px; }
.db-judul { display: flex; justify-content: space-between; align-items: baseline; gap: 10px; margin: 22px 0 10px; }
.db-judul-t { font-size: 17px; font-weight: 700; color: var(--navy); }
.db-judul-s { font-size: 12px; font-weight: 600; color: var(--biru); text-align: right; }
.db-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; }
.db-k { grid-column: span 2; border-radius: 16px; padding: 12px 14px; }
.db-k.lebar { grid-column: span 3; }
.db-k-n { font-size: 28px; font-weight: 700; line-height: 1.1; }
.db-k-t { font-size: 13px; font-weight: 600; color: var(--navy2); margin-top: 2px; }
.db-hadir { background: #d9f7e8; } .db-hadir .db-k-n { color: #12a15f; }
.db-telat { background: #fff4d6; } .db-telat .db-k-n { color: #b45309; }
.db-izin { background: #dcecf7; } .db-izin .db-k-n { color: #0a5a8f; }
.db-cuti { background: #dff3f0; } .db-cuti .db-k-n { color: #0e8a7d; }
.db-belum { background: #eaf0f5; } .db-belum .db-k-n { color: #3f6785; }
.db-belum.merah { background: #fde2e2; } .db-belum.merah .db-k-n { color: #c0392b; }
.db-libur { background: #fff4d6; color: #b45309; border-radius: 14px; padding: 10px 14px; font-size: 13px; font-weight: 600; margin-bottom: 10px; }
.db-catatan { font-size: 12px; color: var(--muted); margin-top: 12px; line-height: 1.6; }
@media (max-width: 600px) {
  .db-hero-n { font-size: 34px; } .db-k-n { font-size: 24px; } .db-k { padding: 10px 12px; }
  .st-key-menu_admin [role="radiogroup"] { overflow-x: auto; }      /* 8 menu tidak muat di layar sempit: geser, jangan terhimpit */
  .st-key-menu_admin [data-testid="stRadioOption"] { flex: 1 0 auto; padding: 9px 12px; }
}
</style>
"""


def ringkasan_hari_ini(eng, sekarang=None):
    """Hitung kehadiran satu hari untuk Beranda admin. Setiap peserta terdaftar masuk TEPAT SATU kelompok, jadi
    tepat + terlambat + izin + cuti + belum = aktif. Yang dihitung hanya absen yang berhasil (percobaan gagal lokasi
    diabaikan) dan hanya peserta yang masih terdaftar. Bila satu peserta punya beberapa catatan hari itu, absen masuk
    yang menentukan. 'Belum' = belum ada catatan sama sekali; di Sabtu/Minggu nilainya 0 karena memang tidak ada absen."""
    sekarang = sekarang or datetime.now(WITA)
    hari = sekarang.date()
    libur = hari_libur(hari)
    q = (
        select(absensi_t.c.nama, absensi_t.c.jenis, absensi_t.c.jam)
        .select_from(absensi_t.join(peserta_t, peserta_t.c.nama == absensi_t.c.nama))
        .where(and_(absensi_t.c.tanggal == hari.isoformat(), absensi_t.c.status != "gagal_lokasi",
                    absensi_t.c.jenis.in_(("masuk", "izin", "cuti"))))
    )
    with eng.connect() as c:
        aktif = c.execute(select(func.count()).select_from(peserta_t)).scalar() or 0
        baris = c.execute(q).all()
    status = {}
    for nama, jenis, jam in baris:
        if jenis == "masuk":
            status[nama] = "terlambat" if terlambat(jam) else "tepat"
        elif nama not in status:
            status[nama] = jenis
    hitung = {k: sum(1 for v in status.values() if v == k) for k in ("tepat", "terlambat", "izin", "cuti")}
    belum = 0 if libur else max(int(aktif) - len(status), 0)
    return {
        "tanggal": hari, "libur": libur, "aktif": int(aktif), **hitung, "belum": belum,
        "sudah_absen": hitung["tepat"] + hitung["terlambat"],
        # sebelum jendela absen masuk tutup mereka baru "belum absen"; sesudahnya baru pantas disebut tidak hadir
        "label_belum": "Belum absen" if sekarang.time().replace(second=0, microsecond=0) <= JAM_MASUK_TUTUP else "Tidak hadir",
        "waktu": sekarang.strftime("%H:%M"),
    }


def halaman_beranda_admin(eng):
    """Beranda admin: peserta aktif + kehadiran hari ini."""
    st.markdown(CSS_DASHBOARD, unsafe_allow_html=True)
    try:
        r = ringkasan_hari_ini(eng)
    except SQLAlchemyError as e:
        st.error(pesan_db(e))
        return
    merah = " merah" if r["label_belum"] == "Tidak hadir" and r["belum"] else ""
    kartu = [
        ("db-hadir", r["tepat"], "Hadir", ""), ("db-telat", r["terlambat"], "Terlambat", ""), ("db-izin", r["izin"], "Izin", ""),
        ("db-cuti", r["cuti"], "Cuti", " lebar"), (f"db-belum{merah}", r["belum"], r["label_belum"], " lebar"),
    ]
    html = (
        f'<div class="db-hero"><div class="db-hero-n">{r["aktif"]}</div>'
        '<div><div class="db-hero-t">Peserta Aktif</div><div class="db-hero-s">Terdaftar di sistem</div></div></div>'
        '<div class="db-judul"><div class="db-judul-t">Kehadiran Hari Ini</div>'
        f'<div class="db-judul-s">{esc(tgl_indonesia(r["tanggal"]))}</div></div>'
    )
    if r["libur"]:
        html += '<div class="db-libur">Hari ini libur PKL (Sabtu/Minggu), tidak ada absen yang diharapkan.</div>'
    html += '<div class="db-grid">' + "".join(
        f'<div class="db-k {kelas}{lebar}"><div class="db-k-n">{n}</div><div class="db-k-t">{esc(label)}</div></div>'
        for kelas, n, label, lebar in kartu
    ) + "</div>"
    html += (
        f'<div class="db-catatan">Sudah absen masuk: {r["sudah_absen"]} dari {r["aktif"]} peserta. '
        f'Hadir = absen masuk sampai {BATAS_TEPAT_WAKTU:%H:%M}, Terlambat = sesudahnya. '
        f'Diperbarui {r["waktu"]} WITA.</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# Menu admin: (nama menu, ikon Material, judul di header)
MENU_ADMIN = [
    ("Beranda", ":material/home:", "Dashboard Admin"),
    ("Peserta", ":material/group:", "Peserta PKL"),
    ("Rekap", ":material/table_chart:", "Rekap Absensi"),
    ("Analitik", ":material/bar_chart:", "Analitik"),
    ("Model", ":material/psychology:", "Model"),
    ("Proyek", ":material/assignment:", "Proyek"),
    ("Umpan balik", ":material/rate_review:", "Umpan Balik"),
    ("Foto", ":material/photo_camera:", "Foto Bukti"),
]


def _drawer_admin(menu_aktif):
    """Menu hamburger admin: panel geser dari kiri berisi semua menu. Mengikuti desain sidebar Bank Sulselbar."""
    idx = [m[0] for m in MENU_ADMIN].index(menu_aktif)
    st.markdown(
        f"<style>.st-key-drw_{idx} button {{ background: #e9f5d6 !important; color: #2f7d1e !important; }}"
        f".st-key-drw_{idx} button:hover {{ background: #e9f5d6 !important; }}</style>",
        unsafe_allow_html=True,
    )
    with st.container(key="drawerbg"):               # klik area gelap di luar panel = tutup
        if st.button("tutup", key="drw_bg"):
            st.session_state.drawer_admin = False
            st.rerun()
    with st.container(key="drawer"):
        st.markdown(
            f'<div class="dr-head"><div class="dr-top"><span>{ic("menu", 24)}</span><span>{ic("x", 22)}</span></div>'
            '<div class="dr-t">Bank Sulselbar</div><div class="dr-s">Presensi PKL • Admin</div>'
            '<div class="dr-m">MENU UTAMA</div></div>',
            unsafe_allow_html=True,
        )
        with st.container(key="drawerx"):
            if st.button("\u200e", key="drw_x"):
                st.session_state.drawer_admin = False
                st.rerun()
        for i, (nama, ikon, _) in enumerate(MENU_ADMIN):
            if st.button(nama, key=f"drw_{i}", icon=ikon, width="stretch"):
                st.session_state.menu_admin = nama
                st.session_state.drawer_admin = False
                st.session_state.pop("rekap_detail", None)
                st.rerun()
        st.markdown('<div class="dr-garis"></div>', unsafe_allow_html=True)
        if st.button("Keluar", key="drw_keluar", icon=":material/logout:", width="stretch"):
            keluar()
        st.markdown('<div class="dr-foot">© Bank Sulselbar • Absensi PKL</div>', unsafe_allow_html=True)


def halaman_admin(eng):
    """Area admin dengan gaya visual yang selaras dengan halaman peserta."""
    st.markdown(CSS_ADMIN, unsafe_allow_html=True)
    nama_menu = [m[0] for m in MENU_ADMIN]
    menu = st.session_state.get("menu_admin", "Beranda")
    if menu not in nama_menu:
        menu = "Beranda"
    judul = MENU_ADMIN[nama_menu.index(menu)][2]
    st.markdown(
        f"""<div class="admin-head"><div class="admin-top">
<div class="admin-logo">{ic("menu", 28)}</div>
<div><div class="admin-title">{esc(judul)}</div><div class="admin-sub">Bank Sulselbar • Pengelolaan Absensi PKL</div></div>
<div style="flex:1"></div>
<div class="admin-out">{ic("logout", 24)}</div>
</div></div>""",
        unsafe_allow_html=True,
    )
    with st.container(key="adminmenu"):
        if st.button("\u200e", key="buka_menu_admin", width="stretch"):
            st.session_state.drawer_admin = True
            st.rerun()
    with st.container(key="adminlogout"):
        if st.button("\u200e", key="keluar_admin", width="stretch"):
            keluar()
    if st.session_state.get("drawer_admin"):
        _drawer_admin(menu)

    with st.container(key="adminbody"):
        if getattr(eng, "db_sementara", False):
            st.warning(
                "Database masih bersifat SEMENTARA: penyimpanan aplikasi read-only, jadi data disimpan di folder "
                "sementara dan akan HILANG saat aplikasi restart/reboot. Isi `database_url` (Supabase/Postgres) "
                "di Secrets agar data permanen."
            )
        panel_status(eng)
        realtime_admin()
        if menu == "Beranda":
            halaman_beranda_admin(eng)
        elif menu == "Peserta":
            halaman_kelola_peserta(eng)
        elif menu == "Rekap":
            halaman_rekap(eng)
        elif menu == "Analitik":
            halaman_analitik(eng)
        elif menu == "Model":
            halaman_model(eng)
        elif menu == "Proyek":
            halaman_proyek(eng)
        elif menu == "Umpan balik":
            halaman_umpan_balik(eng)
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
        logging.getLogger("absensi").error("Gagal terhubung ke database: %s: %s", type(asal).__name__, str(asal).strip()[:400])
        st.error(
            "Tidak bisa terhubung ke database. Hubungi pengelola aplikasi. "
            "Pengelola: periksa pengaturan database di Secrets (`database_url`, atau `db_host`, `db_user`, `db_password`; "
            "alamat Transaction pooler, pengguna `postgres.KODE_PROJECT`) dan baca detail kesalahan di Logs aplikasi."
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
    elif peran == "pembimbing" and nama and not pkl:
        halaman_pembimbing(eng, nama)
    elif peran == "admin" and st.session_state.get("admin_ok") and not pkl:
        halaman_admin(eng)
    else:
        halaman_login(eng, hanya_peserta=pkl)


if __name__ == "__main__":
    main()
