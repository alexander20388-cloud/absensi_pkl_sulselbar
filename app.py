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
    Column("universitas", String(150)),
    Column("jurusan", String(150)),
    Column("kelas_semester", String(50)),
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
# SEKALI. Ulasan/saran juga terikat ke peserta (peserta_id) supaya pemberinya bisa melihat status tindak lanjut dan
# catatan admin di menu Profil. Halaman admin dan ekspor CSV tetap hanya menampilkan divisi dan tanggal (tanpa nama
# atau peserta_id). Ulasan lama (sebelum kolom peserta_id ada) tidak bertaut ke siapa pun. Tidak masuk audit saat dibuat.
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
    Column("peserta_id", Integer, ForeignKey("peserta.id", ondelete="SET NULL")),   # NULL: ulasan lama / peserta sudah dihapus
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
        "peserta": [("no_rekening", "VARCHAR(40)"), ("setuju_versi", "VARCHAR(20)"), ("setuju_pada", "VARCHAR(19)"),
                    ("universitas", "VARCHAR(150)"), ("jurusan", "VARCHAR(150)"),
                    ("kelas_semester", "VARCHAR(50)")],
        "absensi": [
            ("foto_path", "VARCHAR(200)"), ("alasan", "VARCHAR(300)"),
            ("peserta_id", "INTEGER REFERENCES peserta(id) ON DELETE SET NULL"),
            ("divisi", "VARCHAR(100)"),
        ],
        "ulasan": [("peserta_id", "INTEGER REFERENCES peserta(id) ON DELETE SET NULL")],
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


def list_peserta_lengkap(eng):
    """Seperti list_peserta, ditambah universitas/sekolah, jurusan & kelas/semester (dipakai halaman Kelola Peserta)."""
    with eng.connect() as c:
        return pd.read_sql(
            select(
                peserta_t.c.nama, peserta_t.c.gedung, peserta_t.c.divisi, peserta_t.c.no_rekening,
                peserta_t.c.universitas, peserta_t.c.jurusan, peserta_t.c.kelas_semester,
            ).order_by(peserta_t.c.nama), c
        )


def get_peserta(eng, nama):
    with eng.connect() as c:
        return c.execute(select(peserta_t).where(peserta_t.c.nama == nama)).mappings().first()


@coba_ulang
def add_peserta(eng, nama, gedung, divisi, pin, no_rekening=None, aktor="admin", peran="admin",
                universitas=None, jurusan=None, kelas_semester=None):
    with eng.begin() as c:
        c.execute(insert(peserta_t).values(
            nama=nama.strip(), gedung=gedung, divisi=divisi.strip(), pin_hash=proyek.hash_pin_kuat(pin),
            no_rekening=(no_rekening or None),
            universitas=((universitas or "").strip() or None), jurusan=((jurusan or "").strip() or None),
            kelas_semester=((kelas_semester or "").strip() or None),
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


@coba_ulang
def set_kampus(eng, nama, universitas, jurusan, kelas_semester=None, aktor="admin", peran="admin"):
    with eng.begin() as c:
        r = c.execute(update(peserta_t).where(peserta_t.c.nama == nama).values(
            universitas=((universitas or "").strip() or None), jurusan=((jurusan or "").strip() or None),
            kelas_semester=((kelas_semester or "").strip() or None),
        ))
        if r.rowcount:
            _audit(c, aktor, peran, "peserta.ubah_kampus", f"peserta:{nama}", {})


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
        if pid is not None:                                     # rating dan ulasan tetap sebagai masukan, tetapi dilepas dari peserta
            c.execute(update(rating_aplikasi_t).where(rating_aplikasi_t.c.peserta_id == pid).values(peserta_id=None))
            c.execute(update(ulasan_t).where(ulasan_t.c.peserta_id == pid).values(peserta_id=None))
        ada = c.execute(delete(peserta_t).where(peserta_t.c.nama == nama)).rowcount
        if dengan_absensi:
            paths = [r[0] for r in c.execute(select(absensi_t.c.foto_path).where(absensi_t.c.nama == nama))]
            n_absen = c.execute(delete(absensi_t).where(absensi_t.c.nama == nama)).rowcount
            n_log = c.execute(delete(logbook_t).where(logbook_t.c.nama == nama)).rowcount
            c.execute(delete(login_gagal_t).where(login_gagal_t.c.kunci == f"peserta:{nama}"))
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


CSS_TABEL_RAPI = """<style>
.tr-wrap { max-height: 360px; overflow: auto; border: 1px solid var(--garis, #d3e2ed); border-radius: 12px; margin: 4px 0 12px; -webkit-overflow-scrolling: touch; }
.tr-wrap table { border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; }
.tr-wrap th:not(:last-child), .tr-wrap td:not(:last-child) { border-right: 1px solid var(--garis, #d3e2ed); }
.tr-wrap th { position: sticky; top: 0; z-index: 2; background: #eef2ff; font-size: 10.5px; font-weight: 700; color: var(--muted, #6b8599);
  padding: 8px 8px; text-align: left; white-space: nowrap; border-bottom: 1px solid var(--garis, #d3e2ed); }
.tr-wrap td { font-size: 12.5px; line-height: 1.3; color: var(--navy, #0b3a5e); padding: 7px 8px; vertical-align: middle;
  border-bottom: 1px solid var(--garis, #d3e2ed); overflow-wrap: break-word; min-width: 48px; }
.tr-wrap td.teks { white-space: normal; }
.tr-wrap td.pendek { white-space: nowrap; }
.tr-wrap tr:last-child td { border-bottom: 0; }
.tr-wrap tbody tr:hover td { background: #f3f8fb; }
.tr-wrap .num { text-align: center; white-space: nowrap; }
.tr-wrap td.idx { color: var(--muted, #6b8599); text-align: center; min-width: 28px; white-space: nowrap; }
</style>"""


def _sel(v):
    """Teks sel tabel: kosong/NaN jadi '-', desimal dirapikan."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else f"{v:.4f}".rstrip("0").rstrip(".")
    return str(v)


def tabel_rapi(df, width=None, hide_index=False, tinggi=360):
    """Tabel bergaya tabel Rekap untuk semua halaman: header menempel, garis tipis, teks membungkus (tidak terpotong),
    gulir sendiri bila panjang/lebar. Pengganti st.dataframe (argumen width/hide_index diterima agar pemanggilan lama tetap jalan)."""
    if df is None or len(df.columns) == 0:
        return
    tampil_idx = (not hide_index) and (df.index.name is not None or not isinstance(df.index, pd.RangeIndex))
    d = df.reset_index() if tampil_idx else df
    if tampil_idx and df.index.name is None:        # indeks tanpa nama (mis. label baris): kepala kolomnya dikosongkan
        d.columns = [""] + [str(c) for c in d.columns[1:]]
    angka = [pd.api.types.is_numeric_dtype(d[c]) for c in d.columns]
    # lebar kolom teks: teks pendek satu baris; teks panjang (nama, divisi) diberi ruang agar muat paling banyak 2 baris
    lebar, pendek = [], []
    for c, a in zip(d.columns, angka):
        n = 0 if (a or d.empty) else int(d[c].map(lambda v: len(_sel(v))).max())
        pendek.append((not a) and n <= 14)
        lebar.append(0 if (a or n <= 14) else int(min(max(n * 3.8, 130), 230)))
    kepala = "".join(
        f'<th class="{"num" if a else ""}"' + (f' style="min-width:{w}px"' if w else "") + f'>{esc(str(c))}</th>'
        for c, a, w in zip(d.columns, angka, lebar)
    )
    baris = []
    for rec in d.itertuples(index=False, name=None):
        sel = []
        for k, (v, a) in enumerate(zip(rec, angka)):
            kelas = "idx" if (tampil_idx and k == 0) else ("num" if a else ("pendek" if pendek[k] else "teks"))
            sel.append(f'<td class="{kelas}">{esc(_sel(v))}</td>')
        baris.append("<tr>" + "".join(sel) + "</tr>")
    st.markdown(
        CSS_TABEL_RAPI
        + f'<div class="tr-wrap" style="max-height:{int(tinggi)}px"><table><thead><tr>{kepala}</tr></thead>'
        + f'<tbody>{"".join(baris)}</tbody></table></div>',
        unsafe_allow_html=True,
    )


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
LOGO_BANK_B64 = "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Sg2uoAFldSepdQALK0kkrqAAldRK6gCQbXUSnuAAldRKW4AFlCV1AASuoldQAJ6lOpU11AAkrqJUagAJ7iZ3AATT1LK6gASU9/0FldQCciSo1EoAoSpiRKtcAJ6rKJKT1ACkrqJXUABK6j1JrUABK6h1LqAA9S1kNrqABZQlbAASVOomnqAAlLcStZAASupZXUACSupZXUACSuoldQAErqWV1AAkrqJXUABK6iVF2AAlRqJT3AAJrqLAAG1rIlTqAAbWsj1LqAQWwlAFCV1JK6gAWUSV1AAsokqdQAHqXUSuoABtRqPUuugAD1LqJWkgAJXUStZAASupZXUABPcjanUAAmtJErSQAnqsrqJXUAKkqNRK6gAPUupZXUACSuuhZXUABK6klPcAA2o1EqdQAErqJTtIACaeuglPcACyupJW7AAsrqSUAA9VM6j1Ke4ACV1LK6gAJXUkrqABZXUnqS1AASuolJgAWV1JbqAAlTqWV1AAnqp6lldQAJK6llaSAA9S6klayAAldR6l1AAsrqRtdf0AAJXUSmtQAEqdRK6gAJWkj1LqABZWsiV1AAkosoABK6j1LqAAldRKjUACTS9xK6gAJU6iU9wAEpasSuoAC3USuoAFldSSuoACV1HqXUACytJJK6gCAlMsqAAJK6llLcACStJErqAAlRqE09wAErqWV1AASupJQAFldRK6gASV1HqUxIAFlE9S6gAWULAASV1LK6gASV1FuoAFldRK6gALdSSgAErWRK6gAJp6lldQAJK6iV1AAKpPcsqdQAJK6iV1AASuoldQALK6iUAAlElPcACyiSuoAFldRKAASiN9GAB//9k="


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
  background: #edf3f8; min-height: 100vh; min-height: 100dvh;
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
.stDeployButton, footer, [class*="viewerBadge"], [class*="StatusWidget"],
[data-testid="stAppDeployButton"], [data-testid="manage-app-button"], [class*="_profileContainer"], [class*="_container_gzau3"],
iframe[title="streamlit_badge"], a[href*="streamlit.io/cloud"] { display: none !important; }
/* ---------- scroll: halaman selalu bisa digulir ke bawah (browser HP maupun WebView) ---------- */
html, body { height: auto !important; min-height: 100%; overflow-x: hidden; overflow-y: auto !important; -webkit-overflow-scrolling: touch; }
/* HP / WebView: bilah gulir klasik memakan lebar di sisi kanan sehingga latar biru tua tampak sebagai garis. Gulir tetap berfungsi. */
@media (max-width: 600px) {
  html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"], section.stMain, section.main {
    scrollbar-width: none !important; scrollbar-gutter: auto !important; -ms-overflow-style: none;
  }
  html::-webkit-scrollbar, body::-webkit-scrollbar, [data-testid="stAppViewContainer"]::-webkit-scrollbar,
  [data-testid="stMain"]::-webkit-scrollbar, section.stMain::-webkit-scrollbar { width: 0 !important; height: 0 !important; display: none !important; }
  [data-testid="stMain"], section.stMain, [data-testid="stMainBlockContainer"], .block-container { width: 100% !important; max-width: 100vw; box-sizing: border-box; }
}
[data-testid="stApp"], [data-testid="stAppViewContainer"] { height: auto !important; min-height: 100dvh; overflow: visible !important; }
[data-testid="stMain"], section.stMain, section.main {
  overflow-y: auto !important; overflow-x: hidden !important; -webkit-overflow-scrolling: touch;
  touch-action: pan-y; overscroll-behavior-y: auto;
}
[data-testid="stMainBlockContainer"], .block-container {
  padding: 0 0 calc(48px + env(safe-area-inset-bottom, 0px)) 0 !important; background: var(--bg); position: relative;
  min-height: 100vh; min-height: 100dvh;
}
/* saat keyboard HP terbuka (input fokus), beri ruang ekstra supaya tombol bisa digulir ke atas keyboard */
[data-testid="stMainBlockContainer"]:has(input:focus), .block-container:has(input:focus) { padding-bottom: 45vh !important; }
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

:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stTextInputRootElement"],
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stNumberInputContainer"] {
  background: #e8f1fa !important; border: 1px solid transparent !important; border-radius: 15px !important;
}
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stTextInputRootElement"]:focus-within {
  border-color: var(--biru) !important; box-shadow: 0 0 0 3px rgba(10,90,143,.16) !important;
}
/* radio -> tombol pil */
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stRadio"] [role="radiogroup"] > div { flex: 1 0 auto; min-width: max-content; }
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stRadioOption"] { width: auto; min-width: max-content; box-sizing: border-box; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadio"] [role="radiogroup"] {
  background: var(--tint); border-radius: 14px; padding: 0; gap: 0; flex-wrap: nowrap; width: 100%; overflow-x: auto; overflow-y: hidden; scrollbar-width: none;
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] {
  flex: 1 0 auto; justify-content: center; align-items: center; border-radius: 14px; padding: 14px 14px; margin: 0 !important;
  cursor: pointer; transition: background .15s ease; min-width: max-content;
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] > div { width: 100%; justify-content: center; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] > div > div:first-child { display: none; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"] p { font-weight: 600; font-size: 15px; color: var(--biru); text-align: center; white-space: nowrap; }
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"][data-selected="true"] {
  background: linear-gradient(135deg, var(--hijau), var(--hijau2)); box-shadow: none;
}
:is(.st-key-loginbody, .st-key-adminbody) [data-testid="stRadioOption"][data-selected="true"] p { color: var(--navy); }

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
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stElementContainer"]:has([data-testid="stRadio"]) { width: 100% !important; }

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
  display: flex; background: var(--tint); border-radius: 14px; padding: 0; gap: 0; flex-wrap: nowrap; width: 100%; overflow-x: auto; overflow-y: hidden; scrollbar-width: none;
}
.st-key-pagebody [data-testid="stRadioOption"] {
  flex: 1 0 auto; justify-content: center; align-items: center; border-radius: 14px; padding: 14px 14px; margin: 0 !important;
  cursor: pointer; min-width: max-content;
}
.st-key-pagebody [data-testid="stRadioOption"] > div { width: 100%; justify-content: center; }
.st-key-pagebody [data-testid="stRadioOption"] > div > div:first-child { display: none; }
.st-key-pagebody [data-testid="stRadioOption"] p { font-weight: 600; font-size: 14px; color: var(--biru); text-align: center; white-space: nowrap; }
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] {
  background: linear-gradient(135deg, var(--hijau), var(--hijau2)); box-shadow: none;
}
.st-key-pagebody [data-testid="stRadioOption"][data-selected="true"] p { color: var(--navy); }
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
/* Tata letak login mengikuti desain: bagian atas putih (~55% layar) dengan tepi bawah melengkung, bagian bawah biru,
   kartu form di perbatasan, teks footer lalu lengkung hijau lime. */
/* seluruh latar halaman login biru tua sampai bawah (warna akhir gradasi), supaya tidak ada pita terang di bawah kartu */
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"], section.stMain {
  background: #063f6b !important;
}
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 460px !important; overflow-x: clip; overflow-y: visible !important;   /* jangan 'hidden': mengunci gulir & menutup tombol */
  background: linear-gradient(180deg, #0a5a8f 0, #0a5a8f 480px, #063f6b 100%) !important;
  padding-bottom: calc(24px + env(safe-area-inset-bottom, 0px)) !important;
}
[data-testid="stMarkdownContainer"]:has(.auth-head), [data-testid="stMarkdownContainer"]:has(.auth-foot) { margin: 0 !important; }
.auth-head { background: #fff; height: 478px; box-sizing: border-box; padding: 48px 30px 0; position: relative; z-index: 1; overflow: hidden; }
.auth-head::after {                                   /* tepi biru melengkung (busur) di bawah area putih */
  content: ""; position: absolute; left: -10%; right: -10%; bottom: -1px; height: 44px; background: #0a5a8f;
  border-radius: 50% 50% 0 0 / 100% 100% 0 0;
}
.auth-logo-img { display: block; width: 78%; max-width: 300px; height: auto; margin: 26px auto 66px; }
.auth-brand { display: flex; align-items: center; gap: 16px; }
.auth-logo { width: 61px; height: 61px; border-radius: 17px; background: #e8f1fa; color: #0a5a8f;
  display: flex; align-items: center; justify-content: center; flex: 0 0 61px; }
.auth-title { font-size: 28px; font-weight: 700; color: #0a4f80; line-height: 1.1; }
.auth-bank { font-size: 11px; color: #6a8499; margin-top: 4px; }
.auth-greet { margin-top: 34px; font-size: 27px; font-weight: 700; color: #0a4f80; line-height: 1.2; }
.auth-desc { margin-top: 6px; font-size: 13px; color: #5f7b92; }
.st-key-loginbody {
  width: calc(100% - 44px) !important; margin: -58px 22px 0; background: #fff; border-radius: 22px;
  padding: 10px 19px 27px; position: relative; z-index: 2; box-shadow: 0 18px 24px -12px rgba(3,30,52,.35);
}
.st-key-loginbody [data-testid="stElementContainer"]:has([data-testid="stRadio"]) { width: 100% !important; }
.st-key-loginbody [data-testid="stRadio"] [role="radiogroup"] { border-radius: 14px; padding: 0; }
.st-key-loginbody [data-testid="stRadio"] { margin-bottom: 8px; }
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stSelectbox"] [data-baseweb="select"] > div {
  background: #e8f1fa !important; border: 1px solid transparent !important; border-radius: 15px !important; min-height: 54px;
}
/* versi Streamlit baru: kotak pilihan tidak lagi memakai data-baseweb */
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stSelectbox"] div:has(> input) {
  background: #e8f1fa !important; border: 1px solid transparent !important; border-radius: 15px !important; min-height: 54px; height: 54px;
}
:is(.st-key-loginbody, .st-key-adminbody, .st-key-pagebody) [data-testid="stSelectbox"] div:has(> div > input) { min-height: 54px; height: 54px; }
[data-testid="stMarkdownContainer"]:has(.seg) { margin: 0 !important; }
.seg { display: flex; height: 54px; margin-bottom: 8px; background: var(--tint); border-radius: 14px; overflow: hidden; }
.seg-on { flex: 1; display: flex; align-items: center; justify-content: center; font-size: 15px; font-weight: 600; color: var(--navy);
  background: linear-gradient(135deg, var(--hijau), var(--hijau2)); }
.st-key-loginbody [data-baseweb="base-input"], .st-key-loginbody [data-baseweb="select"] > div { min-height: 54px; border-radius: 15px !important; }
.st-key-loginbody [data-testid="stForm"] [data-testid="stVerticalBlock"] { gap: 14px !important; }
.st-key-loginbody [data-testid="stWidgetLabel"] p { font-size: 14px; }
.st-key-loginbody [data-testid="stTextInputRootElement"] { height: 54px; }
.st-key-loginbody [data-testid="stTextInputRootElement"] input { height: 52px; font-size: 14px; }
.st-key-loginbody [data-testid="stWidgetLabel"] { margin-bottom: 6px; }
.st-key-loginbody [data-testid="stRadioOption"] { padding: 15px 8px; border-radius: 14px; min-height: 54px; box-sizing: border-box; }
.st-key-loginbody button[kind="primaryFormSubmit"], .st-key-loginbody button[kind="primary"] {
  position: relative; min-height: 60px !important; border-radius: 17px !important; font-size: 15px;
}
.st-key-loginbody button[kind="primaryFormSubmit"]::after, .st-key-loginbody button[kind="primary"]::after {
  content: "\u2192"; position: absolute; right: 20px; top: 50%; transform: translateY(-52%); font-size: 20px; color: #fff;
}
.auth-foot {
  position: relative; z-index: 1; height: 180px; box-sizing: border-box; padding-top: 36px; text-align: center; margin-bottom: 72px;
  color: rgba(255,255,255,.88); font-size: 11px; line-height: 16px;
  background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 180' preserveAspectRatio='none'%3E%3Cpath d='M0 65C120 125 270 130 400 18V40C270 150 120 145 0 83Z' fill='%238cc63f'/%3E%3Cpath d='M0 86C120 148 270 153 400 43' fill='none' stroke='white' stroke-opacity='.9' stroke-width='1.4'/%3E%3C/svg%3E") center top / 100% 180px no-repeat;
}
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

/* ---------- menu hamburger (drawer) ---------- */
.dr-chk { position: absolute; width: 0; height: 0; opacity: 0; pointer-events: none; }
.admin-logo.dr-open { cursor: pointer; }
.dr-bg, .st-key-drawer { display: none; }
.stApp:has(#drtoggle:checked) .dr-bg { display: block; position: fixed; inset: 0; z-index: 9998; background: rgba(4,32,58,.55); cursor: pointer; animation: drFade .18s ease-out; }
.stApp:has(#drtoggle:checked) .st-key-drawer { display: flex; }
.dr-x { position: absolute; top: 14px; right: 14px; width: 40px; height: 40px; display: flex; align-items: center; justify-content: center; cursor: pointer; color: #fff; border-radius: 50%; }
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
.st-key-rt_hidden, .st-key-rt_hidden_js { position: absolute; left: -9999px; width: 1px; height: 1px; overflow: hidden; }
/* elemen yang isinya hanya <style> tidak boleh memakan ruang (tiap elemen menambah jarak antar-elemen) */
[data-testid="stElementContainer"]:has(> [data-testid="stMarkdownContainer"] > style:only-child) { display: none !important; }
.st-key-adminbody {
  width: calc(100% - 36px) !important; margin: -44px 18px 0; background: #fff; border-radius: 26px;
  padding: 22px 22px 26px; position: relative; z-index: 2; box-shadow: 0 8px 28px rgba(6,63,107,.12);
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
            notif("Logbook diperbarui dan diajukan ulang ke pembimbing.")
            st.rerun()
    if k2.button("Hapus", key=f"lb_hapus_{pilih}", width="stretch"):
        try:
            hapus_logbook(eng, int(pilih), nama)
        except ValueError as e:
            st.error(str(e))
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
        else:
            notif("Logbook dihapus.")
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
CSS_PESERTA = """
<style>
/* tabel peserta: gaya sama dengan tabel Rekap (header menempel, garis tipis, gulir sendiri) */
.st-key-tabel_peserta, .st-key-tabel_peserta > [data-testid="stVerticalBlock"] { overflow: auto !important; gap: 0 !important; row-gap: 0 !important; }
.st-key-tabel_peserta { border: 1px solid var(--garis); border-radius: 12px; padding: 0 4px; }
.st-key-tabel_peserta > * { border-bottom: 1px solid var(--garis); }
.st-key-tabel_peserta > :has(.tp-h) { position: sticky; top: 0; z-index: 3; background: #eef2ff; }
.st-key-tabel_peserta > :not(:has(.tp-h)):hover { background: #f3f8fb; }
.st-key-tabel_peserta [data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; min-width: 300px; gap: 0 !important; align-items: stretch !important; padding: 0 2px; margin: 0; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]) { align-self: stretch !important; display: flex !important; flex-direction: column; justify-content: center; padding: 6px 6px !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]):not(:first-child) { border-left: 1px solid var(--garis); }
.st-key-tabel_peserta [data-testid="stElementContainer"] { margin: 0 !important; }
.st-key-tabel_peserta [data-testid="stMarkdownContainer"] { min-width: 0; margin: 0 !important; }
.st-key-tabel_peserta [data-testid="stMarkdownContainer"] p { margin: 0 !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]) { width: auto !important; min-width: 0 !important; flex: 1 1 0 !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]):nth-child(1) { flex: 0.55 1 0 !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]):nth-child(2) { flex: 2.6 1 0 !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]):nth-child(3) { flex: 1.1 1 0 !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]):nth-child(4) { flex: 3 1 0 !important; }
.st-key-tabel_peserta :is([data-testid="stColumn"], [data-testid="column"]):last-child { flex: 0 0 28px !important; }
.tp-h { font-size: 10.5px; font-weight: 700; color: var(--muted); white-space: nowrap; text-transform: uppercase; letter-spacing: .3px; }
.tp-h.tengah, .tp-c.tengah { text-align: center; }
.tp-c { font-size: 12.5px; line-height: 1.3; color: var(--navy); overflow-wrap: anywhere; word-break: break-word; }
/* nama (tombol): tampil seperti teks nama di tabel Rekap, bisa turun baris, tidak terpotong */
.st-key-adminbody .st-key-tabel_peserta button[kind="tertiary"] {
  min-height: 0 !important; height: auto !important; padding: 0 !important; width: 100% !important; justify-content: flex-start !important;
  text-align: left !important; background: transparent !important; border: 0 !important; box-shadow: none !important;
}
.st-key-adminbody .st-key-tabel_peserta button[kind="tertiary"] p {
  white-space: normal !important; overflow-wrap: anywhere; word-break: break-word; text-align: left !important;
  font-size: 12.5px !important; font-weight: 600 !important; line-height: 1.3 !important; color: var(--navy) !important; margin: 0 !important;
}
.st-key-adminbody .st-key-tabel_peserta button[kind="secondary"] {
  min-height: 26px !important; height: 26px !important; padding: 0 !important; width: 100% !important;
  border-radius: 8px !important; border: 1px solid var(--garis) !important; background: var(--tint) !important; box-shadow: none !important;
}
.st-key-adminbody .st-key-tabel_peserta button[kind="secondary"] p { font-size: 16px !important; font-weight: 700; line-height: 1; margin: 0 !important; }
.pf-baris { padding: 10px 0; border-bottom: 1px solid #e3edf4; }
.pf-baris:last-child { border-bottom: 0; }
.pf-l { font-size: 12px; font-weight: 600; color: #6b8599; }
.pf-v { font-size: 15px; font-weight: 600; color: #0b3a5e; overflow-wrap: anywhere; margin-top: 2px; }
</style>
"""


def _kosong(v):
    """Teks siap tampil: nilai kosong/NaN jadi '-'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "-"
    t = str(v).strip()
    return t if t else "-"


def halaman_profil_peserta(df, nama):
    """Halaman profil peserta (halaman penuh, bukan pop-up). Dibuka dengan menekan nama di tabel."""
    if st.button("← Kembali ke daftar peserta", key="profil_kembali"):
        st.session_state.pop("profil_peserta", None)
        st.rerun()
    cocok = df[df["nama"] == nama]
    if cocok.empty:
        st.session_state.pop("profil_peserta", None)
        st.warning("Peserta tidak ditemukan (mungkin sudah dihapus).")
        return
    data = {k: _kosong(v) for k, v in cocok.iloc[0].to_dict().items()}
    st.subheader("Profil Peserta")
    baris = [
        ("Nama", data["nama"]), ("Divisi", data["divisi"]), ("Gedung", data["gedung"]),
        ("No. Rekening", data["no_rekening"]), ("Universitas/Sekolah", data["universitas"]), ("Jurusan", data["jurusan"]),
        ("Kelas/Semester", data["kelas_semester"]),
    ]
    st.markdown(
        "".join(
            f'<div class="pf-baris"><div class="pf-l">{esc(k)}</div><div class="pf-v">{esc(_kosong(v))}</div></div>'
            for k, v in baris
        ),
        unsafe_allow_html=True,
    )


def tabel_peserta(df):
    """Tabel peserta bergaya tabel Rekap: teks membungkus (tidak terpotong), nama/tombol › membuka halaman profil."""
    lebar = [0.55, 2.6, 1.1, 3, 0.4]
    with st.container(key="tabel_peserta", height=min(440, 44 + 52 * len(df))):
        kepala = st.columns(lebar)
        for i, (kol, teks) in enumerate(zip(kepala, ["No", "Nama", "Gedung", "Divisi", ""])):
            kol.markdown(f'<div class="tp-h">{teks}</div>', unsafe_allow_html=True)
        for no, r in enumerate(df.to_dict("records"), start=1):
            kol = st.columns(lebar, vertical_alignment="center")
            kol[0].markdown(f'<div class="tp-c">{no}</div>', unsafe_allow_html=True)
            buka = kol[1].button(str(r["nama"]), key=f"profil_{no}", type="tertiary")
            kol[2].markdown(f'<div class="tp-c">{esc(_kosong(r["gedung"]))}</div>', unsafe_allow_html=True)
            kol[3].markdown(f'<div class="tp-c">{esc(_kosong(r["divisi"]))}</div>', unsafe_allow_html=True)
            buka = kol[4].button("›", key=f"profil_p_{no}", help=f"Profil {r['nama']}") or buka
            if buka:
                st.session_state["profil_peserta"] = r["nama"]
                st.rerun()


def form_tambah_peserta(eng):
    """Form tambah peserta (tampil di halaman, bukan pop-up); hanya muncul setelah tombol Tambah Peserta ditekan."""
    with st.form("tambah_peserta", clear_on_submit=False, border=True):
        nama = st.text_input("Nama peserta")
        divisi = st.selectbox("Divisi", list(DIVISI_GEDUNG.keys()))
        gedung = DIVISI_GEDUNG[divisi]
        univ = st.text_input("Universitas/Sekolah", max_chars=150)
        jurusan = st.text_input("Jurusan", max_chars=150)
        kelas_sem = st.text_input("Kelas/Semester", max_chars=50, placeholder="contoh: Semester 6 atau Kelas XI")
        no_rek = st.text_input("Nomor rekening", placeholder="contoh: 130-210-000554616-9", max_chars=30)
        pin = st.text_input("PIN (4-6 angka)", max_chars=6)
        kirim = st.form_submit_button("Simpan Peserta", width="stretch")

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
                add_peserta(eng, nama, gedung, divisi, pin, rek_rapi, universitas=univ, jurusan=jurusan, kelas_semester=kelas_sem)
            except IntegrityError:
                st.error("Nama itu sudah terdaftar.")
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                catatan = "" if rek_rapi else " Nomor rekening belum diisi, bisa ditambahkan nanti lewat panel Edit di bawah."
                st.session_state["pesan_tambah"] = (
                    f"{nama.strip()} ditambahkan ke Gedung {gedung}. Kasih tahu PIN-nya ke yang bersangkutan.{catatan}"
                )
                st.session_state["tambah_terbuka"] = False      # form disembunyikan lagi
                notif(f"{nama.strip()} berhasil ditambahkan.")
                st.rerun()


def halaman_kelola_peserta(eng):
    st.header("👥 Kelola Peserta PKL")
    if not admin_login():
        return
    st.markdown(CSS_PESERTA, unsafe_allow_html=True)

    profil = st.session_state.get("profil_peserta")
    if profil:
        halaman_profil_peserta(list_peserta_lengkap(eng), profil)
        if st.session_state.get("profil_peserta"):
            return

    terbuka = st.session_state.get("tambah_terbuka", False)
    if st.button("✖ Tutup Form" if terbuka else "➕ Tambah Peserta", key="buka_tambah_peserta"):
        st.session_state["tambah_terbuka"] = not terbuka
        st.rerun()
    pesan_tambah = st.session_state.pop("pesan_tambah", None)
    if pesan_tambah:
        st.success(pesan_tambah)
    if st.session_state.get("tambah_terbuka", False):
        form_tambah_peserta(eng)

    pesan = st.session_state.pop("pesan_hapus", None)
    if pesan:
        st.success(pesan)

    df = list_peserta_lengkap(eng)
    tabel_peserta(df)
    belum = df[df["no_rekening"].isna() | (df["no_rekening"].astype(str).str.strip() == "")]["nama"].tolist()
    if belum:
        with st.container(key="info_rek"):
            with st.popover(f":material/warning: {len(belum)}"):
                st.caption(
                    f"{len(belum)} peserta belum punya nomor rekening: {', '.join(belum)}. "
                    "Isi lewat panel Edit di bawah (di Excel akan tertulis “-” sampai diisi)."
                )

    if not df.empty:
        with st.expander("Edit Divisi / No. Rekening / Universitas/Sekolah, Jurusan & Kelas/Semester / Reset PIN"):
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
                    notif(f"Divisi {target} diubah menjadi {divisi_baru} (Gedung {DIVISI_GEDUNG[divisi_baru]}).")
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
                        notif(f"Nomor rekening {target} diperbarui.")
                        st.rerun()

            univ_lama = _kosong(df.loc[df["nama"] == target, "universitas"].values[0])
            jur_lama = _kosong(df.loc[df["nama"] == target, "jurusan"].values[0])
            ks_lama = _kosong(df.loc[df["nama"] == target, "kelas_semester"].values[0])
            univ_baru = st.text_input("Universitas/Sekolah", value="" if univ_lama == "-" else univ_lama,
                                      max_chars=150, key=f"univ_baru_{target}")
            jur_baru = st.text_input("Jurusan", value="" if jur_lama == "-" else jur_lama,
                                     max_chars=150, key=f"jur_baru_{target}")
            ks_baru = st.text_input("Kelas/Semester", value="" if ks_lama == "-" else ks_lama,
                                    max_chars=50, key=f"ks_baru_{target}")
            if st.button("Simpan Universitas/Sekolah, Jurusan & Kelas/Semester"):
                try:
                    set_kampus(eng, target, univ_baru, jur_baru, ks_baru)
                except SQLAlchemyError as e:
                    st.error(pesan_db(e))
                else:
                    notif(f"Universitas/sekolah, jurusan, dan kelas/semester {target} diperbarui.")
                    st.rerun()

            pin_baru = st.text_input("PIN baru (4-6 angka)", max_chars=6, key="pin_baru")
            if st.button("Reset PIN"):
                if pin_baru.isdigit() and 4 <= len(pin_baru) <= 6:
                    try:
                        reset_pin(eng, target, pin_baru)
                        notif(f"PIN {target} berhasil diperbarui.")
                        st.rerun()
                    except SQLAlchemyError as e:
                        st.error(pesan_db(e))
                else:
                    st.error("PIN harus 4-6 angka.")

    duplikat = [] if getattr(eng, "indeks_unik_ok", True) else cari_duplikat(eng)
    if duplikat:
        with st.expander(f"🧹 Rapikan absensi ganda ({len(duplikat)} kelompok)"):
            st.caption(
                "Ada catatan ganda (nama + tanggal + jenis sama) dari sebelum aturan anti-ganda dipasang. "
                "Yang dipertahankan adalah catatan paling awal; sisanya dihapus beserta fotonya. "
                "Setelah itu indeks anti-ganda dipasang otomatis."
            )
            tabel_rapi(beri_nomor(pd.DataFrame(duplikat, columns=["Nama", "Tanggal", "Jenis", "Jumlah"])), width="stretch")
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
                        notif(f"{n} catatan ganda dihapus. Indeks anti-ganda "
                              + ("terpasang." if eng.indeks_unik_ok else "BELUM terpasang, cek lagi data."))
                        st.rerun()

    with st.expander("🗑️ Hapus peserta dan data absensi (tidak bisa dibatalkan)"):
        nama_semua = sorted(set(df["nama"]) | set(daftar_nama_absensi(eng)))
        semua = "(SEMUA data absensi, peserta tetap ada)"
        if not nama_semua:
            st.info("Belum ada peserta maupun data absensi.")
        else:
            pilihan = st.selectbox("Pilih peserta", [semua] + nama_semua, key="hapus_absen_pilihan")
            if pilihan == semua:
                st.caption("Seluruh riwayat absensi semua peserta dihapus beserta fotonya. Data peserta dan logbook tidak ikut terhapus.")
            else:
                st.caption(f"{pilihan} dihapus dari daftar peserta beserta seluruh riwayat absensi, logbook, dan foto buktinya "
                           "(termasuk foto di Supabase Storage). Ulasan dan penilaian yang pernah ia kirim tetap tersimpan "
                           "sebagai masukan tanpa nama.")
            konfirmasi = st.text_input("Ketik HAPUS (huruf besar) untuk konfirmasi", key="hapus_absen_konfirmasi")
            if st.button("Hapus", key="tombol_hapus_absen"):
                if konfirmasi.strip() != "HAPUS":
                    st.error("Konfirmasi belum benar. Ketik persis: HAPUS")
                else:
                    try:
                        if pilihan == semua:
                            jumlah = hapus_absensi(eng)
                            notif(f"{jumlah} data absensi (beserta fotonya) dihapus.")
                            st.rerun()
                        else:
                            hapus_peserta(eng, pilihan, dengan_absensi=True)
                            notif(f"{pilihan} dihapus beserta seluruh riwayat absensi, logbook, dan fotonya.")
                            st.rerun()
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
.st-key-rekap_tabel [data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; min-width: 300px; gap: 0 !important; align-items: stretch !important; padding: 0 2px; margin: 0; }
.st-key-rekap_tabel :is([data-testid="stColumn"], [data-testid="column"]) { align-self: stretch !important; display: flex !important; flex-direction: column; justify-content: center; padding: 4px 3px !important; }
.st-key-rekap_tabel :is([data-testid="stColumn"], [data-testid="column"]):not(:first-child) { border-left: 1px solid var(--garis); }
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
.rk-h.rk-wrap { white-space: normal; line-height: 1.15; }          /* "Tidak Hadir" boleh turun ke dua baris di layar sempit */
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
    """{nama: tanggal absen pertama} (percobaan gagal lokasi tidak dihitung). Tidak lagi dipakai hitungan 'tidak hadir' rekap."""
    q = select(absensi_t.c.nama, func.min(absensi_t.c.tanggal)).where(absensi_t.c.status != "gagal_lokasi").group_by(absensi_t.c.nama)
    with eng.connect() as c:
        return {n: t for n, t in c.execute(q).all()}


def rekap_peserta_bulan(eng, df, tahun, bulan, sekarang=None):
    """Satu baris per peserta (terdaftar, ditambah nama yang hanya tersisa di riwayat). 'Tidak Hadir' = hari kerja yang
    sudah lewat tanpa absen masuk, izin, atau cuti, dihitung sejak awal bulan; aturannya sama dengan 'Tidak hadir' di Beranda
    (jumlah per hari), hanya dijumlahkan untuk satu bulan. Berlaku untuk peserta yang masih terdaftar, termasuk yang belum
    pernah absen. Nama yang hanya tersisa di riwayat (sudah dihapus) tidak dihitung, sama seperti Beranda.
    Catatan: tanggal daftar tidak disimpan, jadi peserta yang didaftarkan di tengah bulan ikut dihitung sejak tanggal 1."""
    terdaftar = set(list_peserta(eng)["nama"])
    nama_semua = sorted(terdaftar | set(df["nama"].unique()))
    masuk = df[(df["status"] == "berhasil") & (df["jenis"] == "masuk")]
    masuk = masuk.assign(telat=masuk["jam"].map(terlambat).astype(bool))
    pulang = df[(df["status"] == "berhasil") & (df["jenis"] == "pulang")]
    izin, cuti = df[df["jenis"] == "izin"], df[df["jenis"] == "cuti"]
    tgl = lambda d, n: set(d.loc[d["nama"] == n, "tanggal"])
    hari_kerja = [d.isoformat() for d in hari_kerja_periode(tahun, bulan, sekarang)]
    baris = []
    for nama in nama_semua:
        t_hadir, t_izin, t_cuti = tgl(masuk, nama), tgl(izin, nama), tgl(cuti, nama)
        diabsen = t_hadir | t_izin | t_cuti
        baris.append({
            "Nama": nama, "Hari Hadir": len(t_hadir), "Terlambat": len(tgl(masuk[masuk["telat"]], nama)),
            "Izin": len(t_izin), "Cuti": len(t_cuti),
            "Tidak Hadir": sum(1 for d in hari_kerja if d not in diabsen) if nama in terdaftar else 0,
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
    judul = ["Nama", "Hadir", "Telat", "Izin", "Cuti", "Tidak Hadir", "Uang", ""]
    # Tabel punya scroll sendiri (atas-bawah dan kiri-kanan) dengan tinggi dibatasi; halaman tidak ikut memanjang.
    with st.container(key="rekap_tabel", height=min(340, 40 + 35 * len(rekap))):
        for i, (kol, t) in enumerate(zip(st.columns(lebar), judul)):
            kelas = "rk-h kiri" if i == 0 else "rk-h rk-wrap" if t == "Tidak Hadir" else "rk-h"
            bantu = {"Telat": ' title="Terlambat"', "Uang": ' title="Uang makan"'}.get(t, "")
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
        tabel_rapi(beri_nomor(tabel.drop(columns="_tgl")), width="stretch")
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
                tabel_rapi(beri_nomor(ekstrem[["nama", "tanggal", "jam", "jenis", "hasil", "jarak", "gedung"]].rename(columns={
                    "nama": "Nama", "tanggal": "Tanggal", "jam": "Jam", "jenis": "Jenis",
                    "hasil": "Hasil", "jarak": "Jarak (m)", "gedung": "Gedung"})), width="stretch")

    st.subheader("Ringkasan per peserta")
    tabel_rapi(beri_nomor(eda["per_peserta"].rename(columns={
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
        tabel_rapi(pd.DataFrame(cm, index=["Aktual: tidak terlambat", "Aktual: terlambat"],
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
        tabel_rapi(beri_nomor(tampil), width="stretch")

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
        tabel_rapi(tabel, width="stretch", hide_index=True)
        c1, c2 = st.columns([2, 1])
        pilih = c1.selectbox("Aktifkan versi", [v["id"] for v in versi], format_func=lambda i: f"Versi {i}", key="model_pilih")
        if c2.button("Aktifkan", key="model_aktifkan"):
            try:
                aktifkan_model(eng, pilih)
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                notif(f"Model versi {pilih} sekarang aktif.")
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
    tabel_rapi(beri_nomor(pd.DataFrame({
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
    tabel_rapi(pd.DataFrame({
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
                            notif(f"Logbook {r.nama} {keputusan}.")
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
        tabel_rapi(tabel.rename(columns={"kode": "Kode", "nama": "Paket", "mulai": "Mulai", "selesai": "Selesai",
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
                notif(f"Paket kerja {kode.strip()} ditambahkan.")
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
        aksi, pesan_aksi = None, ""
        if b1.button("Simpan perubahan", key=f"pk_ubah_{pid}", type="primary", width="stretch"):
            aksi = lambda: ubah_paket(eng, int(pid), kode_b, nama_b, mulai_b, selesai_b, est_b, aktor, peran, divisi_scope)  # noqa: E731
            pesan_aksi = f"Paket kerja {kode_b.strip()} diperbarui."
        if b2.button("Buka kembali" if cur["ditutup"] else "Tutup (selesai)", key=f"pk_tutup_{pid}", width="stretch"):
            aksi = lambda: tutup_paket(eng, int(pid), not bool(cur["ditutup"]), aktor, peran, divisi_scope)  # noqa: E731
            pesan_aksi = f"Paket kerja {cur['kode']} " + ("dibuka kembali." if cur["ditutup"] else "ditutup (selesai).")
        if b3.button("Hapus", key=f"pk_hapus_{pid}", width="stretch"):
            aksi = lambda: hapus_paket(eng, int(pid), aktor, peran, divisi_scope)  # noqa: E731
            pesan_aksi = f"Paket kerja {cur['kode']} dihapus."
        if aksi is not None:
            try:
                aksi()
            except (ValueError, PermissionError) as e:
                st.error(str(e))
            except SQLAlchemyError as e:
                st.error(pesan_db(e))
            else:
                notif(pesan_aksi)
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
    tabel_rapi(tampil, width="stretch", hide_index=True)
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
            tabel_rapi(saran.assign(
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
    tabel_rapi(df.rename(columns={"nama": "Nama", "divisi": "Divisi", "dibuat": "Dibuat"}), width="stretch", hide_index=True)
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
                notif(f"Pembimbing {target} dihapus.")
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
    tabel_rapi(df.rename(columns={"id": "ID", "waktu": "Waktu (WITA)", "aktor": "Aktor", "peran": "Peran", "aksi": "Aksi",
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
ULASAN_MAKS_PER_SESI = 3                # penahan spam ringan per sesi (belum dibatasi per peserta di DB)


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
    """Rating SEKALI seumur akun (tidak bisa diubah). Ulasan opsional ikut disimpan dan bertaut ke peserta. Salah -> ValueError."""
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
                c.execute(insert(ulasan_t).values(peserta_id=pes[0], divisi=pes[1], kategori=kategori, isi=isi.strip(),
                                                  dibuat=tanggal, status="baru"))
    except IntegrityError:
        raise ValueError("Kamu sudah memberi penilaian. Terima kasih!") from None


@coba_ulang
def simpan_ulasan(eng, nama, isi, kategori, now=None):
    """Ulasan atau saran dari Profil (boleh berkali-kali). Bertaut ke peserta lewat peserta_id; nama tidak disalin ke tabel ulasan."""
    salah = validasi_ulasan(isi, kategori)
    if salah:
        raise ValueError(" ".join(salah))
    tanggal = (now or datetime.now(WITA)).strftime("%Y-%m-%d")
    with eng.begin() as c:
        pes = c.execute(select(peserta_t.c.id, peserta_t.c.divisi).where(peserta_t.c.nama == nama)).first()
        if pes is None:
            raise ValueError("Data kamu tidak ditemukan.")
        c.execute(insert(ulasan_t).values(peserta_id=pes[0], divisi=pes[1], kategori=kategori, isi=isi.strip(),
                                          dibuat=tanggal, status="baru"))


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
    """Untuk admin. peserta_id sengaja tidak ikut dipilih supaya tidak muncul di layar maupun CSV."""
    q = (select(ulasan_t.c.id, ulasan_t.c.divisi, ulasan_t.c.kategori, ulasan_t.c.isi, ulasan_t.c.dibuat,
                ulasan_t.c.status, ulasan_t.c.catatan_admin)
         .order_by(ulasan_t.c.id.desc()).limit(limit))
    if status:
        q = q.where(ulasan_t.c.status == status)
    with eng.connect() as c:
        return pd.read_sql(q, c)


STATUS_ULASAN_PESERTA = {"baru": "Diterima", "diproses": "Sedang diproses", "selesai": "Selesai"}


def ulasan_peserta(eng, nama, limit=50):
    """Ulasan milik peserta sendiri beserta status dan catatan admin, terbaru dulu. Ulasan lama tidak muncul (tidak bertaut)."""
    q = (select(ulasan_t.c.kategori, ulasan_t.c.isi, ulasan_t.c.dibuat, ulasan_t.c.status, ulasan_t.c.catatan_admin)
         .select_from(ulasan_t.join(peserta_t, peserta_t.c.id == ulasan_t.c.peserta_id))
         .where(peserta_t.c.nama == nama)
         .order_by(ulasan_t.c.id.desc()).limit(limit))
    with eng.connect() as c:
        return [dict(r._mapping) for r in c.execute(q)]


@coba_ulang
def ubah_status_ulasan(eng, ulasan_id, status, catatan="", aktor="admin"):
    """Admin menandai tindak lanjut ulasan (baru, diproses, selesai). Tercatat di audit. Status dan catatan terlihat oleh pemberi ulasan di Profil."""
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
    st.caption("Ulasan terhubung ke akunmu supaya kamu bisa melihat tindak lanjutnya di Profil. "
               "Di halaman admin, ulasan hanya tampil dengan divisi dan tanggal, tanpa namamu.")
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
                 "Hanya sebentar, dan hanya diminta satu kali. Setelah ini, ulasan dan saran bisa kamu tulis kapan saja di menu Profil, dan status tindak lanjutnya bisa kamu lihat di sana.")
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
        try:
            milikku = ulasan_peserta(eng, nama)
        except SQLAlchemyError as e:
            st.error(pesan_db(e))
            return
        if milikku:
            with st.expander(f"Ulasan saya ({len(milikku)})"):
                for u in milikku:
                    with st.container(border=True):
                        st.markdown(f"**{esc(KATEGORI_ULASAN.get(u['kategori'], u['kategori']))}** · {esc(u['dibuat'])} · "
                                    f"{esc(STATUS_ULASAN_PESERTA.get(u['status'], u['status']))}")
                        st.write(u["isi"])
                        tanggapan = u["catatan_admin"]
                        if isinstance(tanggapan, str) and tanggapan.strip():
                            st.caption("Tanggapan admin: " + tanggapan.strip())


def halaman_umpan_balik(eng):
    """Admin: ringkasan rating dan daftar ulasan beserta status tindak lanjutnya (bukti iterasi desain)."""
    st.header("💬 Umpan Balik Peserta")
    if not admin_login():
        return
    st.caption("Dikumpulkan lewat aplikasi Android. Di halaman ini ulasan hanya tampil dengan divisi dan tanggal, tanpa nama "
               "pengirim. Status dan catatan yang kamu isi akan terlihat oleh pemberi ulasan di menu Profilnya. "
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
        tabel_rapi(pd.DataFrame({"Penilaian": [_bintang(b) for b in urut], "Jumlah": [ring["sebaran"][b] for b in urut]}),
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
            catatan = kol_catatan.text_input("Catatan tindak lanjut (terlihat oleh pemberi ulasan)", value=catatan_lama, max_chars=300, key=f"ub_cat_{u.id}")
            if st.button("Simpan", key=f"ub_simpan_{u.id}"):
                try:
                    ubah_status_ulasan(eng, int(u.id), status_baru, catatan)
                except ValueError as e:
                    st.error(str(e))
                except SQLAlchemyError as e:
                    st.error(pesan_db(e))
                else:
                    notif(f"Status ulasan disimpan: {status_baru}.")
                    st.rerun()


# ------------------------------------------------------------------
# MAIN
# ------------------------------------------------------------------
def notif(teks, ikon="✅"):
    """Antrikan pemberitahuan (toast) yang tampil di run berikutnya. Wajib dipakai sebelum st.rerun():
    pesan st.success yang ditulis tepat sebelum rerun langsung hilang dan tidak sempat terbaca."""
    st.session_state.setdefault("_notif", []).append((teks, ikon))


def tampilkan_notif():
    for teks, ikon in st.session_state.pop("_notif", []):
        st.toast(teks, icon=ikon)


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
            st.markdown('<div class="seg"><div class="seg-on">Peserta PKL</div></div>', unsafe_allow_html=True)
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
                    nama = st.selectbox("Nama peserta", df["nama"], index=None, placeholder="Pilih nama peserta")
                    pin = st.text_input("PIN", type="password", max_chars=6, placeholder="Masukkan PIN kamu")
                    kirim = st.form_submit_button("Masuk", type="primary", width="stretch")
                if kirim and not nama:
                    st.error("Pilih nama peserta dulu.")
                elif kirim:
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
                    nama_pb = st.selectbox("Nama pembimbing", db["nama"], index=None, placeholder="Pilih nama pembimbing")
                    pin_pb = st.text_input("PIN", type="password", max_chars=8, placeholder="Masukkan PIN kamu")
                    kirim = st.form_submit_button("Masuk", type="primary", width="stretch")
                if kirim and not nama_pb:
                    st.error("Pilih nama pembimbing dulu.")
                elif kirim:
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
        return "🔕 Realtime belum aktif (isi `supabase_url` dan `supabase_anon_key` di Secrets)."
    with st.container(key="rt_hidden"):
        if st.button("segarkan", key="rt_refresh"):
            st.toast("Ada data absensi baru. Dashboard diperbarui.", icon="🔔")
    with st.container(key="rt_hidden_js"):
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
    return "🔔 Realtime aktif: dashboard diperbarui otomatis saat ada absen/izin baru."


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
.st-key-kartu_db[data-testid="stVerticalBlock"], .st-key-kartu_db > [data-testid="stVerticalBlock"] {
  display: grid !important; grid-template-columns: repeat(6, 1fr); gap: 10px !important; }
.st-key-kartu_db [data-testid="stElementContainer"] { grid-column: span 2; margin: 0 !important; width: auto !important; }
.st-key-kartu_db .st-key-kartu_cuti, .st-key-kartu_db .st-key-kartu_belum { grid-column: span 3; }
.st-key-adminbody .st-key-kartu_db button {
  display: flex !important; flex-direction: column; align-items: flex-start !important; justify-content: center; text-align: left;
  width: 100% !important; min-height: 0 !important; height: auto !important; padding: 12px 14px !important;
  border: 0 !important; border-radius: 16px !important; }
.st-key-adminbody .st-key-kartu_db button::before { font-size: 28px; font-weight: 700; line-height: 1.1; }
.st-key-adminbody .st-key-kartu_db button p { font-size: 13px !important; font-weight: 600 !important; color: var(--navy2) !important; margin: 2px 0 0 !important; text-align: left; }
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
        select(absensi_t.c.nama, absensi_t.c.jenis, absensi_t.c.jam, absensi_t.c.alasan)
        .select_from(absensi_t.join(peserta_t, peserta_t.c.nama == absensi_t.c.nama))
        .where(and_(absensi_t.c.tanggal == hari.isoformat(), absensi_t.c.status != "gagal_lokasi",
                    absensi_t.c.jenis.in_(("masuk", "izin", "cuti"))))
    )
    with eng.connect() as c:
        aktif = c.execute(select(func.count()).select_from(peserta_t)).scalar() or 0
        baris = c.execute(q).all()
    status, jam_masuk, alasan_izin = {}, {}, {}
    for nama, jenis, jam, alasan in baris:
        if jenis == "masuk":
            status[nama] = "terlambat" if terlambat(jam) else "tepat"
            jam_masuk[nama] = jam
        elif nama not in status:
            status[nama] = jenis
            alasan_izin[nama] = alasan
    hitung = {k: sum(1 for v in status.values() if v == k) for k in ("tepat", "terlambat", "izin", "cuti")}
    belum = 0 if libur else max(int(aktif) - len(status), 0)
    return {
        "tanggal": hari, "libur": libur, "aktif": int(aktif), **hitung, "belum": belum,
        "sudah_absen": hitung["tepat"] + hitung["terlambat"],
        # sebelum jendela absen masuk tutup mereka baru "belum absen"; sesudahnya baru pantas disebut tidak hadir
        "label_belum": "Belum absen" if sekarang.time().replace(second=0, microsecond=0) <= JAM_MASUK_TUTUP else "Tidak hadir",
        "waktu": sekarang.strftime("%H:%M"),
        "status": status, "jam_masuk": jam_masuk, "alasan": alasan_izin,      # untuk daftar nama saat kotak diketuk
    }


def _daftar_kartu_beranda(eng, r, kode, label):
    """Daftar nama di balik kotak yang diketuk di Beranda (hadir / terlambat / izin / cuti / belum absen)."""
    try:
        pes = list_peserta(eng)
    except SQLAlchemyError as e:
        st.error(pesan_db(e))
        return
    status, jam = r["status"], r["jam_masuk"]
    if kode == "hadir":
        pilih = [n for n in pes["nama"] if status.get(n) in ("tepat", "terlambat")]
    elif kode == "telat":
        pilih = [n for n in pes["nama"] if status.get(n) == "terlambat"]
    elif kode in ("izin", "cuti"):
        pilih = [n for n in pes["nama"] if status.get(n) == kode]
    else:
        pilih = [] if r["libur"] else [n for n in pes["nama"] if n not in status]
    d = pes[pes["nama"].isin(pilih)].copy()
    if kode in ("hadir", "telat"):
        d["Jam masuk"] = d["nama"].map(lambda n: str(jam.get(n, "-")))
        d = d.sort_values(["Jam masuk", "nama"])
        if kode == "hadir":
            d["Status"] = d["nama"].map(lambda n: "Terlambat" if status.get(n) == "terlambat" else "Tepat waktu")
    elif kode in ("izin", "cuti"):
        d["Keterangan"] = d["nama"].map(lambda n: (r["alasan"].get(n) or "-"))
    d = d.rename(columns={"nama": "Nama", "divisi": "Divisi", "gedung": "Gedung"})
    urut = [c for c in ("Nama", "Divisi", "Gedung", "Jam masuk", "Status", "Keterangan") if c in d.columns]
    st.markdown(f"**{esc(label)} ({len(d)})**")
    if d.empty:
        st.caption("Belum ada.")
    else:
        tabel_rapi(beri_nomor(d[urut]))


def halaman_beranda_admin(eng):
    """Beranda admin: peserta aktif + kehadiran hari ini."""
    st.markdown(CSS_DASHBOARD, unsafe_allow_html=True)
    try:
        r = ringkasan_hari_ini(eng)
    except SQLAlchemyError as e:
        st.error(pesan_db(e))
        return
    merah = r["label_belum"] == "Tidak hadir" and bool(r["belum"])
    # (kode, jumlah, label, warna latar, warna angka)
    kartu = [
        ("hadir", r["sudah_absen"], "Hadir", "#d9f7e8", "#12a15f"), ("telat", r["terlambat"], "Terlambat", "#fff4d6", "#b45309"),
        ("izin", r["izin"], "Izin", "#dcecf7", "#0a5a8f"), ("cuti", r["cuti"], "Cuti", "#dff3f0", "#0e8a7d"),
        ("belum", r["belum"], r["label_belum"], "#fde2e2" if merah else "#eaf0f5", "#c0392b" if merah else "#3f6785"),
    ]
    html = (
        f'<div class="db-hero"><div class="db-hero-n">{r["aktif"]}</div>'
        '<div><div class="db-hero-t">Peserta Aktif</div><div class="db-hero-s">Terdaftar di sistem</div></div></div>'
        '<div class="db-judul"><div class="db-judul-t">Kehadiran Hari Ini</div>'
        f'<div class="db-judul-s">{esc(tgl_indonesia(r["tanggal"]))}</div></div>'
    )
    if r["libur"]:
        html += '<div class="db-libur">Hari ini libur PKL (Sabtu/Minggu), tidak ada absen yang diharapkan.</div>'
    st.markdown(html, unsafe_allow_html=True)
    pilih = st.session_state.get("db_kartu")
    css = ["<style>"]
    for kode, n, label, bg, fg in kartu:
        cincin = f"box-shadow: 0 0 0 2px {fg} inset !important;" if pilih == kode else "box-shadow: none !important;"
        css.append(
            f'.st-key-adminbody .st-key-kartu_db .st-key-kartu_{kode} button {{ background: {bg} !important; {cincin} }}'
            f'.st-key-adminbody .st-key-kartu_db .st-key-kartu_{kode} button::before {{ content: "{n}"; color: {fg}; }}'
        )
    css.append("</style>")
    st.markdown("".join(css), unsafe_allow_html=True)
    with st.container(key="kartu_db"):
        for kode, n, label, bg, fg in kartu:
            if st.button(label, key=f"kartu_{kode}", width="stretch"):
                st.session_state["db_kartu"] = None if pilih == kode else kode
                st.rerun()
    st.caption("Ketuk salah satu kotak untuk melihat daftar namanya. Ketuk lagi untuk menutup.")
    with st.container(key="info_catatan"):
        with st.popover(":material/info:"):
            st.caption(
                f'Sudah absen masuk: {r["sudah_absen"]} dari {r["aktif"]} peserta. '
                f'Hadir = semua yang sudah absen masuk (tepat waktu maupun terlambat). '
                f'Terlambat = absen masuk sesudah {BATAS_TEPAT_WAKTU:%H:%M}. '
                f'Diperbarui {r["waktu"]} WITA.'
            )
    if pilih:
        _daftar_kartu_beranda(eng, r, pilih, {k: lb for k, _, lb, _, _ in kartu}[pilih])


CSS_INFO = """
<style>
/* iframe realtime (tinggi 0) jangan memakan ruang/jarak di halaman */
.st-key-adminbody [data-testid="stElementContainer"]:has(iframe[height="0"], iframe[style*="height: 0"]) { position: absolute !important; height: 0 !important; width: 0 !important; margin: 0 !important; overflow: hidden !important; }
/* tombol keterangan: bulat kecil, rata kanan, serasi dengan kartu dashboard */
.st-key-info_sistem, .st-key-info_catatan, .st-key-info_rek { display: flex; justify-content: flex-end; margin: 0 0 4px; }
.st-key-info_catatan, .st-key-info_rek { margin: 8px 0 0; }
.st-key-adminbody .st-key-info_sistem button, .st-key-adminbody .st-key-info_catatan button, .st-key-adminbody .st-key-info_rek button {
  width: auto !important; min-width: 0 !important; min-height: 0 !important; height: 34px !important; padding: 0 11px !important; gap: 4px;
  border-radius: 999px !important; background: var(--tint, #eaf3fa) !important; border: 0 !important; color: var(--biru, #0a5a8f) !important;
  box-shadow: none !important; font-size: 13px !important; font-weight: 700 !important;
}
.st-key-adminbody .st-key-info_sistem button:hover, .st-key-adminbody .st-key-info_catatan button:hover, .st-key-adminbody .st-key-info_rek button:hover {
  background: #d7e8f5 !important;
}
.st-key-info_sistem button:has(> [data-testid="stMarkdownContainer"]) > [data-testid="stIconMaterial"]:last-child,
.st-key-info_catatan button:has(> [data-testid="stMarkdownContainer"]) > [data-testid="stIconMaterial"]:last-child,
.st-key-info_rek button:has(> [data-testid="stMarkdownContainer"]) > [data-testid="stIconMaterial"]:last-child { display: none !important; }
.st-key-info_sistem [data-testid="stMarkdownContainer"] p, .st-key-info_catatan [data-testid="stMarkdownContainer"] p,
.st-key-info_rek [data-testid="stMarkdownContainer"] p { margin: 0 !important; font-size: 13px !important; color: var(--biru, #0a5a8f) !important; }
</style>
"""


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
    """Isi menu hamburger. Buka/tutup panel terjadi di browser (kotak centang CSS), jadi tanpa loading;
    server baru dipanggil saat salah satu menu dipilih."""
    idx = [m[0] for m in MENU_ADMIN].index(menu_aktif)
    st.markdown(
        f"<style>.st-key-drw_{idx} button {{ background: #e9f5d6 !important; color: #2f7d1e !important; }}"
        f".st-key-drw_{idx} button:hover {{ background: #e9f5d6 !important; }}</style>"
        '<label for="drtoggle" class="dr-bg"></label>',
        unsafe_allow_html=True,
    )
    with st.container(key="drawer"):
        st.markdown(
            f'<div class="dr-head"><div class="dr-top"><span>{ic("menu", 24)}</span></div>'
            f'<label for="drtoggle" class="dr-x">{ic("x", 22)}</label>'
            '<div class="dr-t">Bank Sulselbar</div><div class="dr-s">Presensi PKL • Admin</div>'
            '<div class="dr-m">MENU UTAMA</div></div>',
            unsafe_allow_html=True,
        )
        for i, (nama, ikon, _) in enumerate(MENU_ADMIN):
            if st.button(nama, key=f"drw_{i}", icon=ikon, width="stretch"):
                st.session_state.menu_admin = nama
                st.session_state.drawer_n = st.session_state.get("drawer_n", 0) + 1   # membuat panel tertutup kembali
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
    # Tag pembungkus bergantian tiap menu dipilih: React memasang ulang elemen ini sehingga kotak centang kembali kosong (panel tertutup).
    tag = "div" if st.session_state.get("drawer_n", 0) % 2 == 0 else "section"
    st.markdown(
        f"""<{tag} class="admin-head"><input type="checkbox" id="drtoggle" class="dr-chk" aria-hidden="true"><div class="admin-top">
<label for="drtoggle" class="admin-logo dr-open" aria-label="Buka menu">{ic("menu", 28)}</label>
<div><div class="admin-title">{esc(judul)}</div><div class="admin-sub">Bank Sulselbar • Pengelolaan Absensi PKL</div></div>
<div style="flex:1"></div>
<div class="admin-out">{ic("logout", 24)}</div>
</div></{tag}>""",
        unsafe_allow_html=True,
    )
    with st.container(key="adminlogout"):
        if st.button("\u200e", key="keluar_admin", width="stretch"):
            keluar()
    _drawer_admin(menu)

    with st.container(key="adminbody"):
        if getattr(eng, "db_sementara", False):
            st.warning(
                "Database masih bersifat SEMENTARA: penyimpanan aplikasi read-only, jadi data disimpan di folder "
                "sementara dan akan HILANG saat aplikasi restart/reboot. Isi `database_url` (Supabase/Postgres) "
                "di Secrets agar data permanen."
            )
        st.markdown(CSS_INFO, unsafe_allow_html=True)
        ket_realtime = realtime_admin()
        if menu == "Beranda":
            with st.container(key="info_sistem"):       # tombol keterangan sistem: hanya di Beranda
                with st.popover(":material/info:", help="Status sistem"):
                    panel_status(eng)
                    st.caption(ket_realtime)
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
    tampilkan_notif()
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
