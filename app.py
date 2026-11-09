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
from functools import lru_cache, wraps
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
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps
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

# Libur nasional dan cuti bersama (tanggal 'YYYY-MM-DD' -> nama). Pada tanggal ini absen masuk, pulang, izin, dan cuti ditolak,
# logbook tidak bisa diisi, dan hari itu tidak dihitung sebagai hari kerja (rekap, proyeksi anggaran, jadwal paket kerja).
# Sumber: SKB 3 Menteri (Menag, Menaker, MenPANRB). 2026: No. 1497/2025, 2/2025, 5/2025. 2027: No. 1205/2026, 3/2026, 2/2026.
# Tanggal yang jatuh pada Sabtu/Minggu tetap dicantumkan supaya daftarnya sama dengan SKB.
# Tahun berikutnya: tambahkan daftar baru di sini setelah SKB-nya terbit (biasanya September tahun sebelumnya).
LIBUR_NASIONAL = {
    # 2026
    "2026-01-01": "Tahun Baru 2026 Masehi", "2026-01-16": "Isra Mikraj Nabi Muhammad SAW",
    "2026-02-17": "Tahun Baru Imlek 2577 Kongzili", "2026-03-19": "Hari Suci Nyepi (Tahun Baru Saka 1948)",
    "2026-03-21": "Idulfitri 1447 Hijriah", "2026-03-22": "Idulfitri 1447 Hijriah",
    "2026-04-03": "Wafat Yesus Kristus", "2026-04-05": "Kebangkitan Yesus Kristus (Paskah)",
    "2026-05-01": "Hari Buruh Internasional", "2026-05-14": "Kenaikan Yesus Kristus",
    "2026-05-27": "Iduladha 1447 Hijriah", "2026-05-31": "Hari Raya Waisak 2570 BE",
    "2026-06-01": "Hari Lahir Pancasila", "2026-06-16": "Tahun Baru Islam 1448 Hijriah",
    "2026-08-17": "Proklamasi Kemerdekaan", "2026-08-25": "Maulid Nabi Muhammad SAW",
    "2026-12-25": "Kelahiran Yesus Kristus (Natal)",
    # 2027
    "2027-01-01": "Tahun Baru 2027 Masehi", "2027-01-05": "Isra Mikraj Nabi Muhammad SAW",
    "2027-02-06": "Tahun Baru Imlek 2578 Kongzili", "2027-03-08": "Hari Suci Nyepi (Tahun Baru Saka 1949)",
    "2027-03-10": "Idulfitri 1448 Hijriah", "2027-03-11": "Idulfitri 1448 Hijriah",
    "2027-03-26": "Wafat Yesus Kristus", "2027-03-28": "Kebangkitan Yesus Kristus (Paskah)",
    "2027-05-01": "Hari Buruh Internasional", "2027-05-06": "Kenaikan Yesus Kristus",
    "2027-05-17": "Iduladha 1448 Hijriah", "2027-05-20": "Hari Raya Waisak 2571 BE",
    "2027-06-01": "Hari Lahir Pancasila", "2027-06-06": "Tahun Baru Islam 1449 Hijriah",
    "2027-08-15": "Maulid Nabi Muhammad SAW", "2027-08-17": "Proklamasi Kemerdekaan",
    "2027-12-25": "Kelahiran Yesus Kristus (Natal)", "2027-12-26": "Isra Mikraj Nabi Muhammad SAW 1449 Hijriah",
}
# Cuti bersama bersifat wajib bagi ASN dan fakultatif bagi swasta. Ubah IKUT_CUTI_BERSAMA menjadi False bila PKL
# tetap masuk pada cuti bersama (hanya libur nasional yang diliburkan).
IKUT_CUTI_BERSAMA = True
CUTI_BERSAMA = {
    # 2026
    "2026-02-16": "Cuti Bersama Tahun Baru Imlek", "2026-03-18": "Cuti Bersama Hari Suci Nyepi",
    "2026-03-20": "Cuti Bersama Idulfitri", "2026-03-23": "Cuti Bersama Idulfitri", "2026-03-24": "Cuti Bersama Idulfitri",
    "2026-05-15": "Cuti Bersama Kenaikan Yesus Kristus", "2026-05-28": "Cuti Bersama Iduladha",
    "2026-12-24": "Cuti Bersama Natal",
    # 2027
    "2027-02-05": "Cuti Bersama Tahun Baru Imlek", "2027-03-09": "Cuti Bersama Idulfitri",
    "2027-03-12": "Cuti Bersama Idulfitri", "2027-03-15": "Cuti Bersama Idulfitri",
    "2027-03-25": "Cuti Bersama Wafat Yesus Kristus", "2027-05-18": "Cuti Bersama Iduladha",
    "2027-05-19": "Cuti Bersama Waisak", "2027-12-24": "Cuti Bersama Natal",
}
HARI_LIBUR_KHUSUS = {**LIBUR_NASIONAL, **(CUTI_BERSAMA if IKUT_CUTI_BERSAMA else {})}
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
        eng = create_engine(url, pool_pre_ping=True, pool_recycle=300, connect_args={"connect_timeout": 10})   # gagal cepat (10 dtk) bila database tidak menjawab, bukan loading tanpa akhir
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


def nama_libur(tgl):
    """Nama libur nasional/cuti bersama pada tanggal itu (date/datetime), atau None kalau bukan tanggal merah."""
    return HARI_LIBUR_KHUSUS.get(tgl.strftime("%Y-%m-%d"))


def hari_libur(tgl):
    """True kalau tanggal (date/datetime) jatuh di hari libur PKL: Sabtu/Minggu atau libur nasional/cuti bersama."""
    return tgl.weekday() in HARI_LIBUR or nama_libur(tgl) is not None


def pesan_libur(tgl):
    if tgl.weekday() not in HARI_LIBUR:
        nama = nama_libur(tgl)
        if nama:
            return (f"{tgl_indonesia(tgl)} libur: {nama}. Absen masuk, pulang, izin, dan cuti "
                    "tidak bisa dilakukan pada hari libur nasional dan cuti bersama.")
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
                if tgl.weekday() >= 5:
                    teks = "Libur "
                else:
                    teks = f"Libur: {nama_libur(tgl)}" if nama_libur(tgl) else "Tidak hadir"
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



# ---------- keterangan tulisan di dalam foto absen (tanggal-jam + alamat), seperti aplikasi kamera stempel waktu ----------
BULAN_FOTO = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
# Alamat yang dicetak di foto, per gedung (satu baris per elemen). Sesuaikan bila alamat Gedung B berbeda dari Gedung A.
ALAMAT_FOTO = {
    "A": ("Jalan Cendrawasih No. 18", "Makassar", "South Sulawesi 90125", "Indonesia"),
    "B": ("Jalan Cendrawasih No. 18", "Makassar", "South Sulawesi 90125", "Indonesia"),
}


def teks_foto(now, gedung):
    """Baris tulisan untuk foto: '6 Nov 2026 at 10.33.54' lalu alamat gedung peserta."""
    waktu = f"{now.day} {BULAN_FOTO[now.month - 1]} {now.year} at {now.strftime('%H.%M.%S')}"
    return [waktu, *ALAMAT_FOTO.get(gedung, ALAMAT_FOTO["A"])]


@lru_cache(maxsize=64)
def _font_foto(ukuran):
    """Font tebal (DejaVu Sans Bold) yang disertakan di folder fonts/ supaya tampilan sama di server mana pun;
    cadangan: font sistem, lalu font bawaan Pillow."""
    folder = os.path.dirname(os.path.abspath(__file__))
    for nama in (os.path.join(folder, "fonts", "DejaVuSans-Bold.ttf"), "DejaVuSans-Bold.ttf",
                 "LiberationSans-Bold.ttf", "Arial Bold.ttf", "arialbd.ttf"):
        try:
            return ImageFont.truetype(nama, ukuran)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=ukuran)
    except (TypeError, OSError):
        return ImageFont.load_default()


def beri_keterangan_foto(img, baris):
    """Cetak `baris` rata kanan di pojok kanan atas foto: putih bersih, font tebal, dengan garis tepi tipis gelap +
    bayangan halus supaya tetap jelas di latar terang maupun gelap. Ukuran huruf menyesuaikan lebar foto supaya
    baris terpanjang tetap muat. Mengubah `img` (RGB) langsung dan mengembalikannya."""
    if not baris:
        return img
    lebar, _ = img.size
    margin = max(8, int(lebar * 0.04))
    ukur = ImageDraw.Draw(img)
    ukuran = max(9, int(lebar * 0.05))
    while ukuran > 9 and max(ukur.textlength(b, font=_font_foto(ukuran)) for b in baris) > lebar - 2 * margin:
        ukuran -= 1
    font = _font_foto(ukuran)
    jarak = int(ukuran * 1.22)
    tebal = max(1, round(ukuran / 16))                      # tebal garis tepi gelap di sekeliling huruf
    posisi = [(lebar - margin - ukur.textlength(b, font=font), margin + i * jarak, b) for i, b in enumerate(baris)]

    bayang = Image.new("L", img.size, 0)                   # bayangan lembut di belakang huruf
    gambar = ImageDraw.Draw(bayang)
    geser = max(1, ukuran // 14)
    for x, y, b in posisi:
        gambar.text((x + geser, y + geser), b, font=font, fill=255, stroke_width=tebal)
    bayang = bayang.filter(ImageFilter.GaussianBlur(max(1.0, ukuran / 10))).point(lambda v: min(255, int(v * 1.3)))
    img.paste(Image.new("RGB", img.size, (0, 0, 0)), (0, 0), bayang)

    gambar = ImageDraw.Draw(img)
    for x, y, b in posisi:                                   # garis tepi gelap dulu, lalu isi putih murni di atasnya
        gambar.text((x, y), b, font=font, fill=(255, 255, 255), stroke_width=tebal, stroke_fill=(0, 0, 0))
    return img


def kompres_foto(data: bytes, keterangan=None) -> bytes:
    """Kecilkan foto (maks 640px, JPEG q80) supaya database tidak cepat penuh. keterangan = daftar baris tulisan yang
    dicetak di dalam foto (lihat teks_foto); dicetak SETELAH dikecilkan agar tulisannya tetap tajam."""
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail((640, 640))
    if keterangan:
        beri_keterangan_foto(img, keterangan)
    out = io.BytesIO()
    img.save(out, format="JPEG", quality=80, optimize=True)
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
LOGO_BANK_B64 = "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JeiJ+JxjJaBqoq0CIQZDQJPiLkWEyJqXXPOYXC2QBbQxSN0BFi8MDb4X1iSiHqI3XNy2qffWrd2Pduc25GQPk0MU0RgRVAoJld+1YSYey+/e1vG7h/fMX9F04WlL3pPUcydvrEtVyG2DN+8u/tlWNoKQsgi1AGtTG7G9jLtj3d7fPtVz+6D3tCArCmNARiYGNMYMZDds2XT9gsTp5634eE1oiYggyqQmSJHHhgq7BLMoCUJfmIPMJg7V9Ym5pRVa/AKCJ8N37/qf+/Z+H22HBcEIOMTkIansSPbIttcumrWCDapJtyVCMc+PrruWrEUiLiKIiBEBcI5cc4QvBtITSgTkY3TPc5RSS5fPufyyr37oA++eM69FRBBVxRoTMHOuvfa2D37wI0TzRDyltOf1L1vW+rurfl1bU+U7c08ilChF23f2vub8d3Z2DirdwFwg0sbrP+mk4++553eVxt7laWTfvu50uhiJRgAMonIcLxweXHXEYpgijhMQ47HWatMze2+95YZ1j267667tA4ObAQoAEQCllDB71VXzTj31sBNPOuKUU85cvrz1oO0YUaTHxsy9995x220PPPts1yOP7Emn9wOmQUIABCChUM3xx6865ZRlb/yXi+fPa2Nmf7oIsPJE3dauXZ1IYlmajUeABddJVUWXLJkzbY17f1e3MXGlCMQgklPIzFtoH75yIRHddtujt95y4z33bH/q6fXG5Es2KhIOJ486asU5Z61+x7ve1tBQzcyEdKA9gl95VEp7Bp54fOszzzz16Lqndu/u2bxlf1fXHsAeEAsghihIoQULFh933IJzX33quee+Wms+EObwEfzmzbtOP/3CTLaNkJA84401N+urfvOTk04+ulCQa/90w01/vufue7f29O733PTatUsvuuDcyveYARwvorxcSk4RFL/LWVLEHow18xdXhPsKgECkWPYoH29uT9sUUqV4rVLnUwhAyQHJxs9FMBNAhHR65x/+UAuQ6Ooe+uY3N3zzm6lXn119/gUNZ7wqXF8PkfA4xmDhwBg02NMikQIAz+v7/dW7Lv16TeAObxCJgQkIgQrG2HNmtZxwYkDaKPc5mVVdnaqudv3NDYNQ4JECbJRW8PTTu77w+WW//CVEIsAgzELg50uASGb7NhWNRFrb/6rpVn+jcixEKXJmMLdNcBShXsDxrVYMswWx+th8BMRS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"
# Logo Bank Sulselbar + Syariah First (PNG, base64) untuk header halaman admin.
LOGO_ADMIN_B64 = "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HpvgtYumfJRaUjsiJgqJDVEADkzA3n/yVKoz2AEA1DqX9pZU5eQTD1u682K1ypxTDXxrQENhBHrbO7628rFNYSawtuHSi6i1USaT+9lvbvrUf/7A9A4j8gFVIZ/GUQs4aorXK3VljGeCrGpeUYAtTE0V4qkq6+46SIon1256/IkN+bwFMYiIQyBkY/zsUWrWg6rePNNoJdBewIraP5ONqjdm3rW6fYlQ2zyike1fb1Rakbts0P2oqukUyrUcAphNWzYTE5TFn0/yx4uYjKpQ6tZJNkqASNRGRYmKKk7ZMJGoTvX1LFpzX88PPvXExgfm9uX6nE4oshTXPnjmIaHYHSMFaYxwBVRIiYihks1kVj+58ea/3HP2WceKOhOXK9aft3ZoKdpZqea7sWVfj+Zz3hlouPGJaD6w5ItVjd6r91U611h5B03wJWItwtD85JdXrX18NBxY/KufXvu731//gnOOfeubX3Logcs5FuKRU4/uTQd3NfauAGMoDLMT+ei1b/zMigdGTP8ckSJAwiQsRKzQk55zKACoKY2WmaF46Quf81+f/e2mEVDog8SsZKFOrAt6Bh64d+Ql//KR3fdeeughe+22y+L5C4dHNm598KHHb7195b13PAnuNsP9EkUxraj6PnEKMoAqBSaMgQMUt4BQwwzgxS858b+/csnIiOOgqCIQBYpO1AwM/uB/LpvbP/hfn35DEHDM/uBVCokIVMBEQWAip29682e++80rzcA85ywRQ+HJGkAkREk6BwSluH8dCYgVjqACY3xUSkg95ND6OBeCoNQaJvmL1vrNpf9WbekkyleOksc5JGUoiI11btcdh848+5lf/syl3NUHG4EtkBEnJuxacf+G57/oP374nXfvt9cSwDonSfgxvjuRGhMwZ77346vf+P8+m3fEGVYnccdYFVVHYiAWcbfC+rEoheePIvJoR1VSAiKojawkBmwVXzjHPQ3BzmHj5vG167ZMThV97RSpAMQcBuy/Xn7oduqHlCJha3i2t51SbCREahlcmmjHTqFN7bxpvVmKQ/brN2yKCccqqgy51GkSnPRQobhhvapTtU4iZ51zTIChzHDXwjuuCX7xhU12/aLurlzk8oSMx5LGhX4xG6SoJAHiuMQhIaePO1i4KLLXXPPXs886Ns0VtN3ckXYmM71qVe0lZx2q0GTz1NtI6f3QUViirJnY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"


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
  background: #fff; z-index: 999; padding: 8px 6px calc(12px + env(safe-area-inset-bottom, 0px)); border-radius: 22px 22px 0 0;
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
/* ---------- loading (pindah halaman ataupun tidak): layar ditutup PUTIH solid sampai skrip selesai, lalu halaman muncul utuh. Jeda 0,15 dtk
   supaya proses yang sangat cepat tidak berkedip. Dipasang di <html> supaya tidak terpengaruh latar biru/gradasi halaman. ---------- */
html::after {
  content: ""; position: fixed; inset: 0; z-index: 99998; background: #fff; opacity: 0; pointer-events: none; transition: opacity .15s ease;
}
html:has([data-testid="stApp"][data-test-script-state="running"])::after { opacity: 1; pointer-events: auto; transition: none; }
html::before {
  content: ""; position: fixed; top: 0; left: 0; height: 3px; width: 0; z-index: 99999; pointer-events: none;
  background: linear-gradient(90deg, #8cc63f, #0a5a8f); opacity: 0;
}
html:has([data-testid="stApp"][data-test-script-state="running"])::before { opacity: 1; width: 85%; transition: width 6s cubic-bezier(.1,.7,.2,1) .15s, opacity 0s .15s; }
/* Halaman absen Masuk/Pulang: layar loading HANYA muncul saat masuk ke halaman ini (atau pindah ke halaman lain).
   Rerun karena kamera disiapkan / foto diambil / lokasi terdeteksi tidak memunculkan layar loading.
   .pr-absen-marker dipasang di halaman absen; atribut data-pindah-halaman dipasang skrip kecil saat navigasi bawah diklik. */
html:has(.pr-absen-marker):not([data-pindah-halaman]):has([data-testid="stApp"][data-test-script-state="running"])::after { opacity: 0; pointer-events: none; }
html:has(.pr-absen-marker):not([data-pindah-halaman]):has([data-testid="stApp"][data-test-script-state="running"])::before { opacity: 0; width: 0; transition: none; }
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
/* latar halaman: kolom tengah (selebar 460px, sama dengan kontainer login) tetap biru tua sampai bawah supaya footer & lengkung hijau
   tetap serasi; sisi kiri-kanan (tampilan web/desktop) putih keabu-abuan lembut, tidak seputih kartu login. Di HP seluruhnya biru tua. */
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"], section.stMain {
  background: linear-gradient(90deg, #e9eff5 0, #e9eff5 calc(50% - 230px), #063f6b calc(50% - 230px),
              #063f6b calc(50% + 230px), #e9eff5 calc(50% + 230px), #e9eff5 100%) !important;
}
/* sembunyikan bilah gulir di halaman login (gulir tetap jalan): bilah menyempitkan area sehingga kontainer bergeser ~7px dari garis tengah latar */
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"], [data-testid="stMain"], section.stMain {
  scrollbar-width: none !important; -ms-overflow-style: none;
}
html::-webkit-scrollbar, body::-webkit-scrollbar, [data-testid="stApp"]::-webkit-scrollbar, [data-testid="stAppViewContainer"]::-webkit-scrollbar,
[data-testid="stMain"]::-webkit-scrollbar, section.stMain::-webkit-scrollbar { width: 0 !important; height: 0 !important; display: none !important; }
[data-testid="stMainBlockContainer"], .block-container {
  max-width: 460px !important; overflow-x: clip; overflow-y: visible !important;   /* jangan 'hidden': mengunci gulir & menutup tombol */
  background: linear-gradient(180deg, #0a5a8f 0, #0a5a8f 480px, #063f6b 100%) !important;
  padding-bottom: calc(24px + env(safe-area-inset-bottom, 0px)) !important;
}
[data-testid="stMarkdownContainer"]:has(.auth-head), [data-testid="stMarkdownContainer"]:has(.auth-foot) { margin: 0 !important; }
.auth-head { background: #fff; height: 478px; box-sizing: border-box; padding: 48px 30px 0; position: relative; z-index: 1; overflow: hidden; }
.auth-head::after {                                   /* tepi biru melengkung (busur) di bawah area putih */
  content: ""; position: absolute; left: 0; right: 0; bottom: -1px; height: 96px; background-color: #0a5a8f;   /* biru naik sampai tepat di bawah teks 'Silakan masuk...' (gap ~13px) */
  /* garis hijau lengkung (sama gayanya dengan lengkung di footer) tepat di atas kartu */
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 96' preserveAspectRatio='none'%3E%3Cpath d='M0 28C120 54 270 56 400 8V19C270 68 120 64 0 38Z' fill='%238cc63f'/%3E%3Cpath d='M0 38C120 64 270 68 400 19' fill='none' stroke='white' stroke-opacity='.9' stroke-width='1.2'/%3E%3C/svg%3E");
  background-size: 100% 96px; background-repeat: no-repeat; background-position: center top;
  border-radius: 0;
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
  width: calc(100% - 44px) !important; margin: -52px 22px 0; background: #fff; border-radius: 22px;
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
[data-testid="stLayoutWrapper"]:has(> .st-key-rt_hidden), [data-testid="stLayoutWrapper"]:has(> .st-key-rt_hidden_js) { position: absolute !important; left: -9999px; width: 1px; height: 1px; overflow: hidden; }
/* elemen yang isinya hanya <style> tidak boleh memakan ruang (tiap elemen menambah jarak antar-elemen) */
[data-testid="stElementContainer"]:has(> [data-testid="stMarkdownContainer"] > style:only-child) { display: none !important; }
.st-key-adminbody {
  width: calc(100% - 36px) !important; margin: -44px 18px 0; background: #fff; border-radius: 26px;
  padding: 14px 22px 26px; position: relative; z-index: 2; box-shadow: 0 8px 28px rgba(6,63,107,.12);
}
@media (max-width: 600px) {
  .admin-head { padding: 28px 18px 80px; }
  .st-key-adminlogout { right: 18px; top: 28px; }
  .st-key-adminbody { padding: 12px 12px 22px; width: calc(100% - 20px) !important; margin-left: 10px; }
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
        nama = nama_libur(now) if now.weekday() not in HARI_LIBUR else None
        ket = f"Libur: {esc(nama)}." if nama else f"{HARI_ID[now.weekday()]} libur PKL."
        return "Hari libur", f"{ket} Absen masuk, pulang, izin, dan cuti tidak tersedia."
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
        foto_kecil = kompres_foto(foto_bytes, teks_foto(now, gedung))
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
        sebab = nama_libur(tgl) if tgl.weekday() not in HARI_LIBUR else None
        sebab = f"libur: {sebab}" if sebab else "libur (Sabtu/Minggu)"
        return _hasil("error", f"{tgl_indonesia(tgl)} jatuh pada hari {sebab}, jadi tidak perlu "
                               f"mengajukan {jenis}. Pilih hari kerja.")
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
        with st.spinner("Memproses absen..."):
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


def _tandai_halaman_absen():
    """Penanda halaman absen Masuk/Pulang untuk CSS_TEMA (lihat komentar di sana). Skrip kecil (iframe tinggi 0, satu asal dengan
    halaman) memasang data-pindah-halaman di <html> saat navigasi bawah diklik, dan melepasnya begitu skrip Streamlit selesai.
    Jadi layar loading tetap muncul saat masuk ke / keluar dari halaman absen, tapi tidak saat kamera disiapkan atau foto diambil."""
    st.markdown('<div class="pr-absen-marker"></div>', unsafe_allow_html=True)
    components.html(
        """<script>
(function () {
  var w; try { w = window.parent; w.document; } catch (e) { return; }
  if (w.__absenNavHook) { return; }
  w.__absenNavHook = true;
  var d = w.document, h = d.documentElement;
  d.addEventListener('click', function (e) {
    if (e.target && e.target.closest && e.target.closest('.st-key-bottomnav')) {
      h.setAttribute('data-pindah-halaman', '1');
    }
  }, true);
  var app = d.querySelector('[data-testid="stApp"]');
  if (app) {
    new MutationObserver(function () {
      if (app.getAttribute('data-test-script-state') !== 'running') { h.removeAttribute('data-pindah-halaman'); }
    }).observe(app, {attributes: true, attributeFilter: ['data-test-script-state']});
  }
})();
</script>""",
        height=0,
    )


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
            _tandai_halaman_absen()


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
    """Bilah navigasi tidak lagi diangkat (laptop/web = tepat di dasar layar, tanpa jarak; HP mengikuti tepi bawah/safe-area).
    Masih bisa dipaksa naik dengan ?angkat=1 bila ikon 'Manage app' Streamlit Cloud menutupi tombol."""
    try:
        return str(st.query_params.get("angkat", "")).strip() == "1"
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
    st.header("Kelola Peserta PKL")
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
        if not nama_semua:
            st.info("Belum ada peserta maupun data absensi.")
        else:
            def _pilih_semua_hapus():
                for n in nama_semua:
                    st.session_state[f"hapus_cb::{n}"] = st.session_state["hapus_pilih_semua"]

            st.checkbox("Pilih semua", key="hapus_pilih_semua", on_change=_pilih_semua_hapus)
            with st.container(height=260):
                for n in nama_semua:
                    st.checkbox(n, key=f"hapus_cb::{n}")
            terpilih = [n for n in nama_semua if st.session_state.get(f"hapus_cb::{n}")]
            st.caption(f"{len(terpilih)} peserta dipilih. Yang dipilih dihapus dari daftar peserta beserta seluruh riwayat "
                       "absensi, logbook, dan foto buktinya (termasuk foto di Supabase Storage). Ulasan dan penilaian yang "
                       "pernah dikirim tetap tersimpan sebagai masukan tanpa nama.")
            konfirmasi = st.text_input("Ketik HAPUS (huruf besar) untuk konfirmasi", key="hapus_absen_konfirmasi")
            if st.button("Hapus", key="tombol_hapus_absen"):
                if not terpilih:
                    st.error("Belum ada peserta yang dipilih.")
                elif konfirmasi.strip() != "HAPUS":
                    st.error("Konfirmasi belum benar. Ketik persis: HAPUS")
                else:
                    berhasil, gagal = [], []
                    for n in terpilih:
                        try:
                            hapus_peserta(eng, n, dengan_absensi=True)
                            berhasil.append(n)
                        except SQLAlchemyError as e:
                            gagal.append((n, pesan_db(e)))
                    for n in berhasil:
                        st.session_state.pop(f"hapus_cb::{n}", None)
                    st.session_state.pop("hapus_pilih_semua", None)
                    if berhasil:
                        notif(f"{len(berhasil)} peserta dihapus beserta seluruh riwayat absensi, logbook, dan fotonya.")
                    if gagal:
                        for n, pesan in gagal:
                            st.error(f"{n}: {pesan}")
                    else:
                        st.rerun()


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

HARI_LIBUR_KHUSUS_NOTE = "Libur nasional dan cuti bersama (daftar LIBUR_NASIONAL dan CUTI_BERSAMA di app.py) tidak dihitung sebagai hari kerja."
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
        st.header("Detail Absensi")
    else:
        st.header("Rekap Absensi & Uang Makan")
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
    return proyek.progres_paket(paket, jam_disetujui_per_paket(eng, divisi), hari_ini or datetime.now(WITA).date(), HARI_LIBUR, HARI_LIBUR_KHUSUS)


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
        errs = proyek.validasi_logbook(tanggal, now.date(), uraian, float(durasi), _jam_logbook_hari(c, nama, tanggal.isoformat()), HARI_LIBUR,
                                     libur_khusus=HARI_LIBUR_KHUSUS)
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
                                       _jam_logbook_hari(c, nama, lama["tanggal"], kecuali_id=logbook_id), HARI_LIBUR, maks_mundur=10_000,
                                       libur_khusus=HARI_LIBUR_KHUSUS)
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
        hk = len(proyek.hari_kerja_bulan(t, b, HARI_LIBUR, HARI_LIBUR_KHUSUS))
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
    return proyek.saran_anggaran(stats, len(proyek.hari_kerja_bulan(tahun, bulan, HARI_LIBUR, HARI_LIBUR_KHUSUS)), UANG_MAKAN, cadangan_pct)


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
    hk = proyek.hari_kerja_bulan(tahun, bulan, HARI_LIBUR, HARI_LIBUR_KHUSUS)
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
    st.header("Analitik Absensi (EDA)")
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
    st.header("Model Prediksi & Deteksi")
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
    st.header("Proyek PKL")
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
    st.header("Bukti Foto Absensi")
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
    st.header("Umpan Balik Peserta")
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


def _aktifkan_ketik_selectbox():
    """Di HP/WebView, Streamlit mengunci kotak pilihan (readonly + tanpa keyboard) sehingga nama tidak bisa dicari dengan mengetik.
    Skrip kecil ini (iframe tinggi 0, satu asal dengan halaman) melepas kuncinya supaya fitur cari bawaan selectbox bisa dipakai."""
    components.html(
        """<script>
(function () {
  var d; try { d = window.parent.document; } catch (e) { return; }
  function lepas() {
    var daftar = d.querySelectorAll('[data-testid="stSelectbox"] input');
    for (var i = 0; i < daftar.length; i++) {
      var el = daftar[i];
      if (el.hasAttribute('readonly')) { el.removeAttribute('readonly'); }
      if (el.readOnly) { el.readOnly = false; }
      if (el.getAttribute('inputmode') !== 'search') { el.setAttribute('inputmode', 'search'); }
      if (el.getAttribute('autocomplete') !== 'off') { el.setAttribute('autocomplete', 'off'); }
    }
  }
  lepas();
  setInterval(lepas, 400);
})();
</script>""",
        height=0,
    )


def halaman_login(eng, hanya_peserta=False):
    """Layar pertama: pilih peran, lalu login. Desain mengikuti tema Presensi peserta.
    hanya_peserta=True (aplikasi PKL): langsung login peserta, tanpa pilihan Admin."""
    st.markdown(CSS_AUTH_ADMIN, unsafe_allow_html=True)
    st.markdown(
        f"""<div class="auth-head">
{'<img class="auth-logo-img" alt="Bank Sulselbar Syariah First" src="data:image/png;base64,' + LOGO_ADMIN_B64 + '">' if LOGO_ADMIN_B64 else ""}
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
                    nama = st.selectbox("Nama peserta", df["nama"], index=None, placeholder="Ketik atau pilih nama peserta")
                    _aktifkan_ketik_selectbox()
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
.db-k { grid-column: span 2; border-radius: 16px; padding: 12px 14px; display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; }
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
  display: flex !important; flex-direction: column; align-items: center !important; justify-content: center; text-align: center;
  width: 100% !important; min-height: 0 !important; height: auto !important; padding: 12px 14px !important;
  border: 0 !important; border-radius: 16px !important; }
.st-key-adminbody .st-key-kartu_db button::before { font-size: 28px; font-weight: 700; line-height: 1.1; width: 100%; text-align: center; }
.st-key-adminbody .st-key-kartu_db button > div { width: 100% !important; display: flex; justify-content: center; text-align: center; }
.st-key-adminbody .st-key-kartu_db button p { font-size: 13px !important; font-weight: 600 !important; color: var(--navy2) !important; margin: 2px 0 0 !important; text-align: center !important; width: 100%; }
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
        nama_l = nama_libur(r["tanggal"]) if r["tanggal"].weekday() not in HARI_LIBUR else None
        ket_l = f"libur: {esc(nama_l)}" if nama_l else "libur PKL (Sabtu/Minggu)"
        html += f'<div class="db-libur">Hari ini {ket_l}, tidak ada absen yang diharapkan.</div>'
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
