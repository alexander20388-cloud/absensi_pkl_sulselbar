"""Tes: periode PKL per peserta (mulai_pkl / selesai_pkl). 'Tidak Hadir' dan 'Belum absen' hanya dihitung di dalam periode.
Jalankan: pytest -q test_periode_pkl.py"""
from datetime import date, datetime

import pandas as pd
import pytest
from sqlalchemy import create_engine, text

import app as A

# Rabu 11 Nov 2026 pukul 18:00. Hari kerja November sampai hari itu: 2, 3, 4, 5, 6, 9, 10, 11 = 8 hari (1 Nov Minggu; 7-8 Nov Sabtu-Minggu).
SEKARANG = datetime(2026, 11, 11, 18, 0, tzinfo=A.WITA)


def _rekap(eng, nama):
    df = A.get_absensi_bulan(eng, 2026, 11)
    if df.empty:
        df = pd.DataFrame(columns=["nama", "tanggal", "jam", "jenis", "status", "jarak", "uang_makan", "alasan", "tinjau_status"])
    r = A.rekap_peserta_bulan(eng, df, 2026, 11, sekarang=SEKARANG)
    return r[r["Nama"] == nama].iloc[0]


def test_tanpa_periode_dihitung_sejak_awal_bulan_seperti_dulu(eng_kosong):
    assert _rekap(eng_kosong, "P0")["Tidak Hadir"] == 8


def test_mulai_di_tengah_bulan_tidak_menghitung_hari_sebelumnya(eng_kosong):
    A.set_periode_pkl(eng_kosong, "P0", date(2026, 11, 9), None)
    assert _rekap(eng_kosong, "P0")["Tidak Hadir"] == 3                  # 9, 10, 11 Nov
    assert _rekap(eng_kosong, "P1")["Tidak Hadir"] == 8                  # peserta lain tidak terpengaruh


def test_selesai_menghentikan_hitungan(eng_kosong):
    A.set_periode_pkl(eng_kosong, "P0", date(2026, 11, 2), date(2026, 11, 5))
    assert _rekap(eng_kosong, "P0")["Tidak Hadir"] == 4                  # 2, 3, 4, 5 Nov


def test_periode_di_masa_depan_nol(eng_kosong):
    A.set_periode_pkl(eng_kosong, "P0", date(2026, 12, 1), date(2026, 12, 31))
    assert _rekap(eng_kosong, "P0")["Tidak Hadir"] == 0


def test_validasi_dan_audit(eng_kosong):
    with pytest.raises(ValueError):
        A.set_periode_pkl(eng_kosong, "P0", date(2026, 11, 10), date(2026, 11, 1))
    with pytest.raises(ValueError):
        A.add_peserta(eng_kosong, "Baru", "A", "Divisi Umum", "1234", mulai_pkl="bukan-tanggal")
    A.set_periode_pkl(eng_kosong, "P0", "2026-11-02", "2026-11-30")
    p = A.periode_pkl(eng_kosong)["P0"]
    assert p == (date(2026, 11, 2), date(2026, 11, 30))
    assert len(A.daftar_audit(eng_kosong, aksi="peserta.ubah_periode")) == 1
    assert A.verifikasi_audit(eng_kosong)["ok"]
    A.set_periode_pkl(eng_kosong, "P0", None, None)                      # boleh dikosongkan lagi
    assert A.periode_pkl(eng_kosong)["P0"] == (None, None)


def test_add_peserta_menyimpan_periode(eng_kosong):
    A.add_peserta(eng_kosong, "Baru", "A", "Divisi Umum", "1234", mulai_pkl=date(2026, 11, 9), selesai_pkl=date(2027, 1, 8))
    assert A.periode_pkl(eng_kosong)["Baru"] == (date(2026, 11, 9), date(2027, 1, 8))
    lengkap = A.list_peserta_lengkap(eng_kosong).set_index("nama")
    assert lengkap.loc["Baru", "mulai_pkl"] == "2026-11-09"
    assert A._fmt_tgl_pkl(lengkap.loc["Baru", "mulai_pkl"]) == "9 November 2026"
    assert A._fmt_tgl_pkl(None) == "-" and A._fmt_tgl_pkl("-") == "-"


def test_beranda_peserta_di_luar_periode_bukan_belum_absen(eng_kosong):
    hari = datetime(2026, 11, 10, 14, 0, tzinfo=A.WITA)                  # Selasa, jendela absen masuk sudah tutup
    r = A.ringkasan_hari_ini(eng_kosong, sekarang=hari)
    aktif = r["aktif"]
    assert r["belum"] == aktif
    A.set_periode_pkl(eng_kosong, "P0", date(2026, 11, 16), None)         # belum mulai
    A.set_periode_pkl(eng_kosong, "P1", None, date(2026, 11, 6))          # sudah selesai
    r = A.ringkasan_hari_ini(eng_kosong, sekarang=hari)
    assert r["belum"] == aktif - 2 and "P0" not in r["wajib"] and "P1" not in r["wajib"]


def test_migrasi_menambah_kolom_ke_database_lama(tmp_path):
    eng = create_engine(f"sqlite:///{tmp_path / 'lama.db'}")
    with eng.begin() as c:
        c.execute(text("CREATE TABLE peserta (id INTEGER PRIMARY KEY, nama VARCHAR(100) NOT NULL UNIQUE, gedung VARCHAR(1) NOT NULL, "
                       "divisi VARCHAR(100) NOT NULL, pin_hash VARCHAR(200) NOT NULL)"))
        c.execute(text("INSERT INTO peserta (nama, gedung, divisi, pin_hash) VALUES ('Lama', 'A', 'Divisi Umum', 'x')"))
    A._migrasi(eng)
    assert A.periode_pkl(eng)["Lama"] == (None, None)
    A.set_periode_pkl(eng, "Lama", date(2026, 11, 2), None)
    assert A.periode_pkl(eng)["Lama"][0] == date(2026, 11, 2)
