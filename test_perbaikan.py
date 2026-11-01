"""Tes untuk perbaikan keamanan: password admin bawaan, pembatasan login, peningkatan hash PIN, batas akurasi GPS.
Jalankan: pytest -q test_perbaikan.py"""
from datetime import datetime, timedelta

from sqlalchemy import insert, select

import app as A
import test_ml as T


def _secret(nilai):
    return lambda kunci, default=None: nilai.get(kunci, default)


def test_password_bawaan_hanya_untuk_mode_lokal(monkeypatch):
    monkeypatch.setattr(A, "secret", _secret({}))
    assert A.password_admin() == "admin123"                          # lokal tanpa database cloud
    monkeypatch.setattr(A, "secret", _secret({"db_host": "db.contoh.co", "db_password": "x"}))
    assert A.password_admin() is None                                # database cloud lewat db_host: bawaan DIMATIKAN
    monkeypatch.setattr(A, "secret", _secret({"database_url": "postgresql://u:p@h/db"}))
    assert A.password_admin() is None
    monkeypatch.setattr(A, "secret", _secret({"db_host": "h", "admin_password": "rahasia-kuat"}))
    assert A.password_admin() == "rahasia-kuat"


def test_banding_teks_tidak_error_untuk_huruf_non_ascii():
    assert A._sama("sandi-é", "sandi-é") is True
    assert A._sama("sandi-é", "sandi-e") is False


def test_login_dikunci_setelah_lima_gagal_lalu_dibuka(eng_kosong):
    t0 = datetime(2026, 11, 2, 8, 0, tzinfo=A.WITA)
    for i in range(A.MAKS_GAGAL_LOGIN - 1):
        assert A.proses_login(eng_kosong, "peserta:X", lambda: False, now=t0 + timedelta(seconds=i)) == ("salah", 0)
    status, menit = A.proses_login(eng_kosong, "peserta:X", lambda: False, now=t0 + timedelta(seconds=10))
    assert status == "terkunci" and menit == A.KUNCI_LOGIN_MENIT

    dipanggil = []
    sama_dgn_benar = A.proses_login(eng_kosong, "peserta:X", lambda: dipanggil.append(1) or True, now=t0 + timedelta(minutes=5))
    assert sama_dgn_benar[0] == "terkunci" and not dipanggil          # PIN benar pun ditolak, dan tidak diperiksa sama sekali

    assert A.proses_login(eng_kosong, "peserta:Y", lambda: True, now=t0) == ("ok", 0)   # akun lain tidak terpengaruh
    assert A.proses_login(eng_kosong, "peserta:X", lambda: True, now=t0 + timedelta(minutes=16)) == ("ok", 0)
    assert A.sisa_kunci_login(eng_kosong, "peserta:X", t0 + timedelta(minutes=16)) == 0
    assert any(a == "login.terkunci" for a in A.daftar_audit(eng_kosong)["aksi"])
    assert A.verifikasi_audit(eng_kosong)["ok"]


def test_gagal_lama_di_luar_jendela_tidak_menumpuk(eng_kosong):
    t0 = datetime(2026, 11, 2, 8, 0, tzinfo=A.WITA)
    for i in range(A.MAKS_GAGAL_LOGIN - 1):
        A.proses_login(eng_kosong, "admin", lambda: False, now=t0 + timedelta(seconds=i))
    lanjut = t0 + timedelta(minutes=A.KUNCI_LOGIN_MENIT + 1)
    assert A.proses_login(eng_kosong, "admin", lambda: False, now=lanjut) == ("salah", 0)   # hitungan mulai dari 1 lagi


def test_pin_baru_pakai_scrypt_dan_pin_lama_ditingkatkan_saat_login(eng_kosong):
    nama = T.NAMA[0]
    assert A.get_peserta(eng_kosong, nama)["pin_hash"].startswith("scrypt$")      # peserta baru langsung scrypt
    assert A.cek_pin(eng_kosong, nama, "1234") and not A.cek_pin(eng_kosong, nama, "9999")

    with eng_kosong.begin() as c:                                               # peserta lama: hash SHA-256
        c.execute(insert(A.peserta_t).values(nama="Lama", gedung="A", divisi="Divisi Umum", pin_hash=A.hash_pin("Lama", "4321")))
    assert not A.cek_pin(eng_kosong, "Lama", "0000")
    assert not A.get_peserta(eng_kosong, "Lama")["pin_hash"].startswith("scrypt$")  # gagal login tidak mengubah apa pun
    assert A.cek_pin(eng_kosong, "Lama", "4321")
    assert A.get_peserta(eng_kosong, "Lama")["pin_hash"].startswith("scrypt$")      # ditingkatkan
    assert A.cek_pin(eng_kosong, "Lama", "4321") and not A.cek_pin(eng_kosong, "Lama", "0000")
    assert not A.cek_pin(eng_kosong, "Tidak Ada", "1234")


def test_akurasi_gps_buruk_ditolak_tanpa_catatan_dan_tidak_mengubah_radius(eng_kosong):
    nama = T.NAMA[0]
    peserta = A.get_peserta(eng_kosong, nama)
    senin = datetime(2026, 11, 2, 7, 0, tzinfo=A.WITA)
    lat, lon = A.GEDUNG_COORD[peserta["gedung"]]
    h = A.proses_absen(eng_kosong, nama, peserta, "masuk", b"bukan-foto", lat, lon, akurasi=150.0, now=senin)
    assert not h["ok"] and h["alasan"] == "akurasi"
    assert A.get_riwayat_peserta(eng_kosong, nama) == []                       # tidak ada catatan gagal_lokasi
    h = A.proses_absen(eng_kosong, nama, peserta, "masuk", b"bukan-foto", lat, lon, akurasi=100.0, now=senin)
    assert h.get("alasan") == "wajah"                                          # tepat 100 m lolos pemeriksaan akurasi, berhenti di foto
    assert A.RADIUS_METER == 200 and A.MAKS_AKURASI_METER == 100
