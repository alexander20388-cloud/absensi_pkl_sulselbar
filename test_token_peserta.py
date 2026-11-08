"""Tes token aplikasi native (murni stdlib). Jalankan: pytest -q test_token_peserta.py"""
import token_peserta as tk

RAHASIA = "rahasia-uji-yang-cukup-panjang"


def test_token_sah_bisa_dibaca_kembali():
    t = tk.buat("Alexander Prima Deus", "scrypt$abc", RAHASIA, sekarang=1000)
    d = tk.periksa(t, RAHASIA, sekarang=1001)
    assert d["n"] == "Alexander Prima Deus" and d["v"] == tk.versi_pin("scrypt$abc")


def test_token_kedaluwarsa_ditolak():
    t = tk.buat("A", "h", RAHASIA, berlaku_detik=60, sekarang=1000)
    assert tk.periksa(t, RAHASIA, sekarang=1059) is not None
    assert tk.periksa(t, RAHASIA, sekarang=1061) is None


def test_rahasia_salah_atau_diubah_ditolak():
    t = tk.buat("A", "h", RAHASIA, sekarang=1000)
    assert tk.periksa(t, "rahasia-lain-yang-panjang", sekarang=1001) is None
    isi, tanda = t.split(".")
    palsu = tk.buat("B", "h", "kunci-penyerang-panjang-sekali", sekarang=1000).split(".")[0]
    assert tk.periksa(f"{palsu}.{tanda}", RAHASIA, sekarang=1001) is None      # isi diganti, tanda tangan lama
    assert tk.periksa(isi + "x." + tanda, RAHASIA, sekarang=1001) is None


def test_sampah_tidak_melempar():
    for x in ("", "abc", "a.b", ".", "....", "e30.e30", "😀.😀"):
        assert tk.periksa(x, RAHASIA) is None


def test_versi_pin_berubah_saat_pin_berubah_dan_tidak_membocorkan_hash():
    assert tk.versi_pin("hash1") != tk.versi_pin("hash2")
    assert "hash1" not in tk.versi_pin("hash1") and len(tk.versi_pin("hash1")) == 12
