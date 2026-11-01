"""Tes analitik EDA. Jalankan: pytest -q test_analitik.py"""
import pandas as pd
import pytest

import app as A

PESERTA = pd.DataFrame({
    "nama": ["Budi", "Sari"], "gedung": ["A", "B"], "divisi": ["Divisi Umum", "Divisi Syariah"],
    "no_rekening": [None, None],
})


def baris(nama, tanggal, jam, jenis="masuk", status="berhasil", jarak=50.0):
    return dict(nama=nama, tanggal=tanggal, jam=jam, jenis=jenis, status=status, jarak=jarak, akurasi=10.0, uang_makan=0)


def data_contoh():
    # 2026-10-26 Senin, 27 Selasa, 28 Rabu
    return pd.DataFrame([
        baris("Budi", "2026-10-26", "07:30:00"),                                   # tepat waktu
        baris("Sari", "2026-10-26", "08:00:00"),                                   # tepat waktu (batas masih tepat)
        baris("Budi", "2026-10-27", "08:20:00"),                                   # terlambat
        baris("Sari", "2026-10-27", "08:45:00"),                                   # terlambat
        baris("Budi", "2026-10-28", "07:50:00"),                                   # tepat waktu
        baris("Budi", "2026-10-28", "17:00:00", jenis="pulang", jarak=80.0),
        baris("Sari", "2026-10-28", "07:55:00", status="gagal_lokasi", jarak=450.0),  # percobaan gagal
        baris("Sari", "2026-10-28", "09:00:00", jenis="izin", status="izin", jarak=None),
    ])


def test_menit_dan_slot():
    assert A.menit_hari("07:30:00") == 450
    assert A.label_menit(450) == "07:30"
    assert A.label_slot(467, 15) == "07:45"
    assert A.label_slot(467, 30) == "07:30"
    s = A.daftar_slot(30)
    assert s[0] == "06:00" and s[-1] == "12:30" and len(s) == 14


def test_batas_iqr():
    assert A.batas_iqr([1, 2, 3]) is None                    # < 4 data
    q1, q3, lo, hi = A.batas_iqr([10, 11, 12, 13, 14, 100])
    assert 100 > hi and lo < 10


def test_siapkan_eda_angka_dasar():
    eda = A.siapkan_eda(data_contoh(), PESERTA)
    assert eda["n_masuk"] == 5                                # gagal_lokasi, pulang, izin tidak dihitung
    assert eda["n_percobaan"] == 7 and eda["n_gagal"] == 1    # 5 masuk berhasil + 1 pulang + 1 masuk gagal (izin tanpa jarak tidak dihitung)
    assert round(eda["pct_telat"]) == 40                      # 2 dari 5 terlambat; 08:00 masih tepat waktu
    assert eda["median_menit"] == 480.0                       # 450, 470, 480, 500, 525


def test_harian_dan_per_hari():
    eda = A.siapkan_eda(data_contoh(), PESERTA)
    h = eda["harian"].set_index("tanggal")
    assert h.loc["2026-10-26", "Hadir"] == 2 and h.loc["2026-10-26", "Terlambat"] == 0
    assert h.loc["2026-10-27", "Hadir"] == 2 and h.loc["2026-10-27", "Terlambat"] == 2
    assert h.loc["2026-10-28", "Hadir"] == 1 and h.loc["2026-10-28", "Izin"] == 1
    ph = eda["per_hari"].set_index("hari")
    assert ph.loc["Selasa", "persen"] == 100.0 and ph.loc["Senin", "persen"] == 0.0


def test_histogram_grid_lengkap_dan_jumlah_cocok():
    eda = A.siapkan_eda(data_contoh(), PESERTA)
    hist = eda["hist"]
    assert len(hist) == len(A.daftar_slot(15)) * 2            # blok kosong tetap ada sebagai 0
    assert hist["jumlah"].sum() == eda["n_masuk"]
    def n(slot, ket):
        return int(hist[(hist["slot"] == slot) & (hist["ketepatan"] == ket)]["jumlah"].iloc[0])
    assert n("08:00", A.KET_TEPAT_WAKTU) == 1                 # 08:00 pas = masih tepat waktu
    assert n("08:00", A.KET_TERLAMBAT) == 0
    assert n("08:15", A.KET_TERLAMBAT) == 1                   # 08:20
    assert n("08:45", A.KET_TERLAMBAT) == 1                   # 08:45
    assert n("07:45", A.KET_TEPAT_WAKTU) == 1                 # 07:50


def test_slot_di_luar_jam_buka_tidak_hilang():
    df = pd.DataFrame([baris("Budi", "2026-10-26", "05:30:00"), baris("Budi", "2026-10-27", "07:30:00")])
    eda = A.siapkan_eda(df, PESERTA)
    assert eda["hist"]["jumlah"].sum() == 2


def test_peserta_tidak_terdaftar_dan_data_kosong():
    df = pd.DataFrame([baris("Hantu", "2026-10-26", "07:30:00")])
    eda = A.siapkan_eda(df, PESERTA)
    assert (eda["jarak"]["gedung"] == "?").all()
    eda2 = A.siapkan_eda(df, PESERTA.iloc[0:0])               # tidak ada peserta terdaftar sama sekali
    assert eda2["n_masuk"] == 1


def test_insight_dan_grafik_tidak_error():
    eda = A.siapkan_eda(data_contoh(), PESERTA)
    teks = " ".join(A.buat_insight(eda))
    assert "peserta hadir per hari" in teks and "gagal lokasi" in teks
    assert "Jarak absen berhasil" in teks                       # outlier hanya dari absen berhasil, bukan percobaan gagal
    grafik = A.grafik_eda(eda)
    assert {"tren", "hist", "heat", "per_hari", "jarak"} <= set(grafik)
    for g in grafik.values():
        assert g.to_dict()                                     # spesifikasi Vega-Lite valid


def test_insight_tidak_menyimpulkan_hari_terburuk_dari_data_sedikit():
    teks = " ".join(A.buat_insight(A.siapkan_eda(data_contoh(), PESERTA)))
    assert "keterlambatan tertinggi" not in teks               # tiap hari < 5 absen
