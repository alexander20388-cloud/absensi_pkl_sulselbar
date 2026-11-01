"""Tes logika murni proyek.py (angka dihitung manual). Jalankan: pytest -q test_proyek.py"""
import copy
from datetime import date

import pandas as pd
import pytest

import proyek as P


# ---------- kalender ----------
def test_hari_kerja_bulan_dihitung_manual():
    assert len(P.hari_kerja_bulan(2026, 11)) == 21        # 1 Nov = Minggu: 2-6, 9-13, 16-20, 23-27, 30
    assert len(P.hari_kerja_bulan(2026, 10)) == 22        # 1 Okt = Kamis
    assert P.hari_kerja_antara(date(2026, 11, 7), date(2026, 11, 8)) == []        # Sabtu-Minggu saja
    assert P.hari_kerja_antara(date(2026, 11, 9), date(2026, 11, 2)) == []        # terbalik
    assert P.bulan_sebelumnya(2026, 1) == (2025, 12) and P.bulan_sebelumnya(2026, 3, 3) == (2025, 12)


# ---------- proyeksi anggaran ----------
NOV = P.hari_kerja_bulan(2026, 11)
SEPULUH = date(2026, 11, 10)                               # Selasa; hari kerja yang sudah lewat: 2-6 dan 9 = 6 hari


def per_hari(nilai, sampai_sebelum, tambahan=None):
    d = {x: nilai for x in NOV if x < sampai_sebelum}
    d.update(tambahan or {})
    return d


def test_proyeksi_semua_hadir_hari_ini_belum_selesai():
    r = P.proyeksi_anggaran(per_hari(150_000, SEPULUH, {SEPULUH: 60_000}), NOV, SEPULUH, False, 10, 15_000)
    assert (r["hari_selesai"], r["hari_sisa"], r["hari_total"]) == (6, 15, 21)
    assert r["realisasi"] == 960_000 and r["realisasi_selesai"] == 900_000     # hari ini (sebagian) tidak ikut rata-rata
    assert r["rata_harian"] == 150_000 and r["proyeksi"] == 3_150_000 and r["proyeksi_maks"] == 3_150_000


def test_proyeksi_kehadiran_80_persen_dan_batas_atas():
    r = P.proyeksi_anggaran(per_hari(120_000, SEPULUH), NOV, SEPULUH, False, 10, 15_000)
    assert r["proyeksi"] == 720_000 + 120_000 * 15 == 2_520_000
    assert r["proyeksi_maks"] == 720_000 + 150_000 * 15 == 2_970_000           # skenario semua hadir


def test_hari_ini_dihitung_selesai_bila_sudah_lewat_jam_pulang():
    d = per_hari(150_000, SEPULUH, {SEPULUH: 150_000})
    r = P.proyeksi_anggaran(d, NOV, SEPULUH, True, 10, 15_000)
    assert (r["hari_selesai"], r["hari_sisa"]) == (7, 14) and r["proyeksi"] == 3_150_000


def test_bulan_lampau_proyeksi_sama_dengan_realisasi():
    d = {x: 100_000 for x in NOV}
    r = P.proyeksi_anggaran(d, NOV, date(2026, 12, 5), True, 10, 15_000)
    assert r["hari_sisa"] == 0 and r["proyeksi"] == r["realisasi"] == 2_100_000


def test_bulan_depan_memakai_asumsi_awal_atau_nol():
    r = P.proyeksi_anggaran({}, NOV, date(2026, 10, 30), False, 10, 15_000, rate_awal=135_000)
    assert r["hari_selesai"] == 0 and r["proyeksi"] == 135_000 * 21 and "asumsi" in r["sumber"]
    r0 = P.proyeksi_anggaran({}, NOV, date(2026, 10, 30), False, 10, 15_000)
    assert r0["proyeksi"] == 0 and r0["sumber"] == "belum ada data"


def test_proyeksi_tidak_pernah_di_bawah_yang_sudah_keluar():
    r = P.proyeksi_anggaran({SEPULUH: 500_000}, NOV, SEPULUH, False, 10, 15_000)      # tak ada hari selesai, data hari ini besar
    assert r["proyeksi"] >= r["realisasi"] == 500_000


@pytest.mark.parametrize("anggaran,proyeksi,status", [
    (3_000_000, 3_150_000, "Melebihi"), (3_000_000, 3_000_000, "Waspada"), (3_000_000, 2_700_001, "Waspada"),
    (3_000_000, 2_700_000, "Aman"), (3_000_000, 0, "Aman"), (None, 100, "Belum ada anggaran"), (0, 100, "Belum ada anggaran"),
])
def test_status_anggaran(anggaran, proyeksi, status):
    assert P.status_anggaran(anggaran, proyeksi) == status


def test_saran_anggaran_angka_manual():
    stats = {"Div A": {"n_peserta": 4, "tingkat_kehadiran": 0.95}, "Div B": {"n_peserta": 2, "tingkat_kehadiran": None},
             "Div C": {"n_peserta": 0, "tingkat_kehadiran": 0.9}}
    df = P.saran_anggaran(stats, 21, 15_000, cadangan_pct=10).set_index("divisi")
    assert list(df.index) == ["Div A", "Div B"]                                  # divisi tanpa peserta tidak dianggarkan
    a = df.loc["Div A"]
    assert a["estimasi_dasar"] == 1_197_000 and a["cadangan"] == 119_700 and a["saran"] == 1_320_000    # 1.316.700 dibulatkan ke atas
    assert not a["asumsi_kehadiran"]
    b = df.loc["Div B"]
    assert b["estimasi_dasar"] == 567_000 and b["saran"] == 630_000 and b["asumsi_kehadiran"]       # asumsi 90%


def test_saran_anggaran_membatasi_tingkat_kehadiran_tak_wajar():
    df = P.saran_anggaran({"X": {"n_peserta": 1, "tingkat_kehadiran": 0.2}, "Y": {"n_peserta": 1, "tingkat_kehadiran": 1.4}}, 20, 15_000, 0)
    d = df.set_index("divisi")
    assert d.loc["X", "tingkat_kehadiran"] == 0.5 and d.loc["Y", "tingkat_kehadiran"] == 1.0
    assert P.saran_anggaran({}, 20, 15_000).empty


# ---------- paket kerja ----------
def paket(**kw):
    dasar = dict(id=1, kode="1", nama="Analisis kebutuhan", mulai=date(2026, 11, 2), selesai=date(2026, 11, 13),
                 estimasi_jam=40.0, ditutup=0)
    dasar.update(kw)
    return pd.DataFrame([dasar])


@pytest.mark.parametrize("jam,hari_ini,ditutup,status,progres,rencana", [
    (20, date(2026, 11, 6), 0, "Sesuai jadwal", 0.5, 0.5),             # 5 dari 10 hari kerja lewat, 20 dari 40 jam
    (8, date(2026, 11, 6), 0, "Terlambat", 0.2, 0.5),                 # SPI 0,4
    (30, date(2026, 11, 6), 0, "Lebih cepat", 0.75, 0.5),             # SPI 1,5
    (40, date(2026, 11, 6), 0, "Selesai", 1.0, 0.5),
    (5, date(2026, 11, 6), 1, "Selesai", 1.0, 0.5),                   # ditutup manual
    (0, date(2026, 10, 30), 0, "Belum mulai", 0.0, 0.0),
    (10, date(2026, 11, 20), 0, "Lewat jadwal", 0.25, 1.0),
    (50, date(2026, 11, 6), 0, "Selesai", 1.0, 0.5),                  # jam melebihi estimasi dibatasi 100%
])
def test_progres_paket(jam, hari_ini, ditutup, status, progres, rencana):
    r = P.progres_paket(paket(ditutup=ditutup), {1: jam}, hari_ini).iloc[0]
    assert r["status"] == status and r["progres"] == pytest.approx(progres) and r["rencana"] == pytest.approx(rencana)
    if 0 < rencana < 1 and not ditutup and jam < 40:
        assert r["spi"] == pytest.approx(progres / rencana)


def test_paket_tanpa_jam_dan_urutan_kode_numerik():
    df = pd.concat([paket(id=1, kode="1.10"), paket(id=2, kode="1.2"), paket(id=3, kode="1"), paket(id=4, kode="2")], ignore_index=True)
    hasil = P.progres_paket(df, {}, date(2026, 11, 6))
    assert list(hasil["kode"]) == ["1", "1.2", "1.10", "2"]
    assert (hasil["jam_disetujui"] == 0).all()
    assert P.progres_paket(df.iloc[0:0], {}, date(2026, 11, 6)).empty


def test_validasi_paket():
    ok = ("1.1", "Analisis kebutuhan", date(2026, 11, 2), date(2026, 11, 13), 40)
    assert P.validasi_paket(*ok) == []
    for kode in ("", "a", "1.", "1..2", "1.2.3.4.5", "1,2"):
        assert P.validasi_paket(kode, *ok[1:]), kode
    assert P.validasi_paket("1", "ab", *ok[2:])
    assert any("sebelum" in e for e in P.validasi_paket("1", "Nama ok", date(2026, 11, 13), date(2026, 11, 2), 40))
    assert any("hari kerja" in e for e in P.validasi_paket("1", "Nama ok", date(2026, 11, 7), date(2026, 11, 8), 40))   # akhir pekan
    assert P.validasi_paket("1", "Nama ok", *ok[2:4], 0) and P.validasi_paket("1", "Nama ok", *ok[2:4], 5000)


def test_gantt_menghasilkan_spesifikasi_valid():
    hasil = P.progres_paket(pd.concat([paket(id=1, kode="1"), paket(id=2, kode="2", mulai=date(2026, 11, 16),
                                                                    selesai=date(2026, 11, 27))], ignore_index=True), {1: 20}, date(2026, 11, 6))
    spesifikasi = P.grafik_gantt(hasil, date(2026, 11, 6)).to_dict()
    assert spesifikasi and len(spesifikasi["layer"]) == 3


# ---------- validasi logbook ----------
HARI_INI = date(2026, 11, 6)              # Jumat
URAIAN = "Menyusun laporan rekap kehadiran bulanan"


def test_logbook_valid():
    assert P.validasi_logbook(HARI_INI, HARI_INI, URAIAN, 4) == []
    assert P.validasi_logbook(date(2026, 10, 23), HARI_INI, URAIAN, 0.5) == []            # tepat 14 hari ke belakang, Jumat


@pytest.mark.parametrize("tanggal,uraian,durasi,lain,cari", [
    (date(2026, 11, 7), URAIAN, 2, 0, "masa depan"),
    (date(2026, 10, 22), URAIAN, 2, 0, "14 hari"),
    (date(2026, 11, 1), URAIAN, 2, 0, "libur"),                          # Minggu
    (HARI_INI, "terlalu pendek", 2, 0, "minimal 15"),
    (HARI_INI, "satu dua", 2, 0, "minimal 15"),
    (HARI_INI, "abcdefghijklmnopqrstuvwxyz", 2, 0, "minimal 3 kata"),
    (HARI_INI, "x " * 600, 2, 0, "maksimal 1000"),
    (HARI_INI, URAIAN, 0, 0, "kelipatan 0,5"),
    (HARI_INI, URAIAN, 10.5, 0, "kelipatan 0,5"),
    (HARI_INI, URAIAN, 1.25, 0, "kelipatan 0,5"),
    (HARI_INI, URAIAN, 4, 9, "melebihi 12"),
])
def test_logbook_tidak_valid(tanggal, uraian, durasi, lain, cari):
    errs = P.validasi_logbook(tanggal, HARI_INI, uraian, durasi, lain)
    assert any(cari in e for e in errs), errs


def test_logbook_spasi_di_pinggir_tidak_dihitung():
    assert P.validasi_logbook(HARI_INI, HARI_INI, "   a b c   ", 2)                  # 5 karakter sebenarnya


# ---------- kredensial ----------
def test_pin_kuat_benar_salah_dan_garam_acak():
    h1, h2 = P.hash_pin_kuat("123456"), P.hash_pin_kuat("123456")
    assert h1 != h2 and h1.startswith("scrypt$")                                       # garam acak: hash PIN sama berbeda
    assert P.cek_pin_kuat("123456", h1) and P.cek_pin_kuat("123456", h2)
    assert not P.cek_pin_kuat("123457", h1) and not P.cek_pin_kuat("", h1)


@pytest.mark.parametrize("rusak", [None, "", "abc", "sha256$1$2", "scrypt$x$8$1$aa$bb", "scrypt$16384$8$1$zz$bb", "scrypt$16384$8$1$aa"])
def test_pin_kuat_format_rusak_ditolak_tanpa_error(rusak):
    assert P.cek_pin_kuat("123456", rusak) is False


def test_validasi_pin_pembimbing():
    assert P.validasi_pin_pembimbing("123456") == "" and P.validasi_pin_pembimbing("12345678") == ""
    for buruk in ("12345", "123456789", "12345a", "", None):
        assert P.validasi_pin_pembimbing(buruk)


# ---------- rantai audit ----------
def rantai(n=5):
    rows, prev = [], P.GENESIS
    for i in range(1, n + 1):
        r = dict(id=i, waktu=f"2026-11-0{i} 08:00:00", aktor="admin", peran="admin", aksi="peserta.tambah",
                 objek=f"peserta:P{i}", detail=P.detail_json({"i": i}), hash_prev=prev)
        r["hash"] = P.hitung_hash(prev, r["waktu"], r["aktor"], r["peran"], r["aksi"], r["objek"], r["detail"])
        rows.append(r)
        prev = r["hash"]
    return rows


def test_rantai_utuh_dan_kosong():
    rows = rantai()
    v = P.verifikasi_rantai(rows)
    assert v["ok"] and v["jumlah"] == 5 and v["hash_terakhir"] == rows[-1]["hash"]
    assert P.verifikasi_rantai([]) == {"ok": True, "jumlah": 0, "rusak_id": None, "hash_terakhir": None, "alasan": ""}


@pytest.mark.parametrize("kolom,nilai", [("detail", '{"i": 99}'), ("aktor", "penyusup"), ("aksi", "peserta.hapus"),
                                         ("objek", "peserta:X"), ("waktu", "2026-01-01 00:00:00")])
def test_ubah_isi_catatan_terdeteksi_pada_catatan_itu(kolom, nilai):
    rows = rantai()
    rows[2][kolom] = nilai
    v = P.verifikasi_rantai(rows)
    assert not v["ok"] and v["rusak_id"] == 3 and "diubah" in v["alasan"]


def test_hapus_di_tengah_terdeteksi_pada_catatan_berikutnya():
    rows = rantai()
    del rows[1]
    v = P.verifikasi_rantai(rows)
    assert not v["ok"] and v["rusak_id"] == 3 and "terputus" in v["alasan"]


def test_hapus_pertama_dan_urutan_terbalik_terdeteksi():
    rows = rantai()
    assert not P.verifikasi_rantai(rows[1:])["ok"]
    assert not P.verifikasi_rantai(list(reversed(rows)))["ok"]


def test_penyisipan_catatan_palsu_terdeteksi():
    rows = rantai()
    palsu = copy.deepcopy(rows[1])
    palsu["id"], palsu["aksi"] = 99, "peserta.hapus"
    rows.insert(2, palsu)
    assert not P.verifikasi_rantai(rows)["ok"]


def test_batas_yang_diketahui_pemotongan_ekor_tidak_terdeteksi():
    """Didokumentasikan: menghapus catatan TERAKHIR tidak memutus rantai. Mitigasi: catat hash puncak di luar sistem."""
    rows = rantai()
    asli = P.verifikasi_rantai(rows)["hash_terakhir"]
    dipotong = P.verifikasi_rantai(rows[:-1])
    assert dipotong["ok"] and dipotong["hash_terakhir"] != asli                      # hanya terlihat bila dibandingkan dengan hash puncak lama
