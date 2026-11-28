"""Tes: pilihan Cuti dihapus dari halaman absen; izin mencakup tidak datang sama sekali. Pada tanggal izin tanpa absen
masuk, Excel menulis satu baris penuh "Izin: alasan" (seperti cuti dulu), dan hari itu bukan "Tidak Hadir".
Jalankan: pytest -q test_izin_penuh.py"""
import io
from datetime import date, datetime

import openpyxl
import pytest

import app as A


@pytest.fixture(autouse=True)
def _tanpa_deteksi_wajah(monkeypatch):
    monkeypatch.setattr(A, "periksa_wajah", lambda data: (True, "ok"))
    monkeypatch.setattr(A, "simpan_foto_storage", lambda path, data: False)


def jam(tgl, h, m=0):
    return datetime(tgl.year, tgl.month, tgl.day, h, m, tzinfo=A.WITA)


@pytest.fixture
def data(eng_kosong):
    """P0: hadir Senin 9 Nov 2026, izin Selasa 10 Nov (tidak datang), izin Rabu 11 Nov tapi tetap absen masuk."""
    peserta = A.get_peserta(eng_kosong, "P0")
    lat, lon = A.GEDUNG_COORD[peserta["gedung"]]
    senin, selasa, rabu = date(2026, 11, 9), date(2026, 11, 10), date(2026, 11, 11)
    assert A.proses_absen(eng_kosong, "P0", peserta, "masuk", b"x", lat, lon, now=jam(senin, 7, 30))["level"] == "success"
    assert A.proses_tidak_hadir(eng_kosong, "P0", "izin", selasa, "sakit demam", now=jam(senin, 9))["level"] == "success"
    assert A.proses_tidak_hadir(eng_kosong, "P0", "izin", rabu, "ke kampus sebentar", now=jam(senin, 9))["level"] == "success"
    assert A.proses_absen(eng_kosong, "P0", peserta, "masuk", b"x", lat, lon, now=jam(rabu, 8, 5))["level"] == "success"
    return eng_kosong


def test_cuti_tidak_lagi_ditawarkan_di_halaman_absen():
    sumber = open(A.__file__, encoding="utf-8").read()
    assert '["Masuk", "Pulang", "Izin", "Cuti"]' not in sumber
    assert '["Masuk", "Pulang", "Izin"]' in sumber
    assert ("absen", "Absen / Izin", ":material/fingerprint:") in A.MENU_PESERTA


def test_izin_tetap_tidak_menghalangi_absen_masuk_di_tanggal_itu(data):
    sh = A.get_status_hari(data, "P0", "2026-11-11")
    assert sh["izin"] is not None and sh["masuk"] is not None            # datang walau sudah izin: tercatat hadir


def test_hari_izin_tanpa_masuk_bukan_tidak_hadir_dan_terhitung_izin(data):
    df = A.get_absensi_bulan(data, 2026, 11)
    r = A.rekap_peserta_bulan(data, df, 2026, 11, sekarang=datetime(2026, 11, 12, 18, 0, tzinfo=A.WITA))
    p0 = r[r["Nama"] == "P0"].iloc[0]
    assert p0["Izin"] == 2 and p0["Hari Hadir"] == 2
    tabel, ringkas = A.detail_absensi_peserta(df, "P0")
    assert ringkas["hadir"] == 2 and ringkas["izin"] == 1                # 11 Nov: masuk menentukan (hadir); 10 Nov: izin saja
    baris = tabel[tabel["_tgl"] == "2026-11-10"].iloc[0]
    assert baris["Status"] == "Izin" and baris["Keterangan"] == "sakit demam"


@pytest.mark.skipif(not __import__("os").path.exists(A.TEMPLATE_PATH), reason="template_absen.xlsx tidak ada")
def test_excel_izin_tanpa_masuk_jadi_baris_penuh_seperti_cuti(data):
    df = A.get_absensi_bulan(data, 2026, 11)
    isi = A.excel_template_bytes(data, df, 2026, 11, ["P0"])
    ws = openpyxl.load_workbook(io.BytesIO(isi)).worksheets[0]
    baris = {ws[f"B{r}"].value.date(): r for r in range(8, 20) if isinstance(ws[f"B{r}"].value, datetime)}
    r_izin, r_datang = baris[date(2026, 11, 10)], baris[date(2026, 11, 11)]
    merged = {str(m) for m in ws.merged_cells.ranges}
    assert ws[f"C{r_izin}"].value == "Izin: sakit demam" and f"C{r_izin}:G{r_izin}" in merged
    assert "Tidak hadir" not in str(ws[f"C{r_izin}"].value)
    assert ws[f"F{r_datang}"].value == "ke kampus sebentar" and f"C{r_datang}:G{r_datang}" not in merged   # datang + izin: baris biasa
