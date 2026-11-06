"""Tes: batas terlambat ML = aplikasi, libur nasional/cuti bersama, dan keterangan tulisan di dalam foto absen.
Jalankan: pytest -q test_libur_foto.py"""
import io
from datetime import date, datetime

import numpy as np
import pytest
from PIL import Image
from sqlalchemy import select

import app as A
import ml_absensi as ml
import proyek


# ------------------------------------------------------------------
# Batas terlambat: modul ML harus memakai jam yang sama dengan aplikasi
# ------------------------------------------------------------------
def test_batas_terlambat_ml_sama_dengan_aplikasi():
    assert ml.BATAS_MENIT == A.BATAS_TEPAT_WAKTU.hour * 60 + A.BATAS_TEPAT_WAKTU.minute == 8 * 60


@pytest.mark.parametrize("jam, telat", [("07:59:00", False), ("08:00:00", False), ("08:00:59", False), ("08:01:00", True), ("11:30:00", True)])
def test_label_terlambat_ml_dan_aplikasi_selalu_sepakat(jam, telat):
    assert A.terlambat(jam) is telat
    assert (ml.menit_dari_jam(jam) > ml.BATAS_MENIT) is telat      # ML membaca menit, detik diabaikan seperti aplikasi
    # (08:00:59 -> aplikasi menghitung per menit sehingga masih tepat waktu; menit_dari_jam juga membuang detik)


# ------------------------------------------------------------------
# Libur nasional dan cuti bersama
# ------------------------------------------------------------------
def test_jumlah_libur_sesuai_skb():
    def n(daftar, tahun):
        return sum(1 for k in daftar if k.startswith(str(tahun)))
    assert (n(A.LIBUR_NASIONAL, 2026), n(A.CUTI_BERSAMA, 2026)) == (17, 8)      # SKB 1497/2025, 2/2025, 5/2025
    assert (n(A.LIBUR_NASIONAL, 2027), n(A.CUTI_BERSAMA, 2027)) == (18, 8)      # SKB 1205/2026, 3/2026, 2/2026
    assert not set(A.LIBUR_NASIONAL) & set(A.CUTI_BERSAMA)


@pytest.mark.parametrize("tgl, hari", [      # hari dalam seminggu seperti tertulis di SKB (0=Senin): menangkap salah ketik tanggal
    ("2026-01-01", 3), ("2026-03-20", 4), ("2026-05-27", 2), ("2026-08-25", 1), ("2026-12-24", 3), ("2026-12-25", 4),
    ("2027-01-01", 4), ("2027-01-05", 1), ("2027-03-10", 2), ("2027-03-26", 4), ("2027-05-17", 0), ("2027-08-17", 1),
    ("2027-12-24", 4),
])
def test_tanggal_libur_jatuh_di_hari_yang_benar(tgl, hari):
    assert tgl in A.LIBUR_NASIONAL or tgl in A.CUTI_BERSAMA
    assert date.fromisoformat(tgl).weekday() == hari


def test_hari_libur_mengenali_tanggal_merah_dan_akhir_pekan():
    assert A.hari_libur(date(2026, 12, 25))                                   # Natal (Jumat)
    assert A.hari_libur(datetime(2026, 12, 24, 7, 30, tzinfo=A.WITA))         # cuti bersama, objek datetime juga bisa
    assert A.hari_libur(date(2026, 11, 7))                                    # Sabtu biasa
    assert not A.hari_libur(date(2026, 11, 6))                                # Jumat biasa
    assert A.nama_libur(date(2027, 3, 10)) == "Idulfitri 1448 Hijriah"
    assert A.nama_libur(date(2026, 11, 6)) is None


def test_pesan_libur_menyebut_nama_libur_dan_akhir_pekan_tetap_seperti_semula():
    assert "Natal" in A.pesan_libur(date(2026, 12, 25))
    assert "Sabtu dan Minggu" in A.pesan_libur(date(2026, 11, 7))


def test_cuti_bersama_bisa_dimatikan(monkeypatch):
    hanya_nasional = dict(A.LIBUR_NASIONAL)
    monkeypatch.setattr(A, "HARI_LIBUR_KHUSUS", hanya_nasional)
    assert A.hari_libur(date(2026, 12, 25)) and not A.hari_libur(date(2026, 12, 24))


def test_absen_masuk_dan_pulang_ditolak_pada_libur_nasional(eng_kosong):
    peserta = A.get_peserta(eng_kosong, "P0")
    lat, lon = A.GEDUNG_COORD["A"]
    for jenis, jam in (("masuk", 7), ("pulang", 17)):
        now = datetime(2026, 12, 25, jam, 30 if jenis == "masuk" else 5, tzinfo=A.WITA)
        h = A.proses_absen(eng_kosong, "P0", peserta, jenis, b"bukan-foto", lat, lon, now=now)
        assert h["level"] == "warning" and h["alasan"] == "libur" and "Natal" in h["pesan"]


def test_izin_dan_cuti_pada_tanggal_libur_ditolak_dengan_nama_libur(eng_kosong):
    now = datetime(2026, 12, 21, 9, 0, tzinfo=A.WITA)
    h = A.proses_tidak_hadir(eng_kosong, "P0", "izin", date(2026, 12, 24), "urusan keluarga", now=now)
    assert h["level"] == "error" and "Cuti Bersama Natal" in h["pesan"]
    assert A.proses_tidak_hadir(eng_kosong, "P0", "izin", date(2026, 12, 23), "urusan keluarga", now=now)["level"] == "success"


def test_status_beranda_peserta_menyebut_libur():
    now = datetime(2026, 12, 25, 9, 0, tzinfo=A.WITA)
    judul, ket = A.teks_status_hari(now, {"masuk": None, "pulang": None, "cuti": None, "izin": None})
    assert judul == "Hari libur" and "Natal" in ket


def test_libur_nasional_tidak_dihitung_hari_kerja():
    desember = proyek.hari_kerja_bulan(2026, 12, A.HARI_LIBUR)
    assert len(desember) == 23
    tanpa = proyek.hari_kerja_bulan(2026, 12, A.HARI_LIBUR, A.HARI_LIBUR_KHUSUS)
    assert len(tanpa) == 21 and date(2026, 12, 24) not in tanpa and date(2026, 12, 25) not in tanpa
    akhir = datetime(2027, 1, 5, 18, 0, tzinfo=A.WITA)
    assert len(A.hari_kerja_periode(2026, 12, akhir)) == 21               # rekap: libur tidak jadi "Tidak Hadir"


def test_logbook_pada_libur_nasional_ditolak():
    errs = proyek.validasi_logbook(date(2026, 12, 25), date(2026, 12, 28), "Mengerjakan laporan bulanan divisi",
                                   2.0, 0.0, A.HARI_LIBUR, libur_khusus=A.HARI_LIBUR_KHUSUS)
    assert any("hari libur" in e for e in errs)
    ok = proyek.validasi_logbook(date(2026, 12, 23), date(2026, 12, 28), "Mengerjakan laporan bulanan divisi",
                                 2.0, 0.0, A.HARI_LIBUR, libur_khusus=A.HARI_LIBUR_KHUSUS)
    assert ok == []


def test_proyek_tanpa_libur_khusus_tetap_berperilaku_lama():
    assert len(proyek.hari_kerja_antara(date(2026, 12, 21), date(2026, 12, 25))) == 5


# ------------------------------------------------------------------
# Keterangan tulisan di dalam foto
# ------------------------------------------------------------------
def _jpeg(warna=(30, 30, 30), ukuran=(720, 960)):
    buf = io.BytesIO()
    Image.new("RGB", ukuran, warna).save(buf, format="JPEG", quality=95)
    return buf.getvalue()


def _arr(data):
    return np.asarray(Image.open(io.BytesIO(data)).convert("RGB")).astype(int)


def test_format_teks_foto_seperti_contoh():
    t = A.teks_foto(datetime(2026, 11, 6, 10, 33, 54), "A")
    assert t == ["6 Nov 2026 at 10.33.54", "Jalan Cendrawasih No. 18", "Makassar", "South Sulawesi 90125", "Indonesia"]
    assert A.teks_foto(datetime(2026, 12, 1, 7, 5, 9), "B")[0] == "1 Dec 2026 at 07.05.09"     # hari tanpa nol, jam dua digit
    assert A.teks_foto(datetime(2026, 5, 31, 17, 0, 0), "?")[1:] == list(A.ALAMAT_FOTO["A"])     # gedung tak dikenal -> alamat A


def test_foto_dengan_keterangan_tetap_jpeg_kecil_dan_tulisan_ada_di_pojok_kanan_atas():
    data = _jpeg()
    polos = _arr(A.kompres_foto(data))
    cap = A.kompres_foto(data, A.teks_foto(datetime(2026, 11, 6, 10, 33, 54), "A"))
    assert Image.open(io.BytesIO(cap)).format == "JPEG" and max(Image.open(io.BytesIO(cap)).size) <= 640
    a = _arr(cap)
    h, w, _ = a.shape
    assert polos.max() < 80                                              # foto polos gelap: tidak ada piksel terang
    kanan_atas = a[: h // 3, w // 4:]
    assert (kanan_atas.min(axis=2) > 230).sum() > 500                     # ada banyak piksel putih (huruf)
    assert (a[h // 2:, :] < 80).all()                                     # separuh bawah tidak tersentuh
    assert (a[:, : int(w * 0.03)] < 80).all()                             # tidak ada yang menyentuh tepi kiri (muat dalam foto)
    assert (a[:, w - int(w * 0.02):] < 80).all()                          # dan tepi kanan


def test_tulisan_tetap_terbaca_di_foto_terang():
    a = _arr(A.kompres_foto(_jpeg((235, 230, 215)), A.teks_foto(datetime(2026, 11, 6, 10, 33, 54), "A")))
    h, w, _ = a.shape
    assert (a[: h // 3, w // 4:].min(axis=2) < 150).sum() > 500           # bayangan membuat tepi huruf gelap


def test_tanpa_keterangan_hasil_sama_seperti_sebelumnya():
    data = _jpeg((90, 120, 150), (300, 200))
    assert A.kompres_foto(data) == A.kompres_foto(data, None) == A.kompres_foto(data, [])


def test_ukuran_huruf_menyesuaikan_foto_kecil_dan_besar():
    for ukuran in ((200, 260), (1600, 1200)):
        a = _arr(A.kompres_foto(_jpeg(ukuran=ukuran), A.teks_foto(datetime(2026, 11, 6, 10, 33, 54), "A")))
        h, w, _ = a.shape
        assert (a[: h // 2].min(axis=2) > 230).sum() > 50                 # tulisan ada
        assert (a[:, : max(1, int(w * 0.02))] < 80).all()                 # dan tidak terpotong di kiri


def test_absen_berhasil_menyimpan_foto_dengan_keterangan(eng_kosong, monkeypatch):
    monkeypatch.setattr(A, "periksa_wajah", lambda data: (True, "ok"))
    peserta = A.get_peserta(eng_kosong, "P0")
    lat, lon = A.GEDUNG_COORD[peserta["gedung"]]
    now = datetime(2026, 11, 6, 10, 33, 54, tzinfo=A.WITA)
    h = A.proses_absen(eng_kosong, "P0", peserta, "masuk", _jpeg(), lat, lon, now=now)
    assert h["level"] == "success"
    with eng_kosong.connect() as c:
        foto = c.execute(select(A.absensi_t.c.foto).where(A.absensi_t.c.nama == "P0")).scalar_one()
    a = _arr(bytes(foto))
    assert (a[: a.shape[0] // 3, a.shape[1] // 4:].min(axis=2) > 230).sum() > 500


def test_absen_pulang_juga_dicap(eng_kosong, monkeypatch):
    monkeypatch.setattr(A, "periksa_wajah", lambda data: (True, "ok"))
    peserta = A.get_peserta(eng_kosong, "P0")
    lat, lon = A.GEDUNG_COORD[peserta["gedung"]]
    A.proses_absen(eng_kosong, "P0", peserta, "masuk", _jpeg(), lat, lon, now=datetime(2026, 11, 6, 7, 30, tzinfo=A.WITA))
    h = A.proses_absen(eng_kosong, "P0", peserta, "pulang", _jpeg(), lat, lon, now=datetime(2026, 11, 6, 17, 5, 3, tzinfo=A.WITA))
    assert h["level"] == "success" and h["uang"] == A.UANG_MAKAN
    with eng_kosong.connect() as c:
        foto = c.execute(select(A.absensi_t.c.foto).where(A.absensi_t.c.jenis == "pulang")).scalar_one()
    a = _arr(bytes(foto))
    assert (a[: a.shape[0] // 3, a.shape[1] // 4:].min(axis=2) > 230).sum() > 500
