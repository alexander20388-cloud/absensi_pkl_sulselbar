"""Tes: absen pulang dibuka 12:30-18:00. Pulang sebelum 16:50 = "pulang awal": wajib alasan, uang makan menunggu
persetujuan admin (disetujui -> Rp15.000, ditolak -> Rp0). Mulai 16:50 pulang seperti biasa.
Jalankan: pytest -q test_pulang_awal.py"""
from datetime import datetime, time

import pytest
from sqlalchemy import create_engine, inspect, select, text

import app as A

HARI = (2026, 11, 6)         # Jumat biasa (bukan libur)
FOTO = b"bukan-foto"         # deteksi wajah dimatikan oleh fixture di bawah


def jam(h, m=0, s=0):
    return datetime(*HARI, h, m, s, tzinfo=A.WITA)


@pytest.fixture(autouse=True)
def _tanpa_deteksi_wajah(monkeypatch):
    monkeypatch.setattr(A, "periksa_wajah", lambda data: (True, "ok"))
    monkeypatch.setattr(A, "simpan_foto_storage", lambda path, data: False)
    monkeypatch.setattr(A, "kompres_foto", lambda data, keterangan=None: data)   # foto tes bukan gambar sungguhan; cap foto dites di test_libur_foto


@pytest.fixture
def siap(eng_kosong):
    """(engine, peserta P0, lat, lon di gedungnya) dengan absen masuk P0 pukul 07:30 sudah tercatat."""
    peserta = A.get_peserta(eng_kosong, "P0")
    lat, lon = A.GEDUNG_COORD[peserta["gedung"]]
    assert A.proses_absen(eng_kosong, "P0", peserta, "masuk", FOTO, lat, lon, now=jam(7, 30))["level"] == "success"
    return eng_kosong, peserta, lat, lon


def baris_pulang(eng):
    with eng.connect() as c:
        return c.execute(select(A.absensi_t).where(A.absensi_t.c.jenis == "pulang")).mappings().all()


# ------------------------------------------------------------------
# Jam
# ------------------------------------------------------------------
def test_jendela_pulang_baru():
    assert (A.JAM_PULANG_BUKA, A.JAM_PULANG_NORMAL, A.JAM_PULANG_TUTUP) == (time(12, 30), time(16, 50), time(18, 0))


@pytest.mark.parametrize("h, m, s, awal", [(12, 30, 0, True), (16, 49, 59, True), (16, 50, 0, False), (17, 5, 0, False), (18, 0, 0, False)])
def test_batas_pulang_awal_per_menit(h, m, s, awal):
    assert A.pulang_awal(jam(h, m, s)) is awal


def test_pulang_sebelum_buka_dan_setelah_tutup_ditolak(siap):
    eng, peserta, lat, lon = siap
    for h, m in ((12, 29), (18, 1)):
        hasil = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(h, m), alasan="ada keperluan keluarga")
        assert hasil["level"] == "error" and "hanya dibuka" in hasil["pesan"]
    assert baris_pulang(eng) == []


# ------------------------------------------------------------------
# Absen pulang awal
# ------------------------------------------------------------------
@pytest.mark.parametrize("alasan", [None, "", "   ", "abc", "  a b  "])
def test_pulang_awal_tanpa_alasan_ditolak_dan_tidak_tersimpan(siap, alasan):
    eng, peserta, lat, lon = siap
    hasil = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(13, 0), alasan=alasan)
    assert hasil["level"] == "warning" and hasil["alasan"] == "alasan_pulang" and not hasil["ok"]
    assert baris_pulang(eng) == []                       # tidak ada catatan, peserta boleh mencoba lagi


def test_pulang_awal_dengan_alasan_berhasil_uang_makan_menunggu(siap):
    eng, peserta, lat, lon = siap
    hasil = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(13, 5), alasan="  Bimbingan   skripsi di kampus ")
    assert hasil["level"] == "success" and hasil["pulang_awal"] is True and hasil["uang"] == 0
    assert hasil["tinjau_status"] == "menunggu" and "menunggu persetujuan" in hasil["info"]
    (r,) = baris_pulang(eng)
    assert r["status"] == "berhasil" and r["uang_makan"] == 0
    assert r["tinjau_status"] == "menunggu" and r["alasan"] == "Bimbingan skripsi di kampus"      # spasi dirapikan
    assert r["tinjau_oleh"] is None and r["tinjau_pada"] is None


def test_alasan_dipotong_300_karakter(siap):
    eng, peserta, lat, lon = siap
    A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(14, 0), alasan="x" * 500)
    assert len(baris_pulang(eng)[0]["alasan"]) == 300


def test_pulang_awal_tidak_bisa_diulang_di_hari_yang_sama(siap):
    eng, peserta, lat, lon = siap
    A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(13, 0), alasan="urusan keluarga")
    lagi = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(17, 0))
    assert lagi["level"] == "warning" and "sudah absen pulang" in lagi["pesan"]
    assert len(baris_pulang(eng)) == 1


def test_pulang_awal_di_luar_area_gagal_tanpa_alasan_tersimpan(siap):
    eng, peserta, lat, lon = siap
    hasil = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat + 0.02, lon + 0.02, now=jam(13, 0), alasan="urusan keluarga")
    assert hasil["level"] == "error" and hasil["status"] == "gagal_lokasi"
    (r,) = baris_pulang(eng)
    assert r["status"] == "gagal_lokasi" and r["uang_makan"] == 0 and r["tinjau_status"] is None and r["alasan"] is None
    assert A.daftar_pulang_awal(eng, None) == [] and A.jumlah_pulang_awal_menunggu(eng) == 0


# ------------------------------------------------------------------
# Pulang normal (mulai 16:50) tidak berubah
# ------------------------------------------------------------------
@pytest.mark.parametrize("alasan", [None, "alasan yang tidak diperlukan"])
def test_pulang_mulai_1650_tanpa_alasan_uang_makan_langsung(siap, alasan):
    eng, peserta, lat, lon = siap
    hasil = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(16, 50), alasan=alasan)
    assert hasil["level"] == "success" and hasil["pulang_awal"] is False and hasil["uang"] == A.UANG_MAKAN
    (r,) = baris_pulang(eng)
    assert r["uang_makan"] == A.UANG_MAKAN and r["tinjau_status"] is None and r["alasan"] is None   # alasan diabaikan


def test_form_dibuka_1649_dikirim_1651_tidak_butuh_alasan(siap):
    """Yang dipakai jam saat dikirim, bukan saat form dibuka."""
    eng, peserta, lat, lon = siap
    hasil = A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(16, 51), alasan=None)
    assert hasil["level"] == "success" and hasil["uang"] == A.UANG_MAKAN


# ------------------------------------------------------------------
# Peninjauan admin
# ------------------------------------------------------------------
def _pulang_awal(siap, h=13, m=0):
    eng, peserta, lat, lon = siap
    A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(h, m), alasan="ada urusan keluarga")
    return eng, baris_pulang(eng)[0]["id"]


def test_daftar_dan_jumlah_menunggu(siap):
    eng, aid = _pulang_awal(siap)
    assert A.jumlah_pulang_awal_menunggu(eng) == 1 and A.jumlah_pulang_awal_menunggu(eng, "P0") == 1
    assert A.jumlah_pulang_awal_menunggu(eng, "P1") == 0
    (r,) = A.daftar_pulang_awal(eng, "menunggu")
    assert (r["id"], r["nama"], r["alasan"]) == (aid, "P0", "ada urusan keluarga")
    assert A.daftar_pulang_awal(eng, "disetujui") == []


def test_setujui_mengisi_uang_makan_dan_tercatat_di_audit(siap):
    eng, aid = _pulang_awal(siap)
    assert A.tinjau_pulang_awal(eng, aid, "disetujui", aktor="admin") is True
    (r,) = baris_pulang(eng)
    assert r["tinjau_status"] == "disetujui" and r["uang_makan"] == A.UANG_MAKAN
    assert r["tinjau_oleh"] == "admin" and r["tinjau_pada"]
    assert A.jumlah_pulang_awal_menunggu(eng) == 0
    audit = A.daftar_audit(eng, aksi="absensi.tinjau_pulang_awal")
    assert len(audit) == 1
    a = audit.iloc[0]
    assert a["aktor"] == "admin" and a["objek"] == "absensi:P0:2026-11-06"
    assert A.verifikasi_audit(eng)["ok"]                                                  # rantai audit tetap utuh


def test_tolak_uang_makan_tetap_nol_dan_absen_tetap_sah(siap):
    eng, aid = _pulang_awal(siap)
    assert A.tinjau_pulang_awal(eng, aid, "ditolak") is True
    (r,) = baris_pulang(eng)
    assert r["tinjau_status"] == "ditolak" and r["uang_makan"] == 0 and r["status"] == "berhasil"
    assert A.get_status_hari(eng, "P0", "2026-11-06")["pulang"] is not None             # tetap tercatat pulang


def test_keputusan_boleh_diubah_uang_makan_mengikuti(siap):
    eng, aid = _pulang_awal(siap)
    A.tinjau_pulang_awal(eng, aid, "ditolak")
    A.tinjau_pulang_awal(eng, aid, "disetujui")
    assert baris_pulang(eng)[0]["uang_makan"] == A.UANG_MAKAN
    A.tinjau_pulang_awal(eng, aid, "ditolak")
    assert baris_pulang(eng)[0]["uang_makan"] == 0
    assert len(A.daftar_audit(eng, aksi="absensi.tinjau_pulang_awal")) == 3


def test_tinjau_menolak_data_yang_bukan_pulang_awal_dan_keputusan_aneh(siap):
    eng, peserta, lat, lon = siap
    A.proses_absen(eng, "P0", peserta, "pulang", FOTO, lat, lon, now=jam(17, 0))        # pulang normal
    normal = baris_pulang(eng)[0]["id"]
    with eng.connect() as c:
        masuk = c.execute(select(A.absensi_t.c.id).where(A.absensi_t.c.jenis == "masuk")).scalar_one()
    assert A.tinjau_pulang_awal(eng, normal, "ditolak") is False                         # tidak bisa memotong uang makan pulang normal
    assert A.tinjau_pulang_awal(eng, masuk, "disetujui") is False
    assert A.tinjau_pulang_awal(eng, 999999, "disetujui") is False
    assert baris_pulang(eng)[0]["uang_makan"] == A.UANG_MAKAN
    with pytest.raises(ValueError):
        A.tinjau_pulang_awal(eng, normal, "mungkin")


# ------------------------------------------------------------------
# Rekap, anggaran, tampilan
# ------------------------------------------------------------------
def test_total_uang_makan_menunggu_tidak_dihitung_lalu_masuk_setelah_disetujui(siap):
    eng, aid = _pulang_awal(siap)
    rekap, detail = A._susun_rekap_detail(A.get_absensi_bulan(eng, 2026, 11))
    assert int(rekap.loc[rekap["Nama"] == "P0", "Total Uang Makan"].iloc[0]) == 0
    assert "Pulang awal (menunggu)" in set(detail["Keterangan"]) and "tinjau_status" not in detail.columns
    A.tinjau_pulang_awal(eng, aid, "disetujui")
    rekap, detail = A._susun_rekap_detail(A.get_absensi_bulan(eng, 2026, 11))
    assert int(rekap.loc[rekap["Nama"] == "P0", "Total Uang Makan"].iloc[0]) == A.UANG_MAKAN
    assert "Pulang awal (disetujui)" in set(detail["Keterangan"])


def test_detail_per_peserta_menandai_pulang_awal_beserta_alasan(siap):
    eng, aid = _pulang_awal(siap)
    tabel, ringkas = A.detail_absensi_peserta(A.get_absensi_bulan(eng, 2026, 11), "P0")
    ket = tabel["Keterangan"].iloc[0]
    assert "Pulang awal (menunggu)" in ket and "ada urusan keluarga" in ket
    assert ringkas["uang"] == 0 and tabel["Uang makan"].iloc[0] == "Rp0"


def test_status_hari_membawa_tinjau_status(siap):
    eng, aid = _pulang_awal(siap)
    sh = A.get_status_hari(eng, "P0", "2026-11-06")
    assert sh["pulang"]["tinjau_status"] == "menunggu"
    judul, ket = A.teks_status_hari(jam(15, 0), sh)
    assert "pulang awal" in ket and "menunggu persetujuan admin" in ket
    assert A.teks_uang_pulang(sh["pulang"]) == "Menunggu"


def test_teks_uang_pulang():
    assert A.teks_uang_pulang(None) == "-"
    assert A.teks_uang_pulang({"uang_makan": 15000, "tinjau_status": None}) == A.rupiah(15000)
    assert A.teks_uang_pulang({"uang_makan": 15000, "tinjau_status": "disetujui"}) == A.rupiah(15000)
    assert A.teks_uang_pulang({"uang_makan": 0, "tinjau_status": "menunggu"}) == "Menunggu"
    assert A.teks_uang_pulang({"uang_makan": 0, "tinjau_status": "ditolak"}) == "Ditolak"


def _sh_sudah_masuk():
    return {"masuk": {"jam": "07:30:00", "tinjau_status": None}, "pulang": None, "cuti": None, "izin": None,
            "gagal_masuk": False, "gagal_pulang": False}


def test_tombol_dan_status_beranda_mengikuti_jendela_baru():
    sh = _sh_sudah_masuk()
    judul, sub, aktif, _ = A.teks_tombol_absen(jam(12, 29), sh)
    assert not aktif and "12:30" in sub
    judul, sub, aktif, _ = A.teks_tombol_absen(jam(14, 0), sh)
    assert aktif and "alasan" in sub
    judul, sub, aktif, _ = A.teks_tombol_absen(jam(16, 50), sh)
    assert aktif and sub == "Selfie + lokasi GPS"
    judul, sub, aktif, _ = A.teks_tombol_absen(jam(18, 1), sh)
    assert not aktif
    assert "wajib mengisi alasan" in A.teks_status_hari(jam(14, 0), sh)[1]
    assert "wajib mengisi alasan" not in A.teks_status_hari(jam(17, 0), sh)[1]


# ------------------------------------------------------------------
# Migrasi database lama
# ------------------------------------------------------------------
def test_migrasi_menambah_kolom_ke_tabel_absensi_lama(tmp_path):
    eng = create_engine(f"sqlite:///{tmp_path / 'lama.db'}")
    with eng.begin() as c:
        c.execute(text(
            "CREATE TABLE absensi (id INTEGER PRIMARY KEY AUTOINCREMENT, nama VARCHAR(100) NOT NULL, tanggal VARCHAR(10) NOT NULL, "
            "jam VARCHAR(8) NOT NULL, jenis VARCHAR(10) NOT NULL, lat FLOAT, lon FLOAT, akurasi FLOAT, jarak FLOAT, "
            "status VARCHAR(20) NOT NULL, uang_makan INTEGER NOT NULL, foto BLOB)"))
        c.execute(text("INSERT INTO absensi (nama, tanggal, jam, jenis, status, uang_makan) "
                       "VALUES ('Lama', '2026-10-01', '17:00:00', 'pulang', 'berhasil', 15000)"))
    A._migrasi(eng)
    kolom = {k["name"] for k in inspect(eng).get_columns("absensi")}
    assert {"tinjau_status", "tinjau_oleh", "tinjau_pada", "alasan"} <= kolom
    with eng.connect() as c:                                # data lama utuh, bukan pulang awal
        r = c.execute(select(A.absensi_t.c.uang_makan, A.absensi_t.c.tinjau_status)).one()
    assert tuple(r) == (15000, None)
    A._migrasi(eng)                                         # aman diulang
