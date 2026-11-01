"""Tes lapisan database modul proyek: audit, logbook, paket kerja, anggaran, pembimbing, privasi.
Jalankan: pytest -q test_proyek_db.py"""
import json
import threading
from datetime import date, datetime

import pandas as pd
import pytest
from sqlalchemy import text
from sqlalchemy.exc import IntegrityError

import app as A

WITA = A.WITA
JUM = datetime(2026, 11, 6, 10, 0, tzinfo=WITA)            # Jumat
DIV_A, DIV_B = "Divisi Umum", "Divisi Syariah"
URAIAN = "Menyusun laporan rekap kehadiran bulanan"


@pytest.fixture
def eng(eng_kosong):
    A.add_peserta(eng_kosong, "Sari Syariah", "B", DIV_B, "1234")
    return eng_kosong


def absen(eng, nama, tanggal, jenis="masuk", jam="07:30:00", uang=0, status="berhasil"):
    A.simpan_absen(eng, nama=nama, tanggal=tanggal, jam=jam, jenis=jenis, status=status, uang_makan=uang)


def audit_detail(eng, aksi):
    df = A.daftar_audit(eng, limit=500, aksi=aksi)
    return [json.loads(d) for d in df["detail"]]


def semua_teks_audit(eng):
    df = A.daftar_audit(eng, limit=1000)
    return " ".join(df["objek"].tolist() + df["detail"].tolist() + df["aksi"].tolist())


# =================== AUDIT ===================
def test_perubahan_peserta_tercatat_tanpa_membocorkan_rahasia(eng_kosong):
    e = eng_kosong
    A.add_peserta(e, "Rahasia Satu", "A", DIV_A, "987654", "130-210-000554616-9")
    A.set_rekening(e, "Rahasia Satu", "111-222-000999888-7")
    A.reset_pin(e, "Rahasia Satu", "135790")
    A.set_divisi(e, "Rahasia Satu", DIV_B, "B")
    aksi = list(A.daftar_audit(e)["aksi"][::-1])
    assert aksi[-4:] == ["peserta.tambah", "peserta.ubah_rekening", "peserta.reset_pin", "peserta.ubah_divisi"]
    teks = semua_teks_audit(e)
    for rahasia in ("987654", "135790", "130-210-000554616-9", "111-222-000999888-7", "0005546169"):
        assert rahasia not in teks, f"{rahasia} bocor ke jejak audit"
    assert audit_detail(e, "peserta.ubah_rekening")[0] == {"dari_akhir": "6169", "ke_akhir": "8887"}
    assert audit_detail(e, "peserta.ubah_divisi")[0] == {"dari": DIV_A, "ke": DIV_B, "gedung_dari": "A", "gedung_ke": "B"}


def test_aktor_dan_peran_tercatat(eng_kosong):
    A.add_peserta(eng_kosong, "Via Api", "A", DIV_A, "1234", aktor="api", peran="sistem")
    r = A.daftar_audit(eng_kosong, aksi="peserta.tambah", aktor="api").iloc[0]
    assert r["peran"] == "sistem" and r["objek"] == "peserta:Via Api"


def test_hapus_peserta_dan_absensi_tercatat_dengan_jumlah(eng):
    absen(eng, "P0", "2026-11-02")
    absen(eng, "P0", "2026-11-03")
    A.hapus_absensi(eng, "P0")
    A.hapus_peserta(eng, "Sari Syariah", dengan_absensi=True)
    assert audit_detail(eng, "absensi.hapus")[0] == {"jumlah": 2}
    assert audit_detail(eng, "peserta.hapus")[0] == {"dengan_riwayat": True, "absensi_dihapus": 0, "logbook_dihapus": 0}


def test_operasi_tanpa_efek_tidak_menulis_audit(eng):
    sebelum = len(A.daftar_audit(eng, limit=1000))
    A.reset_pin(eng, "Tidak Ada", "1234")
    A.set_rekening(eng, "Tidak Ada", "123-456-789")
    A.set_divisi(eng, "Tidak Ada", DIV_A, "A")
    A.hapus_peserta(eng, "Tidak Ada")
    A.hapus_absensi(eng, "Tidak Ada")
    assert A.set_anggaran(eng, 2026, 11, DIV_A, 1_000_000) is True
    assert A.set_anggaran(eng, 2026, 11, DIV_A, 1_000_000) is False            # nilai sama: tidak ada perubahan
    assert len(A.daftar_audit(eng, limit=1000)) == sebelum + 1


def test_perubahan_dan_jejak_atomik_audit_gagal_membatalkan_perubahan(eng, monkeypatch):
    def gagal(*a, **k):
        raise RuntimeError("audit gagal")
    monkeypatch.setattr(A, "_audit", gagal)
    with pytest.raises(RuntimeError):
        A.add_peserta(eng, "Tidak Boleh Ada", "A", DIV_A, "1234")
    assert A.get_peserta(eng, "Tidak Boleh Ada") is None                         # transaksi dibatalkan utuh


def test_perubahan_gagal_tidak_meninggalkan_jejak(eng):
    sebelum = len(A.daftar_audit(eng, limit=1000))
    with pytest.raises(IntegrityError):
        A.add_peserta(eng, "P0", "A", DIV_A, "1234")                              # nama sudah ada
    assert len(A.daftar_audit(eng, limit=1000)) == sebelum


def test_rantai_utuh_lalu_terdeteksi_bila_dirusak(eng):
    assert A.verifikasi_audit(eng)["ok"]
    A.set_anggaran(eng, 2026, 11, DIV_A, 1_000_000)
    A.set_anggaran(eng, 2026, 11, DIV_B, 500_000)
    A.reset_pin(eng, "P1", "2468")
    v = A.verifikasi_audit(eng)
    assert v["ok"] and v["jumlah"] >= 11 and len(v["hash_terakhir"]) == 64

    with eng.begin() as c:                                                        # penyerang mengubah isi satu catatan
        c.execute(text("UPDATE audit_log SET aktor = 'samar' WHERE aksi = 'anggaran.set' AND objek LIKE '%Umum%'"))
    rusak = A.verifikasi_audit(eng)
    assert not rusak["ok"] and "diubah" in rusak["alasan"]


def test_catatan_audit_dihapus_di_tengah_terdeteksi(eng):
    A.set_anggaran(eng, 2026, 11, DIV_A, 1_000_000)
    A.reset_pin(eng, "P1", "2468")
    A.reset_pin(eng, "P2", "2468")
    with eng.begin() as c:
        c.execute(text("DELETE FROM audit_log WHERE aksi = 'peserta.reset_pin' AND objek = 'peserta:P1'"))
    v = A.verifikasi_audit(eng)
    assert not v["ok"] and "terputus" in v["alasan"]


def test_rantai_tetap_utuh_di_bawah_penulisan_serentak(eng):
    """Beberapa thread menulis audit bersamaan: tidak ada catatan hilang dan rantai tidak bercabang."""
    sebelum = len(A.daftar_audit(eng, limit=5000))
    divisi = list(A.DIVISI_GEDUNG)[:4]
    galat = []

    def kerja(i):
        try:
            for bulan in range(1, 5):
                A.set_anggaran(eng, 2026, bulan, divisi[i], 100_000 * (i + 1) + bulan)
        except Exception as e:                                                    # noqa: BLE001
            galat.append(repr(e))

    th = [threading.Thread(target=kerja, args=(i,)) for i in range(4)]
    [t.start() for t in th]
    [t.join(timeout=60) for t in th]
    assert not galat, galat
    assert len(A.daftar_audit(eng, limit=5000)) == sebelum + 16
    assert A.verifikasi_audit(eng)["ok"]


# =================== PRIVASI ===================
def test_persetujuan_privasi_alur_dan_versi_baru(eng, monkeypatch):
    p = A.get_peserta(eng, "P0")
    assert A.perlu_persetujuan(p) is True
    A.catat_persetujuan(eng, "P0")
    p = A.get_peserta(eng, "P0")
    assert p["setuju_versi"] == A.VERSI_PRIVASI and p["setuju_pada"] and A.perlu_persetujuan(p) is False
    j = A.daftar_audit(eng, aksi="privasi.setuju").iloc[0]
    assert j["aktor"] == "P0" and j["peran"] == "peserta"
    monkeypatch.setattr(A, "VERSI_PRIVASI", "2099-01")                            # isi pemberitahuan berubah
    assert A.perlu_persetujuan(A.get_peserta(eng, "P0")) is True


def test_persetujuan_bisa_dimatikan_lewat_konfigurasi(eng, monkeypatch):
    monkeypatch.setenv("PRIVASI_WAJIB", "0")
    assert A.persetujuan_wajib() is False and A.perlu_persetujuan(A.get_peserta(eng, "P0")) is False


# =================== PEMBIMBING ===================
def test_pembimbing_tambah_login_reset_hapus(eng):
    A.add_pembimbing(eng, "Pak Budi", DIV_A, "123456")
    b = A.get_pembimbing(eng, "Pak Budi")
    assert b["pin_hash"].startswith("scrypt$") and "123456" not in b["pin_hash"]
    assert A.cek_pin_pembimbing(eng, "Pak Budi", "123456") is True
    assert A.cek_pin_pembimbing(eng, "Pak Budi", "654321") is False
    assert A.cek_pin_pembimbing(eng, "Tidak Ada", "123456") is False
    A.reset_pin_pembimbing(eng, "Pak Budi", "777888")
    assert A.cek_pin_pembimbing(eng, "Pak Budi", "123456") is False and A.cek_pin_pembimbing(eng, "Pak Budi", "777888") is True
    assert list(A.list_pembimbing(eng)["nama"]) == ["Pak Budi"]
    A.hapus_pembimbing(eng, "Pak Budi")
    assert A.get_pembimbing(eng, "Pak Budi") is None and A.cek_pin_pembimbing(eng, "Pak Budi", "777888") is False
    assert "123456" not in semua_teks_audit(eng) and "777888" not in semua_teks_audit(eng)


@pytest.mark.parametrize("nama,divisi,pin,cari", [
    ("Pak Budi", DIV_A, "1234", "6-8 angka"), ("Pak Budi", DIV_A, "12345a", "6-8 angka"),
    ("Pb", DIV_A, "123456", "minimal 3"), ("Pak Budi", "Divisi Fiktif", "123456", "tidak dikenal"),
])
def test_pembimbing_validasi(eng, nama, divisi, pin, cari):
    with pytest.raises(ValueError, match=cari):
        A.add_pembimbing(eng, nama, divisi, pin)
    assert A.list_pembimbing(eng).empty


def test_pembimbing_nama_ganda_ditolak(eng):
    A.add_pembimbing(eng, "Pak Budi", DIV_A, "123456")
    with pytest.raises(IntegrityError):
        A.add_pembimbing(eng, "Pak Budi", DIV_B, "123456")


# =================== PAKET KERJA ===================
MULAI, SELESAI = date(2026, 11, 2), date(2026, 11, 13)


def test_paket_tambah_ubah_tutup_hapus_dan_audit(eng):
    pid = A.tambah_paket(eng, DIV_A, "1", "Analisis kebutuhan", MULAI, SELESAI, 40)
    df = A.daftar_paket(eng, DIV_A)
    assert list(df["kode"]) == ["1"] and df.iloc[0]["mulai"] == MULAI and df.iloc[0]["ditutup"] == 0
    A.ubah_paket(eng, pid, "1.1", "Analisis kebutuhan sistem", MULAI, SELESAI, 50)
    assert A.daftar_paket(eng).iloc[0]["estimasi_jam"] == 50
    A.tutup_paket(eng, pid, True)
    assert A.daftar_paket(eng).iloc[0]["ditutup"] == 1
    A.hapus_paket(eng, pid)
    assert A.daftar_paket(eng).empty
    aksi = list(A.daftar_audit(eng, limit=50)["aksi"][::-1])
    assert [a for a in aksi if a.startswith("paket.")] == ["paket.tambah", "paket.ubah", "paket.tutup", "paket.hapus"]
    assert audit_detail(eng, "paket.ubah")[0]["dari"]["kode"] == "1" and audit_detail(eng, "paket.ubah")[0]["ke"]["kode"] == "1.1"


def test_paket_validasi_dan_kode_ganda_per_divisi(eng):
    A.tambah_paket(eng, DIV_A, "1", "Analisis kebutuhan", MULAI, SELESAI, 40)
    with pytest.raises(ValueError, match="sudah dipakai"):
        A.tambah_paket(eng, DIV_A, "1", "Nama lain lagi", MULAI, SELESAI, 40)
    A.tambah_paket(eng, DIV_B, "1", "Kode sama divisi lain", MULAI, SELESAI, 40)           # boleh: unik per divisi
    with pytest.raises(ValueError, match="Kode WBS"):
        A.tambah_paket(eng, DIV_A, "x", "Nama paket", MULAI, SELESAI, 40)
    with pytest.raises(ValueError, match="Divisi tidak dikenal"):
        A.tambah_paket(eng, "Divisi Fiktif", "1", "Nama paket", MULAI, SELESAI, 40)
    with pytest.raises(ValueError, match="tidak ditemukan"):
        A.tutup_paket(eng, 9999, True)


def test_pembimbing_tidak_bisa_menyentuh_paket_divisi_lain(eng):
    pid_b = A.tambah_paket(eng, DIV_B, "1", "Paket milik divisi lain", MULAI, SELESAI, 40)
    kw = dict(aktor="Pak Budi", peran="pembimbing", divisi_scope=DIV_A)
    with pytest.raises(PermissionError):
        A.tambah_paket(eng, DIV_B, "2", "Menyusup ke divisi lain", MULAI, SELESAI, 40, **kw)
    with pytest.raises(PermissionError):
        A.ubah_paket(eng, pid_b, "1", "Diubah penyusup", MULAI, SELESAI, 40, **kw)
    with pytest.raises(PermissionError):
        A.tutup_paket(eng, pid_b, True, **kw)
    with pytest.raises(PermissionError):
        A.hapus_paket(eng, pid_b, **kw)
    with pytest.raises(PermissionError):                                           # pembimbing tanpa divisi = tidak sah
        A.hapus_paket(eng, pid_b, aktor="X", peran="pembimbing", divisi_scope=None)
    p = A.daftar_paket(eng, DIV_B).iloc[0]
    assert p["nama"] == "Paket milik divisi lain" and p["ditutup"] == 0               # tidak ada yang berubah
    assert A.tambah_paket(eng, DIV_A, "1", "Paket divisi sendiri", MULAI, SELESAI, 40, **kw) > 0


def test_hapus_paket_melepas_logbook_tanpa_menghapusnya(eng):
    pid = A.tambah_paket(eng, DIV_A, "1", "Analisis kebutuhan", MULAI, SELESAI, 40)
    lid = A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, pid, now=JUM)
    A.hapus_paket(eng, pid)
    d = A.daftar_logbook(eng, nama="P0")
    assert len(d) == 1 and d.iloc[0]["id"] == lid and pd.isna(d.iloc[0]["paket_id"])
    assert audit_detail(eng, "paket.hapus")[0]["logbook_dilepas"] == 1


# =================== LOGBOOK ===================
def test_logbook_diajukan_lalu_disetujui_dan_masuk_progres(eng):
    pid = A.tambah_paket(eng, DIV_A, "1", "Analisis kebutuhan", MULAI, SELESAI, 40)
    l1 = A.tambah_logbook(eng, "P0", date(2026, 11, 3), URAIAN, 8, pid, now=JUM)
    l2 = A.tambah_logbook(eng, "P1", date(2026, 11, 4), URAIAN, 12 - 4, pid, now=JUM)
    l3 = A.tambah_logbook(eng, "P2", date(2026, 11, 5), URAIAN, 4, pid, now=JUM)       # tidak akan disetujui
    assert A.progres_divisi(eng, DIV_A, JUM.date()).iloc[0]["jam_disetujui"] == 0     # belum ada yang disetujui
    A.tinjau_logbook(eng, l1, "disetujui", "", "Pak Budi", "pembimbing", DIV_A)
    A.tinjau_logbook(eng, l2, "disetujui", "bagus", "Pak Budi", "pembimbing", DIV_A)
    A.tinjau_logbook(eng, l3, "ditolak", "uraian kurang jelas", "Pak Budi", "pembimbing", DIV_A)
    r = A.progres_divisi(eng, DIV_A, date(2026, 11, 6)).iloc[0]
    assert r["jam_disetujui"] == 16 and r["progres"] == pytest.approx(0.4) and r["rencana"] == pytest.approx(0.5)
    assert r["spi"] == pytest.approx(0.8) and r["status"] == "Terlambat"             # 0,8 < 0,9
    assert A.jam_disetujui_per_paket(eng, DIV_B) == {}


@pytest.mark.parametrize("kwargs,cari", [
    (dict(tanggal=date(2026, 11, 7)), "masa depan"),
    (dict(tanggal=date(2026, 11, 1)), "libur"),
    (dict(tanggal=date(2026, 10, 1)), "14 hari"),
    (dict(uraian="pendek"), "minimal 15"),
    (dict(durasi=0), "kelipatan"),
    (dict(durasi=11), "kelipatan"),
])
def test_logbook_ditolak_validasi(eng, kwargs, cari):
    isi = dict(tanggal=date(2026, 11, 5), uraian=URAIAN, durasi=2)
    isi.update(kwargs)
    with pytest.raises(ValueError, match=cari):
        A.tambah_logbook(eng, "P0", isi["tanggal"], isi["uraian"], isi["durasi"], now=JUM)
    assert A.daftar_logbook(eng).empty


def test_logbook_batas_12_jam_per_hari_menghitung_yang_tidak_ditolak(eng):
    t = date(2026, 11, 5)
    A.tambah_logbook(eng, "P0", t, URAIAN, 8, now=JUM)
    with pytest.raises(ValueError, match="melebihi 12"):
        A.tambah_logbook(eng, "P0", t, URAIAN + " tambahan", 5, now=JUM)
    A.tambah_logbook(eng, "P0", t, URAIAN + " tambahan", 4, now=JUM)                 # pas 12
    A.tambah_logbook(eng, "P1", t, URAIAN, 8, now=JUM)                               # peserta lain tidak terpengaruh
    d = A.daftar_logbook(eng, nama="P0")
    A.tinjau_logbook(eng, int(d[d["durasi_jam"] == 8].iloc[0]["id"]), "ditolak", "perlu diperbaiki lagi", "Admin", "admin")
    A.tambah_logbook(eng, "P0", t, URAIAN + " pengganti", 6, now=JUM)               # jam yang ditolak tidak dihitung lagi


def test_logbook_paket_harus_dari_divisi_sendiri(eng):
    pid_b = A.tambah_paket(eng, DIV_B, "1", "Paket divisi lain", MULAI, SELESAI, 40)
    with pytest.raises(ValueError, match="tidak termasuk divisi kamu"):
        A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, pid_b, now=JUM)
    with pytest.raises(ValueError, match="Peserta tidak ditemukan"):
        A.tambah_logbook(eng, "Hantu", date(2026, 11, 5), URAIAN, 2, now=JUM)


def test_ubah_dan_hapus_hanya_milik_sendiri_dan_belum_disetujui(eng):
    lid = A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, now=JUM)
    with pytest.raises(ValueError, match="tidak ditemukan"):                          # milik orang lain
        A.ubah_logbook(eng, lid, "P1", URAIAN + " dibajak", 3, now=JUM)
    with pytest.raises(ValueError, match="tidak ditemukan"):
        A.hapus_logbook(eng, lid, "P1")
    A.ubah_logbook(eng, lid, "P0", URAIAN + " direvisi", 3, now=JUM)
    d = A.daftar_logbook(eng, nama="P0").iloc[0]
    assert d["uraian"].endswith("direvisi") and d["durasi_jam"] == 3 and d["status"] == "diajukan"
    A.tinjau_logbook(eng, lid, "disetujui", "", "Admin", "admin")
    with pytest.raises(ValueError, match="sudah disetujui"):
        A.ubah_logbook(eng, lid, "P0", URAIAN + " lagi", 3, now=JUM)
    with pytest.raises(ValueError, match="sudah disetujui"):
        A.hapus_logbook(eng, lid, "P0")


def test_logbook_ditolak_diperbaiki_kembali_diajukan_dan_catatan_dibersihkan(eng):
    lid = A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, now=JUM)
    A.tinjau_logbook(eng, lid, "ditolak", "tambahkan hasil yang dicapai", "Pak Budi", "pembimbing", DIV_A)
    d = A.daftar_logbook(eng, nama="P0").iloc[0]
    assert d["status"] == "ditolak" and d["catatan_peninjau"] == "tambahkan hasil yang dicapai" and d["ditinjau_oleh"] == "Pak Budi"
    A.ubah_logbook(eng, lid, "P0", URAIAN + " serta hasil yang dicapai", 2, now=JUM)
    d = A.daftar_logbook(eng, nama="P0").iloc[0]
    assert d["status"] == "diajukan" and d["catatan_peninjau"] is None and d["ditinjau_oleh"] is None
    assert any(a["catatan"] == "tambahkan hasil yang dicapai" for a in audit_detail(eng, "logbook.tolak"))     # alasan tetap ada di audit


def test_tinjau_aturan_dan_wewenang(eng):
    lid = A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, now=JUM)          # divisi Umum
    kw = dict(aktor="Bu Sari", peran="pembimbing")
    with pytest.raises(PermissionError):                                               # pembimbing divisi LAIN
        A.tinjau_logbook(eng, lid, "disetujui", "", divisi_scope=DIV_B, **kw)
    with pytest.raises(PermissionError):                                               # pembimbing tanpa divisi
        A.tinjau_logbook(eng, lid, "disetujui", "", divisi_scope=None, **kw)
    assert A.daftar_logbook(eng).iloc[0]["status"] == "diajukan"                       # tidak berubah
    with pytest.raises(ValueError, match="Alasan penolakan"):
        A.tinjau_logbook(eng, lid, "ditolak", "  ", divisi_scope=DIV_A, **kw)
    with pytest.raises(ValueError, match="Keputusan"):
        A.tinjau_logbook(eng, lid, "mungkin", "", divisi_scope=DIV_A, **kw)
    with pytest.raises(ValueError, match="tidak ditemukan"):
        A.tinjau_logbook(eng, 9999, "disetujui", "", divisi_scope=DIV_A, **kw)
    A.tinjau_logbook(eng, lid, "disetujui", "", divisi_scope=DIV_A, **kw)
    with pytest.raises(ValueError, match="sudah disetujui"):                           # tidak bisa ditinjau dua kali
        A.tinjau_logbook(eng, lid, "ditolak", "berubah pikiran", divisi_scope=DIV_A, **kw)
    r = A.daftar_logbook(eng).iloc[0]
    assert r["status"] == "disetujui" and r["ditinjau_oleh"] == "Bu Sari" and r["ditinjau_pada"]
    assert audit_detail(eng, "logbook.setuju")[0]["peserta"] == "P0"


def test_divisi_logbook_adalah_snapshot_saat_dibuat(eng):
    lid = A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, now=JUM)
    A.set_divisi(eng, "P0", DIV_B, "B")                                                # peserta pindah divisi
    assert A.daftar_logbook(eng).iloc[0]["divisi"] == DIV_A                            # tetap ditinjau pembimbing divisi lama
    with pytest.raises(PermissionError):
        A.tinjau_logbook(eng, lid, "disetujui", "", "Bu Sari", "pembimbing", DIV_B)
    A.tinjau_logbook(eng, lid, "disetujui", "", "Pak Budi", "pembimbing", DIV_A)


def test_daftar_logbook_filter_dan_penanda_tanpa_absensi(eng):
    absen(eng, "P0", "2026-11-05")                                                     # P0 hadir 5 Nov, tidak 4 Nov
    A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, now=JUM)
    A.tambah_logbook(eng, "P0", date(2026, 11, 4), URAIAN, 2, now=JUM)
    A.tambah_logbook(eng, "Sari Syariah", date(2026, 11, 5), URAIAN, 2, now=JUM)
    d = A.daftar_logbook(eng)
    flag = {(r.nama, r.tanggal): r.tanpa_absensi for r in d.itertuples()}
    assert flag == {("P0", "2026-11-05"): False, ("P0", "2026-11-04"): True, ("Sari Syariah", "2026-11-05"): True}
    assert len(A.daftar_logbook(eng, divisi=DIV_B)) == 1 and len(A.daftar_logbook(eng, nama="P0")) == 2
    assert len(A.daftar_logbook(eng, awal=date(2026, 11, 5))) == 2 and len(A.daftar_logbook(eng, status="disetujui")) == 0
    kosong = A.daftar_logbook(eng, nama="Tidak Ada")
    assert kosong.empty and "tanpa_absensi" in kosong.columns


def test_hapus_peserta_logbook_ikut_hapus_atau_dipertahankan(eng):
    A.tambah_logbook(eng, "P0", date(2026, 11, 5), URAIAN, 2, now=JUM)
    A.tambah_logbook(eng, "P1", date(2026, 11, 5), URAIAN, 2, now=JUM)
    A.hapus_peserta(eng, "P0", dengan_absensi=False)
    assert len(A.daftar_logbook(eng, nama="P0")) == 1                                  # dipertahankan sebagai riwayat
    A.hapus_peserta(eng, "P1", dengan_absensi=True)
    assert A.daftar_logbook(eng, nama="P1").empty and audit_detail(eng, "peserta.hapus")[0]["logbook_dihapus"] == 1


# =================== ANGGARAN ===================
def test_set_hapus_anggaran_dan_validasi(eng):
    assert A.set_anggaran(eng, 2026, 11, DIV_A, 1_000_000, "awal") is True
    assert A.set_anggaran(eng, 2026, 11, DIV_A, 1_200_000, "revisi") is True
    assert A.get_anggaran(eng, 2026, 11) == {DIV_A: 1_200_000}
    d = audit_detail(eng, "anggaran.set")
    assert [(x["dari"], x["ke"]) for x in d] == [(1_000_000, 1_200_000), (None, 1_000_000)]
    assert A.hapus_anggaran(eng, 2026, 11, DIV_A) is True and A.hapus_anggaran(eng, 2026, 11, DIV_A) is False
    for args in ((2026, 13, DIV_A, 1), (1999, 1, DIV_A, 1), (2026, 11, "Fiktif", 1), (2026, 11, DIV_A, -1), (2026, 11, DIV_A, 10 ** 11)):
        with pytest.raises(ValueError):
            A.set_anggaran(eng, *args)


NOW_5NOV = datetime(2026, 11, 5, 8, 0, tzinfo=WITA)          # Kamis pagi: 2, 3, 4 Nov sudah selesai; 5 Nov belum


def isi_pulang(eng, nama, tanggal_list):
    for t in tanggal_list:
        absen(eng, nama, t, jenis="pulang", jam="17:05:00", uang=A.UANG_MAKAN)


def test_ringkasan_anggaran_angka_manual(eng):
    for n in ("P0", "P1", "P2"):
        isi_pulang(eng, n, ["2026-11-02", "2026-11-03", "2026-11-04"])
    A.set_anggaran(eng, 2026, 11, DIV_A, 1_000_000)
    df, meta = A.ringkasan_anggaran(eng, 2026, 11, NOW_5NOV)
    assert meta == {"hari_total": 21, "hari_selesai": 3, "hari_sisa": 18}
    r = df.set_index("divisi").loc[DIV_A]
    assert r["peserta"] == 8 and r["realisasi"] == 135_000                              # 3 peserta x 3 hari x 15.000
    assert r["proyeksi"] == 135_000 + 45_000 * 18 == 945_000
    assert r["proyeksi_maks"] == 135_000 + 8 * 15_000 * 18 == 2_295_000
    assert r["sisa_anggaran"] == 865_000 and r["selisih"] == 55_000 and r["status"] == "Waspada"      # 0,945 > 0,9
    assert r["serapan_pct"] == pytest.approx(13.5)
    s = df.set_index("divisi").loc[DIV_B]
    assert s["anggaran"] == 0 and s["status"] == "Belum ada anggaran" and s["realisasi"] == 0


def test_hari_ini_belum_selesai_tidak_menekan_rata_rata(eng):
    for n in ("P0", "P1"):
        isi_pulang(eng, n, ["2026-11-02", "2026-11-03", "2026-11-04"])               # 30.000 per hari, 3 hari = 90.000
    isi_pulang(eng, "P2", ["2026-11-05"])                                              # hari ini baru satu orang yang pulang
    pagi = A.ringkasan_anggaran(eng, 2026, 11, NOW_5NOV)[0].set_index("divisi").loc[DIV_A]
    assert pagi["realisasi"] == 105_000                                                # yang sudah keluar tetap tercatat
    assert pagi["proyeksi"] == 90_000 + 30_000 * 18 == 630_000                         # hari ini dihitung SISA, bukan menekan rata-rata
    malam = A.ringkasan_anggaran(eng, 2026, 11, datetime(2026, 11, 5, 18, 30, tzinfo=WITA))[0].set_index("divisi").loc[DIV_A]
    assert malam["proyeksi"] == 105_000 + 26_250 * 17 == 551_250                       # setelah jam pulang, hari ini ikut rata-rata


def test_snapshot_divisi_menjaga_anggaran_saat_peserta_pindah(eng):
    isi_pulang(eng, "P0", ["2026-11-02"])
    A.set_divisi(eng, "P0", DIV_B, "B")
    isi_pulang(eng, "P0", ["2026-11-03"])
    real = A.realisasi_harian(eng, 2026, 11)
    peta = {(r.divisi, r.tanggal): r.jumlah for r in real.itertuples()}
    assert peta == {(DIV_A, date(2026, 11, 2)): 15_000, (DIV_B, date(2026, 11, 3)): 15_000}


def test_tingkat_kehadiran_historis_dan_saran(eng):
    okt = [d.isoformat() for d in A.proyek.hari_kerja_bulan(2026, 10)]                 # 22 hari kerja
    isi_pulang(eng, "P0", okt)                                                          # P0 hadir penuh
    isi_pulang(eng, "P1", okt[:11])                                                     # P1 setengahnya
    assert A.tingkat_kehadiran_historis(eng, 2026, 11) == {DIV_A: pytest.approx((22 + 11) / (2 * 22))}
    saran = A.saran_anggaran_bulan(eng, 2026, 11).set_index("divisi")
    a = saran.loc[DIV_A]
    assert a["peserta"] == 8 and a["tingkat_kehadiran"] == pytest.approx(0.75) and not a["asumsi_kehadiran"]
    assert a["estimasi_dasar"] == 1_890_000 and a["cadangan"] == 189_000 and a["saran"] == 2_080_000     # 2.079.000 dibulatkan
    assert saran.loc[DIV_B]["asumsi_kehadiran"]                                          # tanpa riwayat: asumsi 90%


def test_terapkan_saran_tidak_menimpa_anggaran_manual(eng):
    A.set_anggaran(eng, 2026, 11, DIV_A, 777_000, "diputuskan manual")
    assert A.terapkan_saran_anggaran(eng, 2026, 11) == 1                               # hanya DIV_B yang belum punya
    ang = A.get_anggaran(eng, 2026, 11)
    assert ang[DIV_A] == 777_000 and ang[DIV_B] > 0
    assert A.terapkan_saran_anggaran(eng, 2026, 11) == 0                               # idempoten
    assert A.terapkan_saran_anggaran(eng, 2026, 11, timpa=True) == 1                   # timpa mengganti DIV_A
    assert A.get_anggaran(eng, 2026, 11)[DIV_A] != 777_000


def test_migrasi_mengisi_snapshot_divisi_dan_kolom_privasi_tersedia(tmp_path):
    from sqlalchemy import create_engine, inspect
    e = create_engine(f"sqlite:///{tmp_path / 'm.db'}")
    A.metadata.create_all(e)
    A._migrasi(e)
    assert {"divisi"} <= {c["name"] for c in inspect(e).get_columns("absensi")}
    assert {"setuju_versi", "setuju_pada"} <= {c["name"] for c in inspect(e).get_columns("peserta")}
    assert {"audit_log", "pembimbing", "paket_kerja", "logbook", "anggaran"} <= set(inspect(e).get_table_names())
