"""Tes umpan balik peserta: rating sekali, ulasan bertaut ke akun (tanpa nama di tampilan admin), syarat seminggu,
status oleh admin yang bisa dilihat pemberinya, dan tampilan hanya di aplikasi.
Jalankan: pytest -q test_umpan_balik.py"""
import json
from datetime import date, timedelta
from pathlib import Path

import pytest
import streamlit as st
from sqlalchemy import inspect, text
from streamlit.testing.v1 import AppTest

import app as A

HARI_INI = date(2026, 11, 20)
ULASAN = "Tombol absen pulang agak sulit ditemukan di layar kecil"
APP = str(Path(__file__).parent / "app.py")


def tgl(selisih_hari):
    return (HARI_INI - timedelta(days=selisih_hari)).isoformat()


def absen(eng, nama, tanggal, status="berhasil"):
    A.simpan_absen(eng, nama=nama, tanggal=tanggal, jam="07:30:00", jenis="masuk", status=status, uang_makan=0)


@pytest.fixture
def eng(eng_kosong):
    A.add_peserta(eng_kosong, "Andi", "A", "Divisi Umum", "1234")
    A.add_peserta(eng_kosong, "Sari", "B", "Divisi Syariah", "1234")
    return eng_kosong


# =================== SYARAT SEMINGGU ===================
def test_belum_pernah_absen_tidak_diminta(eng):
    assert A.hari_pemakaian(eng, "Andi", HARI_INI) is None
    assert A.perlu_diminta_umpan_balik(eng, "Andi", HARI_INI) is False


def test_hari_pemakaian_dihitung_dari_absen_pertama_yang_berhasil(eng):
    absen(eng, "Andi", tgl(20), status="gagal_lokasi")       # percobaan gagal tidak dihitung
    absen(eng, "Andi", tgl(8))
    absen(eng, "Andi", tgl(1))
    assert A.hari_pemakaian(eng, "Andi", HARI_INI) == 8


@pytest.mark.parametrize("selisih,diminta", [(6, False), (7, True), (30, True)])
def test_diminta_setelah_tujuh_hari(eng, selisih, diminta):
    absen(eng, "Andi", tgl(selisih))
    assert A.perlu_diminta_umpan_balik(eng, "Andi", HARI_INI) is diminta


# =================== RATING SEKALI ===================
def test_rating_hanya_sekali_dan_tidak_diminta_lagi(eng):
    absen(eng, "Andi", tgl(10))
    A.simpan_rating(eng, "Andi", 4)
    assert A.rating_peserta(eng, "Andi") == 4
    assert A.perlu_diminta_umpan_balik(eng, "Andi", HARI_INI) is False
    with pytest.raises(ValueError, match="sudah memberi penilaian"):
        A.simpan_rating(eng, "Andi", 5)
    assert A.rating_peserta(eng, "Andi") == 4                # nilai pertama tidak berubah


@pytest.mark.parametrize("bintang", [0, 6, -1, "5", 4.5, True, None])
def test_rating_tidak_valid_ditolak(eng, bintang):
    with pytest.raises(ValueError):
        A.simpan_rating(eng, "Andi", bintang)
    assert A.rating_peserta(eng, "Andi") is None


def test_rating_dengan_ulasan_ikut_tersimpan_bertaut_ke_akun_tanpa_nama_di_tampilan_admin(eng):
    A.simpan_rating(eng, "Andi", 5, ULASAN, "absen")
    df = A.daftar_ulasan(eng)
    assert list(df["isi"]) == [ULASAN] and list(df["divisi"]) == ["Divisi Umum"] and list(df["status"]) == ["baru"]
    kolom = {c["name"] for c in inspect(eng).get_columns("ulasan")}
    assert "nama" not in kolom                                # nama tidak pernah disalin ke tabel ulasan
    assert "peserta_id" in kolom                              # tetapi ulasan bertaut ke akun (agar pemberinya bisa melihat status)
    assert "peserta_id" not in df.columns and "nama" not in df.columns      # tampilan admin tidak memuat penghubung
    assert "peserta_id" not in df.to_csv(index=False)         # ekspor CSV juga tidak
    assert len(df["dibuat"].iloc[0]) == 10                    # hanya tanggal, bukan jam


def test_rating_ulasan_dan_audit_tidak_membuka_identitas(eng):
    A.simpan_rating(eng, "Andi", 5, ULASAN, "absen")
    A.simpan_ulasan(eng, "Andi", ULASAN, "logbook")
    semua = " ".join(A.daftar_audit(eng, limit=500)["aksi"].tolist() + A.daftar_audit(eng, limit=500)["objek"].tolist())
    assert "umpan" not in semua and "ulasan.kirim" not in semua    # pengiriman peserta sengaja tidak diaudit


def test_rating_kosong_ulasan_tidak_membuat_baris_ulasan(eng):
    A.simpan_rating(eng, "Andi", 3, "   ")
    assert A.daftar_ulasan(eng).empty


def test_rating_dengan_ulasan_pendek_gagal_dan_tidak_menyimpan_apa_pun(eng):
    with pytest.raises(ValueError, match="minimal"):
        A.simpan_rating(eng, "Andi", 5, "bagus")
    assert A.rating_peserta(eng, "Andi") is None and A.daftar_ulasan(eng).empty


# =================== ULASAN DARI PROFIL ===================
def test_ulasan_dari_profil_boleh_berkali_kali(eng):
    A.simpan_ulasan(eng, "Andi", ULASAN, "tampilan")
    A.simpan_ulasan(eng, "Andi", "Logbook sebaiknya bisa dilampiri foto", "logbook")
    assert len(A.daftar_ulasan(eng)) == 2


@pytest.mark.parametrize("isi,kategori,potongan", [
    ("", "absen", "minimal"), ("pendek", "absen", "minimal"), ("x" * 501, "absen", "maksimal"), (ULASAN, "ngawur", "kategori"),
])
def test_ulasan_tidak_valid_ditolak(eng, isi, kategori, potongan):
    with pytest.raises(ValueError, match=potongan):
        A.simpan_ulasan(eng, "Andi", isi, kategori)
    assert A.daftar_ulasan(eng).empty


def test_peserta_tidak_dikenal_ditolak(eng):
    with pytest.raises(ValueError, match="tidak ditemukan"):
        A.simpan_ulasan(eng, "Hantu", ULASAN, "absen")


# =================== ULASAN SAYA (peserta melihat tindak lanjut) ===================
def test_peserta_hanya_melihat_ulasannya_sendiri(eng):
    A.simpan_ulasan(eng, "Andi", ULASAN, "absen")
    A.simpan_rating(eng, "Andi", 4, "Logbook sebaiknya bisa dilampiri foto", "logbook")
    A.simpan_ulasan(eng, "Sari", "Rekap bulanan sebaiknya bisa diunduh", "lainnya")
    milik_andi = A.ulasan_peserta(eng, "Andi")
    milik_sari = A.ulasan_peserta(eng, "Sari")
    assert {u["isi"] for u in milik_andi} == {ULASAN, "Logbook sebaiknya bisa dilampiri foto"}
    assert [u["isi"] for u in milik_sari] == ["Rekap bulanan sebaiknya bisa diunduh"]
    assert A.ulasan_peserta(eng, "Hantu") == []
    assert milik_andi[0]["kategori"] == "logbook"             # terbaru dulu


def test_peserta_melihat_status_dan_catatan_admin(eng):
    A.simpan_ulasan(eng, "Andi", ULASAN, "absen")
    assert A.ulasan_peserta(eng, "Andi")[0]["status"] == "baru"
    uid = int(A.daftar_ulasan(eng)["id"].iloc[0])
    A.ubah_status_ulasan(eng, uid, "selesai", "Tombol sudah dipindah di versi 1.2")
    u = A.ulasan_peserta(eng, "Andi")[0]
    assert u["status"] == "selesai" and u["catatan_admin"] == "Tombol sudah dipindah di versi 1.2"
    assert A.ulasan_peserta(eng, "Sari") == []                # peserta lain tidak ikut melihat


def test_ulasan_lama_tanpa_penghubung_tidak_muncul_di_profil_siapa_pun(eng):
    with eng.begin() as c:                                    # meniru baris dari sebelum kolom peserta_id ada
        c.execute(text("INSERT INTO ulasan (divisi, kategori, isi, dibuat, status) "
                       "VALUES ('Divisi Umum', 'absen', 'ulasan lama sebelum fitur', '2026-10-01', 'baru')"))
    assert A.ulasan_peserta(eng, "Andi") == [] and A.ulasan_peserta(eng, "Sari") == []
    assert len(A.daftar_ulasan(eng)) == 1                     # admin tetap melihatnya


def test_migrasi_menambah_kolom_peserta_id_pada_database_lama(eng):
    with eng.begin() as c:
        c.execute(text("DROP TABLE ulasan"))
        c.execute(text("CREATE TABLE ulasan (id INTEGER PRIMARY KEY AUTOINCREMENT, divisi VARCHAR(100), "
                       "kategori VARCHAR(20) NOT NULL, isi VARCHAR(500) NOT NULL, dibuat VARCHAR(10) NOT NULL, "
                       "status VARCHAR(12) NOT NULL, catatan_admin VARCHAR(300))"))
    assert "peserta_id" not in {c["name"] for c in inspect(eng).get_columns("ulasan")}
    A._migrasi(eng)
    A._migrasi(eng)                                           # aman diulang
    assert "peserta_id" in {c["name"] for c in inspect(eng).get_columns("ulasan")}
    A.simpan_ulasan(eng, "Andi", ULASAN, "absen")
    assert len(A.ulasan_peserta(eng, "Andi")) == 1


# =================== ADMIN ===================
def test_ringkasan_rating(eng):
    assert A.ringkasan_rating(eng)["rata"] is None
    A.simpan_rating(eng, "Andi", 5)
    A.simpan_rating(eng, "Sari", 3)
    r = A.ringkasan_rating(eng)
    assert r["n"] == 2 and r["rata"] == pytest.approx(4.0) and r["sebaran"][5] == 1 and r["sebaran"][3] == 1
    assert r["n_peserta"] >= 2


def test_status_ulasan_diubah_admin_dan_tercatat_di_audit(eng):
    A.simpan_ulasan(eng, "Andi", ULASAN, "absen")
    uid = int(A.daftar_ulasan(eng)["id"].iloc[0])
    A.ubah_status_ulasan(eng, uid, "diproses", "Tombol sudah dipindah di versi berikutnya")
    row = A.daftar_ulasan(eng).iloc[0]
    assert row["status"] == "diproses" and row["catatan_admin"].startswith("Tombol sudah")
    detail = [json.loads(d) for d in A.daftar_audit(eng, aksi="ulasan.status")["detail"]]
    assert detail == [{"dari": "baru", "ke": "diproses"}]
    assert A.verifikasi_audit(eng)["ok"]                     # rantai audit tetap utuh


def test_status_ulasan_tidak_dikenal_ditolak(eng):
    A.simpan_ulasan(eng, "Andi", ULASAN, "absen")
    uid = int(A.daftar_ulasan(eng)["id"].iloc[0])
    with pytest.raises(ValueError):
        A.ubah_status_ulasan(eng, uid, "dihapus")
    with pytest.raises(ValueError, match="tidak ditemukan"):
        A.ubah_status_ulasan(eng, 9999, "selesai")


def test_hapus_peserta_rating_tetap_sebagai_statistik_tapi_dilepas(eng):
    A.simpan_rating(eng, "Andi", 4)
    A.hapus_peserta(eng, "Andi")
    assert A.ringkasan_rating(eng)["n"] == 1                  # statistik tidak hilang
    with eng.connect() as c:
        assert c.execute(text("SELECT peserta_id FROM rating_aplikasi")).scalar() is None
    A.add_peserta(eng, "Andi", "A", "Divisi Umum", "1234")    # nama yang sama didaftarkan lagi = akun baru
    assert A.rating_peserta(eng, "Andi") is None


def test_hapus_peserta_ulasan_tetap_sebagai_masukan_tapi_dilepas(eng):
    A.simpan_ulasan(eng, "Andi", ULASAN, "absen")
    A.hapus_peserta(eng, "Andi")
    assert len(A.daftar_ulasan(eng)) == 1                     # masukan tidak hilang
    with eng.connect() as c:
        assert c.execute(text("SELECT peserta_id FROM ulasan")).scalar() is None
    A.add_peserta(eng, "Andi", "A", "Divisi Umum", "1234")    # nama yang sama didaftarkan lagi = akun baru
    assert A.ulasan_peserta(eng, "Andi") == []                # tidak mewarisi ulasan akun lama


# =================== TAMPILAN (AppTest) ===================
@pytest.fixture
def ui(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("PRIVASI_WAJIB", raising=False)
    st.cache_resource.clear()
    eng = A.get_engine()
    A.add_peserta(eng, "Andi", "A", "Divisi Umum", "1234")
    A.catat_persetujuan(eng, "Andi")
    yield eng
    st.cache_resource.clear()


def jalan(di_app, **state):
    at = AppTest.from_file(APP, default_timeout=90)
    if di_app:
        at.query_params["mode"] = "pkl"
    for k, v in state.items():
        at.session_state[k] = v
    return at.run()


def tanpa_error(at):
    assert not at.exception, [e.value[:300] for e in at.exception]


def absen_lama(eng, hari=10):
    absen(eng, "Andi", (date.today() - timedelta(days=hari)).isoformat())


def test_di_browser_fitur_tidak_muncul_sama_sekali(ui):
    absen_lama(ui)
    at = jalan(False, peran="user", nama_login="Andi", tab="profil")
    tanpa_error(at)
    assert [r for r in at.radio if r.key and r.key.startswith("ub_")] == []
    assert [t for t in at.text_area if t.key and t.key.startswith("ub_")] == []


def test_di_aplikasi_diarahkan_setelah_seminggu_lewati_hanya_untuk_sesi(ui):
    absen_lama(ui)
    at = jalan(True, peran="user", nama_login="Andi")
    tanpa_error(at)
    assert at.radio(key="ub_bintang_awal") is not None        # diarahkan ke form
    at.button(key="ub_lewati_awal").click().run()
    tanpa_error(at)
    assert [r for r in at.radio if r.key == "ub_bintang_awal"] == []     # lanjut ke beranda
    assert A.rating_peserta(ui, "Andi") is None                          # belum diisi
    sesi_baru = jalan(True, peran="user", nama_login="Andi")             # buka aplikasi lagi
    assert sesi_baru.radio(key="ub_bintang_awal") is not None            # muncul lagi


def test_kirim_rating_lalu_tidak_muncul_lagi_dan_ulasan_ada_di_profil(ui):
    absen_lama(ui)
    at = jalan(True, peran="user", nama_login="Andi")
    at.button(key="ub_kirim_awal").click().run()
    assert any("Pilih jumlah bintang" in e.value for e in at.error) and A.rating_peserta(ui, "Andi") is None
    at.radio(key="ub_bintang_awal").set_value(5)
    at.text_area(key="ub_isi_awal").set_value(ULASAN)
    at.button(key="ub_kirim_awal").click().run()
    tanpa_error(at)
    assert A.rating_peserta(ui, "Andi") == 5 and len(A.daftar_ulasan(ui)) == 1
    assert [r for r in at.radio if r.key == "ub_bintang_awal"] == []
    sesi_baru = jalan(True, peran="user", nama_login="Andi", tab="profil")
    tanpa_error(sesi_baru)
    assert [r for r in sesi_baru.radio if r.key == "ub_bintang_awal"] == []        # tidak diminta lagi
    assert [r for r in sesi_baru.radio if r.key == "ub_bintang_profil"] == []      # rating tidak bisa diulang
    sesi_baru.text_area(key="ub_isi_profil").set_value("Rekap bulanan sebaiknya bisa diunduh")
    sesi_baru.button(key="ub_kirim_profil").click().run()
    tanpa_error(sesi_baru)
    assert len(A.daftar_ulasan(ui)) == 2


def test_peserta_baru_tidak_diarahkan_tapi_bisa_menilai_dari_profil(ui):
    absen_lama(ui, hari=2)
    at = jalan(True, peran="user", nama_login="Andi", tab="profil")
    tanpa_error(at)
    assert [r for r in at.radio if r.key == "ub_bintang_awal"] == []
    at.radio(key="ub_bintang_profil").set_value(4)
    at.button(key="ub_kirim_profil").click().run()
    tanpa_error(at)
    assert A.rating_peserta(ui, "Andi") == 4


def test_halaman_admin_umpan_balik_tampil(ui):
    A.simpan_rating(ui, "Andi", 5, ULASAN, "absen")
    tanpa_error(jalan(False, peran="admin", admin_ok=True, menu_admin="Umpan balik"))


def test_profil_menampilkan_ulasan_saya_dengan_status_dan_tanggapan_admin(ui):
    A.simpan_rating(ui, "Andi", 5, ULASAN, "absen")
    uid = int(A.daftar_ulasan(ui)["id"].iloc[0])
    A.ubah_status_ulasan(ui, uid, "diproses", "Sedang dibuat lebih besar")
    at = jalan(True, peran="user", nama_login="Andi", tab="profil")
    tanpa_error(at)
    assert [e.label for e in at.expander if e.label.startswith("Ulasan saya")] == ["Ulasan saya (1)"]
    assert any("Sedang diproses" in m.value for m in at.markdown)
    assert any("Tanggapan admin: Sedang dibuat lebih besar" in c.value for c in at.caption)


def test_profil_tanpa_ulasan_tidak_menampilkan_daftar(ui):
    at = jalan(True, peran="user", nama_login="Andi", tab="profil")
    tanpa_error(at)
    assert [e for e in at.expander if e.label.startswith("Ulasan saya")] == []
