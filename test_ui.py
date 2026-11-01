"""Tes alur antarmuka dengan AppTest: privasi, logbook, login pembimbing, panel admin Proyek.
Jalankan: pytest -q test_ui.py"""
from datetime import datetime, timedelta
from pathlib import Path

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

import app as A

APP = str(Path(__file__).parent / "app.py")


def kerja_terakhir():
    d = datetime.now(A.WITA).date()
    while A.hari_libur(d):
        d -= timedelta(days=1)
    return d


@pytest.fixture
def ui(tmp_path, monkeypatch):
    """Database SQLite baru di folder sementara; cache engine Streamlit dikosongkan agar tidak terbawa antar tes."""
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("PRIVASI_WAJIB", raising=False)
    st.cache_resource.clear()
    eng = A.get_engine()
    A.add_peserta(eng, "Andi", "A", "Divisi Umum", "1234")
    A.add_peserta(eng, "Sari", "B", "Divisi Syariah", "1234")
    A.add_pembimbing(eng, "Pak Budi", "Divisi Umum", "123456")
    A.add_pembimbing(eng, "Bu Rina", "Divisi Syariah", "654321")
    yield eng
    st.cache_resource.clear()


def jalan(**state):
    at = AppTest.from_file(APP, default_timeout=90)
    for k, v in state.items():
        at.session_state[k] = v
    return at.run()


def tombol(at, label):
    return [b for b in at.button if b.label == label][0]


def tanpa_error(at):
    assert not at.exception, [e.value[:300] for e in at.exception]


# ---------- peserta ----------
def test_layar_privasi_menghalangi_sampai_disetujui(ui):
    at = jalan(peran="user", nama_login="Andi")
    tanpa_error(at)
    setuju = [b for b in at.button if b.label.startswith("Setuju")][0]
    assert setuju.disabled and not any("Logbook kegiatan" in m.value for m in at.markdown)       # beranda belum terbuka
    assert A.perlu_persetujuan(A.get_peserta(ui, "Andi"))
    at.checkbox(key="privasi_cek").check().run()
    [b for b in at.button if b.label.startswith("Setuju")][0].click().run()
    tanpa_error(at)
    assert A.get_peserta(ui, "Andi")["setuju_versi"] == A.VERSI_PRIVASI
    assert any("Logbook kegiatan" in m.value for m in at.markdown)                                # beranda terbuka
    assert list(A.daftar_audit(ui, aksi="privasi.setuju")["aktor"]) == ["Andi"]


def test_peserta_tulis_logbook_validasi_dan_daftar(ui):
    A.catat_persetujuan(ui, "Andi")
    at = jalan(peran="user", nama_login="Andi", tab="logbook")
    tanpa_error(at)
    at.text_area(key="lb_uraian").set_value("pendek")
    at.button(key="lb_kirim").click().run()
    assert any("minimal 15" in e.value for e in at.error) and A.daftar_logbook(ui).empty        # ditolak, tidak tersimpan
    assert at.text_area(key="lb_uraian").value == "pendek"                                       # ketikan peserta TIDAK hilang
    at.text_area(key="lb_uraian").set_value("Menyusun rekap data nasabah dan memeriksa angkanya")
    at.number_input(key="lb_durasi").set_value(5.0)
    at.button(key="lb_kirim").click().run()
    tanpa_error(at)
    d = A.daftar_logbook(ui).iloc[0]
    assert d["nama"] == "Andi" and d["status"] == "diajukan" and d["durasi_jam"] == 5 and d["divisi"] == "Divisi Umum"
    assert at.text_area(key="lb_uraian").value == "" and any("diajukan" in s.value for s in at.success)   # dikosongkan setelah berhasil
    for sub in ("Daftar", "Progres"):
        at.radio(key="lb_sub").set_value(sub).run()
        tanpa_error(at)
    at.radio(key="lb_sub").set_value("Daftar").run()
    assert [(m.label, m.value) for m in at.metric][0] == ("Menunggu", "1")


def test_peserta_semua_tab_lama_tetap_jalan(ui):
    A.catat_persetujuan(ui, "Andi")
    for tab in ("beranda", "riwayat", "rekap", "profil", "logbook"):
        tanpa_error(jalan(peran="user", nama_login="Andi", tab=tab))


def test_peserta_tidak_bisa_memperbaiki_logbook_orang_lain_lewat_tampilan(ui):
    A.catat_persetujuan(ui, "Andi")
    lid = A.tambah_logbook(ui, "Sari", kerja_terakhir(), "Menyusun rekap data nasabah dan memeriksa angkanya", 2)
    at = jalan(peran="user", nama_login="Andi", tab="logbook")
    at.radio(key="lb_sub").set_value("Daftar").run()
    assert A.daftar_logbook(ui, nama="Andi").empty and "Sari" not in " ".join(m.value for m in at.markdown)
    with pytest.raises(ValueError):
        A.hapus_logbook(ui, lid, "Andi")


# ---------- pembimbing ----------
def test_login_pembimbing_pin_salah_dan_benar(ui):
    at = jalan()
    at.radio(key="login_peran").set_value("Pembimbing").run()
    at.selectbox[0].set_value("Pak Budi")                                    # bawaan urut abjad = "Bu Rina"
    at.text_input[0].set_value("000000")
    tombol(at, "Masuk").click().run()
    assert any("PIN salah" in e.value for e in at.error) and "peran" not in at.session_state
    at.text_input[0].set_value("123456")
    tombol(at, "Masuk").click().run()
    tanpa_error(at)
    assert at.session_state["peran"] == "pembimbing" and at.session_state["nama_login"] == "Pak Budi"
    assert any("Panel Pembimbing" in m.value for m in at.markdown)


def test_pembimbing_hanya_melihat_dan_meninjau_divisinya(ui):
    uraian = "Menyusun rekap data nasabah dan memeriksa angkanya"
    l_andi = A.tambah_logbook(ui, "Andi", kerja_terakhir(), uraian, 3)
    l_sari = A.tambah_logbook(ui, "Sari", kerja_terakhir(), uraian, 3)
    at = jalan(peran="pembimbing", nama_login="Pak Budi")
    tanpa_error(at)
    teks = " ".join(m.value for m in at.markdown)
    assert "Andi" in teks and "Sari" not in teks                                                # divisi lain tidak tampil
    assert not [b for b in at.button if b.key == f"tj_disetujui_{l_sari}"]
    at.button(key=f"tj_disetujui_{l_andi}").click().run()
    tanpa_error(at)
    r = A.daftar_logbook(ui).set_index("id")
    assert r.loc[l_andi, "status"] == "disetujui" and r.loc[l_andi, "ditinjau_oleh"] == "Pak Budi"
    assert r.loc[l_sari, "status"] == "diajukan"                                                # logbook divisi lain tidak tersentuh


def test_pembimbing_menolak_wajib_alasan(ui):
    lid = A.tambah_logbook(ui, "Andi", kerja_terakhir(), "Menyusun rekap data nasabah dan memeriksa angkanya", 3)
    at = jalan(peran="pembimbing", nama_login="Pak Budi")
    at.button(key=f"tj_ditolak_{lid}").click().run()
    assert any("Alasan penolakan" in e.value for e in at.error) and A.daftar_logbook(ui).iloc[0]["status"] == "diajukan"
    at.text_input(key=f"tj_cat_{lid}").set_value("tambahkan hasil yang dicapai")
    at.button(key=f"tj_ditolak_{lid}").click().run()
    d = A.daftar_logbook(ui).iloc[0]
    assert d["status"] == "ditolak" and d["catatan_peninjau"] == "tambahkan hasil yang dicapai"


def test_pembimbing_menambah_paket_untuk_divisinya_dan_halaman_paket_tampil(ui):
    at = jalan(peran="pembimbing", nama_login="Pak Budi", menu_pembimbing="Paket kerja")
    tanpa_error(at)
    hari = kerja_terakhir()
    A.tambah_paket(ui, "Divisi Umum", "1", "Analisis kebutuhan", hari - timedelta(days=7), hari + timedelta(days=7), 40,
                   "Pak Budi", "pembimbing", "Divisi Umum")
    tanpa_error(jalan(peran="pembimbing", nama_login="Pak Budi", menu_pembimbing="Paket kerja"))   # dengan Gantt
    assert list(A.daftar_paket(ui)["divisi"]) == ["Divisi Umum"]


def test_pembimbing_yang_dihapus_saat_sesi_terbuka_tidak_error(ui):
    A.hapus_pembimbing(ui, "Pak Budi")
    at = jalan(peran="pembimbing", nama_login="Pak Budi")
    tanpa_error(at)
    assert any("tidak ditemukan" in e.value for e in at.error)


def test_pembimbing_tidak_bisa_membuka_panel_admin_lewat_sesi(ui):
    at = jalan(peran="pembimbing", nama_login="Pak Budi", admin_ok=True, menu_admin="Proyek")
    assert any("Panel Pembimbing" in m.value for m in at.markdown) and not any("Panel Admin" in m.value for m in at.markdown)


# ---------- admin ----------
@pytest.mark.parametrize("sub", ["Logbook", "Paket", "Anggaran", "Pembimbing", "Audit"])
def test_admin_proyek_semua_bagian_tampil(ui, sub):
    tanpa_error(jalan(peran="admin", admin_ok=True, menu_admin="Proyek", menu_proyek=sub))


@pytest.mark.parametrize("menu", ["Peserta", "Rekap", "Analitik", "Model", "Umpan balik", "Foto"])
def test_admin_halaman_lama_tidak_regresi(ui, menu):
    tanpa_error(jalan(peran="admin", admin_ok=True, menu_admin=menu))


def test_admin_audit_verifikasi_ok_lalu_rusak(ui):
    at = jalan(peran="admin", admin_ok=True, menu_admin="Proyek", menu_proyek="Audit")
    at.button(key="au_verif").click().run()
    assert any("Rantai utuh" in s.value for s in at.success)
    with ui.begin() as c:
        c.execute(A.text("UPDATE audit_log SET objek = 'palsu' WHERE id = (SELECT MIN(id) FROM audit_log)"))
    at.button(key="au_verif").click().run()
    assert any("RANTAI RUSAK" in e.value for e in at.error)


def test_admin_tambah_pembimbing_lewat_form_dan_tercatat(ui):
    at = jalan(peran="admin", admin_ok=True, menu_admin="Proyek", menu_proyek="Pembimbing")
    at.text_input(key="pb_nama").set_value("Pak Anto")
    at.selectbox(key="pb_divisi").set_value("Divisi Kepatuhan")
    at.text_input(key="pb_pin").set_value("12")
    at.button(key="pb_tambah").click().run()
    assert any("6-8 angka" in e.value for e in at.error) and A.get_pembimbing(ui, "Pak Anto") is None
    assert at.text_input(key="pb_nama").value == "Pak Anto"                          # nama tidak hilang karena PIN salah
    at.text_input(key="pb_pin").set_value("246810")
    at.button(key="pb_tambah").click().run()
    assert A.get_pembimbing(ui, "Pak Anto")["divisi"] == "Divisi Kepatuhan"
    assert "pembimbing.tambah" in list(A.daftar_audit(ui, 50)["aksi"])
    assert at.text_input(key="pb_nama").value == ""                                   # dikosongkan setelah berhasil
