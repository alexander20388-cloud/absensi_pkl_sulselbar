"""Tes API aplikasi native peserta. Jalankan: pytest -q test_api_peserta.py"""
import pytest
from fastapi.testclient import TestClient

import api
import app as A
from conftest import buat_engine
import test_ml as T

PIN = "1234"


@pytest.fixture(autouse=True)
def _rahasia(monkeypatch):
    monkeypatch.setenv("TOKEN_SECRET", "rahasia-token-uji-yang-panjang")
    monkeypatch.setenv("PRIVASI_WAJIB", "1")


@pytest.fixture
def klien(tmp_path):
    api.atur_engine(buat_engine(tmp_path))
    yield TestClient(api.api)
    api.atur_engine(None)


def _masuk(klien, nama=None, pin=PIN):
    nama = nama or T.NAMA[0]
    return klien.post("/v1/peserta/login", json={"nama": nama, "pin": pin})


def _auth(klien, setuju=True):
    r = _masuk(klien)
    assert r.status_code == 200, r.text
    h = {"Authorization": "Bearer " + r.json()["token"]}
    if setuju:
        assert klien.post("/v1/peserta/persetujuan", headers=h).status_code == 200
    return h


def test_daftar_nama_terbuka(klien):
    r = klien.get("/v1/peserta/daftar-nama")
    assert r.status_code == 200 and set(T.NAMA) <= set(r.json()["nama"])


def test_login_benar_memberi_token_dan_profil(klien):
    r = _masuk(klien)
    j = r.json()
    assert r.status_code == 200 and j["token"] and j["peserta"]["nama"] == T.NAMA[0]
    assert j["peserta"]["perlu_persetujuan"] is True and "lokasi_gedung" in j["peserta"]


def test_login_pin_salah_401_lalu_terkunci_429(klien):
    for _ in range(A.MAKS_GAGAL_LOGIN - 1):
        assert _masuk(klien, pin="0000").status_code == 401
    assert _masuk(klien, pin="0000").status_code == 429
    assert _masuk(klien, pin=PIN).status_code == 429            # PIN benar pun ditolak selama dikunci


def test_nama_tidak_dikenal_sama_dengan_pin_salah(klien):
    assert _masuk(klien, nama="Tidak Ada", pin=PIN).status_code == 401


def test_tanpa_token_atau_token_palsu_401(klien):
    assert klien.get("/v1/peserta/hari-ini").status_code == 401
    assert klien.get("/v1/peserta/hari-ini", headers={"Authorization": "Bearer abc.def"}).status_code == 401


def test_token_dicabut_saat_pin_direset(klien):
    h = _auth(klien)
    assert klien.get("/v1/peserta/saya", headers=h).status_code == 200
    A.reset_pin(api.eng(), T.NAMA[0], "9999")
    assert klien.get("/v1/peserta/saya", headers=h).status_code == 401


def test_belum_setuju_privasi_ditolak_403_lalu_boleh(klien):
    h = _auth(klien, setuju=False)
    r = klien.get("/v1/peserta/riwayat", headers=h)
    assert r.status_code == 200                                    # baca riwayat tidak butuh persetujuan
    r = klien.post("/v1/peserta/tidak-hadir", headers=h, json={"jenis": "izin", "tanggal": "2030-01-02", "alasan": "ke kampus"})
    assert r.status_code == 403 and r.json()["detail"]["kode"] == "persetujuan"
    assert klien.post("/v1/peserta/persetujuan", headers=h).status_code == 200
    assert klien.get("/v1/peserta/saya", headers=h).json()["perlu_persetujuan"] is False


def test_hari_ini_dan_rekap_kosong(klien):
    h = _auth(klien)
    j = klien.get("/v1/peserta/hari-ini", headers=h).json()
    assert j["masuk"] is None and j["pulang"] is None and j["absen_berikutnya"] == "masuk" and "waktu_server" in j
    r = klien.get("/v1/peserta/rekap", headers=h).json()
    assert r["hari_hadir"] == 0 and r["total_uang_makan"] == 0


def test_absen_meneruskan_data_ke_proses_absen_dan_jam_dari_server(klien, monkeypatch):
    h = _auth(klien)
    tangkap = {}

    def palsu(eng, nama, peserta, jenis, foto, lat, lon, akurasi=None, now=None):
        tangkap.update(nama=nama, jenis=jenis, ukuran=len(foto), lat=lat, lon=lon, akurasi=akurasi, now=now)
        return {"level": "success", "pesan": "ok", "ok": True}

    monkeypatch.setattr(A, "proses_absen", palsu)
    r = klien.post("/v1/peserta/absen", headers=h, data={"jenis": "masuk", "lat": "-5.15", "lon": "119.41", "akurasi": "12"},
                   files={"foto": ("f.jpg", b"123456", "image/jpeg")})
    assert r.status_code == 200 and r.json()["ok"] is True
    assert tangkap == {"nama": T.NAMA[0], "jenis": "masuk", "ukuran": 6, "lat": -5.15, "lon": 119.41, "akurasi": 12.0, "now": None}


def test_absen_validasi_input(klien):
    h = _auth(klien)
    f = {"foto": ("f.jpg", b"1", "image/jpeg")}
    assert klien.post("/v1/peserta/absen", headers=h, data={"jenis": "lembur", "lat": "0", "lon": "0"}, files=f).status_code == 422
    assert klien.post("/v1/peserta/absen", headers=h, data={"jenis": "masuk", "lat": "999", "lon": "0"}, files=f).status_code == 422
    assert klien.post("/v1/peserta/absen", headers=h, data={"jenis": "masuk", "lat": "0", "lon": "0"},
                      files={"foto": ("f.jpg", b"", "image/jpeg")}).status_code == 422


def test_logbook_tambah_daftar_ubah_hapus(klien):
    h = _auth(klien)
    tgl = A._hari_kerja_sebelum(A.datetime.now(A.WITA).date()).isoformat()
    r = klien.post("/v1/peserta/logbook", headers=h, json={"tanggal": tgl, "uraian": "Menyusun rekap data nasabah", "durasi_jam": 4})
    assert r.status_code == 201, r.text
    lid = r.json()["id"]
    j = klien.get("/v1/peserta/logbook", headers=h).json()
    assert j["menunggu"] == 1 and j["logbook"][0]["id"] == lid
    assert klien.put(f"/v1/peserta/logbook/{lid}", headers=h, json={"uraian": "Revisi rekap data", "durasi_jam": 3}).status_code == 200
    assert klien.delete(f"/v1/peserta/logbook/{lid}", headers=h).status_code == 200
    assert klien.get("/v1/peserta/logbook", headers=h).json()["logbook"] == []


def test_logbook_peserta_lain_tidak_bisa_diubah(klien):
    h1 = _auth(klien)
    tgl = A._hari_kerja_sebelum(A.datetime.now(A.WITA).date()).isoformat()
    lid = klien.post("/v1/peserta/logbook", headers=h1, json={"tanggal": tgl, "uraian": "Kegiatan milik peserta pertama", "durasi_jam": 2}).json()["id"]
    r2 = klien.post("/v1/peserta/login", json={"nama": T.NAMA[1], "pin": PIN}).json()["token"]
    h2 = {"Authorization": "Bearer " + r2}
    klien.post("/v1/peserta/persetujuan", headers=h2)
    assert klien.delete(f"/v1/peserta/logbook/{lid}", headers=h2).status_code == 422
    assert klien.put(f"/v1/peserta/logbook/{lid}", headers=h2, json={"uraian": "dibajak", "durasi_jam": 1}).status_code == 422


def test_api_tanpa_token_secret_menolak_503(klien, monkeypatch):
    monkeypatch.delenv("TOKEN_SECRET")
    assert _masuk(klien).status_code == 503
