"""Tes API REST. Jalankan: pytest -q test_api.py"""
import pytest
from fastapi.testclient import TestClient

import api
import app as A
import test_ml as T
from conftest import buat_engine, isi_lokasi, isi_masuk

KUNCI = {"X-API-Key": "rahasia-uji-123"}


@pytest.fixture(autouse=True)
def _kunci(monkeypatch):
    monkeypatch.setenv("API_KEY", "rahasia-uji-123")


@pytest.fixture
def klien(eng_masuk):
    api.atur_engine(eng_masuk)
    yield TestClient(api.api)
    api.atur_engine(None)


@pytest.fixture
def klien_kosong(eng_kosong):
    api.atur_engine(eng_kosong)
    yield TestClient(api.api)
    api.atur_engine(None)


# ---------- keamanan ----------
def test_health_terbuka_tanpa_key(klien):
    r = klien.get("/health")
    assert r.status_code == 200 and r.json()["db"] is True and r.json()["model_aktif"] is False


@pytest.mark.parametrize("metode,url", [("get", "/v1/model"), ("get", "/v1/anomali"),
                                         ("post", "/v1/prediksi/terlambat"), ("post", "/v1/model/latih")])
def test_endpoint_terlindungi_menolak_tanpa_atau_salah_key(klien, metode, url):
    kirim = getattr(klien, metode)
    kw = {"json": {"nama": "P0"}} if url.endswith("terlambat") else {}
    assert kirim(url, **kw).status_code == 401
    assert kirim(url, headers={"X-API-Key": "salah"}, **kw).status_code == 401
    assert kirim(url, headers={"X-API-Key": "ключ".encode("utf-8")}, **kw).status_code == 401   # byte non-ASCII tidak membuat server error


def test_tanpa_api_key_terkonfigurasi_semua_ditolak_503(klien, monkeypatch):
    monkeypatch.delenv("API_KEY")
    r = klien.get("/v1/model", headers=KUNCI)
    assert r.status_code == 503 and "API_KEY" in r.json()["detail"]


# ---------- model ----------
def test_model_belum_ada_404_lalu_latih_dan_info(klien):
    assert klien.get("/v1/model", headers=KUNCI).status_code == 404
    r = klien.post("/v1/model/latih", headers=KUNCI, json={})
    assert r.status_code == 200 and r.json()["aktif"] is True
    assert r.json()["metrik"]["uji"]["f1"] > 0.5 and r.json()["metrik"]["lebih_baik_dari_baseline"] is True
    info = klien.get("/v1/model", headers=KUNCI).json()
    assert info["id"] == r.json()["id"] and info["n_data"] > 100


def test_latih_data_kurang_422_dan_kebijakan_tidak_valid_422(klien_kosong):
    r = klien_kosong.post("/v1/model/latih", headers=KUNCI, json={})
    assert r.status_code == 422 and "terlalu sedikit" in r.json()["detail"]
    assert klien_kosong.post("/v1/model/latih", headers=KUNCI, json={"aktifkan": "ngawur"}).status_code == 422


# ---------- prediksi ----------
def test_prediksi_tanpa_model_409(klien):
    r = klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "P0", "tanggal": "2026-10-05"})
    assert r.status_code == 409 and "Belum ada model" in r.json()["detail"]


def test_prediksi_normal_dan_urutan_risiko(klien):
    klien.post("/v1/model/latih", headers=KUNCI, json={})
    p7 = klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "P7", "tanggal": "2026-10-05"}).json()
    p0 = klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "P0", "tanggal": "2026-10-05"}).json()
    assert set(p7) == {"nama", "tanggal", "probabilitas", "risiko", "faktor", "model_id"}
    assert 0 <= p0["probabilitas"] <= 1 and p7["probabilitas"] > p0["probabilitas"]
    assert isinstance(p7["faktor"], list) and p7["tanggal"] == "2026-10-05"


def test_prediksi_peserta_tak_dikenal_404_hari_libur_422_tanggal_buruk_422(klien):
    klien.post("/v1/model/latih", headers=KUNCI, json={})
    assert klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "Hantu", "tanggal": "2026-10-05"}).status_code == 404
    r = klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "P0", "tanggal": "2026-10-10"})   # Sabtu
    assert r.status_code == 422
    assert klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "P0", "tanggal": "bukan-tanggal"}).status_code == 422
    assert klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": ""}).status_code == 422


def test_prediksi_tanpa_tanggal_memakai_hari_kerja_berikutnya(klien):
    klien.post("/v1/model/latih", headers=KUNCI, json={})
    r = klien.post("/v1/prediksi/terlambat", headers=KUNCI, json={"nama": "P1"})
    assert r.status_code == 200
    from datetime import date
    tgl = date.fromisoformat(r.json()["tanggal"])
    assert not A.hari_libur(tgl) and tgl >= A.datetime.now(A.WITA).date()


# ---------- anomali ----------
def test_anomali_data_kurang_bukan_error(klien):
    r = klien.get("/v1/anomali", headers=KUNCI)
    assert r.status_code == 200 and r.json()["items"] == [] and "terlalu sedikit" in r.json()["catatan"]


def test_anomali_menandai_dan_membatasi(tmp_path):
    eng = buat_engine(tmp_path, "lokasi.db")
    df = T.data_lokasi()
    isi_lokasi(eng, df)
    api.atur_engine(eng)
    try:
        k = TestClient(api.api)
        r = k.get("/v1/anomali", headers=KUNCI, params={"awal": "2026-08-01", "akhir": "2026-12-31", "persen": 5}).json()
        assert r["n_dianalisis"] == len(df) and 0 < r["n_ditandai"] <= len(df) * 0.1
        assert r["items"] and set(r["items"][0]) == {"nama", "tanggal", "jam", "jenis", "status", "jarak", "akurasi", "skor", "alasan"}
        skor = [i["skor"] for i in r["items"]]
        assert skor == sorted(skor, reverse=True)                                # paling mencurigakan di atas
        kecil = k.get("/v1/anomali", headers=KUNCI, params={"awal": "2026-08-01", "akhir": "2026-12-31", "batas": 3}).json()
        assert len(kecil["items"]) == 3
        assert k.get("/v1/anomali", headers=KUNCI, params={"persen": 0}).status_code == 422
        assert k.get("/v1/anomali", headers=KUNCI, params={"persen": 50}).status_code == 422
        assert k.get("/v1/anomali", headers=KUNCI, params={"awal": "2026-12-31", "akhir": "2026-01-01"}).status_code == 422
        assert k.get("/v1/anomali", headers=KUNCI, params={"awal": "2020-01-01", "akhir": "2026-12-31"}).status_code == 422
    finally:
        api.atur_engine(None)


def test_health_503_bila_database_mati(klien, monkeypatch):
    from sqlalchemy.exc import OperationalError

    class Mati:
        def connect(self):
            raise OperationalError("SELECT 1", {}, Exception("db mati"))
    api.atur_engine(Mati())
    r = klien.get("/health")
    assert r.status_code == 503 and r.json()["db"] is False


# ---------- audit ----------
def test_latih_via_api_tercatat_sebagai_aktor_api(klien):
    klien.post("/v1/model/latih", headers=KUNCI, json={})
    j = A.daftar_audit(api.eng(), aksi="model.simpan").iloc[0]
    assert j["aktor"] == "api" and j["peran"] == "sistem"


def test_verifikasi_audit_ok_lalu_409_bila_dirusak(klien):
    assert klien.get("/v1/audit/verifikasi").status_code == 401
    klien.post("/v1/model/latih", headers=KUNCI, json={})
    r = klien.get("/v1/audit/verifikasi", headers=KUNCI)
    assert r.status_code == 200 and r.json()["ok"] is True and len(r.json()["hash_terakhir"]) == 64
    with api.eng().begin() as c:
        c.execute(A.text("UPDATE audit_log SET aktor = 'samar' WHERE id = (SELECT MIN(id) FROM audit_log)"))
    r = klien.get("/v1/audit/verifikasi", headers=KUNCI)
    assert r.status_code == 409 and r.json()["ok"] is False and r.json()["rusak_id"] and "diubah" in r.json()["alasan"]


def test_jangkar_mendeteksi_pemotongan_ekor_yang_lolos_dari_rantai(klien):
    klien.post("/v1/model/latih", headers=KUNCI, json={})
    jangkar = klien.get("/v1/audit/verifikasi", headers=KUNCI).json()["hash_terakhir"]
    r = klien.get("/v1/audit/verifikasi", headers=KUNCI, params={"hash_sebelumnya": jangkar})
    assert r.status_code == 200 and r.json()["hash_sebelumnya_ada"] is True
    with api.eng().begin() as c:                                          # penyerang menghapus catatan TERAKHIR
        c.execute(A.text("DELETE FROM audit_log WHERE id = (SELECT MAX(id) FROM audit_log)"))
    tanpa_jangkar = klien.get("/v1/audit/verifikasi", headers=KUNCI)
    assert tanpa_jangkar.status_code == 200                               # rantai sisanya tetap konsisten: ini batas yang diketahui
    dengan_jangkar = klien.get("/v1/audit/verifikasi", headers=KUNCI, params={"hash_sebelumnya": jangkar})
    assert dengan_jangkar.status_code == 409 and dengan_jangkar.json()["hash_sebelumnya_ada"] is False
    assert "dipotong" in dengan_jangkar.json()["alasan"]


@pytest.mark.parametrize("nilai", ["abc", "G" * 64, "a" * 63, "a" * 65])
def test_jangkar_format_salah_ditolak_422(klien, nilai):
    assert klien.get("/v1/audit/verifikasi", headers=KUNCI, params={"hash_sebelumnya": nilai}).status_code == 422
