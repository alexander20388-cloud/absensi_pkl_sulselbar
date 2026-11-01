"""Tes penyimpanan versi model, prediksi, dan monitoring di atas database. Jalankan: pytest -q test_model_db.py"""
import json
from datetime import date

import pytest
from sqlalchemy import select, update

import app as A
import ml_absensi as ml
import test_ml as T
from conftest import isi_masuk


def test_latih_pertama_langsung_aktif_dan_tersimpan(eng_masuk):
    r = A.latih_dan_simpan(eng_masuk)
    assert r["aktif"] is True and r["ada_model_sebelumnya"] is False
    versi = A.daftar_model(eng_masuk)
    assert len(versi) == 1 and versi[0]["aktif"] == 1 and versi[0]["metrik"]["uji"]["f1"] > 0.5
    m = A.ambil_model(eng_masuk)
    assert m["id"] == r["id"] and len(m["artefak"]) > 100 and "telat_7" in m["baseline"]
    assert ml.dari_bytes(m["artefak"])["fitur"] == ml.FITUR_TERLAMBAT


def test_data_kurang_tidak_menyimpan_apa_pun(eng_kosong):
    with pytest.raises(ml.DataKurang):
        A.latih_dan_simpan(eng_kosong)
    assert A.daftar_model(eng_kosong) == [] and A.ambil_model(eng_kosong) is None


def test_kebijakan_aktivasi(eng_masuk):
    A.latih_dan_simpan(eng_masuk)
    r2 = A.latih_dan_simpan(eng_masuk, "tidak")
    assert r2["aktif"] is False and A.ambil_model(eng_masuk)["id"] != r2["id"]
    r3 = A.latih_dan_simpan(eng_masuk, "selalu")
    assert r3["aktif"] is True and A.ambil_model(eng_masuk)["id"] == r3["id"]
    assert sum(v["aktif"] for v in A.daftar_model(eng_masuk)) == 1          # selalu tepat satu yang aktif
    with pytest.raises(ValueError):
        A.latih_dan_simpan(eng_masuk, "sembarang")


def test_model_yang_lebih_buruk_tidak_menggantikan_yang_sehat(eng_masuk):
    r1 = A.latih_dan_simpan(eng_masuk)
    with eng_masuk.begin() as c:                                             # anggap model aktif sangat bagus (F1 uji 0.99)
        m = A.ambil_model(eng_masuk)["metrik"]
        m["uji"]["f1"] = 0.99
        c.execute(update(A.model_versi_t).where(A.model_versi_t.c.id == r1["id"]).values(metrik=json.dumps(m)))
    r2 = A.latih_dan_simpan(eng_masuk, "jika_lebih_baik")
    assert r2["aktif"] is False and A.ambil_model(eng_masuk)["id"] == r1["id"]


def test_aktifkan_model_manual_dan_id_tak_dikenal(eng_masuk):
    r1 = A.latih_dan_simpan(eng_masuk)
    r2 = A.latih_dan_simpan(eng_masuk, "tidak")
    assert A.aktifkan_model(eng_masuk, r2["id"]) is True and A.ambil_model(eng_masuk)["id"] == r2["id"]
    assert A.aktifkan_model(eng_masuk, 9999) is False and A.ambil_model(eng_masuk)["id"] == r2["id"]
    assert A.ambil_model(eng_masuk, model_id=r1["id"])["id"] == r1["id"]


def test_prediksi_semua_peserta(eng_masuk):
    A.latih_dan_simpan(eng_masuk)
    df, m = A.prediksi_semua(eng_masuk, date(2026, 10, 5))                    # Senin
    assert len(df) == len(T.NAMA) and df["probabilitas"].between(0, 1).all()
    assert df["probabilitas"].is_monotonic_decreasing and set(df["risiko"]) <= {"Rendah", "Sedang", "Tinggi"}
    # peserta yang kebiasaannya paling telat (P7) harus lebih berisiko daripada yang paling disiplin (P0)
    p = df.set_index("nama")["probabilitas"]
    assert p["P7"] > p["P0"]
    satu, _ = A.prediksi_semua(eng_masuk, date(2026, 10, 5), hanya_nama="P3")
    assert list(satu["nama"]) == ["P3"]
    kosong, _ = A.prediksi_semua(eng_masuk, date(2026, 10, 5), hanya_nama="Tidak Ada")
    assert kosong.empty


def test_prediksi_tanpa_model_diberi_pesan_jelas(eng_masuk):
    with pytest.raises(ml.DataKurang, match="Belum ada model"):
        A.prediksi_semua(eng_masuk, date(2026, 10, 5))


def test_pantau_sehat_tepat_setelah_dilatih(eng_masuk):
    A.latih_dan_simpan(eng_masuk)
    r = A.ringkasan_pantau(eng_masuk)
    assert r["n_baru"] == 0 and r["eval_baru"] is None
    assert (r["drift"]["status"] == "Stabil").all()
    assert r["perlu_retrain"] is False and r["umur_hari"] == 0


def test_pantau_mendeteksi_perubahan_perilaku_setelah_deploy(eng_masuk):
    A.latih_dan_simpan(eng_masuk)
    # sesudah model dilatih, semua peserta jadi jauh lebih sering telat (+35 menit)
    isi_masuk(eng_masuk, T.data_masuk(20, seed=9, geser=35, mulai=date(2026, 10, 1)))
    r = A.ringkasan_pantau(eng_masuk)
    assert r["n_baru"] > 100
    d = r["drift"].set_index("fitur")
    assert d.loc["rata_menit_5", "status"] in ("Waspada", "Drift") and d.loc["rata_menit_5", "psi"] > 0.1
    assert r["perlu_retrain"] is True and r["alasan"]
    if r["eval_baru"] is not None:
        assert 0 <= r["eval_baru"]["f1"] <= 1


def test_pantau_tanpa_model(eng_masuk):
    assert A.ringkasan_pantau(eng_masuk) is None
