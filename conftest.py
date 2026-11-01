"""Fixture bersama untuk tes yang memakai database SQLite sementara."""
import copy

import pandas as pd
import pytest
from sqlalchemy import create_engine

import app as A
import ml_absensi as ml
import test_ml as T

_CACHE = {}


@pytest.fixture(autouse=True)
def _memo_latih(request, monkeypatch):
    """Pelatihan deterministik (seed tetap), jadi hasilnya aman di-cache antar tes. test_ml TIDAK memakai ini
    karena ia menguji pelatihan itu sendiri."""
    if request.module.__name__ == "test_ml":
        return
    asli = ml.latih_terlambat

    def memo(dataset, *a, **k):
        kunci = (int(pd.util.hash_pandas_object(dataset, index=True).sum()), tuple(dataset.columns), a, tuple(sorted(k.items())))
        if kunci not in _CACHE:
            _CACHE[kunci] = asli(dataset, *a, **k)
        return copy.deepcopy(_CACHE[kunci])

    monkeypatch.setattr(ml, "latih_terlambat", memo)


def buat_engine(tmp_path, nama="t.db"):
    eng = create_engine(f"sqlite:///{tmp_path / nama}")
    A.metadata.create_all(eng)
    A._migrasi(eng)
    for i, n in enumerate(T.NAMA):
        A.add_peserta(eng, n, "A" if i % 2 == 0 else "B", "Divisi Umum", "1234")
    return eng


def isi_masuk(eng, df):
    for r in df.itertuples():
        A.simpan_absen(eng, nama=r.nama, tanggal=r.tanggal, jam=r.jam, jenis="masuk", status="berhasil", uang_makan=0)


def isi_lokasi(eng, df):
    for r in df.itertuples():
        A.simpan_absen(eng, nama=r.nama, tanggal=r.tanggal, jam=r.jam, jenis=r.jenis, status=r.status, uang_makan=0,
                       lat=r.lat, lon=r.lon, akurasi=r.akurasi, jarak=r.jarak)


@pytest.fixture
def eng_kosong(tmp_path):
    return buat_engine(tmp_path)


@pytest.fixture
def eng_masuk(tmp_path):
    eng = buat_engine(tmp_path)
    isi_masuk(eng, T.data_masuk(40))
    return eng
