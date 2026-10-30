"""API REST Absensi PKL.

Memakai database, aturan libur, dan model yang SAMA dengan aplikasi Streamlit (diimpor dari app.py dan
ml_absensi.py), jadi hasilnya selalu konsisten dengan yang dilihat admin. Cocok untuk aplikasi Android/native,
dashboard lain, atau penjadwal (cron/GitHub Actions) yang memicu pelatihan ulang.

Jalankan  : uvicorn api:api --host 0.0.0.0 --port 8000
Dokumentasi interaktif: http://localhost:8000/docs
Konfigurasi: koneksi database sama seperti app.py (secrets/env). Tambahkan secret/env `api_key` (API_KEY)
berisi string acak panjang. Semua endpoint selain /health wajib header  X-API-Key: <api_key>.
Tanpa api_key diatur, endpoint terlindungi menolak semua permintaan (aman secara default).
"""
from __future__ import annotations

import hmac
from datetime import datetime, timedelta
from datetime import date as Date
from typing import Literal, Optional

import pandas as pd
from fastapi import Depends, FastAPI, Header, HTTPException, Query
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from sqlalchemy import text
from sqlalchemy.exc import SQLAlchemyError

import app as absensi
import ml_absensi as ml

api = FastAPI(
    title="API Absensi PKL", version="1.0.0",
    description="Prediksi risiko terlambat, deteksi absen mencurigakan, dan pengelolaan versi model.",
)

_eng = None


def atur_engine(e):
    """Pakai engine tertentu (untuk tes). Secara normal engine dibuat sekali dari app.get_engine()."""
    global _eng
    _eng = e


def eng():
    global _eng
    if _eng is None:
        _eng = absensi.get_engine()
    return _eng


def wajib_kunci(x_api_key: Optional[str] = Header(default=None)):
    kunci = absensi.secret("api_key")
    if not kunci:
        raise HTTPException(503, "API belum diaktifkan: atur secret/env API_KEY terlebih dulu.")
    if not x_api_key or not hmac.compare_digest(x_api_key.encode(), str(kunci).encode()):
        raise HTTPException(401, "API key salah atau tidak ada.")


def _py(v):
    """Nilai pandas/numpy -> tipe Python yang aman untuk JSON (NaN -> None)."""
    if v is None or (not isinstance(v, (list, dict, str)) and pd.isna(v)):
        return None
    return v.item() if hasattr(v, "item") else v


def _ringkas_metrik(metrik: dict) -> dict:
    return {k: metrik.get(k) for k in ("model", "parameter", "cv_f1", "uji", "baseline", "lebih_baik_dari_baseline",
                                        "n_data", "persen_terlambat", "tanggal_awal", "tanggal_akhir")}


class PrediksiIn(BaseModel):
    nama: str = Field(min_length=1, max_length=100, description="Nama peserta persis seperti terdaftar")
    tanggal: Optional[Date] = Field(default=None, description="YYYY-MM-DD; kosong = hari kerja berikutnya (WITA)")


class LatihIn(BaseModel):
    aktifkan: Literal["jika_lebih_baik", "selalu", "tidak"] = "jika_lebih_baik"


@api.get("/health", tags=["Status"])
def health():
    """Cek hidup (tanpa API key): koneksi database dan keberadaan model aktif. 503 kalau database bermasalah."""
    try:
        with eng().connect() as c:
            c.execute(text("SELECT 1"))
        model_ada = absensi.ambil_model(eng()) is not None
    except SQLAlchemyError:
        return JSONResponse({"status": "error", "db": False, "model_aktif": False}, status_code=503)
    return {"status": "ok", "db": True, "model_aktif": model_ada, "waktu_wita": datetime.now(absensi.WITA).isoformat(timespec="seconds")}


@api.get("/v1/model", tags=["Model"], dependencies=[Depends(wajib_kunci)])
def info_model():
    """Informasi dan metrik evaluasi model prediksi yang sedang aktif."""
    m = absensi.ambil_model(eng())
    if m is None:
        raise HTTPException(404, "Belum ada model aktif.")
    return {"id": m["id"], "dilatih_wita": m["dilatih"], "n_data": m["n_data"], "metrik": _ringkas_metrik(m["metrik"])}


@api.post("/v1/prediksi/terlambat", tags=["Prediksi"], dependencies=[Depends(wajib_kunci)])
def prediksi_terlambat(body: PrediksiIn):
    """Probabilitas seorang peserta terlambat pada suatu hari kerja, dari riwayat kehadirannya sebelum hari itu."""
    tgl = body.tanggal
    if tgl is None:
        tgl = datetime.now(absensi.WITA).date()
        while absensi.hari_libur(tgl):
            tgl += timedelta(days=1)
    if absensi.hari_libur(tgl):
        raise HTTPException(422, absensi.pesan_libur(tgl))
    if absensi.get_peserta(eng(), body.nama) is None:
        raise HTTPException(404, f"Peserta '{body.nama}' tidak ditemukan.")
    try:
        df, m = absensi.prediksi_semua(eng(), tgl, hanya_nama=body.nama)
    except ml.DataKurang as e:
        raise HTTPException(409, str(e))
    r = df.iloc[0]
    return {
        "nama": r["nama"], "tanggal": tgl.isoformat(), "probabilitas": round(float(r["probabilitas"]), 4),
        "risiko": r["risiko"], "faktor": [f for f in str(r["faktor"]).split("; ") if f], "model_id": m["id"],
    }


@api.get("/v1/anomali", tags=["Deteksi"], dependencies=[Depends(wajib_kunci)])
def anomali(
    awal: Optional[Date] = Query(default=None, description="Default: 30 hari terakhir"),
    akhir: Optional[Date] = Query(default=None),
    persen: float = Query(default=5.0, ge=1, le=15, description="Persentase absen yang ditandai"),
    batas: int = Query(default=100, ge=1, le=1000, description="Maksimal item yang dikembalikan"),
):
    """Absen yang ditandai tidak lazim. Ini sinyal untuk ditinjau manusia, bukan bukti kecurangan."""
    akhir = akhir or datetime.now(absensi.WITA).date()
    awal = awal or (akhir - timedelta(days=29))
    if awal > akhir:
        raise HTTPException(422, "awal tidak boleh setelah akhir.")
    if (akhir - awal).days > 366:
        raise HTTPException(422, "Rentang maksimal 366 hari.")
    df = absensi.get_absensi_lengkap(eng(), awal, akhir)
    try:
        hasil = ml.deteksi_anomali(df, absensi.list_peserta(eng()), absensi.GEDUNG_COORD, persen=persen)
    except ml.DataKurang as e:
        return {"awal": awal.isoformat(), "akhir": akhir.isoformat(), "n_dianalisis": 0, "n_ditandai": 0,
                "items": [], "catatan": str(e)}
    tanda = hasil[hasil["ditandai"]]
    kolom = ["nama", "tanggal", "jam", "jenis", "status", "jarak", "akurasi", "skor", "alasan"]
    items = [{k: _py(v) for k, v in row.items()} for row in tanda.head(batas)[kolom].to_dict("records")]
    for it in items:
        for k in ("jarak", "akurasi", "skor"):
            if it[k] is not None:
                it[k] = round(float(it[k]), 1)
    return {"awal": awal.isoformat(), "akhir": akhir.isoformat(), "n_dianalisis": int(len(hasil)),
            "n_ditandai": int(len(tanda)), "items": items, "catatan": None}


@api.post("/v1/model/latih", tags=["Model"], dependencies=[Depends(wajib_kunci)])
def latih_model(body: LatihIn = LatihIn()):
    """Latih versi model baru dari seluruh data (bisa dipicu penjadwal). Butuh beberapa detik."""
    try:
        r = absensi.latih_dan_simpan(eng(), body.aktifkan)
    except ml.DataKurang as e:
        raise HTTPException(422, str(e))
    return {"id": r["id"], "aktif": r["aktif"], "metrik": _ringkas_metrik(r["metrik"])}
