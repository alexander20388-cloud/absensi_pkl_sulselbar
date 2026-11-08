"""API untuk aplikasi native PESERTA (login, beranda, absen, riwayat, rekap, logbook, profil).

Prinsip: SEMUA aturan tetap di server dan memakai fungsi yang sama dengan website Streamlit (app.py):
proses_absen, proses_tidak_hadir, get_status_hari, tambah_logbook, dst. Aplikasi hanya mengirim PIN, foto, koordinat,
dan teks; jam, radius lokasi, hari libur, absen ganda, dan verifikasi wajah diputuskan server.

Autentikasi: login dengan nama + PIN menghasilkan token (lihat token_peserta.py), dikirim sebagai
`Authorization: Bearer <token>`. Butuh secret/env TOKEN_SECRET (acak, panjang). Pembatasan percobaan PIN
(5 gagal = kunci 15 menit) memakai mekanisme yang sama dengan website.
"""
from __future__ import annotations

from datetime import date as Date
from datetime import datetime
from typing import Literal, Optional

import pandas as pd
from fastapi import APIRouter, Depends, File, Form, Header, HTTPException, Query, UploadFile
from pydantic import BaseModel, Field
from sqlalchemy.exc import SQLAlchemyError

import app as absensi
import token_peserta as tk

router = APIRouter(prefix="/v1/peserta", tags=["peserta (aplikasi native)"])

MAKS_FOTO_BYTE = 8 * 1024 * 1024          # foto kamera HP; server mengecilkannya sendiri


# ---------------------------------------------------------------- bantuan
def _eng():
    import api                              # impor lambat: api.py yang memasang router ini
    return api.eng()


def _rahasia() -> str:
    r = absensi.secret("token_secret")
    if not r or len(str(r)) < 16:
        raise HTTPException(503, "API aplikasi belum diaktifkan: atur secret/env TOKEN_SECRET (minimal 16 karakter acak).")
    return str(r)


def _bersih(v):
    """Nilai pandas/numpy -> tipe Python yang aman untuk JSON (NaN/NaT -> None)."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(v, "item"):
        return v.item()
    if isinstance(v, (datetime, Date)):
        return v.isoformat()
    return v


def _baris(r) -> dict:
    return {k: _bersih(v) for k, v in dict(r).items()}


def _db(e: SQLAlchemyError):
    raise HTTPException(503, "Database sedang bermasalah. Coba lagi sebentar lagi.")


def peserta_sekarang(authorization: Optional[str] = Header(default=None)):
    """Dependensi: token sah + peserta masih ada + PIN belum diganti sejak token dibuat."""
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(401, "Belum masuk. Silakan login.")
    data = tk.periksa(authorization[7:].strip(), _rahasia())
    if data is None:
        raise HTTPException(401, "Sesi berakhir. Silakan login lagi.")
    try:
        p = absensi.get_peserta(_eng(), data["n"])
    except SQLAlchemyError as e:
        _db(e)
    if p is None or tk.versi_pin(p["pin_hash"]) != data.get("v"):
        raise HTTPException(401, "Sesi tidak berlaku lagi (PIN diganti atau akun dihapus). Silakan login lagi.")
    return p


def peserta_setuju(p=Depends(peserta_sekarang)):
    """Seperti peserta_sekarang, tapi menolak bila persetujuan privasi (versi terbaru) belum diberikan."""
    if absensi.perlu_persetujuan(p):
        raise HTTPException(403, detail={"kode": "persetujuan", "pesan": "Setujui pemberitahuan privasi data terlebih dulu."})
    return p


def _sekarang():
    return datetime.now(absensi.WITA)


def _profil(p) -> dict:
    gedung = p["gedung"]
    lat, lon = absensi.GEDUNG_COORD[gedung]
    return {
        "nama": p["nama"], "divisi": p["divisi"], "gedung": gedung, "inisial": absensi.inisial(p["nama"]),
        "universitas": _bersih(p.get("universitas")), "jurusan": _bersih(p.get("jurusan")),
        "kelas_semester": _bersih(p.get("kelas_semester")),
        "setuju_pada": _bersih(p.get("setuju_pada")), "perlu_persetujuan": bool(absensi.perlu_persetujuan(p)),
        # untuk pra-cek di HP (keputusan akhir tetap di server)
        "lokasi_gedung": {"lat": lat, "lon": lon, "radius_m": absensi.RADIUS_METER, "maks_akurasi_m": absensi.MAKS_AKURASI_METER},
        "jam": {
            "masuk": [absensi.JAM_MASUK_BUKA.strftime("%H:%M"), absensi.JAM_MASUK_TUTUP.strftime("%H:%M")],
            "pulang": [absensi.JAM_PULANG_BUKA.strftime("%H:%M"), absensi.JAM_PULANG_TUTUP.strftime("%H:%M")],
            "batas_tepat_waktu": absensi.BATAS_TEPAT_WAKTU.strftime("%H:%M"),
        },
        "uang_makan_per_hari": absensi.UANG_MAKAN,
    }


# ---------------------------------------------------------------- login & profil
class LoginIn(BaseModel):
    nama: str = Field(min_length=1, max_length=120)
    pin: str = Field(min_length=1, max_length=12)


@router.get("/daftar-nama", summary="Daftar nama peserta untuk layar login (tanpa data lain)")
def daftar_nama():
    try:
        df = absensi.list_peserta(_eng())
    except SQLAlchemyError as e:
        _db(e)
    return {"nama": df["nama"].tolist()}


@router.post("/login", summary="Masuk dengan nama + PIN; mengembalikan token")
def login(body: LoginIn):
    eng = _eng()
    rahasia = _rahasia()
    nama = body.nama.strip()
    try:
        status, menit = absensi.proses_login(eng, f"peserta:{nama}", lambda: absensi.cek_pin(eng, nama, body.pin))
        if status == "terkunci":
            raise HTTPException(429, absensi.pesan_terkunci(menit))
        if status != "ok":
            raise HTTPException(401, "PIN salah. Silakan periksa kembali.")
        p = absensi.get_peserta(eng, nama)
    except SQLAlchemyError as e:
        _db(e)
    if p is None:
        raise HTTPException(401, "PIN salah. Silakan periksa kembali.")
    return {"token": tk.buat(p["nama"], p["pin_hash"], rahasia), "berlaku_hari": tk.BERLAKU_DETIK // 86400, "peserta": _profil(p)}


@router.get("/saya", summary="Profil peserta yang sedang masuk")
def saya(p=Depends(peserta_sekarang)):
    return _profil(p)


@router.get("/privasi", summary="Teks pemberitahuan privasi")
def privasi(p=Depends(peserta_sekarang)):
    return {"versi": absensi.VERSI_PRIVASI,
            "teks": absensi.TEKS_PRIVASI.format(kontak=absensi.kontak_privasi()),
            "perlu_persetujuan": bool(absensi.perlu_persetujuan(p))}


@router.post("/persetujuan", summary="Setuju pada pemberitahuan privasi")
def persetujuan(p=Depends(peserta_sekarang)):
    try:
        absensi.catat_persetujuan(_eng(), p["nama"])
    except SQLAlchemyError as e:
        _db(e)
    return {"ok": True, "versi": absensi.VERSI_PRIVASI}


# ---------------------------------------------------------------- beranda, riwayat, rekap
@router.get("/hari-ini", summary="Status presensi hari ini (beranda)")
def hari_ini(p=Depends(peserta_sekarang)):
    now = _sekarang()
    try:
        sh = absensi.get_status_hari(_eng(), p["nama"], now.strftime("%Y-%m-%d"))
    except SQLAlchemyError as e:
        _db(e)
    judul, ket = absensi.teks_status_hari(now, sh, html=False)

    def jam(r):
        return r["jam"][:5] if r else None

    return {
        "waktu_server": now.isoformat(timespec="seconds"),
        "tanggal": now.strftime("%Y-%m-%d"), "tanggal_teks": absensi.tgl_indonesia(now.date()),
        "salam": absensi.salam(now.hour), "libur": bool(absensi.hari_libur(now)),
        "masuk": jam(sh["masuk"]), "pulang": jam(sh["pulang"]),
        "terlambat": bool(sh["masuk"] and absensi.terlambat(sh["masuk"]["jam"])),
        "uang_makan": int(sh["pulang"]["uang_makan"] or 0) if sh["pulang"] else 0,
        "izin": _bersih(sh["izin"]["alasan"]) if sh["izin"] else None,
        "cuti": _bersih(sh["cuti"]["alasan"]) if sh["cuti"] else None,
        "gagal_masuk": bool(sh["gagal_masuk"]), "gagal_pulang": bool(sh["gagal_pulang"]),
        "status_judul": judul, "status_teks": ket,
        # jenis absen yang masuk akal sekarang (pilihan awal di layar absen)
        "absen_berikutnya": "pulang" if (sh["masuk"] is not None and sh["pulang"] is None) else "masuk",
    }


@router.get("/riwayat", summary="Riwayat absen (percobaan gagal tidak ditampilkan)")
def riwayat(limit: int = Query(60, ge=1, le=200), p=Depends(peserta_sekarang)):
    try:
        rows = absensi.get_riwayat_peserta(_eng(), p["nama"], limit=limit, tanpa_gagal=True)
    except SQLAlchemyError as e:
        _db(e)
    hasil = []
    for r in rows:
        d = datetime.strptime(r["tanggal"], "%Y-%m-%d").date()
        item = {"tanggal": r["tanggal"], "tanggal_teks": absensi.tgl_indonesia(d), "jenis": r["jenis"], "jam": r["jam"][:5],
                "jarak_m": round(r["jarak"] or 0), "uang_makan": int(r["uang_makan"] or 0), "alasan": _bersih(r["alasan"])}
        if r["jenis"] == "masuk":
            item["ketepatan"] = absensi.ketepatan(r["jam"])
            item["terlambat"] = bool(absensi.terlambat(r["jam"]))
        hasil.append(item)
    return {"jumlah": len(hasil), "riwayat": hasil}


@router.get("/rekap", summary="Rekap satu bulan (bawaan: bulan ini)")
def rekap(tahun: Optional[int] = Query(None, ge=2020, le=2100), bulan: Optional[int] = Query(None, ge=1, le=12),
          p=Depends(peserta_sekarang)):
    now = _sekarang()
    tahun, bulan = tahun or now.year, bulan or now.month
    prefix = f"{tahun:04d}-{bulan:02d}-"
    try:
        rows = [r for r in absensi.get_riwayat_peserta(_eng(), p["nama"], limit=1000) if r["tanggal"].startswith(prefix)]
    except SQLAlchemyError as e:
        _db(e)
    masuk_ok = [r for r in rows if r["jenis"] == "masuk" and r["status"] == "berhasil"]
    pulang_ok = [r for r in rows if r["jenis"] == "pulang" and r["status"] == "berhasil"]
    return {
        "tahun": tahun, "bulan": bulan, "bulan_teks": f"{absensi.BULAN_ID[bulan - 1]} {tahun}",
        "hari_hadir": len({r["tanggal"] for r in masuk_ok}),
        "total_uang_makan": int(sum(r["uang_makan"] or 0 for r in pulang_ok)),
        "terlambat": len({r["tanggal"] for r in masuk_ok if absensi.terlambat(r["jam"])}),
        "izin": len({r["tanggal"] for r in rows if r["jenis"] == "izin"}),
        "cuti": len({r["tanggal"] for r in rows if r["jenis"] == "cuti"}),
    }


# ---------------------------------------------------------------- absen
async def _baca_foto(foto: UploadFile) -> bytes:
    data = await foto.read(MAKS_FOTO_BYTE + 1)
    if not data:
        raise HTTPException(422, "Foto kosong.")
    if len(data) > MAKS_FOTO_BYTE:
        raise HTTPException(413, "Foto terlalu besar (maksimal 8 MB).")
    return data


@router.post("/cek-wajah", summary="Pra-cek wajah di bingkai oval (tidak menyimpan apa pun)")
async def cek_wajah(foto: UploadFile = File(...), p=Depends(peserta_setuju)):
    ok, pesan = absensi.periksa_wajah(await _baca_foto(foto))
    return {"ok": bool(ok), "pesan": pesan}


@router.post("/absen", summary="Absen masuk / pulang (GPS + selfie). Keputusan akhir ada di server.")
async def absen(jenis: Literal["masuk", "pulang"] = Form(...), lat: float = Form(..., ge=-90, le=90),
                lon: float = Form(..., ge=-180, le=180), akurasi: Optional[float] = Form(None, ge=0),
                foto: UploadFile = File(...), p=Depends(peserta_setuju)):
    data = await _baca_foto(foto)
    try:
        hasil = absensi.proses_absen(_eng(), p["nama"], p, jenis, data, lat, lon, akurasi)
    except SQLAlchemyError as e:
        _db(e)
    return hasil


class TidakHadirIn(BaseModel):
    jenis: Literal["izin", "cuti"]
    tanggal: Date
    alasan: str = Field(min_length=1, max_length=300)


@router.post("/tidak-hadir", summary="Ajukan izin / cuti")
def tidak_hadir(body: TidakHadirIn, p=Depends(peserta_setuju)):
    try:
        return absensi.proses_tidak_hadir(_eng(), p["nama"], body.jenis, body.tanggal, body.alasan)
    except SQLAlchemyError as e:
        _db(e)


# ---------------------------------------------------------------- logbook
class LogbookIn(BaseModel):
    tanggal: Date
    uraian: str = Field(min_length=1, max_length=1000)
    durasi_jam: float = Field(ge=0.5, le=10)
    paket_id: Optional[int] = None


class LogbookUbahIn(BaseModel):
    uraian: str = Field(min_length=1, max_length=1000)
    durasi_jam: float = Field(ge=0.5, le=10)
    paket_id: Optional[int] = None


def _ralat(e: ValueError):
    raise HTTPException(422, str(e))


@router.get("/logbook/paket", summary="Daftar paket kerja divisi (untuk pilihan di formulir logbook)")
def logbook_paket(p=Depends(peserta_setuju)):
    try:
        df = absensi.daftar_paket(_eng(), p["divisi"])
    except SQLAlchemyError as e:
        _db(e)
    return {"paket": [{"id": int(r.id), "kode": r.kode, "nama": r.nama} for r in df.itertuples()]}


@router.get("/logbook", summary="Logbook milikmu (terbaru dulu)")
def logbook_daftar(limit: int = Query(60, ge=1, le=300), p=Depends(peserta_setuju)):
    try:
        df = absensi.daftar_logbook(_eng(), nama=p["nama"], limit=limit)
    except SQLAlchemyError as e:
        _db(e)
    item = [_baris(r) for r in df.to_dict("records")]
    return {
        "menunggu": sum(1 for r in item if r["status"] == "diajukan"),
        "disetujui": sum(1 for r in item if r["status"] == "disetujui"),
        "ditolak": sum(1 for r in item if r["status"] == "ditolak"),
        "logbook": item,
    }


@router.post("/logbook", summary="Ajukan logbook baru", status_code=201)
def logbook_tambah(body: LogbookIn, p=Depends(peserta_setuju)):
    try:
        lid = absensi.tambah_logbook(_eng(), p["nama"], body.tanggal, body.uraian, body.durasi_jam, body.paket_id, now=_sekarang())
    except ValueError as e:
        _ralat(e)
    except SQLAlchemyError as e:
        _db(e)
    return {"id": lid, "pesan": "Logbook diajukan. Pembimbing akan meninjaunya."}


@router.put("/logbook/{logbook_id}", summary="Ubah logbook milikmu yang belum disetujui (diajukan ulang)")
def logbook_ubah(logbook_id: int, body: LogbookUbahIn, p=Depends(peserta_setuju)):
    try:
        absensi.ubah_logbook(_eng(), logbook_id, p["nama"], body.uraian, body.durasi_jam, body.paket_id, now=_sekarang())
    except ValueError as e:
        _ralat(e)
    except SQLAlchemyError as e:
        _db(e)
    return {"ok": True, "pesan": "Logbook diperbarui dan diajukan ulang."}


@router.delete("/logbook/{logbook_id}", summary="Hapus logbook milikmu yang belum disetujui")
def logbook_hapus(logbook_id: int, p=Depends(peserta_setuju)):
    try:
        absensi.hapus_logbook(_eng(), logbook_id, p["nama"])
    except ValueError as e:
        _ralat(e)
    except SQLAlchemyError as e:
        _db(e)
    return {"ok": True}


@router.get("/logbook-progres", summary="Jam logbook disetujui + progres paket kerja divisi")
def logbook_progres(p=Depends(peserta_setuju)):
    eng = _eng()
    try:
        saya_df = absensi.daftar_logbook(eng, nama=p["nama"], status="disetujui", limit=1000)
        hasil = absensi.progres_divisi(eng, p["divisi"], _sekarang().date())
    except SQLAlchemyError as e:
        _db(e)
    paket = [] if hasil.empty else [
        {"kode": r.kode, "nama": r.nama, "jam_disetujui": float(r.jam_disetujui), "estimasi_jam": float(r.estimasi_jam),
         "rencana": float(r.rencana), "progres": float(r.progres), "spi": _bersih(r.spi), "status": r.status}
        for r in hasil.itertuples()
    ]
    return {"jam_disetujui_saya": float(saya_df["durasi_jam"].sum()), "paket": paket}
