"""Logika murni untuk modul manajemen proyek Absensi PKL.

Tanpa Streamlit dan tanpa database, supaya bisa dites dengan angka yang dihitung manual. Penyimpanan ke database
dan tampilan ada di app.py.

Isi:
1. Kalender kerja
2. Anggaran uang makan: proyeksi (estimate cost), saran anggaran (determine budget), status kendali biaya
3. Paket kerja (WBS): progres terhadap rencana, SPI (schedule performance), diagram Gantt
4. Validasi logbook
5. Kredensial pembimbing (scrypt, garam acak per pengguna)
6. Rantai hash audit (tamper-evident change log)
"""
from __future__ import annotations

import calendar
import hashlib
import hmac
import json
import math
import os
import re
from datetime import date, timedelta

import pandas as pd

HARI_LIBUR_DEFAULT = (5, 6)          # Sabtu, Minggu (sama dengan HARI_LIBUR di app.py)


# ------------------------------------------------------------------
# 1. KALENDER KERJA
# ------------------------------------------------------------------
def hari_kerja_antara(mulai: date, selesai: date, hari_libur=HARI_LIBUR_DEFAULT) -> list[date]:
    """Semua hari kerja dari `mulai` sampai `selesai` (keduanya termasuk)."""
    if selesai < mulai:
        return []
    return [mulai + timedelta(days=i) for i in range((selesai - mulai).days + 1)
            if (mulai + timedelta(days=i)).weekday() not in hari_libur]


def hari_kerja_bulan(tahun: int, bulan: int, hari_libur=HARI_LIBUR_DEFAULT) -> list[date]:
    akhir = calendar.monthrange(tahun, bulan)[1]
    return hari_kerja_antara(date(tahun, bulan, 1), date(tahun, bulan, akhir), hari_libur)


def bulan_sebelumnya(tahun: int, bulan: int, n: int = 1) -> tuple[int, int]:
    idx = tahun * 12 + (bulan - 1) - n
    return idx // 12, idx % 12 + 1


# ------------------------------------------------------------------
# 2. ANGGARAN UANG MAKAN
# ------------------------------------------------------------------
BATAS_WASPADA = 0.9          # proyeksi > 90% anggaran = Waspada
ASUMSI_KEHADIRAN = 0.9       # dipakai bila belum ada riwayat kehadiran


def proyeksi_anggaran(per_hari: dict, hari_kerja: list, hari_ini: date, hari_ini_selesai: bool,
                      n_peserta: int, tarif: int, rate_awal: float | None = None) -> dict:
    """Perkiraan biaya uang makan sampai akhir bulan (estimate cost / estimate at completion).

    per_hari   : {tanggal: rupiah terealisasi} untuk bulan itu (uang makan tercatat saat absen pulang).
    hari_kerja : semua hari kerja bulan itu.
    hari_ini_selesai : True bila hari ini sudah lewat jam absen pulang; bila belum, hari ini masih dianggap
                       "sisa" supaya data hari ini yang belum lengkap tidak menekan rata-rata.
    Rumus: proyeksi = realisasi hari-hari yang SUDAH selesai + rata-rata harian x hari kerja tersisa.
    proyeksi_maks = skenario semua peserta hadir tiap hari tersisa (batas atas).
    """
    selesai = [d for d in hari_kerja if d < hari_ini or (d == hari_ini and hari_ini_selesai)]
    sisa = len(hari_kerja) - len(selesai)
    real_total = int(sum(per_hari.values()))                  # per_hari sudah dibatasi pemanggil pada bulan yang dimaksud
    real_selesai = int(sum(per_hari.get(d, 0) for d in selesai))
    if selesai:
        rate, sumber = real_selesai / len(selesai), "rata-rata harian bulan ini"
    elif rate_awal is not None:
        rate, sumber = float(rate_awal), "asumsi awal (belum ada hari yang selesai)"
    else:
        rate, sumber = 0.0, "belum ada data"
    proyeksi = max(real_total, int(round(real_selesai + rate * sisa)))        # tidak mungkin di bawah yang sudah keluar
    maks = max(proyeksi, int(real_selesai + n_peserta * tarif * sisa))
    return {"realisasi": real_total, "realisasi_selesai": real_selesai, "hari_selesai": len(selesai),
            "hari_sisa": sisa, "hari_total": len(hari_kerja), "rata_harian": rate,
            "proyeksi": proyeksi, "proyeksi_maks": maks, "sumber": sumber}


def status_anggaran(anggaran, proyeksi) -> str:
    if not anggaran:
        return "Belum ada anggaran"
    rasio = proyeksi / anggaran
    if rasio > 1:
        return "Melebihi"
    return "Waspada" if rasio > BATAS_WASPADA else "Aman"


def saran_anggaran(stats: dict, n_hari_kerja: int, tarif: int, cadangan_pct: float = 10.0,
                   pembulatan: int = 10_000) -> pd.DataFrame:
    """Menentukan anggaran (determine budget) per divisi untuk satu bulan.

    stats: {divisi: {"n_peserta": int, "tingkat_kehadiran": float | None}}.
    estimasi dasar = peserta x hari kerja x tarif x tingkat kehadiran; lalu ditambah cadangan (contingency)
    dan dibulatkan ke atas. Tingkat kehadiran tak diketahui diganti ASUMSI_KEHADIRAN dan ditandai.
    """
    baris = []
    for divisi in sorted(stats):
        n = int(stats[divisi].get("n_peserta") or 0)
        if n <= 0:
            continue
        rate = stats[divisi].get("tingkat_kehadiran")
        asumsi = rate is None
        rate = ASUMSI_KEHADIRAN if asumsi else min(1.0, max(0.5, float(rate)))
        dasar = n * n_hari_kerja * tarif * rate
        cadangan = dasar * cadangan_pct / 100.0
        saran = int(math.ceil((dasar + cadangan) / pembulatan) * pembulatan)
        baris.append({"divisi": divisi, "peserta": n, "hari_kerja": n_hari_kerja, "tingkat_kehadiran": rate,
                      "estimasi_dasar": int(round(dasar)), "cadangan": int(round(cadangan)), "saran": saran,
                      "asumsi_kehadiran": asumsi})
    return pd.DataFrame(baris, columns=["divisi", "peserta", "hari_kerja", "tingkat_kehadiran", "estimasi_dasar",
                                        "cadangan", "saran", "asumsi_kehadiran"])


# ------------------------------------------------------------------
# 3. PAKET KERJA (WBS), PROGRES, GANTT
# ------------------------------------------------------------------
KODE_WBS = re.compile(r"^\d{1,3}(\.\d{1,3}){0,3}$")


def validasi_kode_wbs(kode: str) -> bool:
    return bool(KODE_WBS.fullmatch((kode or "").strip()))


def urut_kode(kode: str) -> tuple:
    """'1.10' harus setelah '1.2' (urutan numerik, bukan abjad)."""
    return tuple(int(x) for x in kode.split(".") if x.isdigit())


def validasi_paket(kode, nama, mulai, selesai, estimasi_jam) -> list[str]:
    errs = []
    if not validasi_kode_wbs(kode):
        errs.append("Kode WBS berupa angka bertingkat, mis. 1, 1.1, 2.3.")
    if len((nama or "").strip()) < 3:
        errs.append("Nama paket kerja minimal 3 karakter.")
    if len((nama or "").strip()) > 120:
        errs.append("Nama paket kerja maksimal 120 karakter.")
    if selesai < mulai:
        errs.append("Tanggal selesai tidak boleh sebelum tanggal mulai.")
    elif not hari_kerja_antara(mulai, selesai):
        errs.append("Rentang tanggal tidak memuat satu pun hari kerja.")
    if not (1 <= float(estimasi_jam) <= 2000):
        errs.append("Estimasi jam harus antara 1 dan 2000.")
    return errs


def progres_paket(paket: pd.DataFrame, jam_disetujui: dict, hari_ini: date,
                  hari_libur=HARI_LIBUR_DEFAULT) -> pd.DataFrame:
    """Progres tiap paket kerja terhadap rencana jadwal.

    Rencana (%) = hari kerja yang sudah lewat / total hari kerja paket.
    Progres (%) = jam logbook DISETUJUI / estimasi jam (proksi berbasis usaha; 100% bila ditutup manual).
    SPI = progres / rencana. SPI < 0,9 terlambat, > 1,1 lebih cepat. Ini proksi: jam bukan ukuran keluaran pekerjaan.
    paket: kolom id, kode, nama, mulai, selesai, estimasi_jam, ditutup (0/1).
    """
    baris = []
    for p in paket.to_dict("records"):
        hk = hari_kerja_antara(p["mulai"], p["selesai"], hari_libur)
        total = len(hk)
        lewat = len([d for d in hk if d <= hari_ini])
        rencana = (lewat / total) if total else (1.0 if hari_ini >= p["selesai"] else 0.0)
        jam = float(jam_disetujui.get(p["id"], 0.0))
        ditutup = bool(p.get("ditutup"))
        progres = 1.0 if ditutup else min(1.0, jam / float(p["estimasi_jam"]))
        spi = (progres / rencana) if rencana > 0 else None
        if progres >= 1.0:
            status = "Selesai"
        elif hari_ini < p["mulai"] or spi is None:
            status = "Belum mulai"
        elif hari_ini > p["selesai"]:
            status = "Lewat jadwal"
        elif spi < 0.9:
            status = "Terlambat"
        elif spi > 1.1:
            status = "Lebih cepat"
        else:
            status = "Sesuai jadwal"
        baris.append({**{k: p[k] for k in ("id", "kode", "nama", "mulai", "selesai", "estimasi_jam")},
                      "jam_disetujui": jam, "rencana": rencana, "progres": progres,
                      "deviasi_pp": (progres - rencana) * 100, "spi": spi, "status": status})
    df = pd.DataFrame(baris, columns=["id", "kode", "nama", "mulai", "selesai", "estimasi_jam", "jam_disetujui",
                                      "rencana", "progres", "deviasi_pp", "spi", "status"])
    if len(df):
        df = df.iloc[sorted(range(len(df)), key=lambda i: urut_kode(df.iloc[i]["kode"]))].reset_index(drop=True)
    return df


def grafik_gantt(df: pd.DataFrame, hari_ini: date):
    """Diagram Gantt Altair: batang rencana, batang progres (isi), dan garis hari ini."""
    import altair as alt
    d = df.copy()
    d["label"] = d["kode"] + "  " + d["nama"]
    d["mulai_dt"] = pd.to_datetime(d["mulai"])
    d["selesai_dt"] = pd.to_datetime(d["selesai"]) + pd.Timedelta(days=1)            # inklusif sampai akhir hari
    d["isi_dt"] = d["mulai_dt"] + (d["selesai_dt"] - d["mulai_dt"]) * d["progres"]
    d["progres_txt"] = (d["progres"] * 100).round(0).astype(int).astype(str) + "%"
    urut = list(d["label"])
    y = alt.Y("label:N", sort=urut, title=None, axis=alt.Axis(labelLimit=260))
    dasar = alt.Chart(d).mark_bar(color="#cfdcfa", cornerRadius=4, height=18).encode(
        x=alt.X("mulai_dt:T", title=None), x2="selesai_dt:T", y=y,
        tooltip=["label:N", alt.Tooltip("mulai:N", title="Mulai"), alt.Tooltip("selesai:N", title="Selesai"),
                 alt.Tooltip("progres_txt:N", title="Progres"), alt.Tooltip("status:N", title="Status")])
    isi = alt.Chart(d).mark_bar(color="#2563eb", cornerRadius=4, height=18).encode(
        x="mulai_dt:T", x2="isi_dt:T", y=y)
    garis = alt.Chart(pd.DataFrame({"t": [pd.Timestamp(hari_ini)]})).mark_rule(
        color="#e04a4a", strokeDash=[5, 4], size=2).encode(x="t:T")
    return (dasar + isi + garis).properties(height=max(120, 34 * len(d)))


# ------------------------------------------------------------------
# 4. LOGBOOK
# ------------------------------------------------------------------
def validasi_logbook(tanggal: date, hari_ini: date, uraian: str, durasi: float, jam_lain: float = 0.0,
                     hari_libur=HARI_LIBUR_DEFAULT, maks_mundur: int = 14, maks_per_hari: float = 12.0) -> list[str]:
    """Daftar pesan kesalahan (kosong = valid)."""
    errs = []
    if tanggal > hari_ini:
        errs.append("Tanggal tidak boleh di masa depan.")
    elif (hari_ini - tanggal).days > maks_mundur:
        errs.append(f"Logbook hanya bisa diisi sampai {maks_mundur} hari ke belakang.")
    if tanggal.weekday() in hari_libur:
        errs.append("Tanggal itu hari libur PKL.")
    u = (uraian or "").strip()
    if len(u) < 15:
        errs.append(f"Uraian kegiatan minimal 15 karakter (sekarang {len(u)}).")
    elif len(u) > 1000:
        errs.append(f"Uraian kegiatan maksimal 1000 karakter (sekarang {len(u)}).")
    elif len(u.split()) < 3:
        errs.append("Uraian kegiatan minimal 3 kata.")
    if not (0.5 <= durasi <= 10) or durasi * 2 != round(durasi * 2):
        errs.append("Durasi 0,5 sampai 10 jam, kelipatan 0,5.")
    elif jam_lain + durasi > maks_per_hari:
        errs.append(f"Total jam pada tanggal itu melebihi {maks_per_hari:g} jam (sudah ada {jam_lain:g} jam).")
    return errs


# ------------------------------------------------------------------
# 5. KREDENSIAL PEMBIMBING
# ------------------------------------------------------------------
_SCRYPT_N, _SCRYPT_R, _SCRYPT_P = 2 ** 14, 8, 1


def validasi_pin_pembimbing(pin: str) -> str:
    """Pesan kesalahan, atau '' bila valid. Pembimbing dapat melihat data banyak peserta, jadi PIN lebih panjang."""
    return "" if (pin or "").isdigit() and 6 <= len(pin) <= 8 else "PIN pembimbing harus 6-8 angka."


def hash_pin_kuat(pin: str, garam: bytes | None = None) -> str:
    """scrypt dengan garam acak per pengguna: 'scrypt$N$r$p$garam$hash'. Beda dari SHA-256 bergaram tetap."""
    garam = garam if garam is not None else os.urandom(16)
    h = hashlib.scrypt(pin.encode(), salt=garam, n=_SCRYPT_N, r=_SCRYPT_R, p=_SCRYPT_P, dklen=32)
    return f"scrypt${_SCRYPT_N}${_SCRYPT_R}${_SCRYPT_P}${garam.hex()}${h.hex()}"


def cek_pin_kuat(pin: str, tersimpan: str | None) -> bool:
    try:
        skema, n, r, p, garam, h = (tersimpan or "").split("$")
        if skema != "scrypt":
            return False
        ref = bytes.fromhex(h)
        hh = hashlib.scrypt(pin.encode(), salt=bytes.fromhex(garam), n=int(n), r=int(r), p=int(p), dklen=len(ref))
    except (ValueError, TypeError):
        return False
    return hmac.compare_digest(hh, ref)


# ------------------------------------------------------------------
# 6. RANTAI HASH AUDIT (tamper-evident)
# ------------------------------------------------------------------
GENESIS = "GENESIS"


def hitung_hash(hash_prev: str, waktu: str, aktor: str, peran: str, aksi: str, objek: str, detail: str) -> str:
    """Hash catatan = SHA-256(hash catatan sebelumnya + seluruh isi catatan). Mengubah satu catatan mana pun
    mengubah hash-nya dan memutus rantai sesudahnya."""
    isi = "\x1f".join([waktu, aktor, peran, aksi, objek, detail])
    return hashlib.sha256(f"{hash_prev}\x1e{isi}".encode()).hexdigest()


def detail_json(detail) -> str:
    return json.dumps(detail or {}, ensure_ascii=False, sort_keys=True, default=str)


def verifikasi_rantai(rows: list[dict]) -> dict:
    """rows terurut menurut id menaik, tiap elemen: id, waktu, aktor, peran, aksi, objek, detail, hash_prev, hash.

    Mendeteksi: isi catatan diubah, catatan di tengah dihapus atau disisipkan, urutan dibalik.
    TIDAK dapat mendeteksi penghapusan catatan paling akhir (pemotongan ekor); karena itu hash puncak sebaiknya
    dicatat di luar sistem secara berkala dan dibandingkan.
    """
    prev = GENESIS
    for r in rows:
        if r["hash_prev"] != prev:
            return {"ok": False, "jumlah": len(rows), "rusak_id": r["id"], "hash_terakhir": None,
                    "alasan": "rantai terputus: ada catatan yang dihapus atau disisipkan sebelum catatan ini"}
        if hitung_hash(r["hash_prev"], r["waktu"], r["aktor"], r["peran"], r["aksi"], r["objek"], r["detail"]) != r["hash"]:
            return {"ok": False, "jumlah": len(rows), "rusak_id": r["id"], "hash_terakhir": None,
                    "alasan": "isi catatan telah diubah setelah dicatat"}
        prev = r["hash"]
    return {"ok": True, "jumlah": len(rows), "rusak_id": None, "hash_terakhir": prev if rows else None, "alasan": ""}
