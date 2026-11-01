"""Tes modul ML. Jalankan: pytest -q test_ml.py"""
from datetime import date, timedelta

import numpy as np
import pandas as pd
import pytest

import ml_absensi as ml

GEDUNG = {"A": (-5.150070542826675, 119.41568236911743), "B": (-5.150163347858429, 119.41329137592547)}
NAMA = [f"P{i}" for i in range(8)]
PESERTA = pd.DataFrame({"nama": NAMA, "gedung": ["A", "B"] * 4, "divisi": ["Divisi Umum"] * 8})


def hari_kerja(n, mulai=date(2026, 8, 3)):
    out, d = [], mulai
    while len(out) < n:
        if d.weekday() < 5:
            out.append(d)
        d += timedelta(days=1)
    return out


def data_masuk(n_hari=40, seed=1, geser=0, mulai=date(2026, 8, 3)):
    """Tiap peserta punya kebiasaan jam masuk sendiri (+ Senin lebih telat). geser = menit tambahan (simulasi drift)."""
    rng = np.random.default_rng(seed)
    mu = {n: 455 + 6 * i for i, n in enumerate(NAMA)}            # peserta makin ke belakang makin sering telat
    baris = []
    for tgl in hari_kerja(n_hari, mulai):
        for n in NAMA:
            if rng.random() < 0.05:
                continue                                          # kadang tidak hadir
            m = int(rng.normal(mu[n] + (12 if tgl.weekday() == 0 else 0) + geser, 14))
            m = min(max(m, 365), 750)
            baris.append({"nama": n, "tanggal": tgl.isoformat(), "jam": f"{m // 60:02d}:{m % 60:02d}:00"})
    return pd.DataFrame(baris)


# ---------- dataset & fitur ----------
def test_fitur_tidak_bocor_dari_masa_depan():
    masuk = data_masuk(10)
    ds = ml.bangun_dataset(masuk, PESERTA)
    pertama = ds.sort_values("tanggal").groupby("nama").head(1)
    assert (pertama["n_hadir"] == 0).all() and pertama["telat_7"].isna().all()   # hari pertama: tak ada riwayat
    # ubah hasil hari terakhir -> fitur hari terakhir tidak boleh berubah
    n0 = NAMA[0]
    asli = ds[ds["nama"] == n0].iloc[-1]
    m2 = masuk.copy()
    idx = m2[(m2["nama"] == n0) & (m2["tanggal"] == asli["tanggal"])].index
    m2.loc[idx, "jam"] = "11:59:00"
    baru = ml.bangun_dataset(m2, PESERTA)
    row = baru[(baru["nama"] == n0) & (baru["tanggal"] == asli["tanggal"])].iloc[0]
    for f in ml.FITUR_TERLAMBAT:
        assert (row[f] == asli[f]) or (pd.isna(row[f]) and pd.isna(asli[f]))
    assert row["y"] == 1


def test_fitur_prediksi_sama_dengan_dataset():
    masuk = data_masuk(15)
    ds = ml.bangun_dataset(masuk, PESERTA)
    n0 = NAMA[1]
    baris = ds[ds["nama"] == n0].iloc[-1]
    riwayat = masuk[masuk["nama"] == n0][["tanggal", "jam"]]
    f = ml.fitur_prediksi(riwayat, date.fromisoformat(baris["tanggal"]), "B")
    for k in ml.FITUR_TERLAMBAT:
        assert f[k] == pytest.approx(baris[k], nan_ok=True)


def test_ambang_terlambat_0800_tepat_tidak_terlambat():
    m = pd.DataFrame([{"nama": "P0", "tanggal": "2026-08-03", "jam": "08:00:00"},
                      {"nama": "P0", "tanggal": "2026-08-04", "jam": "08:01:00"}])
    ds = ml.bangun_dataset(m, PESERTA)
    assert ds["y"].tolist() == [0, 1]


# ---------- pelatihan ----------
@pytest.fixture(scope="module")
def ds():
    return ml.bangun_dataset(data_masuk(40), PESERTA)


@pytest.fixture(scope="module")
def h(ds):
    return ml.latih_terlambat(ds)


def test_latih_terlambat_dan_metrik(h):
    m = h["metrik"]
    assert m["model"] in ("Logistic Regression", "Random Forest")
    assert set(m["kandidat"]) == {"Logistic Regression", "Random Forest"}
    assert 0 <= m["uji"]["f1"] <= 1 and m["uji"]["roc_auc"] > 0.6
    assert m["lebih_baik_dari_baseline"] is True
    assert sum(m["kepentingan"].values()) == pytest.approx(1.0)
    assert len(m["uji"]["confusion_matrix"]) == 2
    # data uji benar-benar setelah data latih (uji temporal)
    assert m["uji"]["tanggal_mulai"] > m["tanggal_awal"]
    import json
    json.dumps(m); json.dumps(h["baseline"])                      # semuanya serializable ke JSON


def test_model_tersimpan_bisa_dimuat_dan_konsisten(ds, h):
    paket = ml.dari_bytes(h["artefak"])
    p1 = ml.prediksi_proba(paket, ds.tail(20))
    p2 = ml.prediksi_proba(ml.dari_bytes(h["artefak"]), ds.tail(20))
    assert np.allclose(p1, p2) and ((p1 >= 0) & (p1 <= 1)).all()


def test_data_kurang_ditolak_dengan_pesan_jelas():
    kecil = ml.bangun_dataset(data_masuk(3), PESERTA)
    with pytest.raises(ml.DataKurang, match="terlalu sedikit"):
        ml.latih_terlambat(kecil)
    semua_tepat = ml.bangun_dataset(data_masuk(40, geser=-120), PESERTA)       # hampir tak ada yang terlambat
    with pytest.raises(ml.DataKurang, match="kelas"):
        ml.latih_terlambat(semua_tepat)


def test_hasil_deterministik_dengan_seed_sama(ds, h):
    a, b = h["metrik"], ml.latih_terlambat(ds)["metrik"]
    assert a["uji"]["f1"] == b["uji"]["f1"] and a["model"] == b["model"]


# ---------- drift & monitoring ----------
def test_psi_nol_untuk_data_sama_dan_besar_untuk_data_bergeser(ds, h):
    sama = ml.pantau_drift(h["baseline"], ds)
    assert (sama["psi"].dropna() < 0.05).all() and (sama["status"] == "Stabil").all()
    bergeser = ml.bangun_dataset(data_masuk(40, seed=2, geser=45), PESERTA)
    drift = ml.pantau_drift(h["baseline"], bergeser).set_index("fitur")
    assert drift.loc["rata_menit_5", "status"] == "Drift"
    assert drift.loc["telat_7", "psi"] > 0.1
    assert "telat_semua" not in drift.index and "n_hadir" not in drift.index   # fitur kumulatif memang berubah sendiri


def test_psi_kategori_baru_tidak_error():
    spec = ml.spesifikasi_baseline([0, 0, 1, 1, 1])
    assert ml.hitung_psi(spec, [0, 1, 7, 7]) > 0.25
    assert ml.hitung_psi(None, [1, 2]) is None and ml.hitung_psi(spec, []) is None


def test_evaluasi_pascadeploy_dan_data_terlalu_sedikit(ds, h):
    paket = ml.dari_bytes(h["artefak"])
    assert ml.evaluasi(paket, ds.head(5)) is None
    e = ml.evaluasi(paket, ds)
    assert e["n"] == len(ds) and 0 <= e["accuracy"] <= 1


def test_rekomendasi_retrain():
    drift_ok = pd.DataFrame([{"fitur": "telat_7", "label": "x", "psi": 0.02, "status": "Stabil"}])
    drift_buruk = pd.DataFrame([{"fitur": "telat_7", "label": "Tingkat terlambat", "psi": 0.4, "status": "Drift"}])
    metrik = {"uji": {"f1": 0.8}}
    perlu, _ = ml.rekomendasi_retrain(metrik, drift_ok, {"f1": 0.78}, 5, 10, 400)
    assert perlu is False
    perlu, alasan = ml.rekomendasi_retrain(metrik, drift_buruk, None, 5, 10, 400)
    assert perlu and "Drift" not in alasan[0] and "Tingkat terlambat" in alasan[0]
    perlu, alasan = ml.rekomendasi_retrain(metrik, drift_ok, {"f1": 0.55}, 5, 10, 400)
    assert perlu and "turun" in alasan[0]
    perlu, alasan = ml.rekomendasi_retrain(metrik, drift_ok, None, 20, 150, 400)
    assert perlu and "data baru" in alasan[0]


def test_faktor_risiko_dan_tingkat():
    f = ml.hitung_fitur([490, 495, 500], [1, 1, 1], 0, 0)
    teks = " ".join(ml.faktor_risiko(f))
    assert "berturut-turut" in teks and "Senin" in teks
    assert ml.tingkat_risiko(0.1) == "Rendah" and ml.tingkat_risiko(0.5) == "Sedang" and ml.tingkat_risiko(0.9) == "Tinggi"
    assert "belum ada riwayat" in " ".join(ml.faktor_risiko(ml.hitung_fitur([], [], 2, 0)))


# ---------- anomali ----------
def data_lokasi(seed=3):
    rng = np.random.default_rng(seed)
    baris = []
    for tgl in hari_kerja(25):
        for n in NAMA:
            g = "A" if int(n[1:]) % 2 == 0 else "B"
            lat0, lon0 = GEDUNG[g]
            for jenis, jam in (("masuk", 7 * 60 + int(rng.normal(45, 12))), ("pulang", 17 * 60 + int(rng.normal(5, 4)))):
                lat, lon = lat0 + rng.normal(0, 0.0002), lon0 + rng.normal(0, 0.0002)
                jarak = float(np.hypot((lat - lat0) * 111320, (lon - lon0) * 111320 * np.cos(np.radians(lat0))))
                baris.append(dict(nama=n, tanggal=tgl.isoformat(), jam=f"{jam // 60:02d}:{jam % 60:02d}:00", jenis=jenis,
                                  status="berhasil", lat=lat, lon=lon, akurasi=float(rng.uniform(8, 40)), jarak=jarak))
    return pd.DataFrame(baris)


def test_anomali_menangkap_pola_yang_ditanam():
    df = data_lokasi()
    lat0, lon0 = GEDUNG["A"]
    # peserta "Zed" memakai GPS palsu: koordinat persis sama berulang + akurasi 1 m
    palsu = [dict(nama="P0", tanggal=t.isoformat(), jam="07:55:00", jenis="masuk", status="berhasil",
                  lat=lat0 + 0.00005, lon=lon0 + 0.00005, akurasi=1.0, jarak=7.0) for t in hari_kerja(5, date(2026, 12, 7))]
    # absen jauh dari gedung pada jam janggal
    janggal = [dict(nama="P2", tanggal="2026-12-14", jam="03:10:00", jenis="masuk", status="gagal_lokasi",
                    lat=lat0 + 0.009, lon=lon0 + 0.009, akurasi=900.0, jarak=1400.0)]
    gabung = pd.concat([df, pd.DataFrame(palsu + janggal)], ignore_index=True)
    hasil = ml.deteksi_anomali(gabung, PESERTA, GEDUNG, persen=5)
    tanda = hasil[hasil["ditandai"]]
    ditanam = tanda[(tanda["tanggal"] >= "2026-12-07")]
    assert len(ditanam) >= 5                                       # 5 GPS palsu + 1 janggal, minimal 5 tertangkap
    assert any("identik persis" in a for a in tanda["alasan"])
    assert any("akurasi GPS hanya" in a for a in tanda["alasan"])
    jang = hasil[(hasil["tanggal"] == "2026-12-14") & (hasil["nama"] == "P2")].iloc[0]
    assert jang["ditandai"] and jang["skor"] > 80
    normal = hasil[hasil["tanggal"] < "2026-12-01"]
    assert normal["ditandai"].mean() < 0.10                         # tidak menandai sembarangan
    assert hasil["skor"].between(0, 100).all()


def test_anomali_data_sedikit_ditolak():
    with pytest.raises(ml.DataKurang):
        ml.deteksi_anomali(data_lokasi().head(10), PESERTA, GEDUNG)


def test_jam_pulang_normal_tidak_ditandai_hanya_karena_beda_dari_jam_masuk():
    hasil = ml.deteksi_anomali(data_lokasi(), PESERTA, GEDUNG, persen=5)
    pulang = hasil[hasil["jenis"] == "pulang"]
    assert pulang["ditandai"].mean() < 0.10
    assert not any("jam absen" in a and "17:0" in a for a in hasil["alasan"])      # jam pulang ~17:05 itu normal


def test_absen_pulang_jam_janggal_tetap_tertangkap():
    df = data_lokasi()
    lat0, lon0 = GEDUNG["A"]
    ekstra = pd.DataFrame([dict(nama="P0", tanggal="2026-12-14", jam="03:30:00", jenis="pulang", status="berhasil",
                                lat=lat0 + 0.0001, lon=lon0 + 0.0001, akurasi=20.0, jarak=15.0)])
    hasil = ml.deteksi_anomali(pd.concat([df, ekstra], ignore_index=True), PESERTA, GEDUNG, persen=5)
    row = hasil[(hasil["tanggal"] == "2026-12-14") & (hasil["nama"] == "P0")].iloc[0]
    assert row["ditandai"] and "jam absen 03:30 jauh lebih awal dari jam tipikal" in row["alasan"]


def test_drift_tidak_dinilai_bila_data_baru_terlalu_sedikit(ds, h):
    d = ml.pantau_drift(h["baseline"], ds.tail(10))
    assert (d["status"] == "-").all() and d["psi"].isna().all()
    d0 = ml.pantau_drift(h["baseline"], ds.iloc[0:0])
    assert (d0["status"] == "-").all()


def test_tidak_ada_alarm_drift_palsu_pada_jendela_terbaru_data_yang_sama(ds, h):
    # jendela 14 hari kerja terakhir dari data YANG SAMA dengan data latih tidak boleh dianggap drift
    tgl = sorted(ds["tanggal"].unique())[-14:]
    d = ml.pantau_drift(h["baseline"], ds[ds["tanggal"].isin(tgl)])
    assert (d["status"] == "Stabil").all(), d


def test_psi_terkalibrasi_pada_sampel_kecil_tanpa_perubahan(ds, h):
    """Simulasi: ambil berulang 80 baris acak dari populasi yang SAMA dengan data latih. Hampir semuanya harus 'Stabil'."""
    rng = np.random.default_rng(0)
    matang = ds[ds["n_hadir"] >= ml.MIN_RIWAYAT_PANTAU]
    salah = total = 0
    for _ in range(60):
        d = ml.pantau_drift(h["baseline"], matang.sample(80, replace=True, random_state=int(rng.integers(1e9))))
        salah += int((d["status"] != "Stabil").sum())
        total += len(d)
    assert salah / total < 0.05, f"alarm palsu {salah}/{total}"


def test_pergeseran_nyata_tetap_terdeteksi_pada_sampel_kecil(h):
    bergeser = ml.bangun_dataset(data_masuk(16, seed=5, geser=40, mulai=date(2026, 10, 1)), PESERTA)
    d = ml.pantau_drift(h["baseline"], bergeser).set_index("fitur")
    assert d.loc["rata_menit_5", "status"] == "Drift"


def test_alasan_menyebut_arah_dan_nilai_tipikal_bukan_kalimat_generik():
    df = data_lokasi()
    lat0, lon0 = GEDUNG["A"]
    jauh = pd.DataFrame([dict(nama="P0", tanggal="2026-12-14", jam="07:50:00", jenis="masuk", status="gagal_lokasi",
                              lat=lat0 + 0.004, lon=lon0, akurasi=15.0, jarak=450.0)])
    hasil = ml.deteksi_anomali(pd.concat([df, jauh], ignore_index=True), PESERTA, GEDUNG, persen=5)
    row = hasil[(hasil["tanggal"] == "2026-12-14") & (hasil["nama"] == "P0")].iloc[0]
    assert row["ditandai"] and "lebih besar dari biasanya" in row["alasan"] and "tipikal" in row["alasan"]
    tanda = hasil[hasil["ditandai"]]
    generik = tanda["alasan"].str.contains("tidak ada satu faktor").mean()
    assert generik < 0.5                                       # mayoritas penanda punya alasan yang konkret
