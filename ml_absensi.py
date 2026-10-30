"""Modul machine learning untuk Absensi PKL.

Murni pandas + scikit-learn: TIDAK memakai Streamlit maupun database, supaya bisa dites, dijalankan dari
skrip/cron, dan dipakai ulang oleh API. Penyimpanan model ke database ada di app.py.

Isi:
1. Prediksi risiko terlambat (klasifikasi): fitur dari riwayat, dua kandidat model, tuning GridSearchCV +
   cross-validation, uji temporal, metrik lengkap, pembanding baseline.
2. Deteksi absen mencurigakan (Isolation Forest + aturan transparan).
3. Monitoring: drift fitur (PSI), evaluasi pascadeploy, rekomendasi retrain.
"""
from __future__ import annotations

import io
import math
from datetime import date

import joblib
import numpy as np
import pandas as pd
import sklearn
from sklearn.base import clone
from sklearn.ensemble import IsolationForest, RandomForestClassifier
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, confusion_matrix, f1_score, precision_score, recall_score, roc_auc_score
from sklearn.model_selection import GridSearchCV, StratifiedKFold
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler

BATAS_MENIT = 8 * 60            # terlambat kalau jam masuk > 08:00 (sama dengan aturan di app.py)
MIN_DATA = 60                   # minimal baris untuk melatih model
MIN_PER_KELAS = 10              # minimal contoh tiap kelas (terlambat / tidak)
MIN_ANOMALI = 30                # minimal catatan untuk deteksi anomali

FITUR_TERLAMBAT = [
    "hari_idx", "is_senin", "is_jumat", "telat_7", "telat_semua",
    "rata_menit_5", "streak_telat", "n_hadir", "gedung_b",
]
LABEL_FITUR = {
    "hari_idx": "Hari dalam minggu", "is_senin": "Hari Senin", "is_jumat": "Hari Jumat",
    "telat_7": "Tingkat terlambat 7 hari hadir terakhir", "telat_semua": "Tingkat terlambat sepanjang riwayat",
    "rata_menit_5": "Rata-rata jam masuk 5 hari hadir terakhir", "streak_telat": "Terlambat berturut-turut",
    "n_hadir": "Jumlah hari hadir sebelumnya", "gedung_b": "Ditempatkan di Gedung B",
}
# Fitur kumulatif (n_hadir, telat_semua) berubah sendiri seiring riwayat bertambah, bukan tanda drift: dikecualikan
# dari alarm. Sisanya dibandingkan hanya pada baris dengan riwayat sudah matang (jendela 7 hari hadir sudah terisi),
# dan tidak dinilai bila barisnya terlalu sedikit (PSI di data kecil hanya derau).
FITUR_PANTAU = [f for f in FITUR_TERLAMBAT if f not in ("n_hadir", "telat_semua")]
MIN_RIWAYAT_PANTAU = 7
MIN_BARIS_PANTAU = 30


class DataKurang(ValueError):
    """Data belum cukup untuk melatih / menganalisis."""


def menit_dari_jam(jam) -> int:
    j = str(jam)
    return int(j[0:2]) * 60 + int(j[3:5])


def label_menit(m) -> str:
    m = max(0, int(round(m)))
    return f"{m // 60:02d}:{m % 60:02d}"


# ------------------------------------------------------------------
# 1. PREDIKSI RISIKO TERLAMBAT
# ------------------------------------------------------------------
def hitung_fitur(menit_riwayat, telat_riwayat, hari_idx, gedung_b=0) -> dict:
    """Fitur untuk satu hari, HANYA dari riwayat sebelum hari itu (tidak ada kebocoran data masa depan)."""
    n = len(menit_riwayat)
    streak = 0
    for t in reversed(telat_riwayat):
        if t:
            streak += 1
        else:
            break
    return {
        "hari_idx": int(hari_idx), "is_senin": int(hari_idx == 0), "is_jumat": int(hari_idx == 4),
        "telat_7": float(np.mean(telat_riwayat[-7:])) if n else np.nan,
        "telat_semua": float(np.mean(telat_riwayat)) if n else np.nan,
        "rata_menit_5": float(np.mean(menit_riwayat[-5:])) if n else np.nan,
        "streak_telat": int(streak), "n_hadir": int(n), "gedung_b": int(gedung_b),
    }


def _masuk_harian(masuk_df):
    """Satu baris per (nama, tanggal): absen masuk paling awal. Menambah kolom menit & telat."""
    d = masuk_df[["nama", "tanggal", "jam"]].copy().sort_values(["nama", "tanggal", "jam"])
    d = d.drop_duplicates(["nama", "tanggal"], keep="first")
    d["menit"] = d["jam"].map(menit_dari_jam)
    d["telat"] = (d["menit"] > BATAS_MENIT).astype(int)
    return d


def bangun_dataset(masuk_df, peserta_df) -> pd.DataFrame:
    """masuk_df: absen masuk berhasil (nama, tanggal, jam). Hasil: satu baris per hari hadir + fitur + y (1 = terlambat)."""
    kolom = ["nama", "tanggal", *FITUR_TERLAMBAT, "y"]
    if masuk_df is None or len(masuk_df) == 0:
        return pd.DataFrame(columns=kolom)
    d = _masuk_harian(masuk_df)
    gedung = peserta_df.drop_duplicates("nama").set_index("nama")["gedung"] if len(peserta_df) else pd.Series(dtype=object)
    baris = []
    for nama, g in d.groupby("nama", sort=False):
        gb = int(gedung.get(nama, "A") == "B")
        hm, ht = [], []
        for tgl, menit, telat in zip(g["tanggal"], g["menit"], g["telat"]):
            f = hitung_fitur(hm, ht, date.fromisoformat(tgl).weekday(), gb)
            f.update(nama=nama, tanggal=tgl, y=int(telat))
            baris.append(f)
            hm.append(int(menit))
            ht.append(int(telat))
    return pd.DataFrame(baris)[kolom].sort_values(["tanggal", "nama"]).reset_index(drop=True)


def fitur_prediksi(riwayat_df, tanggal, gedung="A") -> dict:
    """Fitur untuk memprediksi hari `tanggal` (date) dari riwayat masuk satu peserta (kolom tanggal, jam)."""
    iso = tanggal.isoformat()
    if riwayat_df is None or len(riwayat_df) == 0:
        hm, ht = [], []
    else:
        r = riwayat_df.copy()
        r["nama"] = "x"
        d = _masuk_harian(r)
        d = d[d["tanggal"] < iso]
        hm, ht = d["menit"].astype(int).tolist(), d["telat"].astype(int).tolist()
    return hitung_fitur(hm, ht, tanggal.weekday(), int(gedung == "B"))


def _metrik(y, pred, proba=None) -> dict:
    y = np.asarray(y)
    out = {
        "n": int(len(y)),
        "accuracy": float(accuracy_score(y, pred)),
        "precision": float(precision_score(y, pred, zero_division=0)),
        "recall": float(recall_score(y, pred, zero_division=0)),
        "f1": float(f1_score(y, pred, zero_division=0)),
        "roc_auc": None,
        "confusion_matrix": confusion_matrix(y, pred, labels=[0, 1]).tolist(),   # [[TN, FP], [FN, TP]]
    }
    if proba is not None and len(np.unique(y)) == 2:
        out["roc_auc"] = float(roc_auc_score(y, proba))
    return out


def spesifikasi_baseline(seri):
    """Ringkasan distribusi fitur saat pelatihan (untuk PSI)."""
    x = pd.Series(seri).dropna().astype(float)
    if x.empty:
        return None
    uniq = np.sort(x.unique())
    if len(uniq) <= 6:
        return {"tipe": "kategori", "nilai": [float(v) for v in uniq], "proporsi": [float((x == v).mean()) for v in uniq], "n": int(len(x))}
    batas = np.unique(np.quantile(x, np.linspace(0, 1, 11)[1:-1]))
    idx = np.searchsorted(batas, x, side="right")
    prop = np.bincount(idx, minlength=len(batas) + 1) / len(x)
    return {"tipe": "angka", "batas": [float(b) for b in batas], "proporsi": [float(p) for p in prop], "n": int(len(x))}


def hitung_psi(spec, nilai):
    """Population Stability Index antara distribusi saat latih (spec) dan data baru. None kalau tak bisa dihitung.

    PSI mentah punya bias positif pada sampel kecil (kira-kira (jumlah_bin - 1) x (1/n_baru + 1/n_latih)): data yang
    sama persis pun tampak 'bergeser'. Bias itu dikurangkan (hasil tidak negatif) agar ambang 0,1 / 0,25 tetap bermakna."""
    x = pd.Series(nilai).dropna().astype(float)
    if spec is None or x.empty:
        return None
    if spec["tipe"] == "kategori":
        aktual = [float(np.isclose(x, v).mean()) for v in spec["nilai"]]
        aktual.append(max(0.0, 1.0 - sum(aktual)))                  # nilai yang belum pernah ada saat latih
        harap = list(spec["proporsi"]) + [0.0]
    else:
        idx = np.searchsorted(spec["batas"], x, side="right")
        aktual = (np.bincount(idx, minlength=len(spec["batas"]) + 1) / len(x)).tolist()
        harap = list(spec["proporsi"])
    eps = 1e-4
    psi = sum((max(a, eps) - max(e, eps)) * math.log(max(a, eps) / max(e, eps)) for a, e in zip(aktual, harap))
    n_latih = spec.get("n")
    bias = (len(spec["proporsi"]) - 1) * (1.0 / len(x) + (1.0 / n_latih if n_latih else 0.0))
    return float(max(0.0, psi - bias))


def status_psi(psi) -> str:
    if psi is None:
        return "-"
    return "Stabil" if psi < 0.1 else ("Waspada" if psi < 0.25 else "Drift")


def _pipeline_lr():
    return Pipeline([
        ("imp", SimpleImputer(strategy="median")), ("sc", StandardScaler()),
        ("m", LogisticRegression(class_weight="balanced", max_iter=1000)),
    ])


def _pipeline_rf(seed):
    return Pipeline([
        ("imp", SimpleImputer(strategy="median")),
        ("m", RandomForestClassifier(n_estimators=100, class_weight="balanced_subsample", random_state=seed)),
    ])


def ke_bytes(model) -> bytes:
    buf = io.BytesIO()
    joblib.dump({"model": model, "fitur": FITUR_TERLAMBAT, "sklearn": sklearn.__version__}, buf, compress=3)
    return buf.getvalue()


def dari_bytes(data: bytes) -> dict:
    """PERHATIAN: joblib/pickle hanya boleh memuat data dari sumber tepercaya (database milik sendiri)."""
    return joblib.load(io.BytesIO(bytes(data)))


def latih_terlambat(dataset, seed=42, uji_frac=0.2) -> dict:
    """Latih & evaluasi model prediksi terlambat.

    Alur: pisah data secara temporal (hari-hari terakhir = data uji) -> tuning dua kandidat dengan GridSearchCV +
    stratified k-fold CV (skor F1) hanya pada data latih -> pilih terbaik -> ukur pada data uji + bandingkan dengan
    baseline 'selalu tebak kelas mayoritas' -> latih ulang pada SEMUA data untuk dipakai.
    """
    d = dataset.dropna(subset=["y"]).copy()
    if len(d) < MIN_DATA:
        raise DataKurang(f"Data terlalu sedikit untuk melatih model: {len(d)} baris (minimal {MIN_DATA}).")
    y_all = d["y"].astype(int)
    hitung = y_all.value_counts()
    if len(hitung) < 2 or hitung.min() < MIN_PER_KELAS:
        raise DataKurang(f"Tiap kelas butuh minimal {MIN_PER_KELAS} contoh. Sekarang: terlambat={int(hitung.get(1, 0))}, "
                         f"tidak terlambat={int(hitung.get(0, 0))}.")
    tanggal = sorted(d["tanggal"].unique())
    if len(tanggal) < 5:
        raise DataKurang("Butuh data dari minimal 5 hari berbeda.")
    potong = tanggal[int(len(tanggal) * (1 - uji_frac))]
    latih, uji = d[d["tanggal"] < potong], d[d["tanggal"] >= potong]
    ytr, yte = latih["y"].astype(int), uji["y"].astype(int)
    minoritas = int(ytr.value_counts().min()) if ytr.nunique() == 2 else 0
    if minoritas < 3:
        raise DataKurang("Data latih (sebelum periode uji) belum punya cukup contoh dari kedua kelas.")
    k = min(5, minoritas)
    cv = StratifiedKFold(n_splits=k, shuffle=True, random_state=seed)

    kandidat = {
        "Logistic Regression": (_pipeline_lr(), {"m__C": [0.1, 1.0, 10.0]}),
        "Random Forest": (_pipeline_rf(seed), {"m__max_depth": [3, 5, None], "m__min_samples_leaf": [1, 5]}),
    }
    hasil_cv, terbaik, nama_terbaik, skor_terbaik = {}, None, None, -1.0
    for nama, (est, grid) in kandidat.items():
        gs = GridSearchCV(est, grid, scoring="f1", cv=cv, n_jobs=1, refit=True)
        gs.fit(latih[FITUR_TERLAMBAT], ytr)
        hasil_cv[nama] = {"cv_f1": float(gs.best_score_), "parameter": {k_: v for k_, v in gs.best_params_.items()}}
        if gs.best_score_ > skor_terbaik:            # seri -> pertahankan yang lebih sederhana (urutan: LR dulu)
            terbaik, nama_terbaik, skor_terbaik = gs.best_estimator_, nama, float(gs.best_score_)

    pred = terbaik.predict(uji[FITUR_TERLAMBAT])
    proba = terbaik.predict_proba(uji[FITUR_TERLAMBAT])[:, 1]
    m_uji = _metrik(yte, pred, proba)
    mayoritas = int(ytr.mean() >= 0.5)
    m_base = _metrik(yte, np.full(len(yte), mayoritas))
    m_uji["tanggal_mulai"] = str(potong)

    final = clone(terbaik).fit(d[FITUR_TERLAMBAT], y_all)           # model yang dipakai: dilatih pada semua data
    m = final.named_steps["m"]
    if hasattr(m, "coef_"):
        penting = np.abs(m.coef_[0])
        arah = {f: float(c) for f, c in zip(FITUR_TERLAMBAT, m.coef_[0])}
    else:
        penting, arah = m.feature_importances_, None
    penting = penting / penting.sum() if penting.sum() > 0 else penting

    metrik = {
        "model": nama_terbaik, "parameter": hasil_cv[nama_terbaik]["parameter"], "cv_k": int(k), "cv_f1": skor_terbaik,
        "kandidat": hasil_cv, "uji": m_uji,
        "baseline": {"strategi": "selalu menebak kelas mayoritas", "accuracy": m_base["accuracy"], "f1": m_base["f1"]},
        "lebih_baik_dari_baseline": bool(m_uji["f1"] > m_base["f1"] or m_uji["accuracy"] > m_base["accuracy"]),
        "kepentingan": {f: float(v) for f, v in sorted(zip(FITUR_TERLAMBAT, penting), key=lambda t: -t[1])},
        "koefisien": arah, "n_data": int(len(d)), "n_latih": int(len(latih)),
        "persen_terlambat": float(y_all.mean() * 100),
        "tanggal_awal": str(tanggal[0]), "tanggal_akhir": str(tanggal[-1]), "versi_sklearn": sklearn.__version__,
    }
    matang = d[d["n_hadir"] >= MIN_RIWAYAT_PANTAU]
    baseline = {f: spesifikasi_baseline(matang[f]) for f in FITUR_PANTAU}
    return {"artefak": ke_bytes(final), "metrik": metrik, "baseline": baseline,
            "n_data": int(len(d)), "tanggal_akhir": str(tanggal[-1])}


def prediksi_proba(paket, fitur_df) -> np.ndarray:
    return paket["model"].predict_proba(fitur_df[paket["fitur"]])[:, 1]


def tingkat_risiko(p) -> str:
    return "Rendah" if p < 0.33 else ("Sedang" if p < 0.66 else "Tinggi")


def faktor_risiko(f: dict) -> list[str]:
    """Alasan yang bisa dibaca manusia, dari nilai fitur (bukan dari internal model)."""
    out = []
    if f.get("streak_telat", 0) >= 2:
        out.append(f"terlambat {int(f['streak_telat'])} hari hadir berturut-turut")
    if pd.notna(f.get("telat_7")) and f["telat_7"] >= 0.5:
        out.append(f"{f['telat_7'] * 100:.0f}% dari 7 hari hadir terakhir terlambat")
    if pd.notna(f.get("rata_menit_5")) and f["rata_menit_5"] > BATAS_MENIT - 10:
        out.append(f"rata-rata jam masuk {label_menit(f['rata_menit_5'])} (mendekati/melewati batas)")
    if f.get("is_senin"):
        out.append("hari Senin")
    if f.get("n_hadir", 0) == 0:
        out.append("belum ada riwayat kehadiran")
    return out


def evaluasi(paket, dataset) -> dict | None:
    """Ukur model pada dataset (mis. data sesudah tanggal latih). None kalau datanya belum layak diukur."""
    if dataset is None or len(dataset) < 10 or dataset["y"].nunique() < 2:
        return None
    proba = prediksi_proba(paket, dataset)
    return _metrik(dataset["y"].astype(int), (proba >= 0.5).astype(int), proba)


def pantau_drift(baseline: dict, dataset_baru) -> pd.DataFrame:
    """PSI tiap fitur terpantau: data baru (baris dengan riwayat matang) vs distribusi saat latih.
    Status '-' berarti datanya belum cukup untuk dinilai."""
    baru = dataset_baru[dataset_baru["n_hadir"] >= MIN_RIWAYAT_PANTAU] if len(dataset_baru) else dataset_baru
    cukup = len(baru) >= MIN_BARIS_PANTAU
    baris = []
    for f in FITUR_PANTAU:
        psi = hitung_psi((baseline or {}).get(f), baru[f]) if cukup else None
        baris.append({"fitur": f, "label": LABEL_FITUR[f], "psi": psi, "status": status_psi(psi)})
    return pd.DataFrame(baris)


def rekomendasi_retrain(metrik_aktif, drift_df, eval_baru, hari_sejak_latih, n_baru, n_latih) -> tuple[bool, list[str]]:
    """Aturan sederhana & transparan kapan model perlu dilatih ulang."""
    alasan = []
    if drift_df is not None and len(drift_df):
        parah = drift_df[drift_df["status"] == "Drift"]["label"].tolist()
        if parah:
            alasan.append("Distribusi data berubah jauh (PSI >= 0,25) pada: " + ", ".join(parah))
    if eval_baru and metrik_aktif:
        f1_uji = metrik_aktif.get("uji", {}).get("f1")
        if f1_uji is not None and eval_baru["f1"] < f1_uji - 0.10:
            alasan.append(f"F1 pada data baru ({eval_baru['f1']:.2f}) turun lebih dari 0,10 dari F1 uji saat dilatih ({f1_uji:.2f})")
    if n_latih and n_baru >= 0.3 * n_latih and hari_sejak_latih >= 14:
        alasan.append(f"Sudah ada {n_baru} data baru (>= 30% dari data latih) dan model berumur {hari_sejak_latih} hari")
    return bool(alasan), alasan


# ------------------------------------------------------------------
# 2. DETEKSI ABSEN MENCURIGAKAN
# ------------------------------------------------------------------
FITUR_ANOMALI = ["jarak_f", "akurasi_f", "dx", "dy", "menit_rel", "jarak_dari_median", "n_kembar_sendiri"]
LABEL_ANOMALI = {
    "jarak_f": "jarak ke gedung", "akurasi_f": "akurasi GPS", "dx": "posisi timur-barat", "dy": "posisi utara-selatan",
    "menit_rel": "jam absen", "jarak_dari_median": "jarak dari lokasi biasa peserta", "n_kembar_sendiri": "koordinat identik berulang",
}
M_PER_DERAJAT = 111_320.0


def _format_nilai(fitur, v, menit_asli=None):
    if fitur == "menit_rel":
        return label_menit(menit_asli)          # tampilkan jam sebenarnya, bukan selisihnya
    if fitur == "n_kembar_sendiri":
        return f"{v:.0f}x"
    return f"{v:.0f} m"


def _kalimat_penyimpangan(fitur, nilai, tipikal, z, menit_asli):
    """Satu alasan yang bisa dibaca admin: apa yang menyimpang, ke arah mana, dan seberapa jauh dari kebiasaan."""
    if fitur == "menit_rel":
        arah = "lebih lambat" if z > 0 else "lebih awal"
        return f"jam absen {label_menit(menit_asli)} jauh {arah} dari jam tipikal"
    if fitur in ("dx", "dy"):                    # koordinat bertanda: sebut arah mata angin, bukan "lebih kecil/besar"
        arah = ("lebih ke timur" if z > 0 else "lebih ke barat") if fitur == "dx" else ("lebih ke utara" if z > 0 else "lebih ke selatan")
        return f"posisi {arah} dari biasanya ({nilai:+.0f} m dari titik gedung; tipikal {tipikal:+.0f} m)"
    arah = "lebih besar" if z > 0 else "lebih kecil"
    return (f"{LABEL_ANOMALI[fitur]} {arah} dari biasanya ({_format_nilai(fitur, nilai)}; "
            f"tipikal {_format_nilai(fitur, tipikal)})")


def deteksi_anomali(df, peserta_df, gedung_coord, persen=5.0, seed=42) -> pd.DataFrame:
    """Tandai absen yang tidak lazim. df: nama, tanggal, jam, jenis, status, lat, lon, akurasi, jarak.

    Dua lapis: (1) Isolation Forest (tanpa label, mencari kombinasi nilai yang jarang) dan (2) aturan transparan:
    koordinat identik persis berulang >= 3 catatan pada peserta yang sama, atau akurasi GPS <= 1 m.
    Hasilnya SINYAL untuk ditinjau manusia, bukan bukti kecurangan."""
    d = df[df["jenis"].isin(["masuk", "pulang"]) & df["lat"].notna() & df["lon"].notna()].copy().reset_index(drop=True)
    if len(d) < MIN_ANOMALI:
        raise DataKurang(f"Catatan berlokasi terlalu sedikit untuk analisis anomali: {len(d)} (minimal {MIN_ANOMALI}).")
    gedung = peserta_df.drop_duplicates("nama").set_index("nama")["gedung"] if len(peserta_df) else pd.Series(dtype=object)
    d["gedung"] = d["nama"].map(gedung).fillna("A")
    lat0 = d["gedung"].map(lambda g: gedung_coord[g][0])
    lon0 = d["gedung"].map(lambda g: gedung_coord[g][1])
    d["dy"] = (d["lat"] - lat0) * M_PER_DERAJAT
    d["dx"] = (d["lon"] - lon0) * M_PER_DERAJAT * np.cos(np.radians(lat0))
    d["menit"] = d["jam"].map(menit_dari_jam)
    d["menit_rel"] = d["menit"] - d.groupby("jenis")["menit"].transform("median")   # jam relatif terhadap jam tipikal jenis itu
    d["akurasi_f"] = d["akurasi"].fillna(d["akurasi"].median() if d["akurasi"].notna().any() else 0.0)
    d["jarak_f"] = d["jarak"].fillna(np.sqrt(d["dx"] ** 2 + d["dy"] ** 2))
    med = d.groupby("nama")[["lat", "lon"]].transform("median")
    d["jarak_dari_median"] = np.sqrt(
        ((d["lat"] - med["lat"]) * M_PER_DERAJAT) ** 2
        + ((d["lon"] - med["lon"]) * M_PER_DERAJAT * np.cos(np.radians(d["lat"]))) ** 2
    )
    kunci = d["lat"].round(6).astype(str) + "|" + d["lon"].round(6).astype(str)
    d["n_kembar_sendiri"] = d.groupby(["nama", kunci])["lat"].transform("size") - 1
    d["n_kembar_lintas"] = d.groupby(kunci)["nama"].transform("nunique") - 1      # informasi saja (Wi-Fi positioning bisa kembar)

    X = d[FITUR_ANOMALI].astype(float)
    Xs = StandardScaler().fit_transform(X)
    kontaminasi = float(np.clip(persen / 100.0, 0.001, 0.5))
    iso = IsolationForest(n_estimators=200, contamination=kontaminasi, random_state=seed)
    flag_model = iso.fit_predict(Xs) == -1
    mentah = -iso.score_samples(Xs)
    rentang = mentah.max() - mentah.min()
    d["skor"] = ((mentah - mentah.min()) / rentang * 100) if rentang > 0 else 0.0

    med_x, mad = X.median(), (X - X.median()).abs().median() * 1.4826
    skala = mad.where(mad > 1e-9, X.std(ddof=0)).where(lambda s: s > 1e-9)
    z = (X - med_x) / skala

    alasan, aturan = [], []
    for i in range(len(d)):
        a = []
        if d.at[i, "n_kembar_sendiri"] >= 2:
            a.append(f"koordinat GPS identik persis pada {int(d.at[i, 'n_kembar_sendiri']) + 1} catatan peserta ini")
        if pd.notna(d.at[i, "akurasi"]) and d.at[i, "akurasi"] <= 1.0:
            a.append(f"akurasi GPS hanya {d.at[i, 'akurasi']:.1f} m (tidak lazim)")
        aturan.append(bool(a))
        if flag_model[i]:
            zi = z.iloc[i].abs().dropna().sort_values(ascending=False)
            ada_aturan_kembar = any("identik persis" in t for t in a)
            kuat = [f for f in zi.index if zi[f] > 2.0 and not (f == "n_kembar_sendiri" and ada_aturan_kembar)][:2]
            for f in kuat:
                a.append(_kalimat_penyimpangan(f, X.at[i, f], med_x[f], z.at[i, f], d.at[i, "menit"]))
            if not kuat and not ada_aturan_kembar:
                a.append("beberapa nilai sedikit menyimpang sekaligus (tidak ada satu faktor yang dominan)")
        alasan.append("; ".join(a))
    d["ditandai"] = flag_model | np.array(aturan)
    d["alasan"] = alasan
    kolom = ["nama", "tanggal", "jam", "jenis", "status", "jarak", "akurasi", "dx", "dy", "skor",
             "ditandai", "alasan", "n_kembar_lintas"]
    return d[kolom].sort_values(["ditandai", "skor"], ascending=[False, False]).reset_index(drop=True)
