"""Latih ulang model prediksi keterlambatan dari seluruh data absensi. Cocok untuk cron / GitHub Actions.

Pemakaian:
    python jalankan_latih.py                       # aktifkan hanya jika tidak lebih buruk dari model aktif
    python jalankan_latih.py --aktifkan selalu     # selalu aktifkan versi baru
    python jalankan_latih.py --aktifkan tidak      # simpan saja
    python jalankan_latih.py --json                # keluaran JSON (untuk dibaca skrip lain)

Koneksi database sama seperti app.py (secrets/env: DB_HOST, DB_USER, dst).
Kode keluar: 0 = berhasil, 2 = data belum cukup (bukan error), 1 = error lain.
"""
from __future__ import annotations

import argparse
import json
import sys

import app as absensi
import ml_absensi as ml


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--aktifkan", choices=absensi.KEBIJAKAN_AKTIF, default="jika_lebih_baik")
    ap.add_argument("--json", action="store_true", help="cetak hasil sebagai JSON")
    args = ap.parse_args(argv)
    try:
        r = absensi.latih_dan_simpan(absensi.get_engine(), args.aktifkan)
    except ml.DataKurang as e:
        print(f"Data belum cukup: {e}", file=sys.stderr)
        return 2
    except Exception as e:                                                   # noqa: BLE001 - pesan jelas untuk log CI
        print(f"Gagal melatih model: {type(e).__name__}: {e}", file=sys.stderr)
        return 1
    k = r["metrik"]
    if args.json:
        print(json.dumps({"id": r["id"], "aktif": r["aktif"], "model": k["model"], "uji": k["uji"],
                          "baseline": k["baseline"]}, ensure_ascii=False))
    else:
        u = k["uji"]
        print(f"Versi {r['id']} tersimpan ({k['model']}), data {k['n_data']} baris.")
        print(f"Uji temporal ({u['n']} data): accuracy {u['accuracy']:.2f}, precision {u['precision']:.2f}, "
              f"recall {u['recall']:.2f}, F1 {u['f1']:.2f}, ROC-AUC {u['roc_auc'] if u['roc_auc'] is None else round(u['roc_auc'], 2)}")
        print(f"Baseline: accuracy {k['baseline']['accuracy']:.2f}, F1 {k['baseline']['f1']:.2f}")
        print("Status: " + ("AKTIF" if r["aktif"] else "tersimpan, TIDAK diaktifkan (tidak lebih baik dari model aktif)"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
