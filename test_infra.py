"""Tes konsistensi konfigurasi (dependensi, rahasia, Docker, CI) agar berkas-berkasnya tidak saling bertentangan.
Jalankan: pytest -q test_infra.py"""
import ast
import re
import sys
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).parent
KODE = ["app.py", "api.py", "ml_absensi.py", "proyek.py", "jalankan_latih.py", "splash_awal.py"]


def baca(p):
    return (ROOT / p).read_text(encoding="utf-8")


def yml(p):
    return yaml.safe_load(baca(p))


def pemicu(w):
    """PyYAML (YAML 1.1) membaca kunci `on:` sebagai boolean True; GitHub membacanya sebagai 'on'."""
    return w["on"] if "on" in w else w[True]


# ---------- dependensi ----------
MODUL_KE_PAKET = {
    "streamlit": "streamlit", "streamlit_js_eval": "streamlit-js-eval", "openpyxl": "openpyxl", "requests": "requests",
    "pandas": "pandas", "PIL": "pillow", "sqlalchemy": "sqlalchemy", "numpy": "numpy", "cv2": "opencv-python-headless",
    "sklearn": "scikit-learn", "joblib": "joblib", "altair": "altair", "fastapi": "fastapi", "uvicorn": "uvicorn",
}
TURUNAN_FASTAPI = {"pydantic", "starlette"}          # selalu ikut terpasang bersama fastapi


def paket_requirements():
    out = {}
    for baris in baca("requirements.txt").splitlines():
        baris = baris.split("#")[0].strip()
        if baris:
            out[re.split(r"[=<>!~\[ ]", baris, maxsplit=1)[0].lower().replace("_", "-")] = baris
    return out


def modul_yang_diimpor():
    lokal = {Path(k).stem for k in KODE}
    ada = set()
    for f in KODE:
        for node in ast.walk(ast.parse(baca(f))):          # termasuk import di dalam fungsi (lazy)
            if isinstance(node, ast.Import):
                ada |= {a.name.split(".")[0] for a in node.names}
            elif isinstance(node, ast.ImportFrom) and node.level == 0:
                ada.add(node.module.split(".")[0])
    return {m for m in ada if m not in sys.stdlib_module_names and m not in lokal}


def test_semua_modul_pihak_ketiga_tercantum_di_requirements():
    req = paket_requirements()
    tak_dikenal = modul_yang_diimpor() - set(MODUL_KE_PAKET) - TURUNAN_FASTAPI
    assert not tak_dikenal, f"modul baru perlu dipetakan & dicantumkan di requirements.txt: {tak_dikenal}"
    hilang = {MODUL_KE_PAKET[m] for m in modul_yang_diimpor() if m in MODUL_KE_PAKET} - set(req)
    assert not hilang, f"diimpor kode tetapi tidak ada di requirements.txt: {hilang}"


def test_driver_postgres_dan_batas_opencv():
    req = paket_requirements()
    assert "psycopg2-binary" in req and "postgresql+psycopg2" in baca("app.py")
    assert re.search(r"<\s*5", req["opencv-python-headless"]), "OpenCV 5 menghapus CascadeClassifier: wajib dibatasi < 5"


def test_requirements_dev_mengikuti_requirements_dan_tes_butuh_httpx_dan_yaml():
    dev = baca("requirements-dev.txt")
    for butuh in ("-r requirements.txt", "pytest", "httpx", "ruff", "PyYAML"):
        assert butuh in dev, butuh


# ---------- rahasia ----------
def test_semua_secret_yang_dibaca_kode_ada_di_env_example():
    kunci = set()
    for f in KODE:
        kunci |= set(re.findall(r'secret\("([a-z_]+)"', baca(f)))
    contoh = {b.split("=", 1)[0].lstrip("# ").strip() for b in baca(".env.example").splitlines() if "=" in b}
    hilang = {k.upper() for k in kunci} - contoh
    assert not hilang, f"dibaca kode tetapi tidak didokumentasikan di .env.example: {hilang}"


def test_rahasia_tidak_bisa_masuk_image_atau_git():
    di, gi = baca(".dockerignore").splitlines(), baca(".gitignore").splitlines()
    for pola in (".env", "*.db"):
        assert pola in di and pola in gi
    assert "!.env.example" in di
    contoh = baca(".env.example")
    assert "GANTI-" in contoh and not re.search(r"(?m)^(ADMIN_PASSWORD|SALT|API_KEY|DB_PASSWORD)=(?!GANTI-)\S", contoh)


# ---------- Docker ----------
def test_dockerfile_praktik_dasar_dan_berkas_yang_dirujuk_ada():
    d = baca("Dockerfile")
    assert re.search(r"(?m)^USER app$", d) and "HEALTHCHECK" in d
    assert "template_absen.xlsx" in d and "template_absen.xlsx" in baca("app.py")      # app butuh file ini saat ekspor rekap
    assert "app.py" in d and (ROOT / "app.py").exists()
    assert "test_*.py" in baca(".dockerignore")


def test_compose_struktur_aman_dan_jalan_tanpa_env():
    tanpa_komentar = "\n".join(b.split("#")[0] for b in baca("docker-compose.yml").splitlines())
    assert ":?" not in tanpa_komentar, "variabel wajib membuat `docker compose up` gagal tanpa .env, walau profil db tidak dipakai"
    c = yml("docker-compose.yml")
    assert set(c["services"]) == {"web", "api", "db"}
    for nama, port in (("web", "8501"), ("api", "8000")):
        assert c["services"][nama]["ports"] == [f"127.0.0.1:{port}:{port}"], "hanya dibuka ke komputer sendiri"
    assert c["services"]["db"]["profiles"] == ["dbsendiri"] and "ports" not in c["services"]["db"]
    a = c["x-app"]
    assert a["read_only"] is True and a["cap_drop"] == ["ALL"] and "no-new-privileges:true" in a["security_opt"]
    assert a["env_file"] == [{"path": ".env", "required": False}]


# ---------- CI ----------
def test_ci_hanya_tes_dan_build_tanpa_rahasia_atau_deploy():
    ci = yml(".github/workflows/ci.yml")
    jobs = ci["jobs"]
    assert set(jobs) == {"tes", "image"} and jobs["image"]["needs"] == "tes"
    assert ci["permissions"] == {"contents": "read"}                        # tanpa izin tulis apa pun
    assert "pull_request_target" not in pemicu(ci)
    teks = baca(".github/workflows/ci.yml")
    assert "secrets." not in teks and "ghcr" not in teks.lower() and "ssh" not in teks.lower()
    for m in re.findall(r"uses:\s*(\S+)", teks):
        assert re.search(r"@v\d+(\.\d+)*$", m), f"action harus dipin ke versi mayor: {m}"


def test_ci_menjalankan_lint_dan_tes_dan_menguji_image_hidup():
    langkah = " ".join(s.get("run", "") for j in yml(".github/workflows/ci.yml")["jobs"].values() for s in j["steps"])
    assert "ruff check" in langkah and "pytest" in langkah
    assert "--read-only" in langkah and "_stcore/health" in langkah and "/health" in langkah


def test_tidak_ada_sisa_rujukan_ke_berkas_vps_yang_sudah_dihapus():
    for f in ("Dockerfile", ".dockerignore", ".gitignore", "docker-compose.yml", ".env.example", ".github/workflows/ci.yml", "PANDUAN.md"):
        t = baca(f).lower()
        for kata in ("nginx", "systemd", "deploy.sh", "backup.sh", "certbot", "vps_host", "ghcr.io"):
            assert kata not in t, f"{f} masih menyebut {kata}"
    for hilang in ("deploy", "DEPLOY.md", "test_deploy.py"):
        assert not (ROOT / hilang).exists(), hilang
    assert sorted(p.name for p in (ROOT / ".github" / "workflows").glob("*.yml")) == ["ci.yml"]
