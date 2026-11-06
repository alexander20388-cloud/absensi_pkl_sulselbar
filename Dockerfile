# syntax=docker/dockerfile:1
# Satu image untuk web (Streamlit) dan API (FastAPI); perintah jalannya ditentukan di docker-compose.yml.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# Dependensi dipasang lebih dulu supaya layer ini ter-cache selama requirements.txt tidak berubah.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Kode + aset (mis. template_absen.xlsx). Berkas yang dikecualikan diatur di .dockerignore.
COPY . .

# Layar pembuka: sisipkan ke index.html Streamlit saat build (masih root, paket bisa ditulis) supaya tampil sebelum loading bawaan.
RUN python splash_awal.py

# app.py membaca template_absen.xlsx untuk ekspor rekap; tanpa file itu ekspor Excel error saat dipakai.
RUN test -f /app/template_absen.xlsx \
    || (echo "GAGAL BUILD: template_absen.xlsx tidak ada di repo (dibutuhkan ekspor rekap)." >&2; exit 1)

# Jalan sebagai pengguna biasa, bukan root.
RUN useradd --system --create-home --uid 10001 app && chown -R app:app /app
USER app

EXPOSE 8501 8000

# Healthcheck bawaan untuk layanan web. Layanan api menimpanya di docker-compose.yml.
HEALTHCHECK --interval=30s --timeout=5s --start-period=40s --retries=3 \
    CMD ["python", "-c", "import urllib.request as u; u.urlopen('http://127.0.0.1:8501/_stcore/health', timeout=4)"]

CMD ["streamlit", "run", "app.py", \
     "--server.address=0.0.0.0", "--server.port=8501", \
     "--server.headless=true", "--browser.gatherUsageStats=false"]
