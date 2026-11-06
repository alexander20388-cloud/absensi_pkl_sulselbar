# Panduan Singkat Absensi PKL

Aplikasi tetap berjalan di **Streamlit Cloud**. Tidak ada server sewaan (VPS). Paket ini hanya menambah file di sekitar `app.py`.

## 1. Yang harus Anda lakukan (3 langkah)

1. **Salin semua file dari zip ke folder repositori GitHub Anda** (timpa yang lama). Pastikan `template_absen.xlsx` tetap ada.
2. **Pastikan `requirements.txt` yang baru ikut ter-commit.** Ini penting: isinya membatasi OpenCV ke versi di bawah 5. Tanpa batas itu,
   verifikasi wajah bisa mati diam-diam.
3. **Commit dan push.** Selesai.

Setelah itu semuanya berjalan sendiri:
- Streamlit Cloud memperbarui aplikasi (HTTPS sudah otomatis dari Streamlit).
- GitHub menjalankan pengecekan otomatis (lihat tab **Actions**). Hasilnya centang hijau atau silang merah; tidak mengubah aplikasi.

Tidak ada yang perlu diubah di Settings > Secrets. Opsional: tambahkan `kontak_privasi` (mis. email Human Capital) agar tampil di pemberitahuan privasi.

## 2. Apa fungsi tiap file

| File | Fungsi | Perlu untuk Streamlit Cloud? |
|---|---|---|
| `app.py`, `proyek.py`, `ml_absensi.py`, `template_absen.xlsx`, `requirements.txt` | Aplikasi itu sendiri | **Wajib** |
| `.github/workflows/ci.yml`, `ruff.toml`, `requirements-dev.txt`, `test_*.py`, `conftest.py` | Pengecekan otomatis di GitHub (CI) | Tidak, jalan sendiri di GitHub |
| `Dockerfile`, `docker-compose.yml`, `.dockerignore`, `.env.example` | Menjalankan aplikasi lewat Docker di komputer sendiri | Tidak (opsional) |
| `api.py`, `jalankan_latih.py` | API dan skrip latih model | Tidak: **tidak bisa berjalan di Streamlit Cloud**, hanya di komputer sendiri/Docker |
| `.gitignore`, `.github/dependabot.yml` | Merapikan repositori dan menawarkan pembaruan pustaka | Tidak |

Tidak satu pun dari Docker maupun CI ada di dalam `app.py`; semuanya file terpisah. `app.py` tidak perlu diubah untuk keduanya.

## 3. CI (pengecekan otomatis), singkatnya

Setiap kali Anda push ke GitHub, robot GitHub menjalankan: pemeriksaan kode (ruff), semua tes, lalu membangun image Docker dan memastikan ia bisa hidup.
Kalau ada yang gagal, Anda diberi tahu sebelum masalahnya sampai ke pengguna.

## 4. Docker (opsional, hanya bila ingin menjalankan di laptop)

```
docker compose up --build
```
Web di http://localhost:8501 dan API di http://localhost:8000/docs. Tanpa file `.env`, aplikasi memakai database sementara (cukup untuk mencoba).
Untuk menambah database Postgres lokal: salin `.env.example` menjadi `.env`, isi `DB_PASSWORD`, lalu `docker compose --profile dbsendiri up --build`.

## 5. Yang tidak otomatis karena tanpa server

- **Latih ulang model:** manual, lewat Admin > Model > Prediksi > *Latih sekarang*.
- **Verifikasi jejak audit:** manual, lewat Admin > Proyek > Audit > *Verifikasi integritas rantai*.
  Catat hash puncak yang ditampilkan (mis. sebulan sekali) di tempat di luar aplikasi. Penghapusan catatan paling akhir tidak memutus rantai,
  tetapi mengubah hash puncak itu.
- **API:** hanya bila dijalankan sendiri (Docker atau `uvicorn api:api`).

## 6. Fitur baru dan aturannya

**Peran pengguna**

| Peran | Masuk lewat | Dapat melakukan |
|---|---|---|
| Peserta | PIN | absen, menulis dan memperbaiki logbook sendiri |
| Pembimbing | pilihan *Pembimbing* di layar masuk; PIN 6-8 angka | meninjau logbook dan mengelola paket kerja **hanya divisinya**; tidak melihat foto atau GPS |
| Admin | password | semuanya, termasuk membuat akun pembimbing (Proyek > Pembimbing) dan menetapkan anggaran |

Pembimbing tidak tersedia pada mode `?mode=pkl` (aplikasi Android peserta).

**Jejak audit.** Perubahan data penting dicatat dalam transaksi yang sama dengan perubahannya. PIN tidak pernah dicatat; rekening hanya 4 angka
terakhir. Tiap catatan terkait dengan catatan sebelumnya lewat hash, jadi mengubah atau menghapus catatan di tengah akan terdeteksi.

**Privasi.** Peserta menyetujui pemberitahuan privasi (GPS dan foto) sekali sebelum memakai aplikasi, dan lagi bila `VERSI_PRIVASI` dinaikkan.
Teks `TEKS_PRIVASI` di `app.py` adalah **draf**: sesuaikan dengan kebijakan instansi dan mintalah tinjauan unit hukum atau Human Capital.

**Anggaran uang makan.** Realisasi = uang makan yang tercatat saat absen pulang, per divisi peserta *saat absen*. Proyeksi akhir bulan = realisasi
hari yang sudah selesai + rata-rata harian x hari kerja tersisa. Saran anggaran = peserta x hari kerja x tarif x tingkat kehadiran + cadangan 10%.

## 7. Batasan yang perlu diketahui

- **Pembatasan login:** 5 kali salah berturut-turut (dalam 15 menit) mengunci akun/peran itu selama 15 menit; penguncian tercatat di jejak audit. Konsekuensinya, orang lain bisa sengaja mengunci akun seseorang (termasuk admin) selama 15 menit dengan menebak sengaja; tunggu masa kunci habis.
- PIN peserta kini di-hash dengan scrypt. PIN lama (SHA-256) tetap bisa dipakai dan otomatis ditingkatkan saat peserta berhasil login. Karena PIN hanya 4-6 angka, tetap lebih aman bila database tidak bocor.
- **Akurasi GPS:** absen ditolak bila ketelitian lokasi perangkat lebih buruk dari 100 m (`MAKS_AKURASI_METER`). Toleransi jarak ke gedung tetap 200 m (`RADIUS_METER`).
- Password admin bawaan `admin123` hanya berlaku di mode lokal (tanpa database cloud). Bila memakai database cloud, `admin_password` WAJIB diisi; kalau tidak, login admin ditolak.
- **Hari libur nasional dan cuti bersama** sudah dikenali untuk 2026 dan 2027 (daftar `LIBUR_NASIONAL` dan `CUTI_BERSAMA` di `app.py`, sumber SKB 3 Menteri). Pada tanggal itu absen masuk, pulang, izin, dan cuti ditolak, logbook tidak bisa diisi, dan hari itu tidak dihitung sebagai hari kerja (rekap, anggaran, jadwal paket kerja).
  **Setiap tahun daftar harus ditambah manual** setelah SKB tahun berikutnya terbit (biasanya September). Tanpa itu, tahun 2028 dan seterusnya hanya mengenal Sabtu-Minggu. Cuti bersama bersifat fakultatif bagi swasta: bila PKL tetap masuk pada cuti bersama, ubah `IKUT_CUTI_BERSAMA = False`. Libur mendadak (mis. penetapan pemerintah di luar SKB) ditambahkan ke `LIBUR_NASIONAL`.
- **Keterangan tulisan di foto absen.** Foto yang tersimpan (masuk dan pulang) memuat tanggal-jam WITA dan alamat gedung di pojok kanan atas. Alamat diatur di `ALAMAT_FOTO` (`app.py`); saat ini Gedung A dan B memakai alamat yang sama, sesuaikan bila alamat Gedung B berbeda. Tulisan hanya ada pada foto yang disimpan, bukan pada tampilan kamera saat memotret, dan tidak ada pada foto lama.
- Tampilan kartu Logbook di Beranda memakai gaya yang sudah ada tetapi belum diperiksa secara visual di layar ponsel; mohon dicek.
- Model ML disimpan dengan joblib: versi scikit-learn saat memuat harus sama dengan saat melatih. `requirements.txt` menjaga itu.
