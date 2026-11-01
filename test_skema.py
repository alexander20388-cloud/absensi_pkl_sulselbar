"""Tes skema & anti-absen-ganda. Jalankan: pytest -q test_skema.py"""
import io
import sqlite3
from datetime import datetime, date

import pytest
from PIL import Image
from sqlalchemy import create_engine, select, func, inspect
from sqlalchemy.exc import IntegrityError

import app as A

SENIN = datetime(2026, 10, 26, 7, 30, tzinfo=A.WITA)      # Senin, di dalam jendela absen masuk


def mesin_baru(tmp_path):
    eng = create_engine(f"sqlite:///{tmp_path / 'baru.db'}")
    A.metadata.create_all(eng)
    A._migrasi(eng)
    return eng


def jpeg_bytes():
    buf = io.BytesIO()
    Image.new("RGB", (64, 64), "gray").save(buf, format="JPEG")
    return buf.getvalue()


def baris_absen(eng, **kw):
    dasar = dict(nama="Budi", tanggal="2026-10-26", jam="07:30:00", jenis="masuk", status="berhasil", uang_makan=0)
    dasar.update(kw)
    A.simpan_absen(eng, **dasar)


def jumlah(eng):
    with eng.connect() as c:
        return c.execute(select(func.count()).select_from(A.absensi_t)).scalar()


def test_database_baru_punya_indeks_dan_kolom_fk(tmp_path):
    eng = mesin_baru(tmp_path)
    assert eng.indeks_unik_ok is True
    nama_idx = {i["name"] for i in inspect(eng).get_indexes("absensi")}
    assert {A.NAMA_INDEKS_UNIK, A.NAMA_INDEKS_GABUNGAN} <= nama_idx
    fk = inspect(eng).get_foreign_keys("absensi")
    assert fk and fk[0]["referred_table"] == "peserta"


def test_absen_ganda_ditolak_database(tmp_path):
    eng = mesin_baru(tmp_path)
    baris_absen(eng)
    with pytest.raises(IntegrityError):
        baris_absen(eng, jam="07:31:00")               # masuk berhasil kedua di hari yang sama


def test_gagal_lokasi_boleh_berkali_kali_dan_masuk_pulang_boleh_sama_hari(tmp_path):
    eng = mesin_baru(tmp_path)
    baris_absen(eng, status="gagal_lokasi")
    baris_absen(eng, status="gagal_lokasi", jam="07:32:00")
    baris_absen(eng)                                    # masuk berhasil
    baris_absen(eng, jenis="pulang", jam="17:00:00", uang_makan=15000)
    assert jumlah(eng) == 4


def test_izin_dan_cuti_tidak_boleh_dobel_per_tanggal(tmp_path):
    eng = mesin_baru(tmp_path)
    baris_absen(eng, jenis="izin", status="izin", alasan="ke kampus")
    with pytest.raises(IntegrityError):
        baris_absen(eng, jenis="izin", status="izin", alasan="lagi")
    baris_absen(eng, jenis="cuti", status="cuti", alasan="sakit")   # jenis lain tetap boleh


def test_peserta_id_terisi_otomatis_dan_dilepas_saat_peserta_dihapus(tmp_path):
    eng = mesin_baru(tmp_path)
    A.add_peserta(eng, "Budi", "A", "Divisi Umum", "1234", "130-210-000554616-9")
    pid = A.get_peserta(eng, "Budi")["id"]
    baris_absen(eng)
    with eng.connect() as c:
        assert c.execute(select(A.absensi_t.c.peserta_id)).scalar() == pid
    A.hapus_peserta(eng, "Budi", dengan_absensi=False)   # riwayat dipertahankan
    with eng.connect() as c:
        assert c.execute(select(A.absensi_t.c.peserta_id)).scalar() is None
    assert jumlah(eng) == 1


def test_proses_tidak_hadir_kalah_balapan_jadi_peringatan(tmp_path, monkeypatch):
    eng = mesin_baru(tmp_path)
    A.add_peserta(eng, "Budi", "A", "Divisi Umum", "1234")
    besok = date(2026, 10, 27)
    h1 = A.proses_tidak_hadir(eng, "Budi", "izin", besok, "ke kampus bimbingan", now=SENIN)
    assert h1["level"] == "success"
    monkeypatch.setattr(A, "ada_catatan", lambda *a, **k: False)   # lewati pengecekan awal = simulasi balapan
    h2 = A.proses_tidak_hadir(eng, "Budi", "izin", besok, "ke kampus bimbingan", now=SENIN)
    assert h2["level"] == "warning" and "sudah mengajukan" in h2["pesan"]
    assert jumlah(eng) == 1


def test_proses_absen_kalah_balapan_jadi_peringatan_bukan_error(tmp_path, monkeypatch):
    eng = mesin_baru(tmp_path)
    A.add_peserta(eng, "Budi", "A", "Divisi Umum", "1234")
    peserta = A.get_peserta(eng, "Budi")
    monkeypatch.setattr(A, "periksa_wajah", lambda data: (True, "ok"))
    lat, lon = A.GEDUNG_COORD["A"]
    h1 = A.proses_absen(eng, "Budi", peserta, "masuk", jpeg_bytes(), lat, lon, now=SENIN)
    assert h1["level"] == "success"
    monkeypatch.setattr(A, "sudah_berhasil", lambda *a, **k: False)   # simulasi permintaan kedua yang lolos pengecekan
    h2 = A.proses_absen(eng, "Budi", peserta, "masuk", jpeg_bytes(), lat, lon, now=SENIN)
    assert h2["level"] == "warning" and "sudah absen masuk" in h2["pesan"]
    assert jumlah(eng) == 1


def buat_db_lama(path):
    """Database versi lama: tanpa peserta_id, tanpa indeks unik, dan sudah berisi data ganda."""
    con = sqlite3.connect(path)
    con.executescript("""
        CREATE TABLE peserta (id INTEGER PRIMARY KEY AUTOINCREMENT, nama VARCHAR(100) NOT NULL UNIQUE,
            gedung VARCHAR(1) NOT NULL, divisi VARCHAR(100) NOT NULL DEFAULT '', pin_hash VARCHAR(64) NOT NULL);
        CREATE TABLE absensi (id INTEGER PRIMARY KEY AUTOINCREMENT, nama VARCHAR(100) NOT NULL,
            tanggal VARCHAR(10) NOT NULL, jam VARCHAR(8) NOT NULL, jenis VARCHAR(10) NOT NULL DEFAULT 'masuk',
            lat FLOAT, lon FLOAT, akurasi FLOAT, jarak FLOAT, status VARCHAR(20) NOT NULL,
            uang_makan INTEGER NOT NULL DEFAULT 0, foto BLOB);
        CREATE TABLE notifikasi (id INTEGER PRIMARY KEY AUTOINCREMENT, jenis VARCHAR(20), dibuat VARCHAR(19));
        INSERT INTO peserta (nama, gedung, divisi, pin_hash) VALUES ('Budi', 'A', 'Divisi Umum', 'x');
        INSERT INTO absensi (nama, tanggal, jam, jenis, status, uang_makan) VALUES
            ('Budi', '2026-10-20', '17:00:00', 'pulang', 'berhasil', 15000),
            ('Budi', '2026-10-20', '17:01:00', 'pulang', 'berhasil', 15000),
            ('Budi', '2026-10-20', '07:30:00', 'masuk', 'berhasil', 0),
            ('Budi', '2026-10-21', '07:30:00', 'masuk', 'gagal_lokasi', 0),
            ('Budi', '2026-10-21', '07:31:00', 'masuk', 'gagal_lokasi', 0),
            ('Hantu', '2026-10-20', '07:30:00', 'masuk', 'berhasil', 0);
    """)
    con.commit()
    con.close()


def test_migrasi_database_lama_dengan_data_ganda(tmp_path):
    path = tmp_path / "lama.db"
    buat_db_lama(path)
    eng = create_engine(f"sqlite:///{path}")
    A.metadata.create_all(eng)          # urutan sama dengan get_engine: tabel baru dibuat dulu, lalu kolom baru ditambahkan
    A._migrasi(eng)

    kolom = {k["name"] for k in inspect(eng).get_columns("absensi")}
    assert {"peserta_id", "foto_path", "alasan", "divisi"} <= kolom
    assert {"setuju_versi", "setuju_pada"} <= {k["name"] for k in inspect(eng).get_columns("peserta")}
    assert eng.indeks_unik_ok is False                       # data ganda menghalangi indeks, tapi app tidak crash
    assert A.cari_duplikat(eng) == [("Budi", "2026-10-20", "pulang", 2)]

    with eng.connect() as c:                                 # backfill: Budi terisi, "Hantu" (bukan peserta) tetap kosong
        isi = dict(c.execute(select(A.absensi_t.c.nama, A.absensi_t.c.peserta_id).distinct()).all())
    assert isi["Budi"] is not None and isi["Hantu"] is None

    with eng.connect() as c:                                 # snapshot divisi terisi untuk peserta yang masih terdaftar
        div = dict(c.execute(select(A.absensi_t.c.nama, A.absensi_t.c.divisi).distinct()).all())
    assert div["Budi"] == "Divisi Umum" and div["Hantu"] is None

    dihapus = A.rapikan_duplikat(eng)
    assert dihapus == 1 and eng.indeks_unik_ok is True
    jejak = A.daftar_audit(eng)
    assert list(jejak["aksi"]) == ["absensi.rapikan_duplikat"] and '"dihapus": 1' in jejak.iloc[0]["detail"]
    assert A.cari_duplikat(eng) == []
    with eng.connect() as c:                                 # uang makan tidak dobel; gagal_lokasi tidak ikut disentuh
        uang = c.execute(select(func.sum(A.absensi_t.c.uang_makan))).scalar()
        gagal = c.execute(select(func.count()).select_from(A.absensi_t).where(A.absensi_t.c.status == "gagal_lokasi")).scalar()
    assert uang == 15000 and gagal == 2
    with pytest.raises(IntegrityError):                       # setelah dirapikan, absen ganda baru ditolak
        baris_absen(eng, tanggal="2026-10-20", jenis="pulang", jam="17:05:00", uang_makan=15000)


def test_migrasi_idempoten(tmp_path):
    eng = mesin_baru(tmp_path)
    A._migrasi(eng)
    A._migrasi(eng)
    assert eng.indeks_unik_ok is True
