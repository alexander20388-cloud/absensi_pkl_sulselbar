"""Tes token sesi login (sesi.py)."""
import sesi

KUNCI = "kunci-rahasia"
SIDIK = sesi.sidik_pin("scrypt$abc")


def _sidik(nama):
    return SIDIK if nama == "Budi" else None


def test_token_sah():
    t = sesi.buat_token("Budi", KUNCI, SIDIK)
    assert sesi.baca_token(t, KUNCI, _sidik) == "Budi"


def test_kedaluwarsa():
    t = sesi.buat_token("Budi", KUNCI, SIDIK, sekarang=1_000_000)
    assert sesi.baca_token(t, KUNCI, _sidik, sekarang=1_000_000 + 29 * 86400) == "Budi"
    assert sesi.baca_token(t, KUNCI, _sidik, sekarang=1_000_000 + 31 * 86400) is None


def test_kunci_salah_atau_dipalsukan():
    t = sesi.buat_token("Budi", KUNCI, SIDIK)
    assert sesi.baca_token(t, "kunci-lain", _sidik) is None
    muatan, tanda = t.split(".")
    palsu = sesi.buat_token("Ani", KUNCI, SIDIK).split(".")[0] + "." + tanda   # nama diganti, tanda tangan lama
    assert sesi.baca_token(palsu, KUNCI, lambda n: SIDIK) is None


def test_pin_diganti_membatalkan_token():
    t = sesi.buat_token("Budi", KUNCI, SIDIK)
    assert sesi.baca_token(t, KUNCI, lambda n: sesi.sidik_pin("scrypt$baru")) is None


def test_peserta_dihapus():
    t = sesi.buat_token("Budi", KUNCI, SIDIK)
    assert sesi.baca_token(t, KUNCI, lambda n: None) is None


def test_sampah_tidak_melempar_galat():
    for x in (None, "", "abc", "a.b", "a.b.c", 123, "x" * 5000, "!!!.???"):
        assert sesi.baca_token(x, KUNCI, _sidik) is None


def test_js_aman():
    assert "pkl_sesi" in sesi.JS_BACA and "try" in sesi.JS_BACA
    assert sesi.js_simpan("abc.def").startswith("<script>")
    assert "removeItem" in sesi.js_hapus()
