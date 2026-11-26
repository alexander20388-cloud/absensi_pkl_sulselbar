"""Sesi login yang bertahan (hanya untuk peserta di aplikasi PKL, ?mode=pkl).

Setelah PIN benar, server membuat token bertanda tangan (HMAC) yang disimpan di localStorage perangkat. Saat aplikasi dibuka lagi,
token dibaca dan diperiksa; kalau sah, peserta langsung masuk. Token dihapus saat peserta menekan Keluar.

Modul ini sengaja TIDAK mengimpor streamlit supaya mudah diuji. Penyambungannya ada di app.py.
"""
import base64
import hashlib
import hmac
import json
import time

KUNCI_BROWSER = "pkl_sesi"       # nama item di localStorage
MASA_BERLAKU_HARI = 30           # token kedaluwarsa setelah ini sejak terakhir dipakai (diperbarui tiap aplikasi dibuka)


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode().rstrip("=")


def _b64_balik(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def sidik_pin(pin_hash) -> str:
    """Sidik jari singkat dari hash PIN. Token ikut tidak sah bila PIN diganti/direset admin."""
    return hashlib.sha256(("sesi|" + (pin_hash or "")).encode()).hexdigest()[:16]


def _tanda(kunci: str, muatan: str, sidik: str) -> str:
    return _b64(hmac.new(str(kunci).encode(), f"{muatan}|{sidik}".encode(), hashlib.sha256).digest())


def buat_token(nama: str, kunci: str, sidik: str, sekarang: float | None = None, hari: int = MASA_BERLAKU_HARI) -> str:
    t = time.time() if sekarang is None else sekarang
    muatan = _b64(json.dumps({"n": nama, "e": int(t + hari * 86400)}, separators=(",", ":")).encode())
    return f"{muatan}.{_tanda(kunci, muatan, sidik)}"


def baca_token(token, kunci: str, sidik_dari_nama, sekarang: float | None = None):
    """Kembalikan nama peserta bila token sah dan belum kedaluwarsa, selain itu None.
    sidik_dari_nama(nama) -> sidik_pin() peserta itu, atau None bila peserta sudah tidak ada."""
    try:
        if not token or not isinstance(token, str) or token.count(".") != 1 or len(token) > 2000:
            return None
        muatan, tanda = token.split(".")
        data = json.loads(_b64_balik(muatan))
        nama, kedaluwarsa = data["n"], int(data["e"])
        if not isinstance(nama, str) or not nama:
            return None
        sidik = sidik_dari_nama(nama)
        if sidik is None:
            return None
        if not hmac.compare_digest(tanda, _tanda(kunci, muatan, sidik)):
            return None
        if (time.time() if sekarang is None else sekarang) > kedaluwarsa:
            return None
        return nama
    except Exception:
        return None


# --- potongan JavaScript (semuanya dibungkus try/catch: WebView tertentu melempar galat saat localStorage diblokir) ---
JS_BACA = f"(function(){{try{{return localStorage.getItem({json.dumps(KUNCI_BROWSER)})||''}}catch(e){{return ''}}}})()"


_COOKIE_ATRIBUT = '"; path=/; max-age=" + %d + "; SameSite=Lax" + (location.protocol === "https:" ? "; Secure" : "")'


def js_simpan(token: str) -> str:
    """Simpan token di COOKIE (server membacanya langsung dari permintaan pertama: tanpa menunggu komponen JavaScript) dan juga di localStorage."""
    nama = json.dumps(KUNCI_BROWSER)
    return ("<script>try{document.cookie=" + nama + '+"="+encodeURIComponent(' + json.dumps(token) + ")+"
            + (_COOKIE_ATRIBUT % (MASA_BERLAKU_HARI * 86400)) + "}catch(e){}"
            f"try{{localStorage.setItem({nama},{json.dumps(token)})}}catch(e){{}}</script>")


def js_hapus() -> str:
    nama = json.dumps(KUNCI_BROWSER)
    return ("<script>try{document.cookie=" + nama + '+"=; path=/; max-age=0; SameSite=Lax"}catch(e){}'
            f"try{{localStorage.removeItem({nama})}}catch(e){{}}</script>")
