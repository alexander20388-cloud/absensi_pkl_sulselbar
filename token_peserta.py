"""Token masuk untuk aplikasi native peserta (tanpa dependensi tambahan).

Bentuk: base64url(JSON{n: nama, v: versi_pin, e: kedaluwarsa}) + "." + base64url(HMAC-SHA256).
- Stateless: server tidak menyimpan sesi.
- `v` diturunkan dari hash PIN, jadi reset PIN oleh admin otomatis mencabut semua token lama peserta itu.
- Perbandingan tanda tangan memakai hmac.compare_digest (waktu konstan).
"""
import base64
import hashlib
import hmac
import json
import time

BERLAKU_DETIK = 14 * 24 * 3600          # 14 hari; setelah itu peserta masuk ulang dengan PIN


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode()


def _unb64(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def versi_pin(pin_hash: str) -> str:
    """Sidik jari singkat dari hash PIN (bukan hash itu sendiri), dipakai untuk mencabut token saat PIN berubah."""
    return hashlib.sha256(("pin|" + (pin_hash or "")).encode()).hexdigest()[:12]


def _tanda(isi: str, rahasia: str) -> str:
    return _b64(hmac.new(rahasia.encode(), isi.encode(), hashlib.sha256).digest())


def buat(nama: str, pin_hash: str, rahasia: str, berlaku_detik: int = BERLAKU_DETIK, sekarang: float | None = None) -> str:
    t = time.time() if sekarang is None else sekarang
    isi = _b64(json.dumps({"n": nama, "v": versi_pin(pin_hash), "e": int(t + berlaku_detik)},
                          separators=(",", ":"), ensure_ascii=False).encode())
    return f"{isi}.{_tanda(isi, rahasia)}"


def periksa(token: str, rahasia: str, sekarang: float | None = None):
    """Kembalikan {'n','v','e'} bila token sah dan belum kedaluwarsa; selain itu None. Tidak pernah melempar."""
    try:
        isi, tanda = token.split(".", 1)
        if not hmac.compare_digest(tanda.encode(), _tanda(isi, rahasia).encode()):
            return None
        data = json.loads(_unb64(isi))
        if not isinstance(data, dict) or not isinstance(data.get("n"), str):
            return None
        t = time.time() if sekarang is None else sekarang
        if int(data.get("e", 0)) < t:
            return None
        return data
    except Exception:
        return None
