"""Layar pembuka yang tampil LANGSUNG saat halaman dibuka, sebelum loading bawaan Streamlit.

Cara kerja: potongan HTML/CSS/JS disisipkan ke berkas index.html milik paket streamlit (static/index.html), yaitu halaman
kosong pertama yang dikirim server sebelum aplikasi dimuat. Karena itu tidak ada loading lain yang tampil lebih dulu.

Dijalankan dua cara:
  1. Saat build Docker:   RUN python splash_awal.py     (paling andal, karena saat build kita root dan paket bisa ditulis)
  2. Saat aplikasi mulai: app.py memanggil pasang()      (cadangan, untuk jalan tanpa Docker, mis. Streamlit Cloud)
Aman dijalankan berulang (ada penanda). Gagal menulis tidak membuat aplikasi error.
"""
import base64
import os
import re
import sys

PENANDA = "splash-awal"
_DIR = os.path.dirname(os.path.abspath(__file__))

_CSS = """<style id="splash-awal-css">
#splash-awal{position:fixed;inset:0;z-index:2147483647;background:linear-gradient(180deg,#fff 0%,#fff 62%,#eef5fb 100%);
display:flex;flex-direction:column;align-items:center;justify-content:center;transition:opacity .45s ease,visibility .45s;
font-family:Inter,system-ui,sans-serif}
#splash-awal.keluar{opacity:0;visibility:hidden;pointer-events:none}
#splash-awal .sa-logo{width:min(72vw,300px);aspect-ratio:560/151;background:url(__LOGO__) center/contain no-repeat;
margin-top:-16vh;animation:sa-denyut 1.4s ease-in-out infinite}
#splash-awal .sa-bar{width:140px;height:4px;border-radius:4px;margin-top:34px;background:#e4f0f8;overflow:hidden;position:relative}
#splash-awal .sa-bar::after{content:"";position:absolute;top:0;left:0;height:100%;width:60px;border-radius:4px;
background:linear-gradient(90deg,#8cc63f,#0a5a8f);animation:sa-gerak 1.1s ease-in-out infinite}
#splash-awal .sa-teks{position:absolute;left:0;right:0;bottom:calc(44px + env(safe-area-inset-bottom,0px));text-align:center;
font:600 12px/18px Inter,system-ui,sans-serif;letter-spacing:.3px;color:#5f7b92}
#splash-awal .sa-garis{position:absolute;left:0;right:0;bottom:0;height:6px;background:linear-gradient(90deg,#0a5a8f,#8cc63f)}
@keyframes sa-denyut{0%,100%{transform:scale(.96)}50%{transform:scale(1.04)}}
@keyframes sa-gerak{0%{left:-60px}100%{left:140px}}
</style>"""

_HTML = """<div id="splash-awal"><div class="sa-logo"></div><div class="sa-bar"></div>
<div class="sa-teks">Absensi Peserta PKL<br>Bank Sulselbar Syariah</div><div class="sa-garis"></div></div>
<script>
(function(){var el=document.getElementById('splash-awal');if(!el){return;}var t0=Date.now(),lihat=0;
function selesai(){el.classList.add('keluar');setTimeout(function(){if(el.parentNode){el.parentNode.removeChild(el);}},600);}
function cek(){var a=document.querySelector('[data-testid="stApp"]'),m=document.querySelector('[data-testid="stMainBlockContainer"]'),n=Date.now();
if(m&&!lihat){lihat=n;}
var st=a?a.getAttribute('data-test-script-state'):null;
if(m&&((st&&st!=='running'&&st!=='initial')||(n-lihat>2500))){setTimeout(selesai,Math.max(0,900-(n-t0)));return;}
if(n-t0>20000){selesai();return;}
setTimeout(cek,120);}
cek();})();
</script>"""


def _lokasi_index():
    import streamlit
    return os.path.join(os.path.dirname(streamlit.__file__), "static", "index.html")


def _logo_uri():
    with open(os.path.join(_DIR, "logo_splash.png"), "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def pasang(path=None):
    """Sisipkan layar pembuka ke index.html. Mengembalikan 'sudah' (sebelumnya sudah terpasang), 'baru' (baru dipasang
    sekarang, berlaku mulai pemuatan halaman berikutnya), atau 'gagal'."""
    try:
        path = path or _lokasi_index()
        with open(path, "r", encoding="utf-8") as f:
            html = f.read()
        if PENANDA in html:
            return "sudah"
        css = _CSS.replace("__LOGO__", _logo_uri())
        # CSS di akhir <head>, elemen + skrip tepat setelah <body>
        html2, n1 = re.subn(r"</head>", lambda m: css + "</head>", html, count=1, flags=re.I)
        html2, n2 = re.subn(r"<body[^>]*>", lambda m: m.group(0) + _HTML, html2, count=1, flags=re.I)
        if not (n1 and n2):
            return "gagal"
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(html2)
        os.replace(tmp, path)
        return "baru"
    except Exception:
        return "gagal"


if __name__ == "__main__":
    hasil = pasang(sys.argv[1] if len(sys.argv) > 1 else None)
    print("splash_awal:", hasil)
    sys.exit(0)      # jangan gagalkan build kalau tidak bisa dipasang; app.py punya cadangan
