"""Membuka aplikasi Streamlit dengan browser sungguhan supaya dianggap ada pengunjung (ping HTTP biasa tidak cukup).
Bila aplikasi sedang tidur, tombol "Yes, get this app back up!" ditekan."""
import os
import sys

from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

URL = os.environ.get("APP_URL", "https://absensipklsulselbar.streamlit.app/")


def coba_sekali(pg):
    pg.goto(URL, wait_until="domcontentloaded", timeout=90_000)
    try:                                           # aplikasi tidur: bangunkan
        tombol = pg.get_by_role("button", name="Yes, get this app back up!")
        tombol.wait_for(timeout=15_000)
        tombol.click()
        print("Aplikasi sedang tidur: tombol bangunkan ditekan.")
    except PWTimeout:
        print("Aplikasi tidak menampilkan tombol bangunkan (sudah hidup).")
    # tunggu aplikasi benar-benar tampil (di streamlit.app ia berada di dalam iframe)
    app = pg.frame_locator('iframe[title="streamlitApp"]').locator('[data-testid="stApp"]')
    app.wait_for(timeout=180_000)
    pg.wait_for_timeout(20_000)                    # tahan sebentar supaya sambungan websocket terbentuk
    print("Aplikasi tampil dan sambungan terbentuk.")


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        pg = browser.new_page()
        for i in range(1, 4):
            try:
                coba_sekali(pg)
                browser.close()
                return 0
            except Exception as e:                 # noqa: BLE001
                print(f"Percobaan {i} gagal: {e}")
        browser.close()
    return 1


if __name__ == "__main__":
    sys.exit(main())
