package id.sulselbar.presensipkl;

import android.Manifest;
import android.app.Activity;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.view.ViewGroup;
import android.view.Window;
import android.webkit.CookieManager;
import android.webkit.GeolocationPermissions;
import android.webkit.PermissionRequest;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.FrameLayout;
import android.widget.ProgressBar;

/**
 * Pembungkus WebView untuk Presensi PKL.
 * - Membuka halaman aplikasi Streamlit LANGSUNG (/~/+/), tanpa halaman pembungkus Streamlit Cloud: lebih cepat, dan cookie/localStorage
 *   dihitung milik aplikasi sendiri (bukan pihak ketiga di dalam bingkai), jadi login otomatis bekerja.
 * - Cookie & penyimpanan lokal aktif dan disimpan ke disk setiap aplikasi berhenti.
 * - Izin kamera (foto absen) dan lokasi (GPS) diteruskan ke halaman.
 */
public class MainActivity extends Activity {
    // GANTI di sini bila alamat aplikasi berubah. "/~/+/" = halaman aplikasi tanpa pembungkus Streamlit Cloud.
    private static final String ALAMAT_APLIKASI = "https://absensipklsulselbar.streamlit.app/~/+/?mode=pkl";
    private static final int KODE_IZIN = 101;

    private WebView web;
    private ProgressBar bar;
    private PermissionRequest izinKameraTertunda;
    private String lokasiOriginTertunda;
    private GeolocationPermissions.Callback lokasiCallbackTertunda;

    @Override
    protected void onCreate(Bundle b) {
        super.onCreate(b);
        Window w = getWindow();
        w.setStatusBarColor(Color.parseColor("#0A5A8F"));
        w.setNavigationBarColor(Color.parseColor("#EEF2FB"));

        FrameLayout root = new FrameLayout(this);
        root.setBackgroundColor(Color.parseColor("#EEF2FB"));
        web = new WebView(this);
        web.setBackgroundColor(Color.parseColor("#EEF2FB"));
        root.addView(web, new FrameLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.MATCH_PARENT));
        bar = new ProgressBar(this, null, android.R.attr.progressBarStyleHorizontal);
        bar.setMax(100);
        FrameLayout.LayoutParams lp = new FrameLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 8);
        root.addView(bar, lp);
        setContentView(root);

        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setMediaPlaybackRequiresUserGesture(false);
        s.setCacheMode(WebSettings.LOAD_DEFAULT);
        s.setUserAgentString(s.getUserAgentString() + " PresensiPKLApp/2.0");

        CookieManager cm = CookieManager.getInstance();
        cm.setAcceptCookie(true);
        cm.setAcceptThirdPartyCookies(web, true);

        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView v, WebResourceRequest r) {
                Uri u = r.getUrl();
                String skema = u.getScheme() == null ? "" : u.getScheme();
                if (skema.equals("app")) {                      // tombol "Coba lagi" di halaman galat
                    v.loadUrl(ALAMAT_APLIKASI);
                    return true;
                }
                String host = u.getHost() == null ? "" : u.getHost();
                if (host.endsWith(".streamlit.app") || host.endsWith(".streamlit.io")) {
                    return false;                               // tetap di dalam aplikasi
                }
                try {                                           // tautan luar (mis. unduhan) dibuka di browser
                    startActivity(new Intent(Intent.ACTION_VIEW, u));
                } catch (Exception e) { /* tidak ada aplikasi pembuka */ }
                return true;
            }

            @Override
            public void onReceivedError(WebView v, WebResourceRequest r, WebResourceError e) {
                if (r.isForMainFrame()) {
                    v.loadDataWithBaseURL(null, HALAMAN_GALAT, "text/html", "UTF-8", null);
                }
            }
        });

        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onProgressChanged(WebView v, int p) {
                bar.setProgress(p);
                bar.setVisibility(p >= 100 ? View.GONE : View.VISIBLE);
            }

            @Override
            public void onPermissionRequest(final PermissionRequest req) {      // kamera untuk foto absen
                runOnUiThread(new Runnable() {
                    @Override public void run() {
                        if (checkSelfPermission(Manifest.permission.CAMERA) == PackageManager.PERMISSION_GRANTED) {
                            req.grant(req.getResources());
                        } else {
                            izinKameraTertunda = req;
                            requestPermissions(new String[]{Manifest.permission.CAMERA}, KODE_IZIN);
                        }
                    }
                });
            }

            @Override
            public void onGeolocationPermissionsShowPrompt(String origin, GeolocationPermissions.Callback cb) {   // GPS
                if (checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) == PackageManager.PERMISSION_GRANTED) {
                    cb.invoke(origin, true, false);
                } else {
                    lokasiOriginTertunda = origin;
                    lokasiCallbackTertunda = cb;
                    requestPermissions(new String[]{Manifest.permission.ACCESS_FINE_LOCATION, Manifest.permission.ACCESS_COARSE_LOCATION}, KODE_IZIN);
                }
            }
        });

        // Minta izin di awal supaya dialognya tidak muncul di tengah absen.
        if (checkSelfPermission(Manifest.permission.CAMERA) != PackageManager.PERMISSION_GRANTED
                || checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) != PackageManager.PERMISSION_GRANTED) {
            requestPermissions(new String[]{Manifest.permission.CAMERA, Manifest.permission.ACCESS_FINE_LOCATION,
                    Manifest.permission.ACCESS_COARSE_LOCATION}, KODE_IZIN);
        }

        if (b != null) {
            web.restoreState(b);
        } else {
            web.loadUrl(ALAMAT_APLIKASI);
        }
    }

    @Override
    public void onRequestPermissionsResult(int kode, String[] izin, int[] hasil) {
        super.onRequestPermissionsResult(kode, izin, hasil);
        if (izinKameraTertunda != null) {
            if (checkSelfPermission(Manifest.permission.CAMERA) == PackageManager.PERMISSION_GRANTED) {
                izinKameraTertunda.grant(izinKameraTertunda.getResources());
            } else {
                izinKameraTertunda.deny();
            }
            izinKameraTertunda = null;
        }
        if (lokasiCallbackTertunda != null) {
            boolean ok = checkSelfPermission(Manifest.permission.ACCESS_FINE_LOCATION) == PackageManager.PERMISSION_GRANTED;
            lokasiCallbackTertunda.invoke(lokasiOriginTertunda, ok, false);
            lokasiCallbackTertunda = null;
        }
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        web.saveState(out);
    }

    @Override
    protected void onPause() {
        super.onPause();
        CookieManager.getInstance().flush();        // simpan cookie ke disk, supaya login bertahan walau aplikasi ditutup paksa
        web.onPause();
    }

    @Override
    protected void onResume() {
        super.onResume();
        web.onResume();
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) {
            web.goBack();
        } else {
            super.onBackPressed();
        }
    }

    private static final String HALAMAN_GALAT =
            "<html><head><meta name='viewport' content='width=device-width, initial-scale=1'>"
            + "<style>body{font-family:sans-serif;background:#eef2fb;color:#1e2a5a;text-align:center;padding:64px 24px}"
            + "h2{margin:0 0 8px}p{color:#4a5788}a{display:inline-block;margin-top:18px;background:#0a5a8f;color:#fff;"
            + "text-decoration:none;padding:14px 28px;border-radius:14px;font-weight:bold}</style></head><body>"
            + "<h2>Tidak ada koneksi</h2><p>Aplikasi tidak bisa membuka halaman presensi.<br>Periksa internet kamu, lalu coba lagi.</p>"
            + "<a href='app://ulang'>Coba lagi</a></body></html>";
}
