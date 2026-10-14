# Social Scraper — versi lokal (tanpa Google Colab)

Port dari notebook `Social_Scraper_ADVANCED_v58_EN.ipynb` menjadi **aplikasi web lokal**:
backend Python (FastAPI) + UI di browser. Semua fitur notebook tetap ada — mode keyword /
BY LINK / BY PROFILE / MIXED, filter waktu WIB, strict keyword, filter bahasa, negative keyword,
komentar, enrichment profil, retweet, impor data, sentimen & emosi (AI / leksikon), dan output
Excel layout Sample_Crawling (+ sheet Rekap & Codebook).

Bedanya dengan Colab:
- **Tidak ada timeout / disconnect** — jalan di laptop sendiri selama jendela server terbuka.
- Hasil tersimpan di folder `outputs/` (pengganti Google Drive), termasuk autosave `checkpoint_partial.csv`.
- Form **diingat otomatis** di browser — tidak perlu isi ulang setiap kali.
- Tombol **■ Stop** menghentikan request berikutnya; data yang sudah terkumpul tetap disimpan ke Excel.
- Refresh halaman saat scraping berjalan aman — progress & log tetap terpantau.

## Cara pakai

1. Install **Python 3.10+** (Windows: https://www.python.org/downloads/ — centang *Add Python to PATH*).
2. Download / clone folder ini.
3. Jalankan:
   - **Windows**: klik dua kali `start-windows.bat`
   - **Mac / Linux**: `./start-mac-linux.sh`

   Pertama kali akan memasang dependensi (±1–2 menit). Browser terbuka otomatis di
   http://localhost:8765
4. Isi API key di bagian **🔐 API key** → **Simpan** (tersimpan ke file `.env`, cukup sekali).
5. Pilih skenario / isi form → **Scrape**. File Excel bisa diunduh dari halaman, atau ambil di folder `outputs/`.

Cara manual (tanpa script start):

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt      # Windows: .venv\Scripts\pip ...
.venv/bin/python run.py
```

### Mode analisis "AI model" (opsional)

Mode AI butuh `transformers` + `torch` (unduhan model pertama ±2 GB). Pasang sekali:

```bash
.venv/bin/pip install -r requirements-ai.txt
```

Tanpa itu, pilih **Leksikon** (instan). Jika mode AI dipilih tapi paket belum ada, aplikasi
mencoba memasangnya otomatis seperti di Colab; bila gagal, otomatis jatuh ke leksikon.
Di laptop tanpa GPU, mode AI untuk ratusan baris bisa makan beberapa menit.

## Struktur

| File | Isi |
| --- | --- |
| `backend/core.py` | Logika scraping/parsing/filter/sentimen/Excel dari notebook (hampir apa adanya) |
| `backend/runner.py` | Alur `on_click()` notebook: menerima nilai form dari UI, bukan ipywidgets |
| `backend/app.py` | API lokal: jalankan job di background, log & progress, stop, unduh file, upload impor |
| `backend/static/index.html` | Antarmuka web |
| `run.py` | Launcher (port 8765, hanya bisa diakses dari laptop ini) |

Pengaturan lewat `.env` (lihat `.env.example`): `SCRAPECREATORS_API_KEY`, `TWITTER_API_KEY`,
`OUTPUT_DIR` (opsional). Port bisa diganti dengan env `PORT`.

> Catatan: password gate di notebook tidak dipakai lagi karena aplikasi hanya bisa diakses dari
> laptop sendiri (`127.0.0.1`). API key tidak lagi ditulis di kode — simpan di `.env` (tidak ikut ter-commit).
