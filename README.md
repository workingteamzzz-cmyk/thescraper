# Social Scraper (local version)

Aplikasi web lokal hasil port dari notebook `Social_Scraper_ADVANCED_v58_EN.ipynb`.
Berjalan di laptop sendiri (tanpa Google Colab), tampilan English, alur Step 1–6 sama seperti form notebook.

## Instalasi (sekali saja)

1. Install **Python 3.10+** — Windows: https://www.python.org/downloads/ (centang **Add Python to PATH**).
2. Download ZIP repo ini lalu ekstrak.
3. Jalankan:
   - **Windows:** klik dua kali `start-windows.bat`
   - **Mac:** klik dua kali `start-mac.command` (lihat catatan Mac di bawah)
   - **Linux:** `./start-mac-linux.sh`
4. Browser terbuka otomatis di http://localhost:8765. Jendela hitam (server) harus tetap terbuka selama dipakai.
5. Step 1 → isi API key → **Save keys**. Key disimpan di file `.env` di laptop (tidak ikut ke GitHub).

### Analisis AI (sentiment & emotion)

Sudah termasuk: saat pertama kali dijalankan, start script otomatis memasang paket AI dan mengunduh
modelnya (±2–3 GB, 5–15 menit tergantung internet). Setelah itu tidak diulang lagi.
Kalau gagal (misalnya internet putus), aplikasi tetap jalan memakai Lexicon; ulangi kapan saja dengan
`install-ai-windows.bat` (Windows) atau `install-ai-mac.command` (Mac).

### Catatan Mac

- Install Python dari https://www.python.org/downloads/macos/ (file .pkg), jangan pakai `python3` bawaan.
- Pertama kali klik dua kali `start-mac.command`, macOS bisa menolak ("cannot be opened because it is from an unidentified developer").
  Solusi: **klik kanan → Open → Open**. Cukup sekali.
- Kalau muncul "permission denied", buka Terminal lalu jalankan:
  `cd ~/Downloads/SocialScraper && chmod +x *.command *.sh && ./start-mac.command`

## Anti-timeout

- Scraping berjalan di server lokal, bukan di tab browser: tab boleh ditutup/di-refresh, proses tetap jalan.
- Koneksi putus / server penyedia error sementara (502/503/504, timeout) otomatis dicoba ulang hingga 3x.
- Laptop dicegah masuk *sleep* selama scraping berjalan (Windows & Mac).
- Autosave `outputs/checkpoint_partial.csv` tiap ±50 baris; tombol **Stop** tetap menyimpan data yang sudah terkumpul.

## Output

Excel/CSV di folder `outputs/` (bisa juga diunduh dari halaman). Semua kolom selalu ada:
No, Sentiment, Username, User ID, Follower, Text, Date, Time, Likes, Reply, Shares, Url, Emotion, User Image,
Platform, Sub Type, Views Count, Gender, Age Category, Content Location, Religion, Ethnic, Phsychography,
Parent ID, ID, User ID Parent, Query — ditambah sheet **Summary** dan **Codebook**.
Kolom Gender s/d Phsychography sengaja kosong untuk coding manual.

## Struktur

| File | Isi |
| --- | --- |
| `backend/core.py` | Logika scraping, filter, sentimen/emosi, layout Excel (dari notebook) |
| `backend/runner.py` | Alur tombol Scrape (`on_click` notebook) |
| `backend/app.py` | Server lokal: job di background, log/progress, stop, unduh, upload impor |
| `backend/static/index.html` | Tampilan |
| `run.py` | Launcher (port 8765, hanya bisa diakses dari laptop ini) |
