# Bimbel Kotagede — GitHub + Vercel

Aplikasi tetap berjalan di **Google Apps Script**. Vercel hanya jadi alamat/pintu masuk (alihan ke URL Web app). Kode Apps Script di `apps_script/` TIDAK diubah.

## Isi repo
```
apps_script/     cadangan Code.gs, Index.html, appsscript.json (tidak dipakai Vercel)
public/index.html   halaman cadangan
vercel.json      alihan ke Web app
```

## Langkah (±10 menit)

### 1. Ambil URL Web app
Apps Script → **Deploy > Manage deployments** → salin **Web app URL** (berakhiran `/exec`).

### 2. Isi `vercel.json`
Ganti `GANTI_DENGAN_URL_WEB_APP_ANDA` dengan URL tadi (tetap di dalam tanda kutip). Contoh:
`"destination": "https://script.google.com/macros/s/AKfy.../exec"`

### 3. Buat repo GitHub
1. github.com → **New repository** → nama `bimbel-kotagede` → pilih **Private** → Create.
2. **uploading an existing file** → seret SEMUA isi folder ini (`apps_script`, `public`, `vercel.json`, `README.md`, `.gitignore`) → **Commit changes**.

### 4. Deploy di Vercel
1. vercel.com → **Sign up with GitHub**.
2. **Add New > Project** → pilih repo `bimbel-kotagede` → **Import**.
3. **Framework Preset: Other**. Kosongkan Build Command. → **Deploy**.
4. Buka alamat `xxxx.vercel.app`. Harus beralih ke aplikasi.

### 5. Ganti URL web app nanti
Edit `vercel.json` langsung di GitHub (ikon pensil) → Commit. Vercel deploy ulang otomatis.

### Domain sendiri (opsional)
Vercel → project → **Settings > Domains** → tambah domain, ikuti petunjuk DNS.

## PERINGATAN
- Setelah dibuka, alamat browser berubah ke `script.google.com`. Ini alihan, bukan hosting `Index.html` di Vercel.
- Pengguna tetap wajib login Google dan terdaftar di tab `user_account`. Orang tua belum bisa memakai.
- Mengubah file di `apps_script/` di GitHub **tidak mengubah aplikasi**. Perubahan harus ditempel manual ke Apps Script, lalu **Deploy > Manage deployments > Edit > New version**.
- Jaga repo **Private**. Jangan unggah file Sheets/data siswa (`.gitignore` memblokir `.xlsx` dan `.csv`).
- Tidak bisa dibuat tampil di dalam halaman Vercel (iframe) atau lewat proxy tanpa mengubah Apps Script: Google memblokir pembingkaian, dan login pengguna tidak ikut terbawa.
- Menjalankan `Index.html` langsung dari Vercel (tampilan di Vercel, data di Sheets) butuh perubahan Apps Script: API JSON + login. Belum dibuat.
