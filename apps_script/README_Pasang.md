# Pasang Aplikasi Bimbel (Google Apps Script) — ±15 menit

## Yang dibutuhkan
- Akun Google (Gmail biasa cukup, gratis).
- File `Bimbel_Kotagede_Fase1_v0.1.xlsx`.
- 3 file di folder ini: `Code.gs`, `Index.html`, `appsscript.json`.

## Langkah
1. **Upload xlsx ke Google Drive** → klik kanan → Buka dengan → Google Sheets → menu **File > Save as Google Sheets**.
2. Di Sheets: **File > Settings > Time zone** → pilih `(GMT+07:00) Jakarta`.
3. **Extensions > Apps Script**.
4. Di kiri ada `Code.gs`. Hapus isinya, tempel isi `Code.gs` dari folder ini.
5. Klik **+** (kiri) > **HTML** > beri nama `Index` (tepat, huruf besar I). Hapus isi, tempel isi `Index.html`.
6. Klik ikon roda gigi (**Project Settings**) → centang **Show "appsscript.json" manifest file**. Kembali ke Editor, buka `appsscript.json`, hapus isi, tempel isi file `appsscript.json`.
7. Klik **Save** (ikon disket).
8. Kembali ke Sheets, **muat ulang halaman**. Muncul menu **Bimbel > Setup awal**. Klik, lalu izinkan akses (Advanced > Go to … > Allow). Ini membuat tab `user_account` dan `audit_log`, dan mendaftarkan email Anda sebagai pemilik (PEM).
9. Di Apps Script: **Deploy > New deployment > ikon roda gigi > Web app**.
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone with Google account**
   - Klik Deploy, salin **URL Web app**. Buka URL itu → aplikasi jalan.
10. Tambah staf: isi tab `user_account` (kolom `user_id`, `email`, `role`, `tutor_id`, `status`). Lalu **Share** Sheets ke email mereka sebagai **Editor**.

## Peran
| role | Bisa buka |
|---|---|
| PEM (pemilik) | semua |
| AL (akademik) | Dasbor, Pantau, Presensi, Siswa (lihat) |
| ADM (admin) | Dasbor, Pantau, Siswa, Bayar |
| TUT (tutor) | Presensi, Pantau (isi `tutor_id` agar hanya kelasnya, mis. `T001`) |

## Pemakaian harian
- **Presensi**: pilih kelas → tanggal → set status tiap siswa → isi exit ticket/PR → Simpan.
- **Siswa**: daftar siswa baru (ortu dicek lewat nomor WhatsApp supaya tidak dobel).
- **Bayar**: awal bulan klik "Buat tagihan", lalu catat tiap pembayaran.
- **Dasbor / Pantau**: otomatis dari rumus di Sheets.

## PERINGATAN
- **Peran hanya membatasi tampilan aplikasi.** Siapa pun yang diberi akses Editor ke Sheets bisa melihat/mengubah SEMUA data langsung. Berbagi Sheets = berbagi data anak. Bagikan seperlunya.
- Jangan edit tab `invoice` kolom abu-abu (rumus). Aplikasi menulis ke baris kosong pertama; jangan hapus baris rumus.
- Jangan ubah nama tab dan nama kolom.
- Batas rumus ±200 baris/tab (Rekap_Siswa ±100 siswa). Lewat itu salin rumus ke bawah.
- Kode di sini belum diuji langsung di akun Google Anda (hanya uji logika + sintaks). Coba dulu dengan data dummy dan salin file sebagai cadangan sebelum data asli.
- Kuota gratis Apps Script cukup untuk puluhan siswa. Jika lambat atau >50 siswa: waktunya Fase 2.
- Data pribadi anak: tinjau aturan PDP sebelum memakai data nyata.
- Setelah ubah kode: **Deploy > Manage deployments > Edit > New version** agar URL yang sama terbarui.

## Belum ada
Rapor PDF, WhatsApp/QRIS otomatis, portal ortu, bank soal, kompetensi, komplain, consent.
