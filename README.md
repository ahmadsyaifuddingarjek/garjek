# GARJEK Driver Registration

Sistem pendaftaran mitra driver GARJEK dengan:
- Frontend HTML responsive untuk GitHub Pages.
- Backend Google Apps Script.
- Google Spreadsheet sebagai database.
- Google Drive untuk dokumen.
- Nomor pendaftaran otomatis.
- Cek status menggunakan NIK (data publik dimasking).
- Dashboard admin + statistik + verifikasi.
- Login admin berbasis token sesi.

## 1. Buat backend Google Apps Script

1. Buka `https://script.google.com/`.
2. Buat project baru.
3. Masukkan seluruh isi `Code.gs`.
4. Simpan.
5. Jalankan fungsi `setupSystem()` satu kali dari editor.
6. Berikan izin yang diminta Google.
7. Lihat hasil eksekusi/log. Setup akan membuat:
   - Spreadsheet database.
   - Folder Google Drive.
   - Sheet `Drivers`, `Documents`, `Settings`, `Admin`.
   - Username admin.
   - Password awal yang dibuat otomatis.
8. Simpan password awal di tempat aman. Password tidak ditampilkan di website.

## 2. Deploy Web App

Di Apps Script:
- Deploy → New deployment
- Type: Web app
- Execute as: Me
- Who has access: Anyone (atau sesuai kebijakan organisasi)
- Deploy
- Salin URL `/exec`.

Jangan menaruh username/password admin di HTML atau GitHub.

## 3. Hubungkan frontend

Buka `index.html` dan ubah:

`const API='PASTE_APPS_SCRIPT_WEB_APP_URL_DI_SINI';`

menjadi URL Web App Apps Script, contoh:

`const API='https://script.google.com/macros/s/XXXXXXXX/exec';`

## 4. GitHub Pages

1. Buat repository baru, misalnya `garjek-driver`.
2. Upload `index.html`.
3. Upload `README.md` jika diinginkan.
4. Repository → Settings → Pages.
5. Source: Deploy from a branch.
6. Pilih branch `main` dan folder `/root`.
7. Save.
8. Buka URL GitHub Pages.

## 5. Struktur

- `Code.gs`: backend/database/upload/auth.
- `index.html`: pendaftaran, cek status, dashboard admin.

## Catatan keamanan

1. Data KTP/SIM/STNK adalah data sensitif. Pastikan akun Google pemilik backend menggunakan keamanan kuat dan 2FA.
2. Folder Drive dibuat private secara default. Admin harus memiliki akses ke akun Google pemilik Drive.
3. Jangan membuat folder dokumen menjadi "Anyone with the link" kecuali benar-benar diperlukan.
4. GitHub repository hanya berisi frontend. Jangan commit credential.
5. Password admin disimpan sebagai hash SHA-256 di Script Properties, bukan plaintext.
6. Untuk produksi skala besar, pertimbangkan backend/database yang lebih khusus karena Apps Script memiliki batas eksekusi, ukuran request, dan kuota.
7. Upload saat ini dibatasi 5 MB per file dan enam file per pendaftaran.
8. "Realtime" pada sistem ini menggunakan pembacaan data terbaru ketika dashboard dibuka/di-refresh; bukan WebSocket realtime. Tombol Refresh disediakan agar admin dapat mengambil data terbaru.
