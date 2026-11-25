# FortMess — Sistem Manajemen Kas & Iuran Mes Putra Fortress

Aplikasi manajemen kas & iuran mes **100% gratis** menggunakan **Google Sheets** sebagai database dan **Google Apps Script (GAS)** sebagai backend & web app.

## Cara Setup (5 Menit)
1. Buat **Google Spreadsheet** baru di Google Drive (misal: `Database Kas FortMess`).
2. Buka menu **Extensions > Apps Script**.
3. Copy-paste isi file [`Code.gs`](./Code.gs) ke editor Apps Script.
4. Buat file HTML baru di Apps Script bernama `index` (`index.html`), lalu copy-paste isi [`index.html`](./index.html).
5. Jalankan fungsi **`setupDatabase()`** sekali untuk membuat 4 tab (`Master_Penghuni`, `Matriks_Iuran`, `Buku_Kas`, `Log_Audit`) beserta data awal.
6. Jalankan fungsi **`setAdminPin('123456')`** (ganti dengan PIN rahasia Bendahara) untuk menyimpan hash PIN di server (`PropertiesService`).
7. Klik **Deploy > New deployment > Web app**:
   * **Execute as:** `Me` (akun pemilik/Bendahara)
   * **Who has access:** `Anyone` (agar penghuni bisa buka dashboard *read-only*).
8. **Proteksi Google Sheet:**
   * Klik **Share** pada Google Sheet -> set **Anyone with the link = Viewer**.
   * Hanya email Bendahara yang diberi akses **Editor**, dan aktifkan **Protect sheet** pada tab `Log_Audit`.
