# SIMADUN - Sistem Informasi Manajemen Arsip Dokumen Inspektorat Madiun

SIMADUN adalah aplikasi manajemen arsip digital yang dirancang khusus untuk Inspektorat Madiun. Aplikasi ini mengintegrasikan Google Sheets sebagai database dan Google Drive sebagai penyimpanan dokumen, serta dapat diinstal sebagai PWA (Progressive Web App).

## Fitur Utama
- **Dashboard Terintegrasi**: Rekapitulasi otomatis untuk Surat Masuk, Keluar, Undangan, SPT, dan Arsip khusus (DUMAS, LHP, MCSP KPK, PDTT, PKKN, Lainnya).
- **Manajemen Arsip**: Upload dokumen langsung ke Google Drive dengan kategori yang terorganisir.
- **PWA Ready**: Dapat diinstal di perangkat Android, iOS, dan Desktop seperti aplikasi native.
- **Keamanan**: Autentikasi berbasis Role dan penyimpanan password terenkripsi.

---

## Panduan Instalasi (Vercel & Google Cloud)

### 1. Persiapan Google Cloud (Backend)
1.  Buka [Google Cloud Console](https://console.cloud.google.com/).
2.  Buat Project Baru (contoh: `Simadun-Project`).
3.  **Enable API**: Aktifkan **Google Sheets API** dan **Google Drive API**.
4.  **Service Account**:
    - Pergi ke menu `IAM & Admin` > `Service Accounts`.
    - Klik `Create Service Account`, beri nama, dan simpan.
    - Klik pada akun yang baru dibuat, buka tab `Keys`, lalu `Add Key` > `Create New Key` (pilih format **JSON**).
    - Simpan file JSON tersebut. **Isi file ini akan digunakan sebagai `GOOGLE_SERVICE_ACCOUNT_KEY`**.
    - Salin alamat email Service Account (contoh: `asdf@project.iam.gserviceaccount.com`).

### 2. Persiapan Google Sheets & Drive (Database & Storage)
1.  **Google Sheets**:
    - Buat Google Spreadsheet baru.
    - Ambil **Sheet ID** dari URL (teks di antara `/d/` dan `/edit`).
    - Bagikan (Share) spreadsheet ini ke email Service Account tadi dengan akses **Editor**.
    - Pastikan Anda membuat tab / sheet baru dengan nama **PERSIS** seperti berikut (perhatikan spasi dan huruf besarnya):
      1. `Users`
      2. `Surat Masuk`
      3. `Surat Keluar`
      4. `Undangan`
      5. `SPT`
      6. `Arsip`
    - (Wajib membuat semua 6 sheet tersebut agar tidak terjadi `GaxiosError: Unable to parse range`).

2.  **Google Drive & Bypass Limitasi Akun Instansi (PENTING)**:
    Seringkali akun Google Workspace Pemerintah (`@madiunkab.go.id` dsb.) memberikan batas **kapasitas 0 Byte** untuk robot Service Account yang berakibat terjadinya *Error 403 (Service Accounts do not have storage quota)*. Untuk mengatasi hal tersebut dan sukses menyimpan file secara tidak terbatas:
    - **Opsi A (Wajib Coba Pertama - Drive Bersama)**:
        - Buka akun Google Anda dan di panel kiri pilih menu **Drive Bersama (Shared Drives)** (Bukan "Drive Saya").
        - Buat Folder di dalam Drive Bersama tersebut.
        - Salin **Folder ID** dari URL.
        - Akses folder tersebut lalu **Kelola Akses**, dan tambahkan Email Service Account Anda sebagai **Pengelola** atau **Kontributor/Editor**.
    - **Opsi B (Rekomendasi Alternatif - Gunakan Gmail Biasa)**:
        - Apabila Google Instansi Anda tidak memiliki opsi *Drive Bersama*, gunakan akun murni `@gmail.com`.
        - Buat Folder di sana dan salin Foldernya.
        - Share folder ke Email Service Account dengan Akses **Editor**. (Gmail gratis tidak melarang Service Account beroperasi secara standar).

### 3. Deploy ke Vercel
1.  Hubungkan repositori GitHub Anda ke [Vercel](https://vercel.com/).
2.  Pada bagian **Environment Variables**, tambahkan variabel berikut:
    - `GOOGLE_SERVICE_ACCOUNT_KEY`: Tempel seluruh isi file JSON Service Account yang diunduh tadi.
    - `GOOGLE_SHEET_ID`: ID Spreadsheet Anda.
    - `GOOGLE_DRIVE_FOLDER_ID`: ID Folder Drive Anda.
3.  Klik **Deploy**.

---

## Pengembangan Lokal
Jika ingin menjalankan aplikasi di komputer lokal:

1.  Clone repositori.
2.  Instal dependensi:
    ```bash
    npm install
    ```
3.  Buat file `.env.local` di root folder dan isi dengan variabel lingkungan di atas.
4.  Jalankan server pengembangan:
    ```bash
    npm run dev
    ```
5.  Buka `http://localhost:3000` di browser.

---

## Struktur Folder
- `pages/api/[action].js`: Backend logic menggunakan Next.js API Routes.
- `public/app/index.html`: Frontend Single Page Application (SPA).
- `public/manifest.json`: Konfigurasi PWA.
- `public/assets/`: Icon dan aset gambar.

## Lisensi
Aplikasi ini dikembangkan untuk penggunaan internal Inspektorat Madiun.
