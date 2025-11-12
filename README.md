# Kalender Kelas - Google Apps Script

Aplikasi kalender berbasis Google Apps Script untuk manajemen agenda kelas dengan fitur upload file lampiran.

## 📋 Fitur

- Tampilan kalender interaktif
- Tambah, edit, dan hapus agenda
- Upload file lampiran
- Integrasi dengan Google Spreadsheet sebagai database
- Akses real-time untuk semua pengguna

## 🚀 Instalasi

### 1. Download Source Code

1. Kunjungi repository: [https://github.com/dhiksdhiky/kalender_kelas](https://github.com/dhiksdhiky/kalender_kelas)
2. Download file berikut:
   - `kode.gs`
   - `webapp.html`
   - Template spreadsheet (bersihkan isinya sebelum digunakan)

### 2. Setup Google Apps Script

1. Buka [https://script.google.com/](https://script.google.com/)
2. Klik **New project**
3. Beri nama project (contoh: "Agenda 505")
4. Copy-paste isi `kode.gs` ke dalam `code.gs`, lalu ganti nama file menjadi `kode.gs`
5. Buat file HTML baru dengan klik tombol **(+)**
6. Copy-paste isi `WebApp.html` dan ganti nama file sesuai

### 3. Organisasi File di Google Drive

1. Di Google Drive, cari file script yang baru dibuat
2. Buat folder baru dan beri nama sesuai keinginan
3. Pindahkan file script ke folder tersebut
4. Pindahkan juga file spreadsheet ke folder yang sama
   
   > **Catatan:** Folder lampiran akan dibuat otomatis, tidak perlu dibuat manual

### 4. Buat Spreadsheet Database

1. Di folder yang sama, buat Google Spreadsheet baru (jangan upload file Excel)
2. Beri nama sesuai keinginan
3. Sesuaikan nama sheet dengan template yang sudah disediakan
4. Salin **Spreadsheet ID** dari URL

   **Contoh URL:**
   ```
   https://docs.google.com/spreadsheets/d/1Rv1x6eeY6DXiZKX-hOSTkH9LcsD07ZymeOV51CTt9gI/edit
   ```
   
   Copy bagian yang ditebalkan: **1Rv1x6eeY6DXiZKX-hOSTkH9LcsD07ZymeOV51CTt9gI**

## ⚙️ Konfigurasi

### 1. Update Spreadsheet ID di kode.gs

Buka file `kode.gs` dan ganti Spreadsheet ID dengan ID yang sudah disalin:

```javascript
var SPREADSHEET_ID = "PASTE_ID_SPREADSHEET_DISINI";
```

### 2. Update Link Spreadsheet di WebApp.html

Buka file `WebApp.html`, cari baris 366, dan ganti link spreadsheet:

```javascript
// Baris 366
var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/PASTE_ID_SPREADSHEET_DISINI/edit";
```

### 3. Deploy Web App

1. Klik **Deploy** > **New deployment**
2. Pilih **Web app**
3. Pada **Who has access**, pilih **Anyone** atau **Everyone**
4. Klik **Deploy**
5. Copy URL yang diberikan
6. Jika diminta authorization, klik **Allow** dan berikan izin yang diperlukan

## 🧪 Testing

Sebelum membagikan aplikasi ke pengguna lain, lakukan testing berikut:

### 1. Tampilan Awal

Pastikan aplikasi menampilkan kalender dengan benar

### 2. Test Fitur Utama

- ✅ **Tambah agenda** (termasuk upload file lampiran)
- ✅ **Edit agenda**
- ✅ **Hapus agenda**

### 3. Verifikasi Database

1. Klik tombol **"Buka Sheet"** di aplikasi
2. Cek apakah data berhasil masuk ke spreadsheet

### 4. Verifikasi Tampilan Detail

1. Klik salah satu agenda di kalender
2. Periksa deskripsi di bagian bawah kalender
3. Jika ada file lampiran, klik untuk memastikan terbuka di tab baru

## 📝 Catatan

- Pastikan menggunakan Google Spreadsheet, **bukan** file Excel yang diupload ke Drive
- Nama sheet harus sesuai dengan template yang disediakan
- Folder lampiran akan dibuat otomatis saat pertama kali upload file

## 🤝 Kontribusi

Jika menemukan bug atau ada pertanyaan, silakan buka issue di repository ini.

## 📄 Lisensi

Project ini dibuat untuk keperluan manajemen agenda kelas.

---

**Selamat menggunakan!** Jika ada yang kurang jelas atau butuh bantuan, jangan ragu untuk bertanya.
