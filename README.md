# Kalender Kelas - Google Apps Script

Aplikasi kalender buat manajemen agenda kelas, ada fitur upload file lampiran juga.

## Fitur

- Kalender interaktif
- Tambah, edit, hapus agenda
- Upload file lampiran
- Database pakai Google Spreadsheet
- Real-time, semua orang bisa akses

## Cara Install

### 1. Download Source Code

a. Di GitHub: [https://github.com/dhiksdhiky/kalender_kelas](https://github.com/dhiksdhiky/kalender_kelas), download:
   - `kode.gs` 
   - `webapp.html`
   - Template spreadsheetnya (bersihkan dulu isinya sebelum dipakai)

b. Buka [https://script.google.com/](https://script.google.com/) > **New project** > kasih nama projek, bebas (contoh: Agenda 505)

c. Copas isi `kode.gs` di `code.gs`, ganti nama jadi `kode.gs`

d. Buat file HTML, klik **(+)**, copas dari file `WebApp.html` (ganti nama juga)

e. Di Drive paling luar, cari file scriptnya, buat folder baru kasih nama, pindah file script ke folder baru (file spreadsheetnya juga). Folder lampiran nanti otomatis, ga perlu bikin manual.

### 2. Buat Spreadsheet

a. Buat spreadsheet di folder, beri nama bebas. **Wajib spreadsheet ya**, bukan upload Excel yang diupload di Drive. Klik new file aja di Drive. Nama sheetnya disesuaikan sama template.

b. Salin (Copy) **ID Spreadsheet** dari URL. ID ini adalah bagian panjang di URL, contoh:

   ```
   https://docs.google.com/spreadsheets/d/1Rv1x6eeY6DXiZKX-hOSTkH9LcsD07ZymeOV51CTt9gI/edit
   ```
   
   Copy text yang **tebel** (1Rv1x6eeY6DXiZKX-hOSTkH9LcsD07ZymeOV51CTt9gI)

### 3. Konfigurasi

a. Buka lagi `kode.gs`, ganti ID spreadsheet dari URL spreadsheet (tahap 2b)

   ```javascript
   var SPREADSHEET_ID = "PASTE_ID_SPREADSHEET_DISINI";
   ```

b. Buka `WebApp.html`, baris 366, ganti link spreadsheet sesuai spreadsheet database

   ```javascript
   // Baris 366
   var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/PASTE_ID_DISINI/edit";
   ```

c. **Deploy** > **New deployment** > **Who has access: Everyone** > **Deploy** > klik URL nya, selesai

d. Kalau ada minta authorization, klik yes yes, lanjut lanjut, allow allow aja

## Tes-tes Sebelum Share

### 1. Tampilan Awal

Kurang lebih tampilannya seperti kalender dengan daftar agenda

### 2. Tes Fitur

- Tes fitur **tambah** (termasuk upload file)
- Tes **edit**
- Tes **hapus**

### 3. Cek Database

- Cek di spreadsheet (tombol "Buka Sheet"), untuk cek apakah data masuk database

### 4. Cek Detail Agenda

- Kalau sudah masuk, klik salah satu agenda di kalender
- Cek di bagian deskripsi bagian bawah kalender, cek apakah sudah sesuai
- Termasuk cek kalau ada file lampiran (klik harus kebuka tab baru)

## Catatan

Kalau ada yang kurang atau bingung, ditanyain aja bang. Semoga bermanfaat! 🙏

---

Selamat pakai! ✨
