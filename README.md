# spx-eha-project

# Shopee Data Automation & Sync (Google Apps Script)

Script ini digunakan untuk mengotomatiskan alur kerja data operasional pada Google Sheets. Menghubungkan data dari sheet **Backlog** dan **Datasuite** untuk menghasilkan laporan terpusat di sheet **Same Data**.

## 🚀 Fitur Utama

1. **Auto Timestamp (Deteksi Input)**
   - **Sheet Backlog:** Mengetik/paste di Kolom A akan otomatis mengisi tanggal & jam (format `dd/MM/yyyy HH:mm:ss`) di Kolom H.
   - **Sheet Datasuite:** Mengetik/paste di Kolom A atau Kolom L akan otomatis mengisi tanggal & jam di Kolom M.

2. **Sinkronisasi Data (Menu 🔄 Sinkron)**
   - Mencocokkan ID (Kolom A) antara sheet Backlog dan Datasuite.
   - Menarik data yang cocok dan menggabungkannya ke sheet **Same Data**.
   - Dilengkapi menu kustom di spreadsheet agar bisa dijalankan manual kapan saja.

3. **Auto Generate Link FMS**
   - Otomatis membuat URL detail pesanan Shopee (`https://spx.shopee.co.id/#/orderDetail/...`) berdasarkan ID pesanan di Kolom A.

4. **Kalkulasi Aging Time (SLA)**
   - Membaca teks di Kolom E (misal: *"60.73 hours"*), mengekstrak angkanya, dan memberikan status otomatis di Kolom J:
     - `>72hour` (Jika > 72 jam)
     - `>48hour` (Jika > 48 jam)
     - Kosong (Jika ≤ 48 jam)

## 🛠️ Cara Penggunaan
1. Buka Google Sheets
2. Paste data Backlog di sheet Backlog
3. Paste data tarikan Datasuite 24jam di sheet Datasuite
4. Gunakan menu **🔄 Sinkron -> Sinkron Same Data** untuk menjalankan proses.
