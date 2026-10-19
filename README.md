# Rekonsiliasi Pajak BKU SKPD SIPD + Coretax

Aplikasi desktop untuk mencocokkan potongan pajak di **BKU SKPD** dengan setoran di **Buku Besar Coretax**:

- NTPN di BKU dicocokkan dengan NTPN yang valid di Coretax (`ReferenceNumber`)
- Kolom **Penerimaan** di BKU dicocokkan dengan kolom **Amount** di Coretax
- Hasilnya file Excel berisi kesimpulan dan daftar yang perlu diperbaiki, lengkap dengan nomor bukti 6 digit

## Isi folder

| File | Fungsi |
|---|---|
| `RekonsiliasiPajakSIPD.exe` | Aplikasi siap pakai (single file, tidak perlu Python) |
| `app.py` | Kode tampilan (GUI tkinter) |
| `konsolidasi.py` | Kode pencocokan dan pembuatan file Excel |
| `jalankan.bat` | Menjalankan aplikasi dari kode Python |
| `build_exe.bat` | Membuat ulang `RekonsiliasiPajakSIPD.exe` setelah kode diubah |
| `requirements.txt` | Library Python yang dibutuhkan (`openpyxl`) |

## Cara pakai

1. Klik dua kali `RekonsiliasiPajakSIPD.exe`.
2. Di bagian **1. Upload file**, klik **Pilih File...**:
   - **BKU SKPD (.xlsx)**: file BKU pajak dari aplikasi keuangan daerah
   - **Buku Besar Coretax (.csv)**: file unduhan buku besar dari Coretax
3. Klik **Proses Rekonsiliasi**. Ringkasan dan daftar perbaikan muncul di bagian **2. Hasil rekonsiliasi**.
4. Jendela simpan muncul dengan nama awal `Rekonsiliasi_Pajak_<Bulan>_<Tahun>.xlsx` di folder file BKU. Klik **Save**, lalu file Excel langsung terbuka.

Tombol **Simpan Excel...**, **Buka File Hasil**, dan **Buka Folder** bisa dipakai setelah proses selesai.

Menu **Bantuan** berisi **Petunjuk Penggunaan** (bisa juga dibuka dengan tombol F1) dan **Tentang Aplikasi** (nama dan versi aplikasi).

### Syarat komputer

- Windows 10 atau 11, 64-bit
- Program untuk membuka `.xlsx` (Excel, WPS, atau LibreOffice), hanya untuk melihat hasil

## File yang dibutuhkan

### BKU SKPD (`.xlsx` atau `.csv`)

Baris judul kolom dicari otomatis di 30 baris pertama, di sheet mana pun.

| Kolom | Wajib | Keterangan |
|---|---|---|
| `NTPN` | Ya | |
| `Uraian` | Ya | Diawali `Penerimaan Pot. Pajak` atau `Pembayaran Pot. Pajak` |
| `Penerimaan` | Ya | Format `2.171.379` atau angka biasa |
| `Pengeluaran` | Ya | |
| `No` | | Baris tanpa nomor (Saldo Sebelumnya, Total) diabaikan |
| `Tanggal` | | Dibandingkan dengan tanggal setor di Coretax |
| `Nomor Bukti` | | 6 digit nomor urut diambil dari sini, mis. `33.26/06.0/`**`002556`**`/UP/...` |
| `Nama Rekening` | | Jenis pajak (PPH 21, PPN, dst.) untuk konsolidasi per jenis |
| `ID Billing`, `Kode Rekening` | | Ditampilkan di hasil |

### Buku Besar Coretax (`.csv` atau `.xlsx`)

Kolom wajib: `ReferenceNumber` (NTPN) dan `Amount`. Kolom lain yang dipakai bila ada: `AmountLeft`, `RevenueCode` (KAP), `PaymentCode` (KJS), `PeriodCode`, `TransactionDate`, `PostingDate`, `IsCancelled`, `ReferenceType`, `TaxpayerName`, `TaxpayerTIN`.

> **Penting:** pakai file CSV asli hasil unduhan Coretax. Jangan buka CSV di Excel lalu klik **Save**. Excel akan mengubah angka panjang ke format ilmiah dan digitnya hilang (mis. NPWP menjadi `3,11661E+12`). Aplikasi akan memberi peringatan di bagian **D. CATATAN** bila mendeteksi hal ini.

## Cara pencocokan

1. Setiap baris **Penerimaan Pot. Pajak** dipasangkan dengan baris **Pembayaran Pot. Pajak**-nya (NTPN dan nilai sama).
2. NTPN BKU dicari di Coretax. Setoran yang dibatalkan (`IsCancelled = true`) tidak dihitung sebagai NTPN valid.
3. NTPN yang tidak ketemu persis dicari padanannya di Coretax dengan beda paling banyak **3 karakter** (salah ketik).
4. Untuk NTPN yang berpadanan, nilai Penerimaan BKU dibandingkan dengan Amount Coretax, dan tanggal BKU dengan tanggal setor.
5. NTPN di Coretax yang tidak dipakai oleh baris BKU mana pun dilaporkan sebagai "tidak tercatat di BKU".

### Arti status

| Status | Warna | Arti |
|---|---|---|
| `COCOK` | Hijau | NTPN, nilai, dan tanggal sama |
| `COCOK - TANGGAL BEDA` | Kuning | NTPN dan nilai sama, tanggal BKU beda dengan tanggal setor |
| `NIHIL` | Abu-abu | Potongan Rp0 tanpa NTPN, tidak perlu tindakan |
| `NTPN SALAH KETIK` | Merah | NTPN BKU salah ketik; NTPN yang benar dari Coretax ditampilkan |
| `NILAI BEDA` | Merah | NTPN sama, nilai BKU berbeda dengan Amount Coretax |
| `TIDAK DITEMUKAN` | Merah | NTPN BKU tidak ada di Coretax (atau kosong padahal ada potongan) |
| `NTPN GANDA DI BKU` | Merah | NTPN yang sama dipakai lebih dari satu transaksi BKU |
| `NTPN DIBATALKAN DI CORETAX` | Merah | NTPN berstatus dibatalkan di Coretax |

## Isi file hasil

| Sheet | Isi |
|---|---|
| **Kesimpulan** | A. Ringkasan angka dan rincian selisih · B. Hasil pencocokan · C. Yang perlu diperbaiki (dengan nomor bukti 6 digit) · D. Catatan |
| **Rekonsiliasi NTPN** | Semua transaksi BKU berdampingan dengan data Coretax, diberi warna sesuai status, bisa difilter |
| **Perlu Perbaikan** | Hanya transaksi yang perlu dikoreksi, beserta tindakannya |
| **Coretax Tidak di BKU** | NTPN yang ada di Coretax tetapi tidak tercatat di BKU |
| **Konsolidasi Jenis Pajak** | Total per jenis pajak (BKU vs Coretax, deposit 411618 vs KAP lain) dan ringkasan per KAP-KJS |

Bagian **C. Yang perlu diperbaiki** juga mengingatkan bila saldo deposit KAP 411618 (`AmountLeft`) belum dipakai, atau bila satu jenis pajak disetor dengan lebih dari satu KAP.

## Menjalankan dari kode Python

Butuh Python 3.10 atau lebih baru ([python.org](https://www.python.org/downloads/), centang **Add python.exe to PATH** saat instalasi).

- **GUI:** klik dua kali `jalankan.bat`. Saat pertama dijalankan, `openpyxl` akan dipasang otomatis (butuh internet).
- **Command line, tanpa GUI:**

  ```
  python konsolidasi.py "BKU SKPD.xlsx" bukubesar_coretax.csv hasil.xlsx
  ```

## Membuat ulang exe

Setelah `app.py` atau `konsolidasi.py` diubah, tutup dulu `RekonsiliasiPajakSIPD.exe` bila sedang terbuka, lalu klik dua kali `build_exe.bat`. File `RekonsiliasiPajakSIPD.exe` yang baru akan menggantikan yang lama di folder ini.

## Masalah yang sering muncul

| Masalah | Solusi |
|---|---|
| "Windows protected your PC" saat membuka exe | Klik **More info** lalu **Run anyway**. Exe tidak punya tanda tangan digital. |
| Exe diblokir antivirus | Minta bagian IT mengizinkan `RekonsiliasiPajakSIPD.exe` |
| "Format file tidak sesuai" | Periksa apakah file BKU dan Coretax tertukar, atau kolom wajib tidak ada |
| "Tidak bisa menyimpan" | File hasil dengan nama yang sama sedang terbuka di Excel. Tutup dulu, lalu klik **Simpan Excel...** |
| File `.xls` tidak bisa dibaca | Buka di Excel lalu **Save As** `.xlsx` |
| NPWP di judul hasil tampil `-` | CSV Coretax pernah disimpan ulang di Excel. Unduh ulang CSV asli dari Coretax |
