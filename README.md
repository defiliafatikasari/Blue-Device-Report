# Laporan Rekapitulasi Kinerja — Program Blue Device
### PT. AKAR DAYA - Indihome by Telkomsel
#### Rekapitulasi Hasil Caring, Visit Lapangan, dan Pelanggan Setuju (Closing 30 November 2026)

Aplikasi web dashboard desktop eksekutif untuk **TIKA (Area Pamekasan)** yang ditujukan kepada **Bapak Choirul Umam (Leader / Supervisor)**.

---

## 📌 4 Komponen Utama Laporan Sesuai Arahan Leader:

1. **Hasil Caring (Panggilan Telepon):**
   - Menghitung seluruh pelanggan yang telah dihubungi oleh **TIKA** (serta opsi filter petugas lain: NURUL, VINDA, ROS, `...` Tanpa Caring, Kosong / Belum di-Caring, dan Semua Data Tim).
   - Rincian: **Diangkat**, **Tidak Diangkat**, **Tidak Menyambung**, serta persentase tersambung.
2. **Hasil Visit Lapangan (Kunjungan Rumah):**
   - Menghitung pelanggan yang dikunjungi langsung oleh **TIKA** pada kolom `NOTE PIC ANTAR`.
   - Menampilkan pemisahan transparan: **Visit Mandiri** vs **Visit Tandem / Tim**.
3. **Jumlah Blue Device yang Pelanggan Setuju:**
   - Menghitung pelanggan yang responnya `BERMINAT? = TERTARIK` atau status order `DONE ACT` / `DATE ACT` terisi.
4. **Realisasi Aktivasi Lapangan (Done Act):**
   - Menghitung perangkat yang sudah aktif di aplikasi **MyTelkomsel & DigiPOS**.

---

## ⚡ Fitur Utama:
- **Logo & Identitas Resmi**: Menggunakan file logo resmi `LOGO-PT.jpeg` dengan nama perusahaan **PT. AKAR DAYA - Indihome by Telkomsel**.
- **100% Data Real Excel**: Seluruh data (10.809 baris) dibaca langsung dari file `BLUE DEVICE TIKA.xlsx` tanpa ada dummy/demo.
- **Deteksi Otomatis Baris Kosong**: Tombol filter **"Kosong / Belum di-Caring"** secara otomatis mendeteksi baris database yang belum pernah di-caring (9.850 baris).
- **Badge Baris Dinamis**: Header menampilkan jumlah total baris asli dari file Excel secara real-time.
- **Opsi Tampilkan Semua (All Rows)**: Pilihan paginasi mencakup opsi `Semua Baris (All)` untuk melihat seluruh data dalam satu tampilan.
- **Bulk Upload Foto Evidence**: Upload puluhan foto bukti lapangan sekaligus dengan pencocokan otomatis ke ID pelanggan / No Telepon.
- **Lembar Siap Cetak**: Tombol cetak / simpan PDF dengan format dokumen resmi yang rapi.

---

## 📁 Berkas Proyek di Folder Ini:
- [`BLUE DEVICE TIKA.xlsx`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/BLUE%20DEVICE%20TIKA.xlsx) — File Excel kerja utama Anda.
- [`LOGO-PT.jpeg`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/LOGO-PT.jpeg) — Logo resmi PT. AKAR DAYA - Indihome by Telkomsel.
- [`Buka_Dashboard.bat`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/Buka_Dashboard.bat) — Shortcut 1-klik untuk membuka dashboard di komputer.
- [`index.html`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/index.html) — File web dashboard laporan utama.
- [`style.css`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/style.css) — Tema korporat & layout cetak lembar laporan.
- [`app.js`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/app.js) — Logika filter caring, visit TIKA, kalkulasi KPI, dan bulk upload foto.
- [`data.js`](file:///e:/ProjectMachineLearning/rekap-bluedevice-tika/data.js) — Data riil cadangan offline (10.809 baris) dari `BLUE DEVICE TIKA.xlsx`.
