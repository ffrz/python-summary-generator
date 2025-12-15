======================================================================== PCM
SUMMARY GENERATOR - TECHNICAL REPOSITORY DOCUMENTATION Version: 1.0.0
========================================================================

1. OVERVIEW

---

PCM Summary Generator adalah aplikasi desktop berbasis Python yang dirancang
untuk mengotomatisasi proses rekapitulasi data Project Cost Management (PCM).
Aplikasi ini membaca data dari file Excel mentah (.xls dan .xlsx), melakukan
normalisasi data, dan menghasilkan laporan Master Summary.

Aplikasi ini menggunakan pola arsitektur modular (refactored) untuk memudahkan
pemeliharaan (maintainability) dan skalabilitas.

2. TECH STACK & DEPENDENCIES

---

Language : Python 3.8+ GUI : PySide6 (Qt for Python) Excel I/O: openpyxl (untuk
.xlsx), xlrd (untuk .xls legacy) System : watchdog (untuk real-time folder
monitoring) Build : PyInstaller (untuk pembuatan file .exe)

3. STRUKTUR FILE (MODULARIZATION)

---

Kode sumber telah dipecah menjadi 5 modul utama untuk penerapan prinsip
Separation of Concerns (SoC):

[1] main.py Titik masuk aplikasi (Entry Point). Hanya bertugas menginisialisasi
QApplication dan memanggil MainWindow.

[2] ui.py Berisi logika antarmuka pengguna (Frontend). Menangani layout, widget
tabel, dialog, dan interaksi tombol.

[3] workers.py Berisi logika Threading (Background Tasks) agar UI tidak
freeze: - WatcherThread : Memantau perubahan file di folder input. -
PreviewWorker : Memindai dan mem-parsing file untuk preview tabel. -
GeneratorWorker : Menyalin file dan menulis file Excel Summary.

[4] parsers.py Berisi logika pembacaan file Excel. Menggunakan "Adapter Pattern"
untuk menstandarisasi antarmuka antara library xlrd dan openpyxl sehingga core
logic ekstraksi data hanya ditulis satu kali (DRY).

[5] helpers.py Fungsi-fungsi utilitas murni (Pure Functions) seperti sanitasi
nama file, konversi currency, konversi tanggal, dan regex.

4. LOGIKA UTAMA (CORE LOGIC)

---

A. Parsing Data Parser membaca cell spesifik berdasarkan alamat absolut:

- Date: B3
- Project No: K4 (Fallback ke H4)
- Customer: K3 (Fallback ke H3)
- Currency: Diditeksi dari teks di cell A5 menggunakan Regex.

Komponen biaya (Sub Total, Penalty, dll) dicari dengan memindai Kolom A mulai
baris ke-9 ke bawah.

B. Mata Uang (Currency) Jika mata uang terdeteksi "IDR", nilai Kurs dipaksa
menjadi 1.0. Jika mata uang asing, nilai Kurs diambil dari cell B4.

C. Output Generation File summary digenerate menggunakan openpyxl. Baris paling
bawah otomatis ditambahkan "GRAND TOTAL" yang berisi rumus Excel (=SUM) untuk
menjumlahkan seluruh kolom numerik.

5. CARA INSTALASI (DEVELOPMENT)

---

Disarankan menggunakan Virtual Environment (venv).

1. Clone repository ini.
2. Buka terminal di folder root project.
3. Buat environment: python -m venv venv
4. Aktifkan environment:
   - Windows: venv\Scripts\activate
   - Mac/Linux: source venv/bin/activate
5. Install requirements: pip install PySide6 openpyxl xlrd watchdog pyinstaller

6. CARA MENJALANKAN APLIKASI

---

Jalankan perintah berikut di terminal:

python main.py

7. CARA BUILD EXE (DEPLOYMENT)

---

Untuk membuat file executable (.exe) standalone:

1. Pastikan file "PCM Summary Generator.spec" ada di folder root.
2. Jalankan perintah PyInstaller:

   pyinstaller "PCM Summary Generator.spec" --clean --noconfirm

3. File .exe akan muncul di folder "dist". (Ukuran estimasi: 30MB - 60MB
   tergantung optimasi).

4. CATATAN PENGEMBANGAN (MAINTENANCE)

---

- Menambah Kolom Baru: Edit file "workers.py" di bagian header list dan mapping
  data pada class GeneratorWorker.

- Mengubah Posisi Cell Input: Edit file "parsers.py" pada fungsi
  extract_common_logic. Gunakan helper adapter.get_by_addr("A1") untuk
  kemudahan.

- Mengubah Format Nama File Output: Edit file "workers.py" pada bagian
  GeneratorWorker -> loop valid_data.

========================================================================
Developer: Fahmi Fauzi Rahman Contact : 0853-1740-4760
========================================================================
