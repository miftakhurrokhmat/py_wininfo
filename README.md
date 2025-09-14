# py_wininfo
**py_wininfo** adalah script Python untuk **mengumpulkan informasi lengkap tentang sistem Windows 11/10**, termasuk:

- Sistem operasi & user
- Processor (CPU)
- RAM & slot RAM
- Hard disk & drive
- VGA / GPU
- Motherboard
- Printer
- Aplikasi populer (Office, PDF reader, Zip/Unzip)

Semua informasi akan **ditampilkan di layar** dan **disimpan ke file** `laporan_sistem.txt`.

---
## Prasyarat

- Python 3.10
- Windows 11 / 10
- Powershell (default di Windows)
- Hak akses normal user sudah cukup

---
## Instalasi
1. Clone repository:

```bash
git clone https://github.com/username/py_wininfo.git
cd py_wininfo
```
2. Tidak ada dependencies tambahan yang diperlukan.

---
## Cara Menggunakan

1. Jalankan script:

```bash
python main.py
```

2. Pilih menu yang tersedia:

```
=== MENU ===
1. Informasi Sistem
2. Informasi Processor
3. Informasi RAM
4. Informasi Hard Disk
5. Informasi VGA
6. Aplikasi Populer
7. Motherboard
8. Printer
0. Keluar
```

3. Informasi akan dicetak di layar dan otomatis disimpan ke `laporan_sistem.txt`.

---
## Contoh Output

```
=== INFORMASI SISTEM ===
OS Name: Windows 11
OS Version: 10.0.22621
User saat ini: User01
Daftar user yang ada: User01, Admin, Guest
...
```
