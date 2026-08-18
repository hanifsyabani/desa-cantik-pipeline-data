# Desa Cantik - Pipeline Data Harga

Project ini dipakai untuk mengolah file Excel mentah hasil download aplikasi
kuisioner menjadi file dashboard siap pakai. Aplikasi membaca sheet raw,
membersihkan kode wilayah, mengubah data 20 komoditas dari format wide ke long,
lalu membuat dua output:

- `Dashboard (Raw)`: semua hasil transformasi.
- `Dashboard (Cleaned)`: hasil yang sudah difilter dari data kosong, wilayah
  gagal lookup, dan harga terlalu rendah.

## Struktur Folder

```text
app/
|-- README.md
|-- requirements.txt
|-- streamlit_app.py
|-- src/
|   `-- desa_cantik/
|       |-- __init__.py
|       |-- pipeline.py
|       |-- config.py
|       `-- excel_utils.py
|-- scripts/
|   |-- generate_full_pipeline.py
|   `-- automate_sheet3_to_4.py
`-- notebooks/
    `-- data_preprocessing.ipynb
```

## Fungsi Masing-Masing File

### `streamlit_app.py`

File utama untuk menjalankan aplikasi web Streamlit.

Fungsinya:

- Menampilkan form upload file `.xlsx`.
- Menyediakan input `Harga minimum`.
- Memanggil pipeline utama dari `src/desa_cantik/pipeline.py`.
- Menampilkan ringkasan jumlah baris hasil proses.
- Menyediakan tombol download output Excel.

Jalankan dengan:

```powershell
streamlit run streamlit_app.py
```

### `requirements.txt`

Daftar library Python yang dibutuhkan project.

Saat ini berisi:

- `streamlit`: untuk UI web.
- `openpyxl`: untuk baca dan tulis file Excel.

Install dependency dengan:

```powershell
python -m pip install -r requirements.txt
```

### `src/desa_cantik/config.py`

Tempat menyimpan konfigurasi utama pipeline.

Fungsinya:

- Nama sheet input dan output.
- Nomor baris/kolom penting dari file Excel.
- Header output dashboard.
- Daftar bulan Indonesia.
- Konfigurasi 20 komoditas di variabel `BLOCKS`.

Kalau ada perubahan kolom komoditas, nama komoditas, satuan, atau mapping
kualitas, edit file ini.

### `src/desa_cantik/pipeline.py`

Logic utama pengolahan data.

Fungsinya:

- Membaca workbook input.
- Mengambil data dari sheet `1. list_data (ori)`.
- Membersihkan wilayah responden.
- Mengubah data 20 komoditas menjadi format dashboard.
- Membuat sheet `Dashboard (Raw)` dan `Dashboard (Cleaned)`.
- Mengembalikan statistik hasil proses.

File ini dipakai oleh:

- `streamlit_app.py`
- `scripts/generate_full_pipeline.py`

### `src/desa_cantik/excel_utils.py`

Kumpulan helper untuk parsing dan lookup Excel.

Fungsinya:

- Menggeser referensi kolom dari format sheet lama ke sheet raw.
- Parsing pilihan kualitas, misalnya `1. Selancar`.
- Parsing tanggal.
- Menghitung minggu dalam bulan.
- Membaca lookup wilayah dari sheet `kd wilayah`.
- Membersihkan nama kecamatan dan desa.

### `src/desa_cantik/__init__.py`

Penanda bahwa folder `desa_cantik` adalah package Python.

File ini sengaja dibuat ringan supaya import package tidak langsung memuat
dependency berat.

### `scripts/generate_full_pipeline.py`

Entrypoint CLI untuk menjalankan pipeline lewat terminal.

Fungsinya sama seperti Streamlit, tapi tanpa UI web.

Jalankan dengan:

```powershell
python scripts\generate_full_pipeline.py input.xlsx output.xlsx --min-harga 1000
```

Contoh:

```powershell
python scripts\generate_full_pipeline.py "..\raw_laporan.xlsx" "..\hasil.xlsx"
```

### `scripts/automate_sheet3_to_4.py`

Script tambahan/legacy untuk otomasi dari sheet 3 ke sheet 4.

Catatan:

- File yang aktif dipakai untuk pipeline penuh adalah
  `scripts/generate_full_pipeline.py`.
- Simpan file ini kalau masih butuh proses lama berbasis sheet 3.
- Kalau workflow sudah sepenuhnya pakai pipeline baru, file ini bisa dianggap
  arsip atau referensi.

### `notebooks/data_preprocessing.ipynb`

Notebook eksplorasi data.

Fungsinya:

- Mencoba langkah preprocessing secara interaktif.
- Mengecek bentuk data.
- Eksperimen sebelum logic dipindahkan ke file Python.

Notebook sebaiknya tidak dijadikan sumber logic utama. Logic produksi tetap
ditaruh di `src/desa_cantik/`.

## Cara Menjalankan Aplikasi Web

Dari folder `app`:

```powershell
python -m pip install -r requirements.txt
streamlit run streamlit_app.py
```

Lalu buka URL yang muncul, biasanya:

```text
http://localhost:8501
```

## Cara Menjalankan Pipeline Lewat Terminal

Dari folder `app`:

```powershell
python scripts\generate_full_pipeline.py path\ke\input.xlsx path\ke\output.xlsx
```

Dengan harga minimum custom:

```powershell
python scripts\generate_full_pipeline.py path\ke\input.xlsx path\ke\output.xlsx --min-harga 500
```

## Format File Input

File Excel input harus berisi sheet berikut:

- `1. list_data (ori)`
- `kd wilayah`
- `Kode Kualitas`

## Panduan Edit Cepat

- Ubah tampilan web: edit `streamlit_app.py`.
- Ubah aturan transformasi data: edit `src/desa_cantik/pipeline.py`.
- Ubah daftar komoditas atau mapping kolom: edit `src/desa_cantik/config.py`.
- Ubah fungsi parsing/lookup Excel: edit `src/desa_cantik/excel_utils.py`.
- Jalankan dari terminal: pakai `scripts/generate_full_pipeline.py`.
