# Universal Excel Translator
## Product Design & Feature Specification

Dokumen ini berisi konsep desain aplikasi desktop untuk melakukan **translasi file Excel/CSV menggunakan LLM**.

Fokus utama adalah **UI/UX yang cepat, intuitif, dan efisien untuk workflow translasi data spreadsheet**.

Dokumen ini hanya menjelaskan **konsep aplikasi dan pengalaman pengguna**, tanpa menentukan teknologi implementasi.

---

# 1. Tujuan Aplikasi

Aplikasi ini bertujuan untuk membantu pengguna menerjemahkan data spreadsheet secara cepat menggunakan LLM.

Target workflow utama:

1. Load file spreadsheet
2. Pilih kolom sumber
3. Pilih kolom hasil terjemahan
4. Jalankan translasi
5. Download hasil

Semua proses harus dapat dilakukan dengan **interaksi seminimal mungkin**.

---

# 2. Konsep UI Utama

Desain UI mengikuti gaya **modern developer tools**.

Layout utama terdiri dari tiga area:

```
Sidebar | Main Panel | File Preview
```

Dengan panel log di bagian bawah.

```
┌─────────────────────────────────────────────┐
│ Universal Excel Translator                  │
├──────────┬─────────────────────┬────────────┤
│ Sidebar  │ Main Panel          │ File       │
│          │                     │ Preview    │
│          │                     │            │
│Trans     │ Settings / Chat     │ Table      │
│Chat      │ / History           │ Preview    │
│Hist      │                     │            │
│ApiMgr    │                     │            │
├──────────┴─────────────────────┴────────────┤
│ Log Terminal                                │
└─────────────────────────────────────────────┘
```

---

# 3. Sidebar Navigasi

Sidebar berisi menu utama aplikasi:

- Translator
- API Manager
- Chat
- History

Karakteristik sidebar:

- icon-only
- minimalis
- selalu terlihat
- tidak collapsible

Tujuannya agar navigasi cepat dan tidak mengganggu workspace utama.

---

# 4. File Preview Panel

Panel kanan menampilkan **preview spreadsheet secara realtime**.

Preview ini merupakan komponen utama UI.

Fitur utama:

- menampilkan hingga 1000 baris
- scroll horizontal
- nomor baris
- header kolom interaktif

---

# 5. Interaksi Kolom (Fitur UX Paling Penting)

Pengguna **tidak memilih kolom dari dropdown**.

Sebaliknya, pengguna **langsung klik header kolom pada tabel**.

Saat header kolom diklik, muncul menu kecil:

```
Set Column Role
• Source Column
• Translation Column
• Context Column
• Repair Column
```

Setelah dipilih, kolom akan memiliki **indikator visual**.

Contoh:

```
Column B → Source
Column C → Translation
Column D → Context
```

Kolom yang dipilih harus:

- diberi warna
- memiliki label kecil di header

Tujuan UX ini adalah agar workflow terasa seperti menggunakan spreadsheet.

---

# 6. Indicator Role Kolom

Di atas tabel ditampilkan ringkasan role kolom.

Contoh:

```
Source: B
Translation: C
Context: D
Repair: E
```

User juga dapat mengubahnya secara manual.

---

# 7. Panel Log

Panel bawah menampilkan log proses.

Karakteristik:

- dapat di-resize
- auto-scroll
- warna log berbeda

Jenis log:

INFO  
SUCCESS  
WARNING  
ERROR  

Tambahkan progress bar untuk batch process.

---

# 8. Tab Translator

Tab ini berisi pengaturan utama proses translasi.

Komponen utama:

Provider selector  
API key selector  
Model selector  

File input

File input mendukung:

- drag and drop
- file picker

Column configuration

Language configuration

Prompt configuration

Terdapat dua jenis prompt:

1. Translation Prompt
2. Repair Prompt

Pengaturan tambahan:

Batch size  
Thread count  
Retry configuration  

LLM parameters:

Temperature  
Max completion tokens

---

# 9. Mode Operasi

Mode yang tersedia:

Translate

Translate + Context

Repair

Repair digunakan untuk memperbaiki hasil translasi yang gagal.

---

# 10. Format File

Input:

- XLSX
- CSV

Jika CSV dimuat, sistem akan memprosesnya seperti spreadsheet.

Output selalu berupa file baru.

```
original_filename_translated.xlsx
```

---

# 11. Chat Interface

Aplikasi juga menyediakan tab chat untuk berinteraksi langsung dengan LLM.

UI chat menyerupai chat modern.

Fitur:

- bubble chat
- markdown rendering
- code block
- tombol copy

Pengguna dapat mengatur parameter LLM secara langsung.

---

# 12. History

Tab history menyimpan aktivitas translasi.

Menampilkan:

- timestamp
- nama file
- mode
- jumlah baris berhasil
- jumlah gagal

History dapat digunakan untuk memuat ulang konfigurasi.

---

# 13. API Manager

Panel untuk mengelola API key.

Fitur:

- tambah API key
- edit API key
- hapus API key

API key disimpan secara terenkripsi.

---

# 14. Sistem Profil User

Aplikasi mendukung beberapa profil pengguna.

Saat pertama kali membuka aplikasi:

pengguna membuat profil.

Profil berisi:

- username
- password
- foto profil
- konfigurasi aplikasi
- API key

Pengguna dapat memilih profil saat login.

---

# 15. Save Configuration

Pengguna dapat menyimpan konfigurasi translasi.

Konfigurasi yang disimpan:

- provider
- model
- kolom
- bahasa
- batch settings
- prompt

Config disimpan dalam folder profil user.

---

# 16. Theme

Aplikasi mendukung:

Dark mode  
Light mode

Default mengikuti sistem.

---

# 17. Prinsip Desain UI

UI harus mengikuti prinsip:

- minimal friction
- spreadsheet-first workflow
- fast column selection
- minimal dropdown usage
- visual feedback yang jelas

Tujuan utama adalah membuat workflow translasi terasa seperti bekerja langsung di spreadsheet.

---

# 18. Output yang Diharapkan dari AI

AI diminta menghasilkan:

1. Arsitektur aplikasi
2. Struktur folder project
3. Desain komponen UI
4. Diagram alur interaksi user
5. Contoh implementasi komponen utama
