# Prompt Pengembangan Aplikasi  
## Universal Excel Translator Desktop App

Anda adalah **Senior Software Engineer** yang ahli dalam pengembangan aplikasi desktop menggunakan **C# dan Avalonia UI**.

Tugas Anda adalah **mendesain dan mengimplementasikan aplikasi desktop profesional** untuk menerjemahkan file Excel/CSV menggunakan berbagai API LLM.

Aplikasi ini merupakan **port dari konsep aplikasi Python sebelumnya**, tetapi harus diimplementasikan ulang menggunakan:

- **C#**
- **.NET 8**
- **Avalonia UI**

Aplikasi harus memiliki arsitektur yang **bersih, modular, scalable, dan production-ready**.

JANGAN menyederhanakan fitur kecuali diminta.

---

# 1. Tech Stack

Gunakan teknologi berikut:

Core:

- C#
- .NET 8
- Avalonia UI
- MVVM Pattern

Library yang direkomendasikan:

- ClosedXML (Excel)
- CsvHelper
- RestSharp / HttpClient
- Newtonsoft.Json atau System.Text.Json
- Polly (retry handling)
- LiteDB atau SQLite untuk history

Packaging:

- Publish sebagai **single executable**

Target OS:

- Windows (utama)
- tetapi tetap cross-platform karena Avalonia

---

# 2. Nama Aplikasi

Universal Excel Translator

Versi:

v2.0

---

# 3. Layout Aplikasi

Gunakan layout modern seperti developer tools.

Struktur layout utama:

```
[Sidebar Kiri] | [Main Content] | [Sidebar Kanan]
```

Dengan panel log di bagian bawah.

```
┌─────────────────────────────────────────────┐
│ Universal Excel Translator v2.0             │
├──────────┬─────────────────────┬────────────┤
│ Sidebar  │ Main Content        │ Preview    │
│ Kiri     │                     │ File       │
│          │                     │            │
│Trans     │ Translator / Chat   │ Table      │
│Chat      │ / History           │ Preview    │
│Hist      │                     │            │
├──────────┴─────────────────────┴────────────┤
│ Log Terminal                                │
└─────────────────────────────────────────────┘
```

---

# 4. Sidebar Kiri

Berisi 3 menu utama:

- Translator
- Chat
- History

Fitur:

Sidebar dapat di-resize.

Jika diperkecil hingga minimum maka otomatis collapse menjadi **icon-only mode**.

Mode:

Expanded:

Icon + Text

Collapsed:

Icon saja.

Lebar minimum:

48px.

---

# 5. Main Content Panel

Konten berubah berdasarkan tab.

Tab yang tersedia:

- Translator
- Chat
- History

---

# 6. Sidebar Kanan (Preview File)

Panel ini menampilkan **preview spreadsheet interaktif**.

Fitur:

- Menampilkan hingga 1000 baris pertama
- Nomor baris
- Scroll horizontal
- Header kolom clickable

Klik header kolom untuk assign role:

- Source Column
- Translated Column
- Repair Column
- Context Column

Header preview menampilkan:

- Nama file
- Tombol fullscreen
- Tombol download

Tambahkan search bar untuk filter baris.

---

# 7. Panel Bawah (Log Terminal)

Terminal log untuk proses aplikasi.

Jenis log:

INFO
SUCCESS
WARNING
ERROR

Tambahkan:

- progress bar
- tombol clear log
- toggle auto-scroll

Gunakan font monospace.

---

# 8. Tab Translator

Form pengaturan utama.

Berisi:

Provider selector  
API key management  
Model selector  
Operation mode  
File input  
Column picker  
Language configuration  
Style preset  
Batch configuration  
Thread configuration  
Retry configuration  
Temperature control  
Algorithm selector  
Action buttons

---

# 9. Provider LLM

Aplikasi harus mendukung berbagai provider:

- OpenAI
- Anthropic
- Google Gemini
- Groq
- Cerebras
- OpenRouter
- Together AI
- DeepSeek
- Mistral

Fitur:

User memasukkan API key.

Aplikasi mengambil **model list dari API provider secara otomatis**.

Dropdown model diperbarui secara dinamis.

---

# 10. Sistem Penyimpanan API Key

API key harus disimpan secara **terenkripsi**.

Gunakan sistem vault:

```
keys.enc
```

Vault harus dilindungi oleh **master password**.

Alur:

Pertama kali:

User membuat master password.

Selanjutnya:

User membuka vault dengan password.

Vault menyimpan:

- Provider
- API key

---

# 11. Mode Operasi

Mode yang harus didukung:

### Translate

Terjemahan dasar.

### Translate + Context

Menggunakan kolom context.

### Repair

Memperbaiki hasil terjemahan yang rusak.

---

# 12. Format File

Input:

- XLSX
- CSV

Jika CSV:

konversi internal ke struktur Excel.

Output:

Selalu membuat file baru:

```
nama_file_translated.xlsx
```

---

# 13. Column Picker

User dapat menentukan:

Source column  
Translated column  
Repair column  
Context column

Kolom bisa dipilih dengan:

- mengetik huruf kolom
- klik header pada preview

---

# 14. Language Configuration

Contoh:

EN → ID  
JP → EN  
AUTO → EN

Gunakan deteksi bahasa otomatis jika diperlukan.

---

# 15. Style Preset

Preset:

- Neutral
- Formal
- Casual
- Marketing
- Technical
- Academic

Preset memodifikasi prompt terjemahan.

---

# 16. Batch Processing

Pengaturan:

- range baris
- batch size
- delay antar batch
- skip baris yang sudah terisi

---

# 17. Multi Thread

User dapat menentukan jumlah thread per API key.

---

# 18. Retry System

Jika gagal:

retry otomatis.

Konfigurasi:

- retry limit
- backoff delay

---

# 19. Temperature Control

Slider:

0.0 – 1.0

Menampilkan nilai realtime.

---

# 20. Tab Chat

Chat interface seperti ChatGPT.

Fitur:

- bubble chat
- markdown rendering
- code block
- tombol copy

Navigasi percakapan.

Input chat:

Enter / Shift+Enter behavior dapat dikonfigurasi.

---

# 21. Chat Advanced Settings

Tambahkan pengaturan:

- Temperature
- Top-P
- Context length
- System prompt
- Custom instructions

---

# 22. History Tab

Menyimpan history proses translate.

Menampilkan:

- timestamp
- nama file
- mode
- statistik

---

# 23. Save / Load Config

User dapat menyimpan konfigurasi ke file JSON.

Yang disimpan:

- provider
- model
- column config
- language
- batch settings
- temperature

API key tidak ikut disimpan.

---

# 24. Theme

Aplikasi harus mendukung:

Dark mode  
Light mode

Default:

Auto detect dari sistem.

---

# 25. Struktur Project

Gunakan struktur modular seperti berikut:

```
UniversalExcelTranslator/

Core/
Services/
Providers/
Vault/
Models/

UI/
Views/
ViewModels/
Controls/

Styles/
Themes/

Data/
History/

Program.cs
App.axaml
```

Gunakan **MVVM pattern**.

---

# 26. Tahapan Implementasi

Implementasikan bertahap:

Phase 1  
Core translation engine

Phase 2  
Provider system

Phase 3  
Vault encryption

Phase 4  
GUI layout

Phase 5  
File preview

Phase 6  
Chat system

Phase 7  
History

Phase 8  
Packaging

---

# 27. Standar Kode

Kode harus:

- modular
- readable
- maintainable
- menggunakan MVVM

Hindari file besar yang monolithic.

---

# 28. Output Yang Diminta

AI harus menghasilkan:

1. Arsitektur project
2. Struktur folder lengkap
3. Penjelasan komponen utama
4. Contoh implementasi class penting
5. Contoh UI Avalonia
6. Contoh ViewModel
7. Contoh service API provider

```
