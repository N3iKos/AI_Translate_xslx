# Universal Excel Translator — Konsep Final Aplikasi

## 1. Tujuan Aplikasi

Aplikasi ini bertujuan untuk membantu pengguna menerjemahkan data spreadsheet secara cepat menggunakan LLM.

Target workflow utama:

1. Load file spreadsheet.
2. Pilih kolom sumber.
3. Pilih kolom hasil terjemahan.
4. Jalankan translasi.
5. Download hasil.

Semua proses harus dapat dilakukan dengan **interaksi seminimal mungkin**.

Prinsip utamanya adalah:
- spreadsheet-first workflow,
- minimal friction,
- fast column selection,
- visual feedback yang jelas,
- minim dropdown untuk tugas yang sebenarnya bisa dilakukan langsung dari preview tabel.

---

## 2. Build yang Direkomendasikan

Build yang paling cocok untuk konsep ini adalah:

- **PySide6** sebagai shell aplikasi desktop.
- **QWebEngineView** untuk menampilkan UI berbasis HTML/CSS/JS di dalam aplikasi desktop.
- **Qt WebChannel** untuk komunikasi dua arah antara Python dan JavaScript.
- **Python** sebagai backend utama untuk logic aplikasi.
- **HTML/CSS/JS** untuk layer UI agar desain modern, fleksibel, dan mudah dikembangkan.
- Packaging output: **Windows `.exe`** sebagai target utama.

### Alasan Pemilihan Build

Konsep aplikasi ini sangat cocok dengan pendekatan hybrid desktop-web karena:

- UI yang diinginkan sangat visual dan modern.
- Layout terdiri dari banyak panel interaktif.
- Preview spreadsheet menjadi komponen utama.
- Klik header kolom adalah inti UX.
- Logic aplikasi sudah sangat cocok dijalankan di Python.
- Python lebih nyaman untuk Excel, request API, retry, batching, vault, dan workflow processing.

Pendekatan terbaik adalah:

- **Desktop shell tetap native**
- **UI dibuat seperti web modern**
- **Sistem utama tetap Python**

Dengan model ini, aplikasi tetap terasa seperti desktop app, tetapi pengembangan UI menjadi jauh lebih fleksibel.

---

## 3. Arsitektur Aplikasi

Arsitektur aplikasi dibagi menjadi 4 lapisan utama:

```
Desktop Shell (PySide6)
    └── Main Window
        └── QWebEngineView
            └── HTML/CSS/JS Frontend
                └── Qt WebChannel Bridge
                    └── Python Core Services
```

### 3.1 Desktop Shell

Lapisan ini bertanggung jawab untuk:
- window utama,
- lifecycle aplikasi,
- native file dialog,
- native window behavior,
- host untuk WebEngine,
- setup theme dasar,
- packaging menjadi executable.

### 3.2 Frontend UI

Lapisan ini menggunakan:
- HTML untuk struktur UI,
- CSS untuk layout dan styling,
- JavaScript untuk interaksi dan state frontend.

Bagian ini menangani:
- sidebar,
- panel translator,
- preview spreadsheet,
- log terminal,
- chat,
- history,
- API manager,
- profile selector,
- theme switching.

### 3.3 Bridge Layer

Bridge layer menghubungkan JavaScript di frontend dengan object Python di backend.

Lapisan ini digunakan untuk:
- load preview file,
- fetch model dari provider,
- start translasi,
- kirim log realtime,
- update progress bar,
- save/load profile,
- save/load config,
- history data.

### 3.4 Python Core Services

Ini adalah inti logic aplikasi.

Service utama:
- `VaultService`
- `ProviderService`
- `ExcelService`
- `TranslationService`
- `HistoryService`
- `ConfigService`

---

## 4. Konsep UI Utama

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
│ Trans    │ Settings / Chat     │ Table      │
│ Chat     │ / History           │ Preview    │
│ Hist     │                     │            │
│ ApiMgr   │                     │            │
├──────────┴─────────────────────┴────────────┤
│ Log Terminal                                │
└─────────────────────────────────────────────┘
```

### Prinsip Desain Layout

- Fokus utama ada pada workflow spreadsheet.
- Preview file harus selalu terlihat.
- Navigasi harus cepat.
- Pengaturan utama harus mudah ditemukan.
- Panel log harus selalu tersedia saat proses berjalan.

---

## 5. Sidebar Navigasi

Sidebar berisi menu utama aplikasi:

- Translator
- API Manager
- Chat
- History

Karakteristik sidebar:
- icon-only,
- minimalis,
- selalu terlihat,
- tidak collapsible.

---

## 6. Main Panel

Main Panel adalah area tengah yang berubah berdasarkan tab aktif.

Tab yang ditampilkan:
- Translator
- API Manager
- Chat
- History

Fokus utama tetap pada **Translator Tab**, karena itu adalah workflow inti aplikasi.

---

## 7. File Preview Panel

Panel kanan menampilkan **preview spreadsheet secara realtime**.

### Fitur Utama Preview
- menampilkan hingga 1000 baris,
- scroll horizontal,
- nomor baris,
- header kolom interaktif,
- pencarian/filter baris,
- visual role indicator per kolom.

---

## 8. Interaksi Kolom

Pengguna **tidak memilih kolom dari dropdown** sebagai alur utama.

Sebaliknya, pengguna **langsung klik header kolom pada tabel**.

Saat header kolom diklik, muncul menu kecil:

```
Set Column Role
─  Source Column
─  Translation Column
─  Context Column
─  Repair Column
```

### Indikator Visual Setelah Role Dipilih

Kolom akan memiliki:
- warna khusus,
- badge/label kecil di header,
- ringkasan role di atas tabel.

Contoh:

```
Column B → Source
Column C → Translation
Column D → Context
Column E → Repair
```

---

## 9. Indicator Role Kolom

Di atas tabel ditampilkan ringkasan role kolom:

```
Source: B     Translation: C     Context: D     Repair: E
```

### Aturan Visual

| Role        | Warna  |
|-------------|--------|
| Source      | Biru   |
| Translation | Hijau  |
| Context     | Ungu   |
| Repair      | Oranye |

---

## 10. Panel Log

Panel bawah menampilkan log proses aplikasi.

Karakteristik:
- dapat di-resize,
- auto-scroll,
- warna log berbeda,
- font monospace,
- memiliki progress bar,
- memiliki tombol clear log,
- memiliki toggle auto-scroll.

| Jenis Log | Keterangan         |
|-----------|--------------------|
| INFO      | Informasi umum     |
| SUCCESS   | Proses berhasil    |
| WARNING   | Peringatan         |
| ERROR     | Kesalahan proses   |

---

## 11. Tab Translator

Tab ini adalah pusat workflow utama.

### 11.1 Provider & Model
- Provider selector
- API key selector
- Model selector

### 11.2 File Input
- Drag and drop
- File picker

### 11.3 Column Configuration
Konfigurasi kolom utamanya dilakukan dari preview tabel. Tersedia juga input manual sebagai fallback:
- Source column
- Translation column
- Context column
- Repair column

### 11.4 Language Configuration
Contoh:
- `EN → ID`
- `JP → EN`
- `AUTO → EN`

### 11.5 Prompt Configuration
1. Translation Prompt
2. Repair Prompt

### 11.6 Batch & Retry
- Batch size
- Thread count
- Retry configuration
- Delay antar batch
- Skip filled rows

### 11.7 LLM Parameters
- Temperature
- Max completion tokens

### 11.8 Action Buttons
- Start Translate
- Stop Process
- Save Config
- Load Config
- Open Output Folder

---

## 12. Mode Operasi

| Mode                   | Deskripsi                                                   |
|------------------------|-------------------------------------------------------------|
| Translate              | Terjemahan dasar dari source ke translation column          |
| Translate + Context    | Menggunakan context column untuk membantu kualitas translasi |
| Repair                 | Memperbaiki hasil translasi yang gagal atau kurang akurat   |

---

## 13. Format File

**Input:**
- XLSX
- CSV

**Output:**

```
original_filename_translated.xlsx
```

File asli tidak akan di-overwrite.

---

## 14. Chat Interface

Tab chat memungkinkan interaksi langsung dengan LLM.

Fitur:
- bubble chat,
- markdown rendering,
- code block,
- tombol copy,
- pengaturan parameter LLM secara langsung.

**Fungsi Chat Tab:**
- tes model sebelum translasi besar,
- cek prompt,
- eksperimen output,
- bantu user memahami perilaku model.

---

## 15. History

Tab history menyimpan aktivitas translasi.

Informasi yang ditampilkan:
- timestamp,
- nama file,
- mode,
- jumlah baris berhasil / gagal,
- provider & model,
- durasi proses.

---

## 16. API Manager

Panel untuk mengelola API key.

Fitur:
- tambah, edit, hapus API key,
- assign key ke provider tertentu.

API key disimpan secara **terenkripsi**.

**Provider awal yang didukung:**
- Google Gemini
- Groq
- Cerebras

Arsitektur backend dibuat extensible untuk:
- OpenAI-compatible endpoint,
- local model provider,
- provider baru lainnya.

---

## 17. Sistem Profil User

Saat pertama kali membuka aplikasi, pengguna membuat profil.

Profil berisi:
- username & password,
- foto profil,
- konfigurasi aplikasi,
- API key,
- history terkait user.

---

## 18. Save Configuration

Konfigurasi yang disimpan:
- provider & model,
- kolom,
- bahasa,
- batch settings,
- prompt,
- temperature & max completion tokens,
- mode operasi.

Fitur:
- Save current config
- Load saved config
- Duplicate config
- Delete config

---

## 19. Theme

| Mode       | Default          |
|------------|------------------|
| Dark mode  | ✓ (ikuti sistem) |
| Light mode | ✓                |

---

## 20. Prinsip Desain UI

- Minimal friction
- Spreadsheet-first workflow
- Fast column selection
- Minimal dropdown usage
- Visual feedback yang jelas
- Fokus pada kecepatan workflow
- Desain modern seperti developer tools

---

## 21. Arsitektur Service Backend

### 21.1 VaultService
- simpan API key terenkripsi,
- login, create, load, update, delete profile.

### 21.2 ProviderService
- ambil daftar model,
- validasi API key,
- memanggil provider LLM,
- abstraction antar provider.

### 21.3 ExcelService
- load file XLSX/CSV,
- preview row,
- baca header,
- mapping kolom,
- write output file,
- export hasil.

### 21.4 TranslationService
- jalankan workflow translate / translate+context / repair,
- retry logic,
- batching,
- progress event,
- stop/cancel process.

### 21.5 HistoryService
- simpan, ambil, filter history,
- load config dari history.

### 21.6 ConfigService
- save, load, delete config,
- manage config per profile.

---

## 22. Struktur Folder Project

```
universal-excel-translator/
├─ main.py
├─ app/
│  ├─ shell/
│  │  ├─ main_window.py
│  │  ├─ web_engine.py
│  │  ├─ resources.qrc
│  │  └─ theme_manager.py
│  ├─ bridge/
│  │  ├─ app_bridge.py
│  │  ├─ translator_bridge.py
│  │  ├─ preview_bridge.py
│  │  ├─ vault_bridge.py
│  │  ├─ history_bridge.py
│  │  └─ chat_bridge.py
│  ├─ core/
│  │  ├─ models/
│  │  │  ├─ job_config.py
│  │  │  ├─ user_profile.py
│  │  │  ├─ history_record.py
│  │  │  └─ provider_config.py
│  │  ├─ services/
│  │  │  ├─ vault_service.py
│  │  │  ├─ provider_service.py
│  │  │  ├─ excel_service.py
│  │  │  ├─ translation_service.py
│  │  │  ├─ history_service.py
│  │  │  └─ config_service.py
│  │  ├─ llm/
│  │  │  ├─ base_provider.py
│  │  │  ├─ gemini_provider.py
│  │  │  ├─ groq_provider.py
│  │  │  ├─ cerebras_provider.py
│  │  │  └─ openai_compatible_provider.py
│  │  ├─ processing/
│  │  │  ├─ engine.py
│  │  │  ├─ rotators.py
│  │  │  ├─ prompts.py
│  │  │  └─ response_parser.py
│  │  └─ utils/
│  │     ├─ file_utils.py
│  │     ├─ lang_utils.py
│  │     └─ crypto_utils.py
│  └─ web/
│     ├─ index.html
│     ├─ css/
│     │  ├─ app.css
│     │  ├─ dark.css
│     │  └─ light.css
│     ├─ js/
│     │  ├─ app.js
│     │  ├─ bridge.js
│     │  ├─ state.js
│     │  ├─ translator.js
│     │  ├─ preview.js
│     │  ├─ logpanel.js
│     │  ├─ chat.js
│     │  ├─ history.js
│     │  └─ apimanager.js
│     └─ assets/
│        ├─ icons/
│        └─ fonts/
├─ user_data/
├─ tests/
├─ requirements.txt
└─ README.md
```

---

## 23. Contoh Implementasi Komponen Utama

### 23.1 Python Bridge

```python
from PySide6.QtCore import QObject, Signal, Slot

class TranslatorBridge(QObject):
    logReceived = Signal(str, str)
    progressChanged = Signal(int, int, str)
    previewLoaded = Signal(list)
    modelsLoaded = Signal(list)
    processFinished = Signal(bool, str)

    def __init__(self, translator_service):
        super().__init__()
        self.translator_service = translator_service

    @Slot(str, result='QVariantList')
    def loadPreview(self, file_path):
        try:
            rows = self.translator_service.load_preview(file_path, limit=1000)
            self.previewLoaded.emit(rows)
            return rows
        except Exception as e:
            self.logReceived.emit("ERROR", str(e))
            return []

    @Slot(str, str, result='QVariantList')
    def fetchModels(self, provider, profile_id):
        try:
            models = self.translator_service.fetch_models(provider, profile_id)
            self.modelsLoaded.emit(models)
            return models
        except Exception as e:
            self.logReceived.emit("ERROR", str(e))
            return []

    @Slot('QVariantMap')
    def startTranslation(self, payload):
        self.translator_service.start_job(payload, self)

    def emit_log(self, level, message):
        self.logReceived.emit(level, message)

    def emit_progress(self, current, total, label):
        self.progressChanged.emit(current, total, label)

    def emit_finished(self, ok, message):
        self.processFinished.emit(ok, message)
```

### 23.2 Main Window

```python
from PySide6.QtWidgets import QMainWindow
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebChannel import QWebChannel

class MainWindow(QMainWindow):
    def __init__(self, translator_bridge, app_bridge):
        super().__init__()
        self.setWindowTitle("Universal Excel Translator")
        self.resize(1600, 950)

        self.view = QWebEngineView()
        self.channel = QWebChannel()

        self.channel.registerObject("translatorBridge", translator_bridge)
        self.channel.registerObject("appBridge", app_bridge)

        self.view.page().setWebChannel(self.channel)
        self.view.setHtml(open("app/web/index.html", "r", encoding="utf-8").read())

        self.setCentralWidget(self.view)
```

### 23.3 JavaScript Bridge Init

```html
<script src="qrc:///qtwebchannel/qwebchannel.js"></script>
<script>
document.addEventListener("DOMContentLoaded", () => {
  new QWebChannel(qt.webChannelTransport, function(channel) {
    window.translatorBridge = channel.objects.translatorBridge;
    window.appBridge = channel.objects.appBridge;

    translatorBridge.logReceived.connect((level, message) => {
      addLog(level, message);
    });

    translatorBridge.progressChanged.connect((current, total, label) => {
      updateProgress(current, total, label);
    });

    translatorBridge.previewLoaded.connect((rows) => {
      renderPreview(rows);
    });

    translatorBridge.modelsLoaded.connect((models) => {
      renderModelOptions(models);
    });

    translatorBridge.processFinished.connect((ok, message) => {
      showProcessResult(ok, message);
    });
  });
});
</script>
```

---

## 24. Diagram Alur Interaksi User

```
Open App
  ↓
Select / Login Profile
  ↓
Open Translator Tab
  ↓
Load XLSX / CSV
  ↓
Spreadsheet Preview Appears
  ↓
Click Column Header → Assign Role
  ↓
Choose Provider / API Key / Model
  ↓
Set Language + Prompt + Batch Config
  ↓
Start Translation
  ↓
Log + Progress Update Realtime
  ↓
Output File Generated
  ↓
History Saved
```

---

## 25. Diagram Alur Sistem

```
Frontend UI (HTML/CSS/JS)
        ↓
Qt WebChannel Bridge
        ↓
Python Bridge Objects
        ↓
Core Services
        ├── VaultService
        ├── ProviderService
        ├── ExcelService
        ├── TranslationService
        ├── HistoryService
        └── ConfigService
        ↓
Output / Saved Data / Logs
```

---

## 26. Prioritas Pengembangan

| Phase   | Fitur                                                               |
|---------|---------------------------------------------------------------------|
| Phase 1 | Shell PySide6, WebEngine + WebChannel, Translator Tab, File Preview, Klik header kolom, Start translation, Log panel |
| Phase 2 | API Manager, Sistem profil user, Vault terenkripsi                  |
| Phase 3 | History, Save / Load Config                                         |
| Phase 4 | Chat interface                                                      |
| Phase 5 | Local model support, Provider plugin system, UI polish, Optimasi packaging |

---

## 27. Kesimpulan

Konsep final aplikasi ini adalah:

- **Desktop app modern** dengan UI berbasis HTML/CSS/JS
- **Backend Python** dengan shell PySide6
- **Komunikasi** melalui Qt WebChannel
- **Fokus utama** pada spreadsheet-first workflow
- **Minim friction** — cepat, visual, dan mudah digunakan

Pendekatan ini adalah titik tengah terbaik antara kenyamanan Python untuk backend, fleksibilitas UI berbasis web, dan pengalaman aplikasi desktop yang rapi. Dengan build ini, aplikasi bisa terasa modern, tetap kuat secara teknis, mudah dipaketkan jadi `.exe`, dan fleksibel untuk dikembangkan ke fitur lanjutan.
