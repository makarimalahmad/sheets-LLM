# Sheets LLM Backend & Dashboard

Backend otomatisasi pembaca data Google Sheets dengan pemetaan semantik nama kolom menggunakan LLM (Groq / `openai/gpt-oss-120b`). Data otomatis distandarisasi ke database SQLite dan disajikan via REST API untuk dashboard perusahaan.

## Fitur Utama
- **Semantic Column Mapping**: Memetakan variasi nama header (misal `"EUP in VAT"` vs `"end user price include vat"`) ke skema standar (`eup_inc_vat`) menggunakan LLM.
- **REST API (FastAPI)**: Endpoint ingest data dari Google Apps Script dan penyaji data agregasi (harga EUP tertinggi, rata-rata margin, total produk).
- **Dashboard Demo Visual**: Tersedia di `/dashboard` untuk simulasi uji coba langsung.

## Struktur Proyek
```text
sheets-llm/
├── app/
│   ├── main.py          # Server FastAPI & endpoints
│   ├── mapper.py        # Logika LLM semantic mapping
│   ├── models.py        # Schema data baku (Pydantic)
│   └── database.py      # SQLite storage & queries
├── dashboard/
│   └── index.html       # Antarmuka web dashboard demo
├── scripts/
│   └── google_apps_script.js # Skrip webhook Google Sheets
├── tests/
│   └── test_simulation.py   # Unit test simulasi
├── requirements.txt
├── run_server.py
└── .env.example
```

## Cara Menjalankan Lokal

1. **Install Dependensi:**
```bash
python -m venv .venv
# Windows:
.\.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate

pip install -r requirements.txt
```

2. **Konfigurasi Environment:**
Salin `.env.example` ke `.env` lalu masukkan API Key:
```env
GROQ_API_KEY=gsk_your_key_here
```

3. **Jalankan Server:**
```bash
python run_server.py
```
- Dashboard: `http://127.0.0.1:8000/dashboard`
- Swagger API Docs: `http://127.0.0.1:8000/docs`
