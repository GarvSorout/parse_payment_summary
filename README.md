## OHIP Payment Summary Parser

Turns scrambled PSSuite/OHIP Payment Summary PDFs into clean, analytics‑ready CSVs.

### What it does
- Decodes the PDF’s scrambled text layer (custom font “caesar shift”).
- OCRs numbers and totals reliably with pdf2image + Tesseract.
- Combines both to produce two structured CSVs:
  - `provider_payments.csv` — per-provider line items and totals
  - `group_payments.csv` — group-level blocks (Group/Exception/All Providers/Summary)
- Also writes human-readable reference text if needed.

### Why this exists
Standard database loaders choke on these PDFs because the text layer isn’t literal text. This tool bypasses that limitation and automates ingestion so finance ops don’t get blocked.

### Prerequisites
- Python 3.10+
- Tesseract OCR (v5 recommended)
- Poppler (for `pdftoppm`, used by `pdf2image`)

macOS (Homebrew):
```bash
brew install tesseract poppler
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows:
- Install Tesseract (add to PATH)
- Install Poppler and set `POPPLER_PATH` or pass it to the script if needed
- Create venv and `pip install -r requirements.txt`

### Quick start
```bash
python3 parse_payment_summary.py "/absolute/path/to/Payment Summary_YYYYMMDD_Group.pdf" \
  --hybrid --renderer pdftoppm --dpi 300 \
  --out-dir "/absolute/path/to/tables_hybrid_YYYYMMDD"
```

Key flags:
- `--hybrid`: use decoded names/structure + OCR numbers (recommended)
- `--renderer`: `pdftoppm` (fast, reliable) or `pdf2image`
- `--dpi`: 300 recommended (balances speed and accuracy)
- `--id-dpi`: optional higher DPI for ID probing

### Outputs (examples)
- `group_payments.csv`
  - Columns: `group_num, period, total_payment_ab_date, total_payment_ab, payment_type, line_item_desc, current_month_amt, year_to_date_amt`
- `provider_payments.csv`
  - Columns: `group_num, period, total_payment_ab_date, total_payment_ab, ohip_number, provider_name, line_item, current_month_amt, year_to_date_amt, section`
- Optional diagnostics: decoded and OCR reference text in the output folder.

### Repo layout
```
.
├── parse_payment_summary.py        # Main parser (hybrid decode + OCR)
├── ocr_pdf_text.py                 # Basic OCR dump for diagnostics
├── inspect_glyphs.py               # Raw glyph inspection (PyMuPDF)
├── requirements.txt
├── README.md
├── .gitignore
└── samples/
    └── .gitkeep                    # No PHI committed by default
```

### What to commit vs ignore
- Commit: code, README, requirements, `.gitignore`, empty `samples/`.
- Do not commit: PDFs with PHI, generated CSVs, OCR images/TSV/MD/TXT artifacts.

### Troubleshooting
- Tesseract not found: ensure it’s installed and on PATH.
- Poppler missing: install and ensure `pdftoppm` is available.
- Numbers missing: raise DPI to 300–400; hybrid mode must be enabled.

### Automation
Designed to be called by n8n / Power Automate on file arrival and push CSVs to Snowflake storage.


