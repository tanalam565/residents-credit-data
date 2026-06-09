# credit-processor

Extracts structured data from TransUnion PDF credit reports and HTML Screening Reports, appends to Excel.

## Project Structure

```
credit-processor/
├── backend/
│   ├── app.py              # FastAPI web app
│   ├── main.py             # CLI batch processor
│   ├── credit_report.py    # Extraction pipeline
│   └── requirements.txt
├── frontend/
│   └── index.html
└── requirements.txt
```

## Setup

```bash
pip install -r requirements.txt
```

Create a `.env` file in `backend/`:

```env
AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
AZURE_DOCUMENT_INTELLIGENCE_KEY=your-key

AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_API_KEY=your-key
AZURE_OPENAI_DEPLOYMENT_NAME=gpt-4o
AZURE_OPENAI_API_VERSION=2025-04-01-preview

# Optional: custom Excel output path
# EXCEL_PATH=/data/credit_reports.xlsx
```

## Run (Web App)

```bash
cd backend
uvicorn app:app --reload --port 8000
```

Open http://localhost:8000

## Run (CLI)

```bash
cd backend
python main.py report1.pdf report2.htm -o credit_reports.xlsx
```

## Supported Formats

- TransUnion PDF credit reports
- HTML Screening Reports (Adara / First Advantage format)
- Images: PNG, JPG, TIFF, BMP (OCR via Azure Document Intelligence)

## Extracted Fields

See `credit_report.py` → `CREDIT_REPORT_SCHEMA` for the full list (~55 fields including):

- Identity & fraud indicators (SSN match, AKA count, address conflict)
- Credit score + key factors
- Record counts (tradelines, collections, public records, inquiries)
- Derogatory items
- Tradeline summary by type (revolving/installment/mortgage/open)
- Late payment totals (30/60/90/120+ day)
- Individual tradelines, collections, inquiries (stored as JSON)
- HTML screening: decision, criteria pass/fail, criminal/eviction/sanctions flags
- Derived flags: `rapid_tradeline_flag`, `aka_count`