# 🏷️ IMDB Auto-Fill Tool
**GDSS-Maverick Hackathon 2026** · AI-Driven Image-to-Item Master Data

## Overview
This tool accepts product images, automatically extracts 13 Item Master Database (IMDB) attributes using a vision AI model, and exports a structured Excel/CSV file ready for database upload. It eliminates manual product data entry for retail teams.

## The 13 IMDB Columns
| Column | Description |
|---|---|
| ITEM NAME | Full descriptive product name |
| BARCODE | Numeric barcode digits only |
| MANUFACTURER | Company that manufactures the product |
| BRAND | Brand name as shown on package |
| WEIGHT | Net weight/volume e.g. 80G, 500ML |
| PACKAGING TYPE | SACHET, BOTTLE, BOX, CAN, POUCH etc. |
| COUNTRY | Country of manufacture/origin |
| VARIANT | e.g. ORIGINAL, LOW FAT, ROSE |
| TYPE | Short category e.g. SOAP, DETERGENT |
| FRAGRANCE FLAVOR | e.g. LEMON, ROSE, CAPPUCCINO |
| PROMOTION | Any on-pack promotion text |
| ADDONS | Additional features e.g. SPOON INCLUDED |
| TAGLINE | Short promotional slogan |

## Tech Stack
- Python 3.10+
- Streamlit (UI)
- Groq API — Llama 4 Scout 17B Vision (primary extraction engine)
- OpenRouter — Gemma 4 26B (fallback engine)
- Pillow (image preprocessing)
- openpyxl (Excel export)

## Prerequisites
- Python 3.10 or higher installed
- A Groq API key (free at console.groq.com — no credit card needed)
- An OpenRouter API key (free at openrouter.ai — no credit card needed)

## Installation

### Step 1 — Clone the repo
```bash
git clone https://github.com/hendrix-llouchi/gdss-hackathon
cd gdss-hackathon
```

### Step 2 — Create a virtual environment
```bash
python -m venv .venv
```

### Step 3 — Activate the virtual environment
**Windows:**
```bash
.venv\Scripts\activate
```
**Mac/Linux:**
```bash
source .venv/bin/activate
```

### Step 4 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 5 — Add your API keys
Open `app.py` and find the config block at the top. Add your keys:
```python
GROQ_API_KEY = "your_groq_api_key_here"
OPENROUTER_API_KEY = "your_openrouter_api_key_here"
```

## Running the App
```bash
streamlit run app.py
```
The app will open automatically at `http://localhost:8501`

## Using the Tool
1. **Upload** — drag and drop product images (JPG/PNG). Multiple images of the same product are grouped automatically by filename prefix
2. **Extract** — click 🚀 Run Pipeline. The AI model processes each image and extracts the 13 fields
3. **Preview & Edit** — review the extracted data in the table. Click any cell to correct it before exporting
4. **Export** — download as Excel (.xlsx) or CSV (.csv)

## Pipeline Architecture
```
Image Upload → Preprocessing (resize, base64)
             → AI Extraction (Groq/OpenRouter vision model)
             → Aggregation (majority vote across multiple images)
             → Validation & Normalization (barcodes, weights, country)
             → Export (Excel/CSV)
```

## Rate Limits (Free Tier)
- Groq: 30 requests/minute, 14,400 requests/day
- OpenRouter: 50 requests/day on free models
- The app handles rate limits automatically with retry logic

## Sample Output
See `IMDB_predictions_submission.xlsx` for the full 48-product output generated from the hackathon dataset using Gemma 4 vision model.

## Project Structure
```
gdss-hackathon/
├── app.py                          # Streamlit UI + pipeline
├── pipeline.py                     # CLI pipeline with checkpoint support
├── sample_images/                  # Sample product images for testing
├── IMDB_predictions_submission.xlsx # Full 48-product output
├── .env.example                    # API key template
├── .gitignore                      # Excludes .env and venv
└── README.md
```

## API Keys for Judges
Groq and OpenRouter API keys are provided separately in the submission form as required by the hackathon guidelines.