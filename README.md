# 🏷️ IMDB Auto-Fill Tool
**GDSS-Maverick Hackathon 2026** · Premium AI-Driven Image-to-Item Master Data Pipeline

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://gdss-hackathon-aw2swnadk2wp8eka4nmb2k.streamlit.app/)

## 🌐 Live Demo
Visit the production app here: [GDSS Hackathon Live App](https://gdss-hackathon-aw2swnadk2wp8eka4nmb2k.streamlit.app/)

---

## ⚡ Key Features

*   **🧠 Multi-Model Vision AI**: Extracts 13 critical Item Master attributes from packaging images using **Groq (Llama 3.2/4 Vision)** and **OpenRouter (Gemma 2/4)**.
*   **📸 Mobile-Optimized Live Camera**: Capture snapshots directly using your phone/device camera with automatic queue accumulation and prefix-based grouping.
*   **⏱️ Rate-Limit Safe Sequential Queue**: Collects all images into a single flat processing queue with a guaranteed 6-second delay between API requests to strictly prevent free-tier rate limiting.
*   **☁️ Supabase Cloud Sync**: Official `supabase` library integration targeting the `imdb_products` table.
*   **🔄 Smart Duplicate Detection & Merging**:
    *   Finds matching products by barcode (if available) or by `brand + weight + packaging_type` (fallback).
    *   Automatically increments `scan_count`, marks the record `is_duplicate = True`, and merges/fills empty fields from the new scan.
*   **📊 Item Master Database View**:
    *   Interactive view displaying all stored items.
    *   Search filter by brand name or product/item name.
    *   Visual alert highlighting duplicate entries in red.
    *   Testing-ready **Clear Database** action.
*   **📥 Premium Exports**: Downloads custom-formatted Excel spreadsheets (pre-styled with borders and headers) and CSVs.

---

## 📋 The 13 IMDB Columns

| Column | Description | Format Rules |
|---|---|---|
| **ITEM NAME** | Full descriptive retail product name | e.g. `KNORR CHICKEN STOCK CUBE 20G` |
| **BARCODE** | Numeric barcode digits | Digits only (no spaces/dashes) |
| **MANUFACTURER** | Producing company name | e.g. `UNILEVER` |
| **BRAND** | Brand name exactly as shown | e.g. `KNORR` |
| **WEIGHT** | Net weight or volume | No space between number and unit (e.g. `80G`, `1L`) |
| **PACKAGING TYPE** | Standard container category | `SACHET`, `BOTTLE`, `BOX`, `CAN`, `POUCH` etc. |
| **COUNTRY** | Origin country as printed | e.g. `MALAYSIA`, `THAILAND` |
| **VARIANT** | Specific flavor/formula variant | e.g. `ORIGINAL`, `LOW FAT`, `REDUCED SALT` |
| **TYPE** | Short product category | e.g. `MAYONNAISE`, `SEASONING`, `DETERGENT` |
| **FRAGRANCE FLAVOR**| Scent or taste descriptor | e.g. `LEMON`, `CHICKEN`, `VANILLA` |
| **PROMOTION** | On-package promotional text | e.g. `BUY 2 FREE 1`, `20% EXTRA FREE` |
| **ADDONS** | Included extras or bonuses | e.g. `SPOON INCLUDED`, `FREE RECIPE BOOK` |
| **TAGLINE** | Short promotional slogan | e.g. `TASTE THE DIFFERENCE` |

---

## 🛠️ Installation & Setup

### Step 1 — Clone the Repository
```bash
git clone https://github.com/hendrix-llouchi/gdss-hackathon
cd gdss-hackathon
```

### Step 2 — Create and Activate Virtual Environment
```bash
# Create environment
python -m venv .venv

# Activate (Windows)
.venv\Scripts\activate

# Activate (Mac/Linux)
source .venv/bin/activate
```

### Step 3 — Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 4 — Configure Local Secrets
Streamlit utilizes a local secrets file for keys. Create a folder named `.streamlit` and add a `secrets.toml` file inside it:

```toml
# .streamlit/secrets.toml
GROQ_API_KEY = "gsk_your_groq_api_key"
OPENROUTER_API_KEY = "sk-or-v1-your_openrouter_api_key"

# Supabase Credentials
SUPABASE_URL = "https://your-project-id.supabase.co"
SUPABASE_KEY = "your-anon-or-service-key"
```

---

## 🚀 Running the Application

Launch the Streamlit dashboard locally:
```bash
streamlit run app.py
```
The application will launch automatically at `http://localhost:8501`.

---

## 📐 Pipeline Architecture

```
[ Upload Files / Camera Snapshots ]
                 │
                 ▼
     [ Flat Sequential Queue ] (Strict 6s Delay)
                 │
                 ▼
     [ Vision AI Extraction ] (Groq / OpenRouter)
                 │
                 ▼
[ Grouping & Aggregation ] (Majority vote on prefix groups)
                 │
                 ▼
     [ Normalization & Validation ] (Barcodes, Weights, etc.)
                 │
                 ▼
     [ Supabase Sync / Merge ] ──► [ 📊 Item Master Database ]
                 │
                 ▼
         [ Excel/CSV Export ]
```

---

## 📁 Project Structure

```
gdss-hackathon/
├── .streamlit/
│   ├── config.toml                  # UI styling settings
│   └── secrets.toml                 # Local API & DB secrets (Git ignored)
├── app.py                           # Main Streamlit application
├── pipeline.py                      # CLI runner with checkpointing support
├── sample_images/                   # Tester images
├── IMDB_predictions_submission.xlsx  # Hackathon dataset output
├── requirements.txt                 # Dependencies (including supabase client)
├── .gitignore                       # Excluded folders, secrets, and environments
└── README.md                        # Documentation
```