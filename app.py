"""
GDSS-Maverick Hackathon
AI-Driven Image-to-IMDB Tool
Streamlit UI
"""

import streamlit as st
import streamlit.components.v1 as components
import json
import re
import base64
import requests
import io
import os
import time
from pathlib import Path
from PIL import Image
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="IMDB Auto-Fill Tool",
    page_icon="🏷️",
    layout="wide"
)

# ─────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400..900;1,400..900&family=Outfit:wght@300;400;500;600;700;800&display=swap');

    .stApp, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        background-color: #f8fafc !important;
        color: #475569 !important;
        font-family: 'Outfit', sans-serif !important;
    }
    .main {
        background-color: transparent !important;
    }
    .block-container {
        padding-top: 6rem !important;
    }

    /* Typography */
    h1, h2, h3, h4, h5, h6, [data-testid="stHeader"] {
        font-family: 'Playfair Display', Georgia, serif !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    
    p, label, caption, small {
        font-family: 'Outfit', sans-serif !important;
        color: #475569 !important;
    }

    /* Sidebar elegant light theme */
    section[data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0 !important;
        box-shadow: 4px 0 20px rgba(0, 0, 0, 0.01) !important;
        z-index: 1000 !important; /* stay below our navbar (z-index: 99999) */
    }
    section[data-testid="stSidebar"] > div {
        background-color: #ffffff !important;
    }
    /* Push sidebar content below the floating navbar capsule */
    section[data-testid="stSidebar"] > div:first-child {
        padding-top: 72px !important;
    }
    section[data-testid="stSidebar"] h3 {
        font-size: 1.1rem !important;
        color: #0f172a !important;
        margin-top: 1rem !important;
    }
    /* ── Hide Streamlit's native sidebar toggle buttons ──
       Moved off-screen (NOT display:none or opacity:0) so they remain
       fully rendered and React's .click() always works. */
    [data-testid="stSidebarCollapsedControl"] {
        position: fixed !important;
        left: -9999px !important;
        top: -9999px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] {
        position: fixed !important;
        left: -9999px !important;
        top: -9999px !important;
    }

    /* Input Fields & Widgets */
    div[data-testid="stTextInput"] [data-baseweb="input"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        height: 42px !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }
    div[data-testid="stTextInput"] input {
        border: none !important;
        background-color: transparent !important;
        color: #0f172a !important;
        font-family: 'Outfit', sans-serif !important;
        height: 100% !important;
        box-shadow: none !important;
        outline: none !important;
    }
    div[data-testid="stSelectbox"] [data-baseweb="select"] {
        background-color: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        color: #0f172a !important;
        border-radius: 10px !important;
        font-family: 'Outfit', sans-serif !important;
        height: 42px !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }
    /* Target selectbox value wrapper to ensure it is white-bg and has dark text */
    div[data-testid="stSelectbox"] [data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #0f172a !important;
    }
    /* Ensure child text and spans inside the selectbox also use the dark color */
    div[data-testid="stSelectbox"] [data-baseweb="select"] * {
        color: #0f172a !important;
    }
    div[data-testid="stTextInput"] [data-baseweb="input"]:hover,
    div[data-testid="stSelectbox"] [data-baseweb="select"]:hover {
        border-color: #94a3b8 !important;
    }
    div[data-testid="stTextInput"] [data-baseweb="input"]:focus-within,
    div[data-testid="stSelectbox"] [data-baseweb="select"]:focus {
        border-color: #0f172a !important;
        box-shadow: 0 0 0 1px #0f172a !important;
    }
    /* Placeholder colors */
    div[data-testid="stTextInput"] input::placeholder {
        color: #94a3b8 !important;
        opacity: 1 !important;
    }
    /* Selectbox dropdown items and portals */
    [data-baseweb="popover"] ul,
    [data-baseweb="menu"],
    div[role="listbox"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
    }
    [data-baseweb="popover"] ul li,
    [data-baseweb="menu"] li,
    div[role="option"],
    li[role="option"] {
        background-color: #ffffff !important;
        color: #0f172a !important;
        padding: 0.5rem 1rem !important;
        transition: background-color 0.15s ease !important;
    }
    [data-baseweb="popover"] ul li:hover,
    [data-baseweb="menu"] li:hover,
    div[role="option"]:hover,
    li[role="option"]:hover {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
    }
    div[data-testid="stSelectbox"] [data-baseweb="select"],
    div[role="option"],
    li[role="option"],
    [data-baseweb="menu"] * {
        cursor: pointer !important;
    }
    div[data-testid="stTextInput"] label, div[data-testid="stSelectbox"] label {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        color: #475569 !important;
        font-size: 0.85rem !important;
    }

    .nav-wrapper {
        container-type: inline-size;
        width: 100%;
    }

    /* Top Navigation Capsule */
    .nav-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: rgba(255, 255, 255, 0.55) !important;
        -webkit-backdrop-filter: blur(40px) saturate(200%) !important;
        backdrop-filter: blur(40px) saturate(200%) !important;
        border: 1px solid rgba(255, 255, 255, 0.75) !important;
        border-radius: 24px !important;
        padding: 0.6rem 1.75rem !important;
        box-shadow: 
            0 20px 40px rgba(15, 23, 42, 0.06), 
            0 5px 15px rgba(15, 23, 42, 0.04), 
            inset 0 1px 0 rgba(255, 255, 255, 0.7) !important;
        position: fixed !important;
        top: 12px !important;
        left: 50% !important;
        transform: translateX(-50%) !important;
        width: 92% !important;
        max-width: 1100px !important;
        z-index: 99999 !important;
        margin: 0 !important;
    }
    /* Mobile/Small Container: compact single-row navbar */
    @container (max-width: 850px) {
        .nav-container {
            width: 96% !important;
            padding: 0.4rem 0.75rem !important;
            border-radius: 16px !important;
        }
        .nav-logo {
            font-size: 0.85rem !important;
            margin-right: 0.5rem !important;
        }
        .nav-hamburger {
            width: 32px !important;
            height: 32px !important;
            margin-right: 0.5rem !important;
            border-radius: 8px !important;
        }
        /* Hide step text labels on mobile/small containers, show only the circled numbers */
        .nav-tab {
            font-size: 0 !important;
        }
        .step-num {
            font-size: 0.65rem !important;
            margin-right: 0 !important;
        }
        .nav-center {
            gap: 0.5rem !important;
        }
        .nav-badge {
            font-size: 0.65rem !important;
            padding: 0.2rem 0.5rem !important;
        }
    }

    /* Tablet & Medium Containers: slightly more compact navbar to prevent overflow */
    @container (min-width: 851px) and (max-width: 1100px) {
        .nav-container {
            padding: 0.5rem 1rem !important;
        }
        .nav-center {
            gap: 0.8rem !important;
        }
        .nav-tab {
            font-size: 0.82rem !important;
        }
        .step-num {
            width: 18px !important;
            height: 18px !important;
            font-size: 0.68rem !important;
            margin-right: 5px !important;
        }
        .nav-logo {
            font-size: 1.1rem !important;
        }
        .nav-badge {
            font-size: 0.75rem !important;
            padding: 0.3rem 0.6rem !important;
        }
    }

    /* Maintain viewport-based padding for the main content block container */
    @media (max-width: 950px) {
        .block-container {
            padding-top: 4.5rem !important;
        }
    }
    /* Hide the zero-height iframe container used for JS injection so it doesn't add blank space */
    div.element-container:has(iframe[height="0"]) {
        display: none !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    /* Hide Streamlit's ⋮ mobile menu button */
    [data-testid="stMainMenuButton"],
    button[data-testid="baseButton-header"] {
        display: none !important;
    }
    /* Hamburger toggle button inside the navbar capsule */
    .nav-hamburger {
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 40px !important;
        height: 40px !important;
        background: #0f172a !important;
        border-radius: 12px !important;
        cursor: pointer !important;
        margin-right: 1.25rem !important;
        flex-shrink: 0 !important;
        transition: background 0.2s ease, transform 0.15s ease !important;
        user-select: none !important;
    }
    .nav-hamburger:hover {
        background: #1e293b !important;
        transform: scale(1.05) !important;
    }
    /* Hide Streamlit's native header bar (Deploy btn, ⋮ menu) */
    header[data-testid="stHeader"] {
        display: none !important;
    }
    .nav-left {
        display: flex;
        align-items: center;
    }
    .nav-logo {

        font-family: 'Playfair Display', Georgia, serif !important;
        font-weight: 800 !important;
        font-size: 1.25rem !important;
        color: #0f172a !important;
    }
    .nav-center {
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .nav-tab {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.9rem !important;
        font-weight: 500 !important;
        color: #000000 !important;
        opacity: 0.55 !important;
        padding: 0.5rem 0.25rem !important;
        cursor: pointer;
        position: relative;
        transition: all 0.2s ease !important;
        display: inline-flex !important;
        align-items: center !important;
    }
    .nav-tab:hover {
        opacity: 0.95 !important;
    }
    .active-tab {
        color: #000000 !important;
        font-weight: 700 !important;
        opacity: 1 !important;
    }
    .active-tab::after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        right: 0;
        height: 2.5px;
        background-color: #000000;
        border-radius: 2px;
    }
    .step-num {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 20px !important;
        height: 20px !important;
        background-color: rgba(0, 0, 0, 0.08) !important;
        color: #000000 !important;
        border-radius: 50% !important;
        font-size: 0.72rem !important;
        font-weight: 600 !important;
        margin-right: 8px !important;
        transition: all 0.2s ease !important;
        opacity: 0.8 !important;
    }
    .active-num {
        background-color: #000000 !important;
        color: #ffffff !important;
        opacity: 1 !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2) !important;
    }
    .nav-right {
        display: flex;
        align-items: center;
        gap: 1rem;
    }
    .nav-badge {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        background-color: #f1f5f9 !important;
        padding: 0.35rem 0.75rem !important;
        border-radius: 20px !important;
        display: flex;
        align-items: center;
        gap: 0.25rem;
        border: 1px solid #e2e8f0 !important;
    }
    .nav-profile {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        background: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        padding: 0.35rem 0.85rem !important;
        border-radius: 20px !important;
    }
    .nav-profile-name {
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #0f172a !important;
    }

    /* Cards & Containers */
    div[data-testid="stMetric"], 
    div[data-testid="stAlert"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 20px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.02) !important;
        color: #475569 !important;
        transition: all 0.3s ease !important;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 15px 35px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stMetric"] {
        padding: 1.25rem !important;
    }
    div[data-testid="stMetricLabel"] > div {
        font-family: 'Outfit', sans-serif !important;
        text-transform: uppercase !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.05em !important;
        color: #64748b !important;
    }
    div[data-testid="stMetricValue"] > div {
        font-family: 'Playfair Display', Georgia, serif !important;
        font-size: 2.25rem !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        margin-top: 0.25rem !important;
    }

    /* Expander & Bordered Containers */
    div[data-testid="stExpander"],
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.02) !important;
        color: #475569 !important;
        transition: all 0.3s ease !important;
        overflow: hidden !important;
        margin-bottom: 1.5rem !important;
    }
    div[data-testid="stExpander"]:hover,
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        border-color: #cbd5e1 !important;
        box-shadow: 0 12px 35px rgba(0, 0, 0, 0.03) !important;
    }

    /* Expander Header (Modern Streamlit uses summary element) */
    div[data-testid="stExpander"] summary,
    .streamlit-expanderHeader {
        background-color: #ffffff !important;
        border-bottom: none !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        color: #0f172a !important;
        padding: 1rem 1.5rem !important;
        transition: background-color 0.2s ease !important;
    }
    /* Expander Header Hover */
    div[data-testid="stExpander"] summary:hover {
        background-color: #f8fafc !important;
    }
    /* Target the text wrapper inside summary */
    div[data-testid="stExpander"] summary > div[data-testid="stMarkdownContainer"] p {
        color: #0f172a !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
    }
    /* Target the toggle icon */
    div[data-testid="stExpander"] summary > svg[data-testid="stExpanderToggleIcon"] {
        fill: #0f172a !important;
    }

    /* Expander Content */
    div[data-testid="stExpander"] div[data-testid="stExpanderDetails"],
    .streamlit-expanderContent {
        background-color: #ffffff !important;
        border-top: 1px solid #f1f5f9 !important;
        padding: 1.5rem !important;
    }

    /* Buttons */
    div.stButton > button, div.stDownloadButton > button {
        background-color: #0f172a !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.65rem 1.75rem !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08) !important;
        transition: all 0.2s ease !important;
        width: 100% !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        background-color: #1e293b !important;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.12) !important;
        transform: translateY(-1px) !important;
    }
    div.stButton > button:active, div.stDownloadButton > button:active {
        transform: translateY(0px) !important;
    }

    /* File Uploader dashed zone */
    div[data-testid="stFileUploader"] section {
        border: 2px dashed #cbd5e1 !important;
        background-color: #ffffff !important;
        border-radius: 20px !important;
        padding: 2rem !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.005) !important;
    }
    div[data-testid="stFileUploader"] section:hover {
        border-color: #0f172a !important;
        background-color: #f8fafc !important;
    }
    /* Fix file uploader text elements that default to white/invisible in dark theme base */
    div[data-testid="stFileUploader"] section span,
    div[data-testid="stFileUploader"] section p,
    div[data-testid="stFileUploader"] section small,
    div[data-testid="stFileUploader"] section div {
        color: #475569 !important;
    }
    /* Ensure the Upload button inside file uploader keeps its dark background and white text */
    div[data-testid="stFileUploader"] section button,
    div[data-testid="stFileUploader"] section button * {
        color: #ffffff !important;
        background-color: #0f172a !important;
    }
    div[data-testid="stFileUploader"] section button:hover {
        background-color: #1e293b !important;
    }

    /* Table & Dataframe */
    div[data-testid="stDataFrame"] {
        border: 1px solid #e2e8f0 !important;
        border-radius: 16px !important;
        overflow: hidden !important;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.01) !important;
    }

    /* Success banner */
    .success-banner {
        background-color: #ecfdf5 !important;
        border: 1px solid #a7f3d0 !important;
        border-radius: 12px !important;
        color: #065f46 !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        padding: 1rem !important;
        text-align: center !important;
        margin-top: 1rem;
        margin-bottom: 1rem;
    }


</style>
""", unsafe_allow_html=True)


def render_navbar(active_step, model_name="Groq API"):
    steps = [
        ("Upload", 1),
        ("Extract", 2),
        ("Aggregate", 3),
        ("Validate", 4),
        ("Export", 5)
    ]
    tabs_html = ""
    for label, step_num in steps:
        if step_num == active_step:
            tabs_html += f'<span class="nav-tab active-tab"><span class="step-num active-num">{step_num}</span>{label}</span>'
        else:
            tabs_html += f'<span class="nav-tab"><span class="step-num">{step_num}</span>{label}</span>'
            
    navbar_html = f"""
    <div class="nav-wrapper">
        <div class="nav-container">
            <div class="nav-left">
                <span class="nav-logo">GDSS</span>
            </div>
            <div class="nav-center">
                {tabs_html}
            </div>
            <div class="nav-right">
                <span class="nav-badge"><span style="color: #22c55e; margin-right: 6px; font-size: 1rem; line-height: 1;">●</span>{model_name}</span>
            </div>
        </div>
    </div>
    """
    st.markdown(navbar_html, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
IMDB_COLS = [
    "ITEM NAME", "BARCODE", "MANUFACTURER", "BRAND", "WEIGHT",
    "PACKAGING TYPE", "COUNTRY", "VARIANT", "TYPE",
    "FRAGRANCE FLAVOR", "PROMOTION", "ADDONS", "TAGLINE",
]


SYSTEM_PROMPT = """You are a product data extraction engine for a retail catalog system.
YOUR ONLY OUTPUT IS A SINGLE VALID JSON OBJECT. NO PREAMBLE. NO EXPLANATION. NO MARKDOWN. NO CODE BLOCKS. JUST THE RAW JSON.

Extract the following 13 fields:
- ITEM_NAME: full descriptive product name for catalog e.g. "KNORR CHICKEN STOCK CUBE 20G"
- BARCODE: numeric digits only, no dashes or spaces
- MANUFACTURER: full company name that manufactures the product
- BRAND: brand name exactly as shown on package e.g. "KNORR", "SUNLIGHT"
- WEIGHT: net weight or volume with unit, no spaces e.g. "80G", "500ML", "1.5KG"
- PACKAGING_TYPE: one of SACHET, BOTTLE, TUB, CAN, BOX, GLASS JAR, CARTON, POUCH, TUBE, PACKET, BAG
- COUNTRY: country of manufacture as printed on package e.g. "MALAYSIA", "THAILAND"
- VARIANT: product variant if present e.g. "ORIGINAL", "LOW FAT", "REDUCED SALT", "PREMIUM"
- TYPE: short product category e.g. "MARGARINE", "MAYONNAISE", "SEASONING", "DETERGENT"
- FRAGRANCE_FLAVOR: scent or flavor descriptor if present e.g. "RICH", "ORIGINAL", "LEMON", "VANILLA"
- PROMOTION: any on-pack promotion text e.g. "BUY 2 FREE 1", "20% EXTRA FREE", "NEW LOWER PRICE"
- ADDONS: additional included items or features e.g. "SPOON INCLUDED", "FREE RECIPE BOOK"
- TAGLINE: short promotional tagline or slogan e.g. "TASTE THE DIFFERENCE", "LOVED BY FAMILIES"

Rules:
- UPPERCASE for all string values
- Use "" for any field not visible or not applicable
- NEVER use null
- BARCODE must be digits only
- WEIGHT must have no space between number and unit"""

TRIGGER = 'Extract the product data from this image and return only the JSON object. Use "" for missing fields. Never use null.'

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def preprocess_image(image_file, max_size=(1024, 1024)):
    """Resize and base64-encode an image. Use smaller max_size for local models."""
    img = Image.open(image_file)
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail(max_size, Image.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)
    return base64.standard_b64encode(buffer.read()).decode("utf-8")


GROQ_API_KEY       = ""
OPENROUTER_API_KEY = ""
GROQ_DELAY_SEC     = 6                       # Minimum seconds between Groq/OpenRouter requests


def extract_via_groq(b64_image, api_key, max_retries=5):
    """
    Send image to Groq chat-completions endpoint (meta-llama/llama-4-scout-17b-16e-instruct).
    Exponential backoff on 429 errors: 30s, 60s, 120s, 240s, 480s.
    """
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "meta-llama/llama-4-scout-17b-16e-instruct",
        "max_tokens": 1000,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
                {"type": "text", "text": TRIGGER}
            ]}
        ]
    }
    for attempt in range(max_retries):
        resp = requests.post(url, json=payload, headers=headers, timeout=90)
        if resp.status_code == 429:
            wait = 30 * (2 ** attempt)
            st.warning(f"⏳ Groq rate limit hit. Waiting {wait}s before retry (attempt {attempt + 1}/{max_retries})...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"].strip()
        content = re.sub(r"^```json\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        return json.loads(content)
    raise Exception(f"Groq API returned 429 after {max_retries} retries.")


def extract_via_openrouter(b64_image, api_key, max_retries=5):
    """
    Send image to OpenRouter chat-completions endpoint (google/gemma-4-26b-a4b-it:free).
    Exponential backoff on 429 errors: 30s, 60s, 120s, 240s, 480s.
    """
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "google/gemma-4-26b-a4b-it:free",
        "max_tokens": 1000,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64_image}"}},
                {"type": "text", "text": TRIGGER}
            ]}
        ]
    }
    for attempt in range(max_retries):
        resp = requests.post(url, json=payload, headers=headers, timeout=90)
        if resp.status_code == 429:
            wait = 30 * (2 ** attempt)
            st.warning(f"⏳ OpenRouter rate limit hit. Waiting {wait}s before retry (attempt {attempt + 1}/{max_retries})...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"].strip()
        content = re.sub(r"^```json\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
        return json.loads(content)
    raise Exception(f"OpenRouter API returned 429 after {max_retries} retries.")


def aggregate(records):
    merged = {}
    fields = [
        "ITEM_NAME", "BARCODE", "MANUFACTURER", "BRAND", "WEIGHT",
        "PACKAGING_TYPE", "COUNTRY", "VARIANT", "TYPE",
        "FRAGRANCE_FLAVOR", "PROMOTION", "ADDONS", "TAGLINE",
    ]
    for col in fields:
        values = [r.get(col, "").strip() for r in records if r.get(col, "").strip()]
        if not values:
            merged[col] = ""
        else:
            count = Counter(values)
            max_count = count.most_common(1)[0][1]
            candidates = [v for v, c in count.items() if c == max_count]
            merged[col] = max(candidates, key=len)
    return merged


def validate(record):
    record["BARCODE"] = re.sub(r"[^\d]", "", record.get("BARCODE", ""))
    weight = record.get("WEIGHT", "").upper()
    weight = re.sub(r"(\d)\s+(G|KG|ML|L|MG)\b", r"\1\2", weight)
    record["WEIGHT"] = weight
    for field in ["ITEM_NAME", "MANUFACTURER", "BRAND", "PACKAGING_TYPE",
                  "COUNTRY", "VARIANT", "TYPE", "FRAGRANCE_FLAVOR",
                  "PROMOTION", "ADDONS", "TAGLINE"]:
        record[field] = record.get(field, "").upper().strip()
    return record


def to_imdb(record):
    return {
        "ITEM NAME":        record.get("ITEM_NAME", ""),
        "BARCODE":          record.get("BARCODE", ""),
        "MANUFACTURER":     record.get("MANUFACTURER", ""),
        "BRAND":            record.get("BRAND", ""),
        "WEIGHT":           record.get("WEIGHT", ""),
        "PACKAGING TYPE":   record.get("PACKAGING_TYPE", ""),
        "COUNTRY":          record.get("COUNTRY", ""),
        "VARIANT":          record.get("VARIANT", ""),
        "TYPE":             record.get("TYPE", ""),
        "FRAGRANCE FLAVOR": record.get("FRAGRANCE_FLAVOR", ""),
        "PROMOTION":        record.get("PROMOTION", ""),
        "ADDONS":           record.get("ADDONS", ""),
        "TAGLINE":          record.get("TAGLINE", ""),
    }


def build_excel(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "IMDB"
    thin  = Side(style='thin', color='000000')
    thick = Side(style='medium', color='000000')
    ws.append(IMDB_COLS)
    for col_idx in range(1, len(IMDB_COLS)+1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name='Calibri', bold=True, size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(left=thick, right=thick, top=thick, bottom=thick)
    ws.row_dimensions[1].height = 20
    for row_idx, row in enumerate(rows, 2):
        for col_idx, col in enumerate(IMDB_COLS, 1):
            val = row.get(col, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val if val else None)
            cell.font = Font(name='Calibri', size=10)
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = Border(left=Side(style='thin', color='000000'),
                                 right=Side(style='thin', color='000000'),
                                 top=Side(style='thin', color='000000'),
                                 bottom=Side(style='thin', color='000000'))
        ws.row_dimensions[row_idx].height = 15
    col_widths = [18, 14, 24, 32, 16, 55, 14, 16, 18, 50]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = []
if "edited_df" not in st.session_state:
    st.session_state.edited_df = None


# ─────────────────────────────────────────────
# ENGINE CONFIG (MOVED FROM SIDEBAR)
# ─────────────────────────────────────────────
with st.expander("⚙️ Model Configuration", expanded=False):
    col1, col2 = st.columns(2)
    with col1:
        engine = st.selectbox(
            "Choose model",
            ["Groq API", "OpenRouter API"]
        )
    with col2:
        api_key = ""
        if engine == "Groq API":
            api_key = st.text_input("Groq API Key", type="password", placeholder="gsk_...", value=GROQ_API_KEY)
        elif engine == "OpenRouter API":
            api_key = st.text_input("OpenRouter API Key", type="password", placeholder="sk-or-v1-...", value=OPENROUTER_API_KEY)

# ─────────────────────────────────────────────
# STATE DETECTION FOR TABS
# ─────────────────────────────────────────────
active_step = 1
uploaded_files_in_state = st.session_state.get("uploader", [])
if uploaded_files_in_state:
    active_step = 2
    if st.session_state.get("results"):
        active_step = 5

# ─────────────────────────────────────────────
# HEADER (NAVBAR & WELCOME GREETING)
# ─────────────────────────────────────────────
navbar_placeholder = st.empty()
with navbar_placeholder.container():
    render_navbar(active_step, engine)

st.markdown("<h2 style='font-family: \"Playfair Display\", Georgia, serif; font-weight: 800; font-size: 2.2rem; color: #0f172a; margin-top: 1rem; margin-bottom: 0.25rem;'>🏷️ IMDB Auto-Fill</h2>", unsafe_allow_html=True)
st.markdown("<p style='font-family: \"Outfit\", sans-serif; color: #64748b; font-size: 1.05rem; margin-bottom: 2rem;'>Welcome back! Streamline retail master data cataloging with AI-driven extraction.</p>", unsafe_allow_html=True)
st.divider()

# ─────────────────────────────────────────────
# STEP 1 — UPLOAD
# ─────────────────────────────────────────────
st.markdown("### Step 1 — Upload Product Images")
st.caption("Upload multiple images of the same product (different sides/angles) or images from different products. The system groups them by product automatically.")

uploaded_files = st.file_uploader(
    "Drop images here or click to browse",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="uploader"
)

if uploaded_files:
    # Group by product ID (filename prefix)
    groups = defaultdict(list)
    for f in uploaded_files:
        product_id = f.name.split("_")[0]
        groups[product_id].append(f)

    col1, col2, col3 = st.columns(3)
    col1.metric("Images uploaded", len(uploaded_files))
    col2.metric("Products detected", len(groups))
    col3.metric("Avg images/product", f"{len(uploaded_files)/len(groups):.1f}")

    # Show image previews
    with st.expander("Preview uploaded images", expanded=False):
        cols = st.columns(5)
        for i, f in enumerate(uploaded_files[:10]):
            with cols[i % 5]:
                st.image(f, caption=f.name, use_container_width=True)

    st.divider()

    # ─────────────────────────────────────────────
    # STEP 2 — EXTRACT
    # ─────────────────────────────────────────────
    st.markdown("### Step 2 — Extract & Process")

    checkpoint_path = Path("./checkpoint.json")
    checkpoint_records = []
    done_ids = set()
    if checkpoint_path.exists():
        try:
            with open(checkpoint_path, "r", encoding="utf-8") as f:
                checkpoint_records = json.load(f)
            done_ids = {r["_product_id"] for r in checkpoint_records if "_product_id" in r}
            if done_ids:
                col_info, col_clear = st.columns([3, 1])
                with col_info:
                    st.info(f"ℹ️ **Found checkpoint on disk:** {len(done_ids)} product(s) already processed.")
                with col_clear:
                    if st.button("🗑️ Clear Checkpoint", use_container_width=True):
                        checkpoint_path.unlink(missing_ok=True)
                        st.rerun()
        except Exception as e:
            st.warning(f"⚠️ Could not read checkpoint file: {e}")

    if st.button("🚀 Run Pipeline", use_container_width=True):
        # Create a dictionary to hold results by product ID to preserve the original order of products
        results_by_product = {}
        for product_id in groups:
            if product_id in done_ids:
                matched_record = next(r for r in checkpoint_records if r.get("_product_id") == product_id)
                results_by_product[product_id] = to_imdb(matched_record)

        # Build a single flat queue of images to process across all non-checkpointed products
        queue = []
        for product_id, files in groups.items():
            if product_id in done_ids:
                continue
            for f in files:
                queue.append((product_id, f))

        total_images = len(queue)

        if total_images > 0:
            progress = st.progress(0)
            status = st.empty()
            
            extracted_by_product = defaultdict(list)
            processed_count_by_product = defaultdict(int)

            for idx, (product_id, f) in enumerate(queue):
                # Update navbar to "Extract" (active_step = 2)
                with navbar_placeholder.container():
                    render_navbar(2, engine)
                
                status.markdown(f"**Processing Image {idx + 1}/{total_images}:** `{f.name}` (Product `{product_id}`)...")
                
                timer_slot = st.empty()
                start = time.time()
                try:
                    b64 = preprocess_image(f, max_size=(1024, 1024))
                    if engine == "Groq API" and api_key:
                        record = extract_via_groq(b64, api_key)
                    elif engine == "OpenRouter API" and api_key:
                        record = extract_via_openrouter(b64, api_key)
                    else:
                        raise ValueError("No valid engine selected or API key missing.")

                    elapsed = time.time() - start
                    timer_slot.success(f"✅ `{f.name}` done in {elapsed:.1f}s")
                    extracted_by_product[product_id].append(record)
                except Exception as e:
                    timer_slot.warning(f"⚠️ Could not extract from `{f.name}`: {e}")

                # Enforce a 6-second sleep between each API call in the queue
                if idx < total_images - 1:
                    status.markdown(f"⏳ Sleeping 6 seconds before next API call...")
                    time.sleep(6)

                processed_count_by_product[product_id] += 1
                if processed_count_by_product[product_id] == len(groups[product_id]):
                    # All images for this product have been processed, let's aggregate and validate!
                    if extracted_by_product[product_id]:
                        with navbar_placeholder.container():
                            render_navbar(3, engine)
                        merged = aggregate(extracted_by_product[product_id])
                        
                        with navbar_placeholder.container():
                            render_navbar(4, engine)
                        merged = validate(merged)
                        
                        merged["_product_id"] = product_id
                        checkpoint_records.append(merged)
                        try:
                            with open(checkpoint_path, "w", encoding="utf-8") as f_out:
                                json.dump(checkpoint_records, f_out, ensure_ascii=False, indent=2)
                        except Exception as e:
                            st.warning(f"⚠️ Could not save checkpoint: {e}")
                            
                        results_by_product[product_id] = to_imdb(merged)

                progress.progress((idx + 1) / total_images)

            status.empty()
            progress.empty()

        # Reassemble the final results in the exact original order of the products
        results = [results_by_product[pid] for pid in groups if pid in results_by_product]
        
        st.session_state.results = results
        st.session_state.edited_df = pd.DataFrame(results, columns=IMDB_COLS)
        
        # Clean up checkpoint after successful full completion
        if checkpoint_path.exists():
            try:
                checkpoint_path.unlink()
            except Exception as e:
                st.warning(f"⚠️ Could not delete checkpoint file: {e}")

        st.markdown('<div class="success-banner">✅ Pipeline complete!</div>', unsafe_allow_html=True)
        
        # Update navbar to "Export/Preview" (active_step = 5)
        with navbar_placeholder.container():
            render_navbar(5, engine)

# ─────────────────────────────────────────────
# STEP 3 — PREVIEW & EDIT
# ─────────────────────────────────────────────
if st.session_state.results:
    st.divider()
    st.markdown("### Step 3 — Preview & Edit")
    st.caption("Click any cell to edit before exporting.")

    edited = st.data_editor(
        st.session_state.edited_df,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True
    )
    st.session_state.edited_df = edited

    # Stats
    df = edited
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total products", len(df))
    col2.metric("With barcode", int(df["BARCODE"].astype(bool).sum()))
    col3.metric("With manufacturer", int(df["MANUFACTURER"].astype(bool).sum()))
    col4.metric("With country", int(df["COUNTRY"].astype(bool).sum()))

    st.divider()

    # ─────────────────────────────────────────────
    # STEP 4 — EXPORT
    # ─────────────────────────────────────────────
    st.markdown("### Step 4 — Export")
    col1, col2 = st.columns(2)

    with col1:
        excel_buffer = build_excel(edited.to_dict("records"))
        st.download_button(
            label="📥 Download as Excel (.xlsx)",
            data=excel_buffer,
            file_name="IMDB_predictions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        csv = edited.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV (.csv)",
            data=csv,
            file_name="IMDB_predictions.csv",
            mime="text/csv",
            use_container_width=True
        )