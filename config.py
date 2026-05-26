import os
from dotenv import load_dotenv

load_dotenv()  # charge .env en local ; sans effet sur Railway (env vars injectées directement)

# ── Credentials — lus depuis les variables d'environnement ───────────────────
SENDER_EMAIL        = os.getenv("SENDER_EMAIL", "")
OUTLOOK_PASSWORD    = os.getenv("OUTLOOK_PASSWORD", "")
RECIPIENT_EMAIL     = os.getenv("RECIPIENT_EMAIL", "")
ANTHROPIC_API_KEY   = os.getenv("ANTHROPIC_API_KEY", "")

# ── Chemins ──────────────────────────────────────────────────────────────────
TEMPLATE_EXCEL = "data/template.xlsx"
OUTPUT_DIR     = "output/"

# ── Constantes métier ────────────────────────────────────────────────────────
CURRENT_MONTH_LABEL = "Octobre 2024"
CURRENT_MONTH_COL   = 13          # colonne M (index openpyxl, 1-based)

SEPT_2024_TOTALS = {
    "RH": 965,
    "Tech": 217,
    "S&M": 153,
    "G&A": 107,
    "TOTAL": 1442,
}

DEPT_TOTAL_ROWS = {
    "RH": 19,
    "Tech": 18,
    "S&M": 19,
    "G&A": 19,
}

DEPT_DATA_START_ROW = 4
