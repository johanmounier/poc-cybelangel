import os
from dotenv import load_dotenv

load_dotenv()  # local : charge .env ; Streamlit Cloud : sans effet (secrets injectés via st.secrets)


def _get(key: str) -> str:
    """Lit un secret dans l'ordre : variable d'env → st.secrets (Streamlit Cloud)."""
    val = os.getenv(key, "")
    if val:
        return val
    try:
        import streamlit as st
        return st.secrets.get(key, "")
    except Exception:
        return ""


# ── Credentials ───────────────────────────────────────────────────────────────
SENDER_EMAIL        = _get("SENDER_EMAIL")
OUTLOOK_PASSWORD    = _get("OUTLOOK_PASSWORD")
RECIPIENT_EMAIL     = _get("RECIPIENT_EMAIL")
ANTHROPIC_API_KEY   = _get("ANTHROPIC_API_KEY")

# ── Chemins ───────────────────────────────────────────────────────────────────
TEMPLATE_EXCEL = "data/template.xlsx"
OUTPUT_DIR     = "output/"

# ── Constantes métier ─────────────────────────────────────────────────────────
CURRENT_MONTH_LABEL = "Octobre 2024"
CURRENT_MONTH_COL   = 12   # col L = Oct (col M = Total YTD Jan-Oct)

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
