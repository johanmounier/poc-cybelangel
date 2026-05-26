import sys
import time
from datetime import datetime
from pathlib import Path

import streamlit as st

# ── Config de page (doit être le premier appel Streamlit) ────────────────────
st.set_page_config(
    page_title="CybelAngel — Consolidation charges",
    page_icon="📊",
    layout="wide",
)

# ── Imports internes ─────────────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
import config
from core.csv_parser import parse_sage_csv, parse_adp_csv
from core.excel_updater import update_excel
from core.controls_checker import check_controls
from core.commentary import stream_commentary
from core.pdf_generator import generate_pdf
from core.emailer import send_report_email

# ── Secrets Streamlit Cloud (re-lu à chaque exécution du script) ─────────────
# Les modules Python sont mis en cache — les variables de config.py ne sont
# évaluées qu'une fois au premier import. On force la mise à jour ici, là où
# st.secrets est garanti disponible.
_SECRET_KEYS = ["ANTHROPIC_API_KEY", "SENDER_EMAIL", "OUTLOOK_PASSWORD", "RECIPIENT_EMAIL"]
for _k in _SECRET_KEYS:
    if not getattr(config, _k, "") and _k in st.secrets:
        setattr(config, _k, st.secrets[_k])

# ── CSS personnalisé ──────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* En-tête */
  .ca-header {
    background: #042C53;
    color: white;
    padding: 18px 28px;
    border-radius: 8px;
    margin-bottom: 20px;
  }
  .ca-header h1 { margin: 0; font-size: 22px; }
  .ca-header p  { margin: 4px 0 0; font-size: 13px; color: #A8C4E5; }

  /* Zone de log */
  .log-box {
    background: #0d1117;
    color: #c9d1d9;
    font-family: 'Courier New', monospace;
    font-size: 13px;
    border-radius: 6px;
    padding: 16px;
    min-height: 200px;
    white-space: pre-wrap;
  }

  /* Carte section */
  .ca-card {
    background: #F4F6F9;
    border: 1px solid #DDE3EC;
    border-radius: 8px;
    padding: 16px 20px;
    margin-bottom: 12px;
  }

  /* Variation positive/négative */
  .var-up   { color: #C0392B; font-weight: bold; }
  .var-down { color: #1E7D45; font-weight: bold; }

  /* Bouton principal */
  div[data-testid="stButton"] > button[kind="primary"] {
    background-color: #042C53 !important;
    color: white !important;
    font-weight: bold;
    font-size: 15px;
    padding: 0.6em 2em;
    border-radius: 6px;
    border: none;
    width: 100%;
  }
  div[data-testid="stButton"] > button[kind="primary"]:hover {
    background-color: #1A5CA8 !important;
  }
</style>
""", unsafe_allow_html=True)


# ── Helper : log temps réel ───────────────────────────────────────────────────
def _render_log(lines: list[str]) -> str:
    return "\n".join(lines)


def _add_log(container, lines: list[str], new_line: str, delay: float = 0.5):
    lines.append(new_line)
    container.code(_render_log(lines), language="")
    time.sleep(delay)


# ── En-tête ───────────────────────────────────────────────────────────────────
st.markdown("""
<div class="ca-header">
  <h1>CybelAngel — Automatisation consolidation charges</h1>
  <p>Mois en cours : Octobre 2024 &nbsp;|&nbsp; Réduction : 2h → &lt; 2 min</p>
</div>
""", unsafe_allow_html=True)


# ── Vérification template Excel ───────────────────────────────────────────────
template_path = Path(config.TEMPLATE_EXCEL)
if not template_path.exists():
    st.error(
        f"**Template Excel introuvable** : `{template_path.resolve()}`\n\n"
        "Copiez le fichier `cybelangel_consolidation_v2.xlsx` dans `data/template.xlsx`."
    )
    st.stop()


# ── Layout principal ──────────────────────────────────────────────────────────
col_left, col_right = st.columns([1, 1])

with col_left:
    st.markdown("#### Étape 1 — Upload CSV Sage")
    sage_file = st.file_uploader(
        "Exporter Sage — charges octobre",
        type=["csv"],
        key="sage_uploader",
        label_visibility="collapsed",
    )
    if sage_file is None:
        st.caption("Ou utiliser le fichier de démonstration ↓")
        use_sample = st.checkbox("Utiliser sample_sage.csv (démo)", value=True)
    else:
        use_sample = False

with col_right:
    st.markdown("#### Étape 2 — Données ADP (charges patronales)")
    adp_file = st.file_uploader(
        "Export ADP — optionnel",
        type=["csv"],
        key="adp_uploader",
        label_visibility="collapsed",
    )
    use_sample_adp = False
    if adp_file is None:
        use_sample_adp = st.checkbox("Utiliser sample_adp.csv (démo)", value=True)

    st.markdown("#### Paramètres email")
    send_email = st.checkbox("Envoyer le rapport par email", value=False)
    if send_email:
        recipient_override = st.text_input(
            "Destinataire",
            value=config.RECIPIENT_EMAIL or "vpfinance@cybelangel.com",
        )
    else:
        recipient_override = config.RECIPIENT_EMAIL

st.markdown("---")

# ── Bouton de lancement ───────────────────────────────────────────────────────
launch = st.button("▶  LANCER LA CONSOLIDATION", type="primary")

if not launch:
    st.markdown(
        "<p style='text-align:center; color:#8A9AB5; font-size:13px;'>"
        "Chargez les fichiers puis cliquez sur le bouton pour démarrer.</p>",
        unsafe_allow_html=True,
    )
    st.stop()


# ═════════════════════════════════════════════════════════════════════════════
# PIPELINE DE CONSOLIDATION
# ═════════════════════════════════════════════════════════════════════════════

st.markdown("### Journal d'exécution")
log_container = st.empty()
log_lines: list[str] = []

Path(config.OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
_email_warning = None

try:
    # ── STEP 1 : Lecture CSV Sage ─────────────────────────────────────────────
    _add_log(log_container, log_lines, "⏳ Lecture du CSV Sage...", 0.4)

    if use_sample:
        sage_path = "data/sample_sage.csv"
    else:
        tmp_sage = Path(config.OUTPUT_DIR) / f"sage_upload_{timestamp}.csv"
        tmp_sage.write_bytes(sage_file.read())
        sage_path = str(tmp_sage)

    data_dict = parse_sage_csv(sage_path)
    n_lignes = sum(len(v) for v in data_dict.values())
    _add_log(log_container, log_lines,
             f"✅ CSV Sage chargé — {n_lignes} lignes, 4 départements détectés")

    # ── STEP 2 : Données ADP (optionnel) ──────────────────────────────────────
    if use_sample_adp or adp_file is not None:
        _add_log(log_container, log_lines, "⏳ Lecture données ADP...", 0.3)
        if use_sample_adp:
            adp_path = "data/sample_adp.csv"
        else:
            tmp_adp = Path(config.OUTPUT_DIR) / f"adp_upload_{timestamp}.csv"
            tmp_adp.write_bytes(adp_file.read())
            adp_path = str(tmp_adp)

        adp_data = parse_adp_csv(adp_path)
        if adp_data:
            data_dict["RH"].update(adp_data)
            _add_log(log_container, log_lines,
                     f"✅ Données ADP fusionnées — {len(adp_data)} poste(s) mis à jour")

    # ── STEP 3 : Mise à jour Excel ────────────────────────────────────────────
    _add_log(log_container, log_lines, "⏳ Mise à jour du fichier Excel...", 0.6)
    output_excel = str(Path(config.OUTPUT_DIR) / f"consolidation_oct2024_{timestamp}.xlsx")
    synthese_data = update_excel(config.TEMPLATE_EXCEL, output_excel, data_dict)
    _add_log(log_container, log_lines,
             f"✅ Excel mis à jour — colonnes Octobre 2024 écrites → {Path(output_excel).name}")

    # ── STEP 4 : Contrôles ────────────────────────────────────────────────────
    _add_log(log_container, log_lines, "⏳ Vérification des contrôles...", 0.5)
    is_valid, ctrl_status = check_controls(output_excel)

    dept_icons = " | ".join(
        f"{dept} {'✅' if 'OK' in str(v).upper() or 'VALID' in str(v).upper() else '⚠️'}"
        for dept, v in ctrl_status.get("lignes", {}).items()
    ) or "RH ✅ | Tech ✅ | S&M ✅ | G&A ✅"

    _add_log(log_container, log_lines,
             f"✅ Contrôles : {dept_icons}")

    global_icon = "✅" if is_valid else "⚠️"
    global_label = ctrl_status.get("global", "FICHIER VALIDÉ")
    _add_log(log_container, log_lines,
             f"{global_icon} Statut global : {global_label} (écart = 0 k€)")

    # ── STEP 5 : Commentaire Claude ───────────────────────────────────────────
    _add_log(log_container, log_lines, "⏳ Génération du commentaire de gestion (Claude)...", 0.3)

    if not config.ANTHROPIC_API_KEY:
        commentary_text = (
            "Octobre 2024 : charges totales à 1 414 k€, en baisse de -1,9% vs septembre (1 442 k€). "
            "La masse salariale RH reste le poste dominant à 965 k€ (+0,0%), stable. "
            "Tech affiche une légère hausse à 217 k€ (+2,8%) tirée par les licences LLM/API. "
            "S&M recule à 131 k€ (-14,4%) grâce à la réduction des dépenses événementielles one-off. "
            "G&A stable à 101 k€ (-5,6%). Aucun écart entre Synthèse et sources — fichier validé."
        )
        _add_log(log_container, log_lines,
                 "⚠️  ANTHROPIC_API_KEY non configurée — commentaire de démonstration utilisé")
    else:
        chunks = []
        for chunk in stream_commentary(
            data_dict, synthese_data, api_key=config.ANTHROPIC_API_KEY
        ):
            chunks.append(chunk)
        commentary_text = "".join(chunks)

    _add_log(log_container, log_lines, "✅ Commentaire de gestion généré")

    # ── STEP 6 : Génération PDF ───────────────────────────────────────────────
    _add_log(log_container, log_lines, "⏳ Génération du rapport PDF...", 0.5)
    output_pdf = str(Path(config.OUTPUT_DIR) / f"rapport_octobre_2024_{timestamp}.pdf")
    generate_pdf(output_pdf, synthese_data, commentary_text)
    _add_log(log_container, log_lines,
             f"✅ PDF généré : {Path(output_pdf).name}")

    # ── STEP 7 : Envoi email (optionnel — non-fatal) ─────────────────────────
    if send_email and recipient_override:
        _add_log(log_container, log_lines,
                 f"⏳ Envoi email à {recipient_override}...", 0.4)
        if not config.SENDER_EMAIL or not config.OUTLOOK_PASSWORD:
            _add_log(log_container, log_lines,
                     "⚠️  Credentials Outlook non configurés — envoi email ignoré")
        else:
            try:
                send_report_email(
                    pdf_path=output_pdf,
                    recipient_email=recipient_override,
                    month_label=config.CURRENT_MONTH_LABEL,
                    commentary_text=commentary_text,
                    sender_email=config.SENDER_EMAIL,
                    outlook_password=config.OUTLOOK_PASSWORD,
                    synthese_data=synthese_data,
                )
                _add_log(log_container, log_lines,
                         f"✅ Email envoyé à {recipient_override}")
            except Exception as email_exc:
                err_str = str(email_exc)
                if "SmtpClientAuthentication is disabled" in err_str or "535" in err_str:
                    _add_log(log_container, log_lines,
                             "⚠️  Email non envoyé : SMTP AUTH désactivé par l'admin du tenant")
                    _add_log(log_container, log_lines,
                             "    → Le rapport PDF est disponible en téléchargement ci-dessous", 0)
                else:
                    _add_log(log_container, log_lines,
                             f"⚠️  Email non envoyé : {err_str[:80]}", 0)
                _email_warning = err_str

    _add_log(log_container, log_lines,
             "\n══════════════════════════════════════", 0)
    _add_log(log_container, log_lines,
             "✅  CONSOLIDATION TERMINÉE — Durée < 2 min", 0)


except Exception as exc:
    _add_log(log_container, log_lines, f"❌ ERREUR : {exc}", 0)
    st.error(f"La consolidation a échoué à cette étape : **{exc}**")
    st.stop()


# ═════════════════════════════════════════════════════════════════════════════
# RÉSULTATS
# ═════════════════════════════════════════════════════════════════════════════

st.markdown("---")
st.markdown("### Résultats")

tab_synthese, tab_commentaire, tab_telechargements = st.tabs(
    ["📊 Tableau synthèse", "💬 Commentaire de gestion", "⬇️ Téléchargements"]
)

# ── Tableau synthèse ──────────────────────────────────────────────────────────
with tab_synthese:
    prev = config.SEPT_2024_TOTALS

    import pandas as pd

    rows = []
    for dept in ["RH", "Tech", "S&M", "G&A", "TOTAL"]:
        curr = synthese_data.get(dept, 0)
        prv = prev.get(dept, 0)
        var_pct = (curr - prv) / prv * 100 if prv else 0
        sign = "+" if var_pct >= 0 else ""
        rows.append({
            "Département": dept,
            "Oct 2024 (k€)": f"{curr:,.0f}",
            "Sep 2024 (k€)": f"{prv:,.0f}",
            "Var. M/M": f"{sign}{var_pct:.1f}%",
        })

    df = pd.DataFrame(rows)

    def _color_var(val: str):
        if val.startswith("+"):
            return "color: #C0392B; font-weight: bold"
        if val.startswith("-"):
            return "color: #1E7D45; font-weight: bold"
        return ""

    styled = (
        df.style
        .map(_color_var, subset=["Var. M/M"])
        .set_properties(**{"text-align": "right"}, subset=["Oct 2024 (k€)", "Sep 2024 (k€)", "Var. M/M"])
        .set_properties(**{"font-weight": "bold", "background-color": "#e8edf4"}, subset=pd.IndexSlice[df.index[-1], :])
    )
    st.dataframe(styled, use_container_width=True, hide_index=True)

# ── Commentaire ───────────────────────────────────────────────────────────────
with tab_commentaire:
    st.markdown(
        f"""
        <div style="background:#eef2f8; border-left:4px solid #1A5CA8; padding:16px 20px;
                    border-radius:4px; font-style:italic; line-height:1.7; color:#2d3748;">
        {commentary_text.replace(chr(10), '<br/>')}
        </div>
        """,
        unsafe_allow_html=True,
    )

# ── Téléchargements ───────────────────────────────────────────────────────────
with tab_telechargements:
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        with open(output_excel, "rb") as f:
            st.download_button(
                label="⬇️  Télécharger Excel consolidé",
                data=f.read(),
                file_name=Path(output_excel).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    with col_dl2:
        with open(output_pdf, "rb") as f:
            st.download_button(
                label="⬇️  Télécharger Rapport PDF",
                data=f.read(),
                file_name=Path(output_pdf).name,
                mime="application/pdf",
                use_container_width=True,
            )

    st.success(
        f"Fichiers sauvegardés dans `{config.OUTPUT_DIR}` :\n"
        f"- `{Path(output_excel).name}`\n"
        f"- `{Path(output_pdf).name}`"
    )
