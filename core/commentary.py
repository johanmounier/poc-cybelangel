import anthropic
from config import SEPT_2024_TOTALS, CURRENT_MONTH_LABEL

MODEL = "claude-sonnet-4-20250514"

SYSTEM_PROMPT = """\
Tu es un contrôleur de gestion senior. Tu rédiges le commentaire mensuel de charges \
pour le VP Finance. Ton style est concis, factuel, financier. Tu utilises des chiffres précis. \
Tu signales les variations notables (>5% vs M-1), les points d'attention, et les éléments one-off. \
Maximum 150 mots. Format : 3 paragraphes courts.
Département | Mois courant | Mois précédent | Variation\
"""


def generate_commentary(
    data_dict: dict,
    synthese_data: dict,
    previous_month_data: dict | None = None,
    api_key: str = "",
) -> str:
    """
    Génère le commentaire de gestion via Claude.

    Args:
        data_dict: données brutes octobre par département (dict de postes → montants)
        synthese_data: totaux par département après mise à jour Excel
        previous_month_data: totaux M-1 (défaut : SEPT_2024_TOTALS)
        api_key: clé API Anthropic

    Returns:
        Texte du commentaire (str)
    """
    prev = previous_month_data or SEPT_2024_TOTALS

    # Construction du prompt utilisateur
    lines = [f"Rapport mensuel — {CURRENT_MONTH_LABEL}\n"]
    lines.append("Totaux par département (k€) :")
    for dept in ["RH", "Tech", "S&M", "G&A", "TOTAL"]:
        curr = synthese_data.get(dept, 0)
        prv = prev.get(dept, 0)
        var_pct = ((curr - prv) / prv * 100) if prv else 0
        sign = "+" if var_pct >= 0 else ""
        lines.append(f"  {dept:6s} | {curr:>6.0f} k€ | {prv:>6.0f} k€ | {sign}{var_pct:.1f}%")

    lines.append("\nDétail postes notables :")
    for dept, postes in data_dict.items():
        top = sorted(postes.items(), key=lambda x: x[1], reverse=True)[:3]
        for poste, montant in top:
            lines.append(f"  [{dept}] {poste} : {montant} k€")

    user_message = "\n".join(lines)

    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model=MODEL,
        max_tokens=400,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
    return message.content[0].text


def stream_commentary(
    data_dict: dict,
    synthese_data: dict,
    previous_month_data: dict | None = None,
    api_key: str = "",
):
    """
    Version streaming — génère le commentaire et yield les chunks de texte.
    Compatible avec st.write_stream().
    """
    prev = previous_month_data or SEPT_2024_TOTALS

    lines = [f"Rapport mensuel — {CURRENT_MONTH_LABEL}\n"]
    lines.append("Totaux par département (k€) :")
    for dept in ["RH", "Tech", "S&M", "G&A", "TOTAL"]:
        curr = synthese_data.get(dept, 0)
        prv = prev.get(dept, 0)
        var_pct = ((curr - prv) / prv * 100) if prv else 0
        sign = "+" if var_pct >= 0 else ""
        lines.append(f"  {dept:6s} | {curr:>6.0f} k€ | {prv:>6.0f} k€ | {sign}{var_pct:.1f}%")

    lines.append("\nDétail postes notables :")
    for dept, postes in data_dict.items():
        top = sorted(postes.items(), key=lambda x: x[1], reverse=True)[:3]
        for poste, montant in top:
            lines.append(f"  [{dept}] {poste} : {montant} k€")

    user_message = "\n".join(lines)

    client = anthropic.Anthropic(api_key=api_key)
    with client.messages.stream(
        model=MODEL,
        max_tokens=400,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    ) as stream:
        for text in stream.text_stream:
            yield text
