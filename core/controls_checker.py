from core.excel_updater import get_controls_status


def check_controls(filepath: str) -> tuple[bool, dict]:
    """
    Lance les contrôles depuis l'onglet '🔍 Contrôles'.

    Returns:
        (is_valid: bool, status_dict: dict)
        is_valid = True si le statut global contient "VALIDÉ" ou "OK"
    """
    status = get_controls_status(filepath)
    global_ok = any(
        kw in status["global"].upper()
        for kw in ("VALID", "OK", "✅")
    )
    return global_ok, status
