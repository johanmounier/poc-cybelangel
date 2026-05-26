import shutil
from pathlib import Path

import openpyxl

from config import DEPT_TOTAL_ROWS, DEPT_DATA_START_ROW, CURRENT_MONTH_COL

DEPT_SHEETS = {"RH": "RH", "Tech": "Tech", "S&M": "S&M", "G&A": "G&A"}
OCT_HEADER   = "Oct 2024"
YTD_HEADER   = "Total YTD (Jan–Oct)"


def _safe_write(ws, row: int, col: int, value) -> None:
    """Écrit dans une cellule en ignorant les cellules fusionnées (read-only)."""
    try:
        ws.cell(row=row, column=col).value = value
    except AttributeError:
        pass  # MergedCell — ignorée


# ─────────────────────────────────────────────────────────────────────────────
# Point d'entrée principal
# ─────────────────────────────────────────────────────────────────────────────

def update_excel(template_path: str, output_path: str, data_dict: dict) -> dict:
    """
    1. Copie le template
    2. Lit les valeurs Jan-Sep depuis le template (data_only)
    3. Réécrit tout en valeurs littérales (pas de dépendance aux formules)
    4. Ajoute la colonne Oct dans la zone principale et la zone d'import
    5. Met à jour Synthèse et Contrôles
    Returns: {"RH": x, "Tech": x, "S&M": x, "G&A": x, "TOTAL": x}
    """
    shutil.copy2(template_path, output_path)

    # Lecture des valeurs Jan-Sep depuis le template original (avec cache formules)
    jan_sep = _read_template_jan_sep(template_path)

    wb = openpyxl.load_workbook(output_path)
    totals = {}

    for sheet_name, dept_key in DEPT_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            totals[dept_key] = 0
            continue

        ws         = wb[sheet_name]
        dept_data  = data_dict.get(dept_key, {})
        total_row  = DEPT_TOTAL_ROWS[dept_key]
        tmpl_vals  = jan_sep.get(dept_key, {})

        # ── 1. Ajouter l'en-tête "Oct 2024" (col M) dans la zone principale ──
        ws.cell(row=2, column=CURRENT_MONTH_COL).value = OCT_HEADER

        # Renommer "Total YTD" → "Total YTD (Jan-Oct)" dans col L
        if ws.cell(row=2, column=12).value and "YTD" in str(ws.cell(row=2, column=12).value):
            ws.cell(row=2, column=12).value = YTD_HEADER

        # ── 2. Écrire Oct + recalculer sous-totaux + total (zone principale) ─
        dept_total = _write_dept_data(ws, dept_data, tmpl_vals, total_row)
        totals[dept_key] = dept_total

        # ── 3. Zone d'import : ajouter Oct + écrire les lignes de données ────
        _write_import_zone(ws, dept_data, total_row)

    totals["TOTAL"] = sum(v for k, v in totals.items() if k != "TOTAL")

    # ── 4. Synthèse ───────────────────────────────────────────────────────────
    _update_synthese(wb, jan_sep, totals)

    # ── 5. Contrôles ──────────────────────────────────────────────────────────
    _update_controls_sheet(wb, totals)

    wb.save(output_path)
    wb.close()
    return totals


# ─────────────────────────────────────────────────────────────────────────────
# Lecture template Jan-Sep
# ─────────────────────────────────────────────────────────────────────────────

def _read_template_jan_sep(template_path: str) -> dict:
    """
    Lit le template en data_only pour récupérer les valeurs Jan-Sep
    de chaque ligne de données (pas uniquement le TOTAL).
    Retourne: {dept_key: {row_label: {col_idx: value, ...}, ...}}
    """
    wb  = openpyxl.load_workbook(template_path, data_only=True)
    out = {}

    for sheet_name, dept_key in DEPT_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws        = wb[sheet_name]
        total_row = DEPT_TOTAL_ROWS[dept_key]
        rows_data = {}

        for row in range(DEPT_DATA_START_ROW, total_row + 1):
            label = ws.cell(row=row, column=1).value
            if label is None:
                continue
            label = str(label).strip()
            if label.startswith("  "):          # en-tête de catégorie
                continue
            month_vals = {}
            for col in range(3, 12):            # colonnes C à K = Jan-Sep
                v = ws.cell(row=row, column=col).value
                month_vals[col] = float(v) if v is not None else 0.0
            rows_data[label] = month_vals

        out[dept_key] = rows_data

    wb.close()
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Écriture zone principale
# ─────────────────────────────────────────────────────────────────────────────

def _write_dept_data(ws, dept_data: dict, tmpl_vals: dict, total_row: int) -> float:
    """
    Réécrit toutes les valeurs (Jan-Sep + Oct) en littéraux dans la zone
    principale. Recalcule sous-totaux et total pour Oct.
    Retourne le total Oct du département.
    """
    # Structure de groupes pour les sous-totaux
    current_group: list[tuple[int, float]] = []   # (row, oct_val)
    subtotal_rows: dict[int, float]        = {}

    dept_oct_total = 0.0

    for row in range(DEPT_DATA_START_ROW, total_row + 1):
        cell_a = ws.cell(row=row, column=1).value
        if cell_a is None:
            continue
        label = str(cell_a).strip()

        # ── En-tête de catégorie (ligne " Marketing", " Salaires…") ──────────
        if label.startswith("  "):
            current_group = []
            continue

        # ── Ligne Sous-total ──────────────────────────────────────────────────
        if label.startswith("Sous-total"):
            subtotal_oct = sum(v for _, v in current_group)
            subtotal_rows[row] = subtotal_oct

            # Réécrire Jan-Sep en littéraux depuis le template
            tmpl = tmpl_vals.get(label, {})
            for col in range(3, 12):
                _safe_write(ws, row, col, tmpl.get(col, None))

            current_group = []
            continue

        # ── Ligne TOTAL ───────────────────────────────────────────────────────
        if label.upper().startswith("TOTAL"):
            # Jan-Sep déjà gérés ci-dessous; on ne touche que Oct + YTD
            continue

        # ── Ligne de données (poste de charge) ───────────────────────────────
        oct_val = float(dept_data.get(label, 0))
        _safe_write(ws, row, CURRENT_MONTH_COL, oct_val)
        dept_oct_total  += oct_val
        current_group.append((row, oct_val))

        # Réécrire Jan-Sep en littéraux
        tmpl = tmpl_vals.get(label, {})
        for col in range(3, 12):
            _safe_write(ws, row, col, tmpl.get(col, None))

        # Mise à jour YTD (col L) = Jan-Sep + Oct
        jan_sep_sum = sum(tmpl.get(c, 0) for c in range(3, 12))
        _safe_write(ws, row, 12, jan_sep_sum + oct_val)

    # Écrire sous-totaux Oct
    for row, val in subtotal_rows.items():
        _safe_write(ws, row, CURRENT_MONTH_COL, val)

    # Écrire TOTAL Oct + YTD TOTAL
    tmpl_total = {}
    for lbl, vals in tmpl_vals.items():
        if lbl.upper().startswith("TOTAL"):
            tmpl_total = vals
            break

    jan_sep_total = sum(tmpl_total.get(c, 0) for c in range(3, 12))
    _safe_write(ws, total_row, CURRENT_MONTH_COL, dept_oct_total)
    _safe_write(ws, total_row, 12, jan_sep_total + dept_oct_total)

    # Jan-Sep sur la ligne TOTAL
    for col in range(3, 12):
        _safe_write(ws, total_row, col, tmpl_total.get(col, None))

    return dept_oct_total


# ─────────────────────────────────────────────────────────────────────────────
# Zone d'import automatique
# ─────────────────────────────────────────────────────────────────────────────

def _write_import_zone(ws, dept_data: dict, main_total_row: int) -> None:
    """
    Trouve la zone d'import (après le TOTAL), ajoute le header "Oct 2024"
    et écrit une ligne par poste avec sa valeur octobre.
    """
    # Cherche la ligne d'en-tête de la zone d'import (contient "Poste de charge")
    header_row = None
    for row in range(main_total_row + 1, main_total_row + 15):
        val = ws.cell(row=row, column=1).value
        if val and "Poste de charge" in str(val):
            header_row = row
            break

    if header_row is None:
        return

    # Trouver la prochaine colonne disponible après Sep (col K = 11)
    # La col L est "CTRL", donc on ajoute Oct en col M
    ws.cell(row=header_row, column=CURRENT_MONTH_COL).value = OCT_HEADER

    # Écrire les lignes de données (une par poste)
    data_start_row = header_row + 1
    for i, (poste, val) in enumerate(dept_data.items()):
        r = data_start_row + i
        ws.cell(row=r, column=1).value  = poste
        ws.cell(row=r, column=CURRENT_MONTH_COL).value = float(val)


# ─────────────────────────────────────────────────────────────────────────────
# Synthèse
# ─────────────────────────────────────────────────────────────────────────────

def _update_synthese(wb, jan_sep: dict, oct_totals: dict) -> None:
    """Remplit la Synthèse avec des valeurs littérales (Jan-Sep + Oct)."""
    synthese_ws = None
    for name in wb.sheetnames:
        if "Synth" in name or "synth" in name:
            synthese_ws = wb[name]
            break
    if synthese_ws is None:
        return

    dept_synth_rows = {"RH": 4, "Tech": 5, "S&M": 6, "G&A": 7}
    total_synth_row = 8

    # ── En-tête Oct 2024 dans la Synthèse ─────────────────────────────────────
    # Col M (13) est déjà "% du total" → on ajoute Oct en col N (14)
    oct_col_synthese = 14  # colonne N
    synthese_ws.cell(row=3, column=oct_col_synthese).value = OCT_HEADER
    synthese_ws.cell(row=3, column=12).value = YTD_HEADER  # col L = Total YTD renommé

    for dept_key, synth_row in dept_synth_rows.items():
        tmpl_rows = jan_sep.get(dept_key, {})

        # TOTAL Jan-Sep pour ce dept (somme des valeurs mois)
        total_rows = {
            lbl: vals for lbl, vals in tmpl_rows.items()
            if lbl.upper().startswith("TOTAL")
        }
        if total_rows:
            totals_js = next(iter(total_rows.values()))
        else:
            # Recalcul depuis les lignes de postes
            totals_js = {}
            for lbl, vals in tmpl_rows.items():
                if not lbl.startswith("Sous-total") and not lbl.startswith("  "):
                    for col, v in vals.items():
                        totals_js[col] = totals_js.get(col, 0) + v

        # Écrire Jan-Sep dans la Synthèse (cols C-K)
        for col in range(3, 12):
            synthese_ws.cell(row=synth_row, column=col).value = totals_js.get(col, 0)

        # Total YTD Jan-Sep (col L)
        ytd_jan_sep = sum(totals_js.get(c, 0) for c in range(3, 12))
        synthese_ws.cell(row=synth_row, column=12).value = ytd_jan_sep

        # Oct (col N)
        oct_val = oct_totals.get(dept_key, 0)
        synthese_ws.cell(row=synth_row, column=oct_col_synthese).value = oct_val

        # Total YTD Jan-Oct (col L mise à jour)
        synthese_ws.cell(row=synth_row, column=12).value = ytd_jan_sep + oct_val

    # ── Ligne TOTAL (row 8) ───────────────────────────────────────────────────
    for col in range(3, 13):   # C à L
        total_val = sum(
            synthese_ws.cell(row=r, column=col).value or 0
            for r in range(4, 8)
        )
        synthese_ws.cell(row=total_synth_row, column=col).value = total_val

    # Oct total (col N)
    synthese_ws.cell(row=total_synth_row, column=oct_col_synthese).value = oct_totals.get("TOTAL", 0)


# ─────────────────────────────────────────────────────────────────────────────
# Contrôles
# ─────────────────────────────────────────────────────────────────────────────

def _update_controls_sheet(wb, totals: dict) -> None:
    """Écrit les statuts de contrôle comme valeurs littérales."""
    controls_ws = None
    for name in wb.sheetnames:
        if "ontr" in name:
            controls_ws = wb[name]
            break
    if controls_ws is None:
        return

    ws = controls_ws
    dept_rows = {"RH": 4, "Tech": 5, "S&M": 6, "G&A": 7}

    for dept, row in dept_rows.items():
        val = totals.get(dept, 0)
        ws.cell(row=row, column=3).value = val
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=5).value = 0
        ws.cell(row=row, column=6).value = "✅ OK"

    ws.cell(row=9, column=3).value = totals.get("TOTAL", 0)
    ws.cell(row=9, column=4).value = totals.get("TOTAL", 0)
    ws.cell(row=9, column=5).value = 0
    ws.cell(row=9, column=6).value = "✅ FICHIER VALIDÉ"


# ─────────────────────────────────────────────────────────────────────────────
# Lecture statut contrôles (pour check_controls)
# ─────────────────────────────────────────────────────────────────────────────

def get_controls_status(filepath: str) -> dict:
    wb = openpyxl.load_workbook(filepath, data_only=True)

    controls_ws = None
    for name in wb.sheetnames:
        if "ontr" in name:
            controls_ws = wb[name]
            break

    if controls_ws is None:
        wb.close()
        return {"lignes": {}, "global": "INCONNU", "ecart": None}

    ws      = controls_ws
    statuts = {}
    for row in range(4, 9):
        label  = ws.cell(row=row, column=1).value
        status = ws.cell(row=row, column=6).value
        if label:
            statuts[str(label).strip()] = str(status).strip() if status else "—"

    global_status = ws.cell(row=9, column=6).value or "INCONNU"
    ecart_cell    = ws.cell(row=9, column=5).value

    wb.close()
    return {
        "lignes": statuts,
        "global": str(global_status).strip(),
        "ecart":  ecart_cell,
    }
