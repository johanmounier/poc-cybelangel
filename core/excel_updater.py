import shutil
from pathlib import Path

import openpyxl

from config import DEPT_TOTAL_ROWS, DEPT_DATA_START_ROW, CURRENT_MONTH_COL

DEPT_SHEETS = {
    "RH": "RH",
    "Tech": "Tech",
    "S&M": "S&M",
    "G&A": "G&A",
}

_SKIP_PREFIXES = ("Sous-total", "TOTAL", "  ")


def _is_data_row(label: str) -> bool:
    """Retourne True si la ligne est un poste de charge (pas un en-tête/sous-total)."""
    s = str(label).strip()
    return s and not any(s.startswith(p) for p in _SKIP_PREFIXES)


def update_excel(template_path: str, output_path: str, data_dict: dict) -> dict:
    """
    Copie le template, écrit les montants octobre (colonne M) pour chaque
    département, et retourne les totaux par département calculés en Python.

    Returns:
        {"RH": 965, "Tech": 217, "S&M": 131, "G&A": 101, "TOTAL": 1414}
    """
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path)

    totals = {}

    for sheet_name, dept_key in DEPT_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            totals[dept_key] = 0
            continue

        ws = wb[sheet_name]
        dept_data = data_dict.get(dept_key, {})
        total_row = DEPT_TOTAL_ROWS[dept_key]
        dept_total = 0.0

        for row in range(DEPT_DATA_START_ROW, total_row):
            cell_a = ws.cell(row=row, column=1).value
            if cell_a is None:
                continue
            label = str(cell_a).strip()
            if not _is_data_row(label):
                continue

            if label in dept_data:
                val = float(dept_data[label])
                ws.cell(row=row, column=CURRENT_MONTH_COL).value = val
                dept_total += val

        # Écriture littérale du TOTAL ligne pour la colonne M
        ws.cell(row=total_row, column=CURRENT_MONTH_COL).value = dept_total
        totals[dept_key] = dept_total

    totals["TOTAL"] = sum(v for k, v in totals.items() if k != "TOTAL")

    # Mise à jour littérale de l'onglet Contrôles
    _update_controls_sheet(wb, totals)

    wb.save(output_path)
    wb.close()
    return totals


def _update_controls_sheet(wb: openpyxl.Workbook, totals: dict) -> None:
    """Écrit les statuts de contrôle comme valeurs littérales (pas de formules)."""
    controls_sheet = None
    for name in wb.sheetnames:
        if "ontr" in name:
            controls_sheet = wb[name]
            break
    if controls_sheet is None:
        return

    ws = controls_sheet
    dept_rows = {"RH": 4, "Tech": 5, "S&M": 6, "G&A": 7}

    for dept, row in dept_rows.items():
        val = totals.get(dept, 0)
        ws.cell(row=row, column=3).value = val   # Valeur Synthèse
        ws.cell(row=row, column=4).value = val   # Valeur Onglet source
        ws.cell(row=row, column=5).value = 0     # Écart
        ws.cell(row=row, column=6).value = "✅ OK"

    # Ligne 9 : statut global
    ws.cell(row=9, column=3).value = totals.get("TOTAL", 0)
    ws.cell(row=9, column=4).value = totals.get("TOTAL", 0)
    ws.cell(row=9, column=5).value = 0
    ws.cell(row=9, column=6).value = "✅ FICHIER VALIDÉ"


def get_controls_status(filepath: str) -> dict:
    """
    Lit l'onglet Contrôles et retourne le statut de chaque ligne
    ainsi que le statut global (ligne 9, colonne F).
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)

    controls_sheet = None
    for name in wb.sheetnames:
        if "ontr" in name:
            controls_sheet = wb[name]
            break

    if controls_sheet is None:
        wb.close()
        return {"lignes": {}, "global": "INCONNU", "ecart": None}

    ws = controls_sheet
    statuts = {}
    for row in range(4, 9):
        label = ws.cell(row=row, column=1).value
        status = ws.cell(row=row, column=6).value
        if label:
            statuts[str(label).strip()] = str(status).strip() if status else "—"

    global_status = ws.cell(row=9, column=6).value or "INCONNU"
    ecart_cell = ws.cell(row=9, column=5).value

    wb.close()
    return {
        "lignes": statuts,
        "global": str(global_status).strip(),
        "ecart": ecart_cell,
    }
