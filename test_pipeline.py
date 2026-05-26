"""Test rapide du pipeline sans Streamlit."""
import sys
sys.path.insert(0, ".")
import config
from core.csv_parser import parse_sage_csv
from core.excel_updater import update_excel
from core.controls_checker import check_controls
from pathlib import Path

Path("output").mkdir(exist_ok=True)

print("1. Parsing CSV Sage...")
data = parse_sage_csv("data/sample_sage.csv")
for dept, postes in data.items():
    total = sum(postes.values())
    print(f"   {dept}: {len(postes)} postes, total = {total:.0f} k€")

print()
print("2. Mise a jour Excel...")
out_excel = "output/test_consolidation.xlsx"
totals = update_excel(config.TEMPLATE_EXCEL, out_excel, data)
print("   Totaux octobre :")
for k, v in totals.items():
    print(f"   {k}: {v:.0f} k€")

print()
print("3. Verification controles...")
is_valid, status = check_controls(out_excel)
print(f"   Statut global : {status['global']}")
print(f"   Valide : {is_valid}")
for label, st in status["lignes"].items():
    print(f"   {label}: {st}")

print()
if is_valid:
    print(">>> Pipeline OK - tout fonctionne!")
else:
    print(">>> ATTENTION: controles non valides")
