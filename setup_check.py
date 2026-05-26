"""
Lance ce script une fois pour vérifier que l'environnement est prêt.
Usage : python setup_check.py
"""
import sys
from pathlib import Path

OK = "\033[92m✅\033[0m"
WARN = "\033[93m⚠️\033[0m"
ERR = "\033[91m❌\033[0m"

errors = 0

print("\n══ Vérification de l'environnement POC CybelAngel ══\n")

# Python version
if sys.version_info < (3, 11):
    print(f"{WARN} Python {sys.version_info.major}.{sys.version_info.minor} détecté — Python 3.11+ recommandé")
else:
    print(f"{OK} Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")

# Packages
packages = ["streamlit", "openpyxl", "anthropic", "reportlab", "pandas"]
for pkg in packages:
    try:
        __import__(pkg)
        print(f"{OK} {pkg}")
    except ImportError:
        print(f"{ERR} {pkg} manquant — lancez : pip install -r requirements.txt")
        errors += 1

# Template Excel
template = Path("data/template.xlsx")
if template.exists():
    print(f"{OK} Template Excel : {template}")
else:
    print(f"{ERR} Template Excel manquant : {template}")
    print("     → Copiez cybelangel_consolidation_v2.xlsx dans data/template.xlsx")
    errors += 1

# CSV de démo
for csv_file in ["data/sample_sage.csv", "data/sample_adp.csv"]:
    p = Path(csv_file)
    if p.exists():
        print(f"{OK} {csv_file}")
    else:
        print(f"{ERR} {csv_file} manquant")
        errors += 1

# config.py
try:
    import config
    checks = {
        "ANTHROPIC_API_KEY": config.ANTHROPIC_API_KEY,
        "SENDER_EMAIL": config.SENDER_EMAIL,
        "OUTLOOK_PASSWORD": config.OUTLOOK_PASSWORD,
        "RECIPIENT_EMAIL": config.RECIPIENT_EMAIL,
    }
    for key, val in checks.items():
        if val:
            print(f"{OK} config.{key} configuré")
        else:
            print(f"{WARN} config.{key} vide (optionnel pour la démo)")
except Exception as e:
    print(f"{ERR} Erreur import config : {e}")
    errors += 1

# Output dir
Path("output").mkdir(exist_ok=True)
print(f"{OK} Dossier output/ prêt")

print("\n" + "══" * 25)
if errors == 0:
    print(f"{OK} Tout est prêt ! Lancez : streamlit run app.py\n")
else:
    print(f"{ERR} {errors} problème(s) à corriger avant le lancement.\n")
