import pandas as pd
from pathlib import Path

KNOWN_DEPTS = {"RH", "Tech", "S&M", "G&A"}
REQUIRED_COLS = {"departement", "poste", "categorie", "montant_octobre"}


def parse_sage_csv(filepath: str) -> dict[str, dict[str, float]]:
    """
    Lit le CSV Sage et retourne un dict structuré par département.

    Returns:
        {
            "RH":   {"Salaires fixes — tech & produit": 442, ...},
            "Tech": {...},
            "S&M":  {...},
            "G&A":  {...},
        }

    Raises:
        ValueError: colonnes manquantes, département inconnu, ou département absent.
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {filepath}")

    df = pd.read_csv(filepath, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans le CSV : {missing}")

    df["departement"] = df["departement"].str.strip()
    df["poste"] = df["poste"].str.strip()
    df["montant_octobre"] = pd.to_numeric(df["montant_octobre"], errors="coerce").fillna(0)

    unknown = set(df["departement"].unique()) - KNOWN_DEPTS
    if unknown:
        raise ValueError(f"Département(s) inconnu(s) dans le CSV : {unknown}")

    missing_depts = KNOWN_DEPTS - set(df["departement"].unique())
    if missing_depts:
        raise ValueError(f"Département(s) absents du CSV : {missing_depts}")

    result: dict[str, dict[str, float]] = {}
    for dept in KNOWN_DEPTS:
        subset = df[df["departement"] == dept]
        result[dept] = dict(zip(subset["poste"], subset["montant_octobre"]))

    return result


def parse_adp_csv(filepath: str) -> dict[str, float]:
    """
    Lit le CSV ADP et retourne un dict poste → montant.
    Utilisé pour surcharger / compléter les charges patronales RH.
    """
    path = Path(filepath)
    if not path.exists():
        return {}

    df = pd.read_csv(filepath, encoding="utf-8-sig")
    df.columns = df.columns.str.strip()

    if "poste" not in df.columns or "montant_octobre" not in df.columns:
        return {}

    df["montant_octobre"] = pd.to_numeric(df["montant_octobre"], errors="coerce").fillna(0)
    return dict(zip(df["poste"].str.strip(), df["montant_octobre"]))
