"""
Parseur de FEC (Fichier des Écritures Comptables).

Lit et nettoie un FEC .txt en détectant automatiquement l'encodage et le séparateur.
Produit un DataFrame pandas avec colonnes typées et la colonne Solde calculée.
"""

import logging
from pathlib import Path
from typing import Union

import pandas as pd

logger = logging.getLogger(__name__)

# Colonnes obligatoires du FEC (art. L.47 A LPF)
COLONNES_OBLIGATOIRES = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib", "Debit", "Credit",
    "EcritureLet", "DateLet", "ValidDate", "Montantdevise", "Idevise",
]

ENCODAGES_A_TESTER = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]
SEPARATEURS_A_TESTER = ["\t", "|", ";"]


def _detecter_encodage(chemin: Path) -> str:
    """Détecte l'encodage du fichier en testant les encodages courants sur le fichier entier."""
    for encodage in ENCODAGES_A_TESTER:
        try:
            chemin.read_text(encoding=encodage)
            logger.info("Encodage détecté : %s", encodage)
            return encodage
        except (UnicodeDecodeError, ValueError):
            continue
    raise ValueError(f"Impossible de décoder {chemin} avec les encodages testés : {ENCODAGES_A_TESTER}")


def _detecter_separateur(chemin: Path, encodage: str) -> str:
    """Détecte le séparateur en analysant la première ligne du fichier."""
    premiere_ligne = chemin.read_text(encoding=encodage).splitlines()[0]
    comptes = {sep: premiere_ligne.count(sep) for sep in SEPARATEURS_A_TESTER}
    separateur = max(comptes, key=comptes.get)
    if comptes[separateur] == 0:
        raise ValueError(f"Aucun séparateur reconnu dans : {premiere_ligne[:200]}")
    logger.info("Séparateur détecté : %r (%d occurrences)", separateur, comptes[separateur])
    return separateur


def _convertir_montant(serie: pd.Series) -> pd.Series:
    """Convertit une colonne montant (virgule française) en float."""
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(",", ".", regex=False)
        .str.replace(r"\s", "", regex=True)
    ).pipe(pd.to_numeric, errors="coerce").fillna(0.0)


def parse(source: Union[str, Path]) -> pd.DataFrame:
    """
    Lit et nettoie un FEC .txt.

    Paramètres
    ----------
    source : str ou Path
        Chemin vers le fichier FEC.

    Retourne
    --------
    pd.DataFrame
        DataFrame avec les 18 colonnes FEC typées + colonne Solde.

    Lève
    ----
    FileNotFoundError
        Si le fichier n'existe pas.
    ValueError
        Si l'encodage, le séparateur ou les colonnes obligatoires sont manquants.
    """
    chemin = Path(source)
    if not chemin.exists():
        raise FileNotFoundError(f"Fichier FEC introuvable : {chemin}")

    encodage = _detecter_encodage(chemin)
    separateur = _detecter_separateur(chemin, encodage)

    df = pd.read_csv(
        chemin,
        sep=separateur,
        encoding=encodage,
        dtype=str,
        keep_default_na=False,
    )

    # Nettoyer les noms de colonnes (BOM résiduel, espaces)
    df.columns = df.columns.str.strip().str.replace("\ufeff", "", regex=False)

    # Nettoyer les valeurs string : strip + suppression des retours chariot Windows
    colonnes_str = df.select_dtypes("object").columns
    df[colonnes_str] = df[colonnes_str].apply(
        lambda col: col.str.strip().str.replace("\r", "", regex=False)
    )

    # Valider les colonnes obligatoires
    manquantes = [c for c in COLONNES_OBLIGATOIRES if c not in df.columns]
    if manquantes:
        raise ValueError(f"Colonnes obligatoires manquantes dans le FEC : {manquantes}")

    # Convertir les montants
    df["Debit"] = _convertir_montant(df["Debit"])
    df["Credit"] = _convertir_montant(df["Credit"])
    df["Montantdevise"] = _convertir_montant(df["Montantdevise"])

    # Ajouter la colonne Solde
    df["Solde"] = df["Debit"] - df["Credit"]

    logger.info(
        "FEC chargé : %d lignes, %d colonnes, encodage=%s",
        len(df), len(df.columns), encodage,
    )
    return df
