"""
Tests unitaires pour src/parsers/fec_parser.py.
Utilise le vrai FEC GILAC dans data/.
"""

import math
from pathlib import Path

import pytest
import pandas as pd

from src.parsers.fec_parser import parse, COLONNES_OBLIGATOIRES

DATA_DIR = Path(__file__).parent.parent / "data"
FEC_PATH = DATA_DIR / "GILAC_2025_12_31_FEC.txt"


@pytest.fixture(scope="module")
def fec() -> pd.DataFrame:
    return parse(FEC_PATH)


# ---------------------------------------------------------------------------
# Structure
# ---------------------------------------------------------------------------

def test_colonnes_obligatoires_presentes(fec):
    for col in COLONNES_OBLIGATOIRES:
        assert col in fec.columns, f"Colonne manquante : {col}"


def test_colonne_solde_presente(fec):
    assert "Solde" in fec.columns


def test_nombre_lignes(fec):
    """Le FEC GILAC doit contenir ~106 272 lignes."""
    assert len(fec) > 100_000


# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------

def test_debit_numerique(fec):
    assert pd.api.types.is_float_dtype(fec["Debit"]), "Debit doit être float"


def test_credit_numerique(fec):
    assert pd.api.types.is_float_dtype(fec["Credit"]), "Credit doit être float"


def test_montantdevise_numerique(fec):
    assert pd.api.types.is_float_dtype(fec["Montantdevise"]), "Montantdevise doit être float"


def test_solde_calcule_correctement(fec):
    """Solde = Debit - Credit sur chaque ligne."""
    diff = (fec["Debit"] - fec["Credit"] - fec["Solde"]).abs()
    assert diff.max() < 1e-6, f"Solde mal calculé, écart max = {diff.max()}"


# ---------------------------------------------------------------------------
# Équilibre comptable
# ---------------------------------------------------------------------------

def test_equilibre_debit_credit(fec):
    """La somme des Débits doit égaler la somme des Crédits (écart < 0,01 €)."""
    ecart = abs(fec["Debit"].sum() - fec["Credit"].sum())
    assert ecart < 0.01, f"Déséquilibre Débit/Crédit : {ecart:.2f} €"


def test_somme_soldes_nulle(fec):
    """La somme de tous les soldes doit être nulle (équilibre comptable)."""
    somme = fec["Solde"].sum()
    assert abs(somme) < 0.01, f"Somme des soldes != 0 : {somme:.2f}"


# ---------------------------------------------------------------------------
# Nettoyage
# ---------------------------------------------------------------------------

def test_pas_de_retour_chariot(fec):
    """Aucune valeur string ne doit contenir de \\r."""
    for col in fec.select_dtypes("object").columns:
        assert not fec[col].str.contains("\r").any(), f"Retour chariot trouvé dans {col}"


def test_pas_de_valeur_na_dans_comptenum(fec):
    """CompteNum ne doit pas contenir de valeurs vides."""
    assert fec["CompteNum"].str.strip().ne("").all(), "CompteNum contient des valeurs vides"


# ---------------------------------------------------------------------------
# Erreurs attendues
# ---------------------------------------------------------------------------

def test_fichier_inexistant():
    with pytest.raises(FileNotFoundError):
        parse("/tmp/fichier_qui_nexiste_pas.txt")
