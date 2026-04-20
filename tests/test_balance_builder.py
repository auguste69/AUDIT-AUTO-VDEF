"""
Tests unitaires pour src/engine/balance_builder.py.
Utilise le vrai FEC GILAC + les données N-1 extraites du FM existant.
"""

from pathlib import Path

import openpyxl
import pytest
import pandas as pd

from src.parsers.fec_parser import parse
from src.engine.balance_builder import build

DATA_DIR = Path(__file__).parent.parent / "data"
FEC_PATH = DATA_DIR / "GILAC_2025_12_31_FEC.txt"
FM_PATH  = DATA_DIR / "FM GILAC.xlsx"


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def df_fec() -> pd.DataFrame:
    return parse(FEC_PATH)


@pytest.fixture(scope="module")
def balance_n1() -> dict:
    """Extrait les soldes N-1 depuis l'onglet 'Balance N Vs N-1' du FM."""
    wb = openpyxl.load_workbook(FM_PATH, read_only=True, data_only=True)
    ws = wb["Balance N Vs N-1"]
    n1: dict = {}
    for row in ws.iter_rows(min_row=10, values_only=True):
        compte_num = row[1]  # colonne B
        libelle    = row[2]  # colonne C
        solde_n1   = row[4]  # colonne E (N-1)
        if compte_num is None:
            continue
        try:
            num_str = str(int(float(compte_num)))
        except (ValueError, TypeError):
            continue
        n1[num_str] = {
            "libelle": str(libelle) if libelle else "",
            "solde_ke": float(solde_n1) if solde_n1 is not None else 0.0,
        }
    return n1


@pytest.fixture(scope="module")
def balance(df_fec, balance_n1) -> pd.DataFrame:
    return build(df_fec, balance_n1)


@pytest.fixture(scope="module")
def balance_sans_n1(df_fec) -> pd.DataFrame:
    return build(df_fec)


# ---------------------------------------------------------------------------
# Structure du DataFrame
# ---------------------------------------------------------------------------

COLONNES_ATTENDUES = [
    "CompteNum", "CompteLib", "Debit", "Credit", "Solde",
    "Solde_KE", "Solde_N1_KE", "Var_KE", "Var_PCT",
]

def test_colonnes_presentes(balance):
    assert list(balance.columns) == COLONNES_ATTENDUES


def test_nb_comptes(balance):
    """Le FEC GILAC 2025 contient 270 comptes distincts."""
    assert len(balance) == 270


def test_trie_par_comptenum(balance):
    nums = balance["CompteNum"].tolist()
    assert nums == sorted(nums)


# ---------------------------------------------------------------------------
# Montants et conversions
# ---------------------------------------------------------------------------

def test_solde_ke_conversion(balance):
    """Solde_KE = Solde / 1000, arrondi à 3 décimales."""
    diff = (balance["Solde"] / 1000 - balance["Solde_KE"]).abs()
    assert diff.max() < 0.001


def test_var_ke_calcul(balance):
    """Var_KE = Solde_KE - Solde_N1_KE."""
    diff = (balance["Solde_KE"] - balance["Solde_N1_KE"] - balance["Var_KE"]).abs()
    assert diff.max() < 0.001


# ---------------------------------------------------------------------------
# Équilibre comptable
# ---------------------------------------------------------------------------

def test_somme_soldes_nulle(balance):
    """La somme des soldes N doit être nulle (FEC équilibré)."""
    assert abs(balance["Solde"].sum()) < 0.01


def test_somme_debits_credits(balance):
    """La somme Débit - Crédit doit reconstruire la somme des Soldes."""
    ecart = abs((balance["Debit"] - balance["Credit"]).sum() - balance["Solde"].sum())
    assert ecart < 0.01


# ---------------------------------------------------------------------------
# Variation N vs N-1
# ---------------------------------------------------------------------------

def test_var_pct_na_quand_n1_nul(balance):
    """Var_PCT = 'n/a' quand Solde_N1_KE ≈ 0 (pas de division par zéro)."""
    masque_n1_nul = balance["Solde_N1_KE"].abs() < 0.001
    vals = balance.loc[masque_n1_nul, "Var_PCT"]
    assert (vals == "n/a").all(), "Var_PCT devrait être 'n/a' quand N-1 ≈ 0"


def test_var_pct_numerique_quand_n1_non_nul(balance):
    """Var_PCT est un float quand Solde_N1_KE != 0."""
    masque = balance["Solde_N1_KE"].abs() >= 0.001
    vals = balance.loc[masque, "Var_PCT"]
    assert vals.apply(lambda v: isinstance(v, float)).all()


def test_n1_bien_charge(balance, balance_n1):
    """Les soldes N-1 connus sont correctement injectés."""
    # 101300 = CAPITAL, N-1 = -1150.001 K€ dans le FM
    ligne = balance[balance["CompteNum"] == "101300"]
    assert not ligne.empty, "Compte 101300 absent de la balance"
    assert abs(ligne.iloc[0]["Solde_N1_KE"] - (-1150.001)) < 0.01


# ---------------------------------------------------------------------------
# Sans données N-1
# ---------------------------------------------------------------------------

def test_sans_n1_solde_n1_zero(balance_sans_n1):
    """Sans balance N-1, Solde_N1_KE = 0 partout."""
    assert (balance_sans_n1["Solde_N1_KE"] == 0.0).all()


def test_sans_n1_var_pct_na(balance_sans_n1):
    """Sans balance N-1, Var_PCT = 'n/a' partout (N-1 = 0)."""
    assert (balance_sans_n1["Var_PCT"] == "n/a").all()


def test_sans_n1_var_ke_egale_solde_ke(balance_sans_n1):
    """Sans N-1, Var_KE = Solde_KE."""
    diff = (balance_sans_n1["Solde_KE"] - balance_sans_n1["Var_KE"]).abs()
    assert diff.max() < 0.001


# ---------------------------------------------------------------------------
# Erreur sur FEC déséquilibré
# ---------------------------------------------------------------------------

def test_fec_desequilibre_leve_erreur(df_fec):
    """Un FEC avec somme des soldes != 0 doit lever ValueError."""
    df_corrompu = df_fec.copy()
    df_corrompu.loc[0, "Solde"] += 1000.0  # déséquilibre artificiel
    with pytest.raises(ValueError, match="déséquilibré"):
        build(df_corrompu)
