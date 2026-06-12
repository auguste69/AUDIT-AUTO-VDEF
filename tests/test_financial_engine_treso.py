"""
Tests unitaires de src/engine/financial_engine.calculer_treso.

Les tests utilisent de PETITES balances synthétiques construites en mémoire
(pd.DataFrame) — jamais le FEC GILAC complet (trop lent). Les préfixes
proviennent de la config réelle src/config/mapping_pcg.yaml (section
liasse_fiscale), chargée une seule fois par module.

Couverture :
- ressources stables présentées en positif (signe passif × -1)
- emplois stables en brut (hors amortissements)
- filtres sur le signe du solde (autres créances >0, dettes <0)
- cohérence fondamentale FRNG = BFR + TN (N et N-1)
- vérification directe TN = trésorerie active − trésorerie passive
- trésorerie passive : 519 et 512 créditeur
- structure TresoSynthetique / as_dict()
"""

from pathlib import Path
from typing import List, Tuple

import pandas as pd
import pytest

from src.engine.financial_engine import calculer_treso
from src.engine.liasse_fiscale_loader import load_liasse_fiscale
from src.models.financial_statements import TresoSynthetique
from src.parsers.mapping_parser import from_pcg_config

PCG_PATH = Path(__file__).parent.parent / "src" / "config" / "mapping_pcg.yaml"


@pytest.fixture(scope="module")
def liasse() -> dict:
    """Charge la section liasse_fiscale de la config PCG réelle."""
    return load_liasse_fiscale(from_pcg_config(PCG_PATH))


def _balance(rows: List[Tuple[str, str, float, float]]) -> pd.DataFrame:
    """Construit une balance synthétique minimale en mémoire."""
    return pd.DataFrame(
        rows, columns=["CompteNum", "CompteLib", "Solde_KE", "Solde_N1_KE"]
    )


@pytest.fixture()
def balance_treso() -> pd.DataFrame:
    """Balance synthétique équilibrée (somme = 0 en N et N-1)."""
    return _balance([
        ("101000", "Capital",             -100.0, -90.0),
        ("213000", "Constructions",         80.0,  70.0),
        ("281300", "Amort constructions",  -20.0, -15.0),
        ("311000", "Stocks MP",             30.0,  25.0),
        ("411000", "Clients",               25.0,  20.0),
        ("401000", "Fournisseurs",         -40.0, -30.0),
        ("512000", "Banque",                25.0,  20.0),
    ])


def test_treso_ressources_stables_signe_passif(balance_treso, liasse):
    """Ressources stables présentées en positif (capitaux propres, amort)."""
    treso = calculer_treso(balance_treso, liasse)
    assert isinstance(treso, TresoSynthetique)
    assert treso.postes["cap_propres"].as_tuple() == (100.0, 90.0)
    # Les amortissements (28x, créditeurs) comptent en ressources, positifs
    assert treso.postes["amort_dep"].as_tuple() == (20.0, 15.0)
    assert treso.postes["total_res"].as_tuple() == (120.0, 105.0)


def test_treso_emplois_stables_en_brut(balance_treso, liasse):
    """Emplois stables = actif immobilisé BRUT (20-27, hors 28x)."""
    treso = calculer_treso(balance_treso, liasse)
    assert treso.postes["actif_immo"].as_tuple() == (80.0, 70.0)
    assert treso.postes["total_emp"].as_tuple() == (80.0, 70.0)


def test_treso_frng(balance_treso, liasse):
    """FRNG = ressources stables − emplois stables."""
    treso = calculer_treso(balance_treso, liasse)
    assert treso.frng.as_tuple() == (40.0, 35.0)


def test_treso_bfr_filtres_signe(balance_treso, liasse):
    """BFR : créances >0 à l'actif circulant, dettes <0 au passif circulant."""
    treso = calculer_treso(balance_treso, liasse)
    assert treso.postes["stocks"].as_tuple() == (30.0, 25.0)
    assert treso.postes["crean_cli"].as_tuple() == (25.0, 20.0)
    # 401000 créditeur : exclu des autres créances (condition >0)…
    assert treso.postes["autres_crean"].as_tuple() == (0.0, 0.0)
    assert treso.postes["total_ac"].as_tuple() == (55.0, 45.0)
    # … mais compté en dettes fournisseurs, présenté positif
    assert treso.postes["det_fourn"].as_tuple() == (40.0, 30.0)
    assert treso.postes["total_pc"].as_tuple() == (40.0, 30.0)
    assert treso.bfr.as_tuple() == (15.0, 15.0)


def test_treso_frng_egale_bfr_plus_tn(balance_treso, liasse):
    """Cohérence fondamentale : FRNG = BFR + TN en N et N-1."""
    treso = calculer_treso(balance_treso, liasse)
    for idx in (0, 1):
        assert treso.frng.as_tuple()[idx] == pytest.approx(
            treso.bfr.as_tuple()[idx] + treso.tn.as_tuple()[idx], abs=0.001
        )
    assert treso.tn.as_tuple() == (25.0, 20.0)


def test_treso_verification_directe(balance_treso, liasse):
    """TN (vérification directe) = trésorerie active − passive = TN."""
    treso = calculer_treso(balance_treso, liasse)
    assert treso.postes["treso_active"].as_tuple() == (25.0, 20.0)
    assert treso.postes["treso_passive"].as_tuple() == (0.0, 0.0)
    assert treso.postes["tn_verif"].as_tuple() == treso.tn.as_tuple()


def test_treso_passive_519_et_512_crediteur(liasse):
    """Trésorerie passive = CBC (519) + banques créditrices (512 < 0)."""
    bal = _balance([
        ("512000", "Banque créditrice", -25.0, -10.0),
        ("519000", "CBC",               -10.0,  -5.0),
    ])
    treso = calculer_treso(bal, liasse)
    # Aucune trésorerie active (pas de solde 50/51/53 positif)
    assert treso.postes["treso_active"].as_tuple() == (0.0, 0.0)
    assert treso.postes["treso_passive"].as_tuple() == (35.0, 15.0)
    assert treso.postes["tn_verif"].as_tuple() == (-35.0, -15.0)


def test_treso_as_dict_format_historique(balance_treso, liasse):
    """as_dict() restitue {cle: (valeur_n, valeur_n1)} avec frng/bfr/tn."""
    d = calculer_treso(balance_treso, liasse).as_dict()
    assert d["frng"] == (40.0, 35.0)
    assert d["bfr"] == (15.0, 15.0)
    assert d["tn"] == (25.0, 20.0)
    assert d["cap_propres"] == (100.0, 90.0)
    assert all(isinstance(v, tuple) and len(v) == 2 for v in d.values())


def test_treso_ne_modifie_pas_la_balance(balance_treso, liasse):
    """La balance d'entrée n'est jamais modifiée en place."""
    copie = balance_treso.copy(deep=True)
    calculer_treso(balance_treso, liasse)
    pd.testing.assert_frame_equal(balance_treso, copie)


# ---------------------------------------------------------------------------
# Corrections audit : résultat en cours, c/c associés, cohérence TN
# ---------------------------------------------------------------------------

def test_treso_resultat_en_cours_dans_ressources(liasse):
    """Le résultat non clôturé (classes 6/7) compte en ressources stables."""
    bal = _balance([
        ("601000", "Achats",  60.0,  50.0),
        ("701000", "Ventes", -100.0, -80.0),
        ("512000", "Banque",  40.0,  30.0),
    ])
    treso = calculer_treso(bal, liasse)
    # Bénéfice de 40 (N) / 30 (N-1) → ressource positive
    assert treso.postes["resultat_enc"].as_tuple() == (40.0, 30.0)
    assert treso.tn.as_tuple() == (40.0, 30.0)
    assert treso.postes["tn_verif"].as_tuple() == (40.0, 30.0)


def test_treso_comptes_courants_associes_crediteurs(liasse):
    """Les c/c d'associés créditeurs (45x) sont capturés au passif circulant."""
    bal = _balance([
        ("455100", "C/C associé", -70.0, -60.0),
        ("512000", "Banque",       70.0,  60.0),
    ])
    treso = calculer_treso(bal, liasse)
    assert treso.postes["autres_det"].as_tuple() == (70.0, 60.0)
    assert treso.tn.as_tuple() == (70.0, 60.0)
    assert treso.postes["tn_verif"].as_tuple() == (70.0, 60.0)


def test_treso_fournisseur_debiteur_pas_de_double_comptage(liasse):
    """Un 401 débiteur va dans les autres créances, PAS en moins des dettes."""
    bal = _balance([
        ("401100", "Fournisseur débiteur",  15.0,  10.0),
        ("401200", "Fournisseurs",         -50.0, -40.0),
        ("512000", "Banque",                35.0,  30.0),
    ])
    treso = calculer_treso(bal, liasse)
    assert treso.postes["autres_crean"].as_tuple() == (15.0, 10.0)
    assert treso.postes["det_fourn"].as_tuple() == (50.0, 40.0)
    assert treso.tn.as_tuple() == treso.postes["tn_verif"].as_tuple()


def test_treso_coherence_tn_balance_complexe(liasse):
    """TN (FRNG − BFR) = trésorerie directe sur une balance hétérogène."""
    bal = _balance([
        ("101000", "Capital",            -100.0,  -90.0),
        ("213000", "Constructions",        80.0,   70.0),
        ("281300", "Amort constructions", -20.0,  -15.0),
        ("311000", "Stocks MP",            30.0,   25.0),
        ("419100", "Clients créditeurs",  -12.0,   -8.0),
        ("455100", "C/C associé",         -25.0,  -20.0),
        ("486000", "CCA",                   5.0,    4.0),
        ("601000", "Achats",               60.0,   50.0),
        ("701000", "Ventes",              -90.0,  -76.0),
        ("519000", "CBC",                 -10.0,   -5.0),
        ("512000", "Banque",               82.0,   65.0),
    ])
    treso = calculer_treso(bal, liasse)
    for idx in (0, 1):
        assert treso.tn.as_tuple()[idx] == pytest.approx(
            treso.postes["tn_verif"].as_tuple()[idx], abs=0.001
        )
    # 519 créditeur en trésorerie passive (classe 5 < 0), jamais en double
    assert treso.postes["treso_passive"].as_tuple() == (10.0, 5.0)
    assert treso.postes["treso_active"].as_tuple() == (82.0, 65.0)
