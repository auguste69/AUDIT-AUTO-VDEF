"""
Tests unitaires de src/engine/financial_engine.py — partie Bilan (+ AACE).

Les tests utilisent de PETITES balances synthétiques construites en mémoire
(pd.DataFrame) — jamais le FEC GILAC complet (trop lent). Les préfixes
proviennent de la config réelle src/config/mapping_pcg.yaml (section
liasse_fiscale), chargée une seule fois par module.

Couverture :
- agrégations par préfixe (_sommer_prefixes, _sommer_prefixes_si)
- signe de présentation du passif (× -1)
- valeurs nettes à l'actif (brut + amortissements négatifs)
- comptes bascule reclassés selon le signe du solde
- égalité Total Actif = Total Passif sur balance équilibrée
- structure BilanSynthetique / as_dict()
- filtrage AACE (filtrer_aace)
"""

from pathlib import Path
from typing import List, Tuple

import pandas as pd
import pytest

from src.engine.financial_engine import (
    _classer_comptes_bascule,
    _sommer_prefixes,
    _sommer_prefixes_si,
    calculer_bilan,
    filtrer_aace,
)
from src.engine.liasse_fiscale_loader import load_liasse_fiscale
from src.models.financial_statements import BilanSynthetique, PosteComptable
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


# ---------------------------------------------------------------------------
# Helpers d'agrégation par préfixe
# ---------------------------------------------------------------------------

def test_sommer_prefixes_agregation():
    """_sommer_prefixes additionne uniquement les comptes matchant un préfixe."""
    bal = _balance([
        ("213000", "Constructions",  60.0, 55.0),
        ("218300", "Matériel",       15.0, 10.0),
        ("411000", "Clients",        30.0, 20.0),
    ])
    assert _sommer_prefixes(bal, ["213", "218"], "Solde_KE") == 75.0
    assert _sommer_prefixes(bal, ["213", "218"], "Solde_N1_KE") == 65.0
    assert _sommer_prefixes(bal, ["999"], "Solde_KE") == 0.0


def test_sommer_prefixes_si_filtre_signe():
    """_sommer_prefixes_si filtre sur le signe de la valeur (>0 / <0)."""
    bal = _balance([
        ("455000", "Associé A",  10.0,  0.0),
        ("455100", "Associé B", -20.0, -4.0),
    ])
    assert _sommer_prefixes_si(bal, ["455"], "Solde_KE", ">0") == 10.0
    assert _sommer_prefixes_si(bal, ["455"], "Solde_KE", "<0") == -20.0
    # N-1 : 0.0 n'est ni >0 ni <0
    assert _sommer_prefixes_si(bal, ["455"], "Solde_N1_KE", ">0") == 0.0
    assert _sommer_prefixes_si(bal, ["455"], "Solde_N1_KE", "<0") == -4.0


def test_classer_comptes_bascule_selon_signe():
    """Les comptes bascule sont ventilés actif/passif selon le signe du solde."""
    bal = _balance([
        ("467000", "Débiteur divers",   5.0,  3.0),
        ("467100", "Créditeur divers", -7.0, -2.0),
    ])
    (actif_n, actif_n1), (passif_n, passif_n1) = _classer_comptes_bascule(
        bal, ["467"]
    )
    assert (actif_n, actif_n1) == (5.0, 3.0)
    # Convention passif : valeur positive
    assert (passif_n, passif_n1) == (7.0, 2.0)


def test_classer_comptes_bascule_colonnes_manquantes():
    """ValueError explicite si Solde_KE / Solde_N1_KE absentes."""
    bal = pd.DataFrame({"CompteNum": ["467000"], "Solde_KE": [5.0]})
    with pytest.raises(ValueError, match="Solde_N1_KE"):
        _classer_comptes_bascule(bal, ["467"])


# ---------------------------------------------------------------------------
# calculer_bilan
# ---------------------------------------------------------------------------

@pytest.fixture()
def balance_equilibree() -> pd.DataFrame:
    """Balance synthétique équilibrée (somme des soldes = 0 en N et N-1)."""
    return _balance([
        ("101000", "Capital",        -100.0, -90.0),
        ("213000", "Constructions",    60.0,  55.0),
        ("411000", "Clients",          30.0,  20.0),
        ("512000", "Banque",           10.0,  15.0),
    ])


def test_bilan_agregation_par_prefixe(balance_equilibree, liasse):
    """Chaque poste agrège les comptes de ses préfixes (N et N-1)."""
    bilan = calculer_bilan(balance_equilibree, liasse)
    assert isinstance(bilan, BilanSynthetique)
    assert bilan.postes["immo_corp"].as_tuple() == (60.0, 55.0)
    assert bilan.postes["crean_cli"].as_tuple() == (30.0, 20.0)
    assert bilan.postes["dispo"].as_tuple() == (10.0, 15.0)
    # Poste sans compte correspondant → zéro
    assert bilan.postes["stocks"].as_tuple() == (0.0, 0.0)


def test_bilan_signe_passif(balance_equilibree, liasse):
    """Les postes passif sont présentés en positif (solde créditeur × -1)."""
    bilan = calculer_bilan(balance_equilibree, liasse)
    assert bilan.postes["capital"].as_tuple() == (100.0, 90.0)


def test_bilan_total_actif_egal_passif(balance_equilibree, liasse):
    """Balance équilibrée → Total Actif = Total Passif en N et N-1."""
    bilan = calculer_bilan(balance_equilibree, liasse)
    assert bilan.total_actif.as_tuple() == (100.0, 90.0)
    assert bilan.total_actif.as_tuple() == bilan.total_passif.as_tuple()


def test_bilan_actif_net_amortissements(liasse):
    """L'actif est en valeur nette : brut + amortissements (soldes négatifs)."""
    bal = _balance([
        ("213000", "Constructions",       80.0,  70.0),
        ("281300", "Amort constructions", -20.0, -15.0),
    ])
    bilan = calculer_bilan(bal, liasse)
    assert bilan.postes["immo_corp"].as_tuple() == (60.0, 55.0)


def test_bilan_bascule_et_interco(liasse):
    """Bascule selon signe + sous-totaux interco (451/455) séparés."""
    bal = _balance([
        ("467000", "Débiteur divers",    5.0,  3.0),
        ("467100", "Créditeur divers",  -7.0, -2.0),
        ("455000", "Associé débiteur",  10.0,  0.0),
        ("451000", "Groupe créditeur", -20.0, -4.0),
    ])
    bilan = calculer_bilan(bal, liasse)
    assert bilan.postes["crean_interco"].as_tuple() == (10.0, 0.0)
    assert bilan.postes["dettes_interco"].as_tuple() == (20.0, 4.0)
    # Reclasses = bascule totale − interco
    assert bilan.postes["bascule_reclasses_actif"].as_tuple() == (5.0, 3.0)
    assert bilan.postes["bascule_reclasses_passif"].as_tuple() == (7.0, 2.0)
    # Les bascules alimentent les totaux : actif = 5+10, passif = 7+20
    assert bilan.total_actif.as_tuple() == (15.0, 3.0)
    assert bilan.total_passif.as_tuple() == (27.0, 6.0)


def test_bilan_as_dict_format_historique(balance_equilibree, liasse):
    """as_dict() restitue {cle: (valeur_n, valeur_n1)} avec les totaux."""
    d = calculer_bilan(balance_equilibree, liasse).as_dict()
    assert d["capital"] == (100.0, 90.0)
    assert d["total_actif"] == (100.0, 90.0)
    assert d["total_passif"] == (100.0, 90.0)
    assert all(isinstance(v, tuple) and len(v) == 2 for v in d.values())


def test_bilan_ne_modifie_pas_la_balance(balance_equilibree, liasse):
    """La balance d'entrée n'est jamais modifiée en place."""
    copie = balance_equilibree.copy(deep=True)
    calculer_bilan(balance_equilibree, liasse)
    pd.testing.assert_frame_equal(balance_equilibree, copie)


def test_poste_comptable_as_tuple():
    """PosteComptable.as_tuple() retourne (valeur_n, valeur_n1)."""
    poste = PosteComptable("capital", 100.0, 90.0)
    assert poste.as_tuple() == (100.0, 90.0)


# ---------------------------------------------------------------------------
# filtrer_aace
# ---------------------------------------------------------------------------

def test_filtrer_aace_prefixes_et_tri(liasse):
    """Seuls les comptes AACE (606-609, 61, 62) sont retenus, triés."""
    bal = _balance([
        ("627000", "Services bancaires", 3.0, 2.0),
        ("606100", "Fournitures",        5.0, 4.0),
        ("701000", "Ventes produits",  -50.0, -45.0),
        ("615000", "Entretien",          8.0, 7.0),
        ("601000", "Achats MP",         12.0, 11.0),
    ])
    df_aace = filtrer_aace(bal, liasse)
    assert df_aace["CompteNum"].tolist() == ["606100", "615000", "627000"]
    # Index réinitialisé
    assert df_aace.index.tolist() == [0, 1, 2]


def test_filtrer_aace_ne_modifie_pas_la_balance(liasse):
    """filtrer_aace retourne une copie — la balance d'entrée est intacte."""
    bal = _balance([
        ("615000", "Entretien", 8.0, 7.0),
        ("701000", "Ventes",  -50.0, -45.0),
    ])
    copie = bal.copy(deep=True)
    filtrer_aace(bal, liasse)
    pd.testing.assert_frame_equal(bal, copie)
