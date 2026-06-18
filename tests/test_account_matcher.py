"""
Tests unitaires de src/engine/account_matcher.py (P6 — rapprochement N/N-1).

Données synthétiques en mémoire + config rapprochements réelle du
mapping_pcg.yaml (seuil, poids, mots_vides).
"""

from pathlib import Path

import pytest

from src.engine.account_matcher import (
    appliquer_rapprochements,
    detecter_orphelins,
    proposer_rapprochements,
    scorer_matching,
)
from src.models.rapprochement import Rapprochement
from src.parsers.mapping_parser import from_pcg_config

PCG_PATH = Path(__file__).parent.parent / "src" / "config" / "mapping_pcg.yaml"


@pytest.fixture(scope="module")
def pcg() -> dict:
    return from_pcg_config(PCG_PATH)


@pytest.fixture(scope="module")
def config_rappro(pcg) -> dict:
    return pcg["rapprochements"]


def _balance_n1(comptes: dict) -> dict:
    """Construit une balance N-1 {num: {libelle, solde_ke}}."""
    return {num: {"libelle": lib, "solde_ke": 100.0}
            for num, lib in comptes.items()}


# ---------------------------------------------------------------------------
# Phase 1 — detecter_orphelins
# ---------------------------------------------------------------------------

def test_detecter_orphelins():
    """Les comptes sans correspondance exacte sont orphelins de chaque côté."""
    comptes_n = {"401000": "FOURNISSEURS", "5123001": "BANQUE CA"}
    balance_n1 = _balance_n1({"401000": "FOURNISSEURS", "512003": "BANQUE CA"})
    orphelins_n1, orphelins_n = detecter_orphelins(comptes_n, balance_n1)
    assert orphelins_n1 == {"512003": "BANQUE CA"}
    assert orphelins_n == {"5123001": "BANQUE CA"}


def test_detecter_orphelins_aucun():
    """Comptes identiques des deux côtés → aucun orphelin."""
    comptes_n = {"401000": "FOURNISSEURS"}
    balance_n1 = _balance_n1({"401000": "FOURNISSEURS"})
    orphelins_n1, orphelins_n = detecter_orphelins(comptes_n, balance_n1)
    assert orphelins_n1 == {} and orphelins_n == {}


# ---------------------------------------------------------------------------
# Phase 2 — scorer_matching
# ---------------------------------------------------------------------------

def test_score_compte_identique_sauf_numero(pcg, config_rappro):
    """512003 → 5123001, même libellé : préfixe 3/7, libellé 1.0, même
    cycle F et même classification Actif → score élevé."""
    r = scorer_matching("512003", "BANQUE CREDIT AGRICOLE",
                        "5123001", "BANQUE CREDIT AGRICOLE",
                        None, pcg, config_rappro)
    assert r.score_prefixe == pytest.approx(3 / 7, abs=1e-4)
    assert r.score_libelle == 1.0
    assert r.meme_cycle is True
    assert r.meme_classification is True
    attendu = 0.40 * (3 / 7) + 0.35 * 1.0 + 0.15 + 0.10
    assert r.score == pytest.approx(attendu, abs=1e-3)


def test_score_libelle_mots_vides(pcg, config_rappro):
    """Les mots vides (de, du, le…) sont ignorés dans la comparaison."""
    r = scorer_matching("606300", "ACHATS DE PETIT EQUIPEMENT",
                        "606310", "ACHATS PETIT EQUIPEMENT",
                        None, pcg, config_rappro)
    assert r.score_libelle == 1.0  # "de" ignoré → tokens identiques


def test_score_comptes_sans_rapport(pcg, config_rappro):
    """Comptes de classes et libellés différents → score < seuil."""
    r = scorer_matching("101000", "CAPITAL SOCIAL",
                        "607100", "ACHATS MARCHANDISES",
                        None, pcg, config_rappro)
    assert r.score < config_rappro["seuil"]
    assert r.meme_cycle is False


# ---------------------------------------------------------------------------
# Phase 2 — proposer_rapprochements
# ---------------------------------------------------------------------------

def test_proposer_filtre_par_seuil(pcg):
    """Seules les paires au-dessus du seuil sont proposées."""
    comptes_n = {"5123001": "BANQUE CREDIT AGRICOLE",
                 "623500": "CADEAUX CLIENTELE"}
    balance_n1 = _balance_n1({"512003": "BANQUE CREDIT AGRICOLE",
                              "101000": "CAPITAL SOCIAL"})
    propositions = proposer_rapprochements(comptes_n, balance_n1, None, pcg)
    assert len(propositions) == 1
    assert propositions[0].compte_n1 == "512003"
    assert propositions[0].compte_n == "5123001"


def test_proposer_affectation_un_pour_un(pcg):
    """Affectation gloutonne : chaque orphelin n'apparaît qu'une fois,
    la paire au meilleur score gagne."""
    comptes_n = {"512300": "BANQUE CREDIT AGRICOLE",
                 "512400": "BANQUE CREDIT AGRICOLE LYON"}
    balance_n1 = _balance_n1({"512003": "BANQUE CREDIT AGRICOLE"})
    propositions = proposer_rapprochements(comptes_n, balance_n1, None, pcg)
    assert len(propositions) == 1
    comptes_n1_proposes = [p.compte_n1 for p in propositions]
    assert comptes_n1_proposes.count("512003") == 1


def test_proposer_sans_orphelins(pcg):
    """Aucun orphelin → aucune proposition (liste vide)."""
    comptes_n = {"401000": "FOURNISSEURS"}
    balance_n1 = _balance_n1({"401000": "FOURNISSEURS"})
    assert proposer_rapprochements(comptes_n, balance_n1, None, pcg) == []


def test_proposer_garde_fou_granularite_collectif(pcg):
    """Un collectif 3 chiffres (À-NOUVEAU) n'est jamais rapproché d'un détail.

    Cas réel (CSVP) : le FEC porte le collectif "411" tandis que la balance
    N-1 a le détail "41110000". Malgré un préfixe et un libellé identiques,
    aucune fusion collectif↔détail ne doit être proposée.
    """
    comptes_n = {"411": "Collectif client"}
    balance_n1 = _balance_n1({"41110000": "Collectif client"})
    assert proposer_rapprochements(comptes_n, balance_n1, None, pcg) == []


# ---------------------------------------------------------------------------
# Phase 3 — appliquer_rapprochements
# ---------------------------------------------------------------------------

def test_appliquer_renumerote_balance_et_mapping():
    """La fusion déplace le solde N-1 et le mapping FM sous le numéro N,
    sans modifier les dicts d'entrée (jamais de modification en place)."""
    balance_n1 = _balance_n1({"512003": "BANQUE CA", "401000": "FOURNISSEURS"})
    mapping_fm = {"512003": {"cycle": "F", "compta": "Actif", "ref": "F0"}}
    r = Rapprochement(
        compte_n1="512003", libelle_n1="BANQUE CA",
        compte_n="5123001", libelle_n="BANQUE CA",
        score=0.8, score_prefixe=0.43, score_libelle=1.0,
        meme_cycle=True, meme_classification=True,
    )
    nouvelle_balance, nouveau_mapping = appliquer_rapprochements(
        balance_n1, [r], mapping_fm)

    assert "512003" not in nouvelle_balance
    assert nouvelle_balance["5123001"]["solde_ke"] == 100.0
    assert nouveau_mapping["5123001"]["cycle"] == "F"
    assert "512003" not in nouveau_mapping
    # Dicts d'entrée intacts
    assert "512003" in balance_n1
    assert "512003" in mapping_fm
