"""
Tests unitaires du contrôle de cohérence du résultat — src/engine/controls.py.

Vérifie que le résultat déduit de la balance (− somme des classes 6 et 7)
est cohérent avec le poste résultat du bilan (12x + résultat en cours).
Utilise des balances synthétiques en mémoire (pas le FEC GILAC complet).
"""

from pathlib import Path

import pandas as pd
import pytest

from src.engine.controls import run_controles_financiers
from src.engine.liasse_fiscale_loader import load_liasse_fiscale
from src.parsers.mapping_parser import from_pcg_config

PCG_PATH = Path(__file__).parent.parent / "src" / "config" / "mapping_pcg.yaml"
NOM_COHERENCE = "Cohérence du résultat"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def liasse_config() -> dict:
    """Section liasse_fiscale réelle, chargée depuis mapping_pcg.yaml."""
    return load_liasse_fiscale(from_pcg_config(PCG_PATH))


def _resultats_dict(resultats):
    """Convertit la liste de résultats en dict {nom: (ok, detail, severity)}."""
    return {nom: (ok, detail, severity) for nom, ok, detail, severity in resultats}


def _balance_coherente() -> pd.DataFrame:
    """
    Balance synthétique cohérente : résultat porté par les classes 6/7
    (exercice non clôturé), compte 12x absent.

    Résultat déduit = −(150 − 200) = 50 K€ = résultat en cours du bilan.
    """
    return pd.DataFrame({
        "CompteNum": ["101000", "411000", "512000", "606000", "706000"],
        "CompteLib": ["Capital", "Clients", "Banque", "Achats", "Ventes"],
        "Solde_KE": [-100.0, 120.0, 30.0, 150.0, -200.0],
        "Solde_N1_KE": [-100.0, 120.0, 30.0, 150.0, -200.0],
    })


def _balance_incoherente() -> pd.DataFrame:
    """
    Balance synthétique incohérente : un résultat de 80 K€ est déjà affecté
    au compte 120000 ALORS QUE les classes 6/7 portent encore un résultat
    en cours de 50 K€.

    Résultat déduit de la balance = 50 K€ ;
    poste résultat du bilan = 80 (12x) + 50 (en cours) = 130 K€ → écart 80 K€.
    """
    return pd.DataFrame({
        "CompteNum": ["101000", "120000", "411000", "512000",
                      "606000", "706000"],
        "CompteLib": ["Capital", "Résultat", "Clients", "Banque",
                      "Achats", "Ventes"],
        "Solde_KE": [-100.0, -80.0, 200.0, 30.0, 150.0, -200.0],
        "Solde_N1_KE": [-100.0, -80.0, 200.0, 30.0, 150.0, -200.0],
    })


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestCoherenceResultat:

    def test_cas_coherent_ok(self, liasse_config):
        """Résultat balance = résultat bilan → ok=True, severity WARNING."""
        res = _resultats_dict(
            run_controles_financiers(_balance_coherente(), liasse_config)
        )
        ok, detail, sev = res[NOM_COHERENCE]
        assert ok is True
        assert sev == "WARNING"
        assert "écart=" in detail

    def test_cas_incoherent_warning_sans_exception(self, liasse_config):
        """Cas incohérent → ok=False, severity WARNING, aucune exception levée."""
        res = _resultats_dict(
            run_controles_financiers(_balance_incoherente(), liasse_config)
        )
        ok, detail, sev = res[NOM_COHERENCE]
        assert ok is False
        assert sev == "WARNING"   # jamais BLOQUANT
        assert "incohérence" in detail

    def test_tolerance_1_ke(self, liasse_config):
        """Écart ≤ 1 K€ → toléré (ok=True)."""
        bal = _balance_coherente()
        # Léger écart : résultat affecté en 12x pour 0.8 K€ + contrepartie banque
        ligne_12 = pd.DataFrame({
            "CompteNum": ["120000"], "CompteLib": ["Résultat"],
            "Solde_KE": [-0.8], "Solde_N1_KE": [-0.8],
        })
        ligne_51 = pd.DataFrame({
            "CompteNum": ["512100"], "CompteLib": ["Banque 2"],
            "Solde_KE": [0.8], "Solde_N1_KE": [0.8],
        })
        bal = pd.concat([bal, ligne_12, ligne_51], ignore_index=True)
        res = _resultats_dict(run_controles_financiers(bal, liasse_config))
        ok, _, _ = res[NOM_COHERENCE]
        assert ok is True

    def test_balance_sans_classes_6_7(self, liasse_config):
        """Exercice clôturé (classes 6/7 soldées, résultat en 12x) →
        WARNING signalé mais aucune exception."""
        bal = pd.DataFrame({
            "CompteNum": ["101000", "120000", "411000"],
            "CompteLib": ["Capital", "Résultat", "Clients"],
            "Solde_KE": [-100.0, -50.0, 150.0],
            "Solde_N1_KE": [-100.0, -50.0, 150.0],
        })
        res = _resultats_dict(run_controles_financiers(bal, liasse_config))
        ok, _, sev = res[NOM_COHERENCE]
        assert ok is False        # résultat déduit (0) ≠ poste bilan (50)
        assert sev == "WARNING"   # signalé, jamais bloquant
