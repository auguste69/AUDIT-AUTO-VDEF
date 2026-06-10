"""
Tests unitaires du contrôle AC-1 (équilibre du bilan) — src/engine/controls.py.

Utilise des balances synthétiques en mémoire (pas le FEC GILAC complet) et
la config liasse_fiscale réelle du mapping_pcg.yaml.
"""

from pathlib import Path

import pandas as pd
import pytest

from src.engine.controls import run_all, run_controles_financiers
from src.engine.liasse_fiscale_loader import load_liasse_fiscale
from src.parsers.mapping_parser import from_pcg_config

PCG_PATH = Path(__file__).parent.parent / "src" / "config" / "mapping_pcg.yaml"
DATE_CLOTURE = "31/12/2025"
NOM_AC1 = "Équilibre du bilan (AC-1)"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def liasse_config() -> dict:
    """Section liasse_fiscale réelle, chargée depuis mapping_pcg.yaml."""
    return load_liasse_fiscale(from_pcg_config(PCG_PATH))


def _balance_equilibree() -> pd.DataFrame:
    """
    Balance synthétique équilibrée (somme des soldes = 0, bilan équilibré).

    Capital 100 au passif, clients 120 + banque 30 à l'actif,
    résultat en cours = 50 (ventes 200 − achats 150).
    Total Actif = 150 = Total Passif (capital 100 + résultat 50).
    """
    return pd.DataFrame({
        "CompteNum": ["101000", "411000", "512000", "606000", "706000"],
        "CompteLib": ["Capital", "Clients", "Banque", "Achats", "Ventes"],
        "Solde_KE": [-100.0, 120.0, 30.0, 150.0, -200.0],
        "Solde_N1_KE": [-100.0, 120.0, 30.0, 150.0, -200.0],
    })


def _resultats_dict(resultats):
    """Convertit la liste de résultats en dict {nom: (ok, detail, severity)}."""
    return {nom: (ok, detail, severity) for nom, ok, detail, severity in resultats}


def _fec_minimal() -> pd.DataFrame:
    """FEC équilibré minimal à 4 lignes pour les tests de run_all."""
    data = {
        "JournalCode": ["ACH", "ACH", "VTE", "VTE"],
        "JournalLib": ["Achats", "Achats", "Ventes", "Ventes"],
        "EcritureNum": ["A1", "A1", "V1", "V1"],
        "EcritureDate": ["20250115", "20250115", "20250320", "20250320"],
        "CompteNum": ["401000", "606000", "411000", "706000"],
        "CompteLib": ["Fourn.", "Achats", "Client", "Ventes"],
        "CompAuxNum": ["", "", "", ""],
        "CompAuxLib": ["", "", "", ""],
        "PieceRef": ["F001", "F001", "F002", "F002"],
        "PieceDate": ["20250115", "20250115", "20250320", "20250320"],
        "EcritureLib": ["Facture fournisseur", "Facture fournisseur",
                        "Facture client", "Facture client"],
        "Debit": [0.0, 1000.0, 2000.0, 0.0],
        "Credit": [1000.0, 0.0, 0.0, 2000.0],
        "EcritureLet": ["", "", "", ""],
        "DateLet": ["", "", "", ""],
        "ValidDate": ["20260110", "20260110", "20260110", "20260110"],
        "Montantdevise": [0.0, 0.0, 0.0, 0.0],
        "Idevise": ["EUR", "EUR", "EUR", "EUR"],
    }
    df = pd.DataFrame(data)
    df["Solde"] = df["Debit"] - df["Credit"]
    return df


# ---------------------------------------------------------------------------
# AC-1 : équilibre N
# ---------------------------------------------------------------------------

class TestBilanEquilibre:

    def test_balance_equilibree_ok(self, liasse_config):
        """Balance équilibrée → AC-1 ok=True."""
        res = _resultats_dict(
            run_controles_financiers(_balance_equilibree(), liasse_config)
        )
        ok, detail, sev = res[NOM_AC1]
        assert ok is True
        assert sev == "BLOQUANT"
        assert "écart=" in detail

    def test_ecart_n_superieur_seuil_bloquant(self, liasse_config):
        """Écart N > seuil → ok=False, severity BLOQUANT."""
        bal = _balance_equilibree()
        # Gonfler les clients de 10 K€ sans contrepartie → bilan N déséquilibré
        bal.loc[bal["CompteNum"] == "411000", "Solde_KE"] = 130.0
        res = _resultats_dict(run_controles_financiers(bal, liasse_config))
        ok, detail, sev = res[NOM_AC1]
        assert ok is False
        assert sev == "BLOQUANT"
        assert "déséquilibré" in detail

    def test_ecart_n_inferieur_seuil_ok(self, liasse_config):
        """Écart N ≤ seuil paramétrable → ok=True."""
        bal = _balance_equilibree()
        bal.loc[bal["CompteNum"] == "411000", "Solde_KE"] = 120.5  # écart 0.5 K€
        res = _resultats_dict(
            run_controles_financiers(bal, liasse_config,
                                     seuil_equilibre_bilan_ke=1.0)
        )
        ok, _, _ = res[NOM_AC1]
        assert ok is True

    def test_ecart_n1_jamais_bloquant(self, liasse_config):
        """Écart N-1 > seuil mais N équilibré → ok=True, WARNING dans le détail."""
        bal = _balance_equilibree()
        # Déséquilibrer uniquement N-1 (cas FM historique non corrigeable)
        bal.loc[bal["CompteNum"] == "411000", "Solde_N1_KE"] = 130.0
        res = _resultats_dict(run_controles_financiers(bal, liasse_config))
        ok, detail, sev = res[NOM_AC1]
        assert ok is True              # jamais bloquant sur N-1
        assert "WARNING N-1" in detail
        assert "non bloquant" in detail

    def test_bilan_non_bloquant_devient_warning(self, liasse_config):
        """Avec bilan_non_bloquant=True, AC-1 échoue en WARNING (pas BLOQUANT)."""
        bal = _balance_equilibree()
        bal.loc[bal["CompteNum"] == "411000", "Solde_KE"] = 130.0
        res = _resultats_dict(
            run_controles_financiers(bal, liasse_config, bilan_non_bloquant=True)
        )
        ok, _, sev = res[NOM_AC1]
        assert ok is False
        assert sev == "WARNING"


# ---------------------------------------------------------------------------
# run_all : rétrocompatibilité 9 contrôles / extension 11 contrôles
# ---------------------------------------------------------------------------

class TestRunAllRetrocompatibilite:

    def test_run_all_sans_balance_9_controles(self):
        """run_all sans balance_mappee → exactement 9 contrôles (inchangé)."""
        resultats = run_all(_fec_minimal(), DATE_CLOTURE)
        assert len(resultats) == 9

    def test_run_all_avec_balance_11_controles(self, liasse_config):
        """run_all avec balance_mappee + liasse_config → 11 contrôles."""
        resultats = run_all(
            _fec_minimal(), DATE_CLOTURE,
            balance_mappee=_balance_equilibree(),
            liasse_config=liasse_config,
        )
        assert len(resultats) == 11
        noms = [r[0] for r in resultats]
        assert NOM_AC1 in noms
        assert "Cohérence du résultat" in noms

    def test_run_all_balance_sans_liasse_9_controles(self):
        """run_all avec balance mais sans liasse_config → 9 contrôles."""
        resultats = run_all(
            _fec_minimal(), DATE_CLOTURE,
            balance_mappee=_balance_equilibree(),
            liasse_config=None,
        )
        assert len(resultats) == 9
