"""
Tests unitaires pour src/engine/controls.py.
Utilise le vrai FEC GILAC pour les tests de bout en bout,
et des DataFrames synthétiques pour tester chaque contrôle en isolation.
"""

from pathlib import Path

import pandas as pd
import pytest

from src.parsers.fec_parser import parse
from src.engine.controls import run_all, COLONNES_OBLIGATOIRES

DATA_DIR = Path(__file__).parent.parent / "data"
FEC_PATH = DATA_DIR / "GILAC_2025_12_31_FEC.txt"
DATE_CLOTURE = "31/12/2025"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _resultats_dict(resultats):
    """Convertit la liste de résultats en dict {nom: (ok, detail, severity)}."""
    return {nom: (ok, detail, severity) for nom, ok, detail, severity in resultats}


def _fec_minimal() -> pd.DataFrame:
    """FEC équilibré minimal à 4 lignes pour les tests synthétiques."""
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
    df["Debit"] = df["Debit"].astype(float)
    df["Credit"] = df["Credit"].astype(float)
    df["Solde"] = df["Debit"] - df["Credit"]
    return df


# ---------------------------------------------------------------------------
# Fixture : résultats sur le vrai FEC GILAC
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def fec_gilac() -> pd.DataFrame:
    return parse(FEC_PATH)


@pytest.fixture(scope="module")
def resultats_gilac(fec_gilac) -> dict:
    return _resultats_dict(run_all(fec_gilac, DATE_CLOTURE))


# ---------------------------------------------------------------------------
# Tests sur le FEC GILAC (bout en bout)
# ---------------------------------------------------------------------------

class TestGilac:

    def test_run_all_retourne_9_controles(self, resultats_gilac):
        assert len(resultats_gilac) == 9

    def test_colonnes_ok(self, resultats_gilac):
        ok, detail, sev = resultats_gilac["Colonnes obligatoires"]
        assert ok is True
        assert sev == "BLOQUANT"

    def test_equilibre_ok(self, resultats_gilac):
        ok, detail, sev = resultats_gilac["Équilibre Débit/Crédit"]
        assert ok is True
        assert sev == "BLOQUANT"
        assert "écart=" in detail

    def test_lignes_zero_ok(self, resultats_gilac):
        ok, detail, _ = resultats_gilac["Lignes à zéro"]
        assert ok is True

    def test_coherence_dates_ok(self, resultats_gilac):
        ok, detail, _ = resultats_gilac["Cohérence des dates"]
        assert ok is True
        assert "01/01/2025" in detail

    def test_dimanches_ko(self, resultats_gilac):
        """GILAC contient des écritures un dimanche (comptabilisation automatique)."""
        ok, detail, sev = resultats_gilac["Écritures un dimanche"]
        assert ok is False
        assert sev == "WARNING"
        assert "dimanche" in detail.lower()

    def test_montants_ronds_ko(self, resultats_gilac):
        """GILAC contient des montants ronds ≥ 10 K€."""
        ok, detail, sev = resultats_gilac["Montants ronds"]
        assert ok is False
        assert sev == "INFO"
        assert "K€" in detail

    def test_doublons_ko(self, resultats_gilac):
        """GILAC contient des doublons potentiels (mêmes écritures dans plusieurs pièces)."""
        ok, detail, sev = resultats_gilac["Doublons potentiels"]
        assert ok is False
        assert sev == "WARNING"

    def test_benford_ok(self, resultats_gilac):
        """Distribution des premiers chiffres GILAC conforme à Benford (MAD < 0.015)."""
        ok, detail, sev = resultats_gilac["Benford (1er chiffre)"]
        assert ok is True
        assert sev == "INFO"
        assert "MAD=" in detail
        assert "χ²=" in detail
        # Extraire MAD et vérifier numériquement
        mad_str = detail.split("MAD=")[1].split(" ")[0]
        assert float(mad_str) < 0.015

    def test_ecritures_tardives_info(self, resultats_gilac):
        """ValidDate unique (export groupé) → contrôle non signifiant → ok=True INFO."""
        ok, detail, sev = resultats_gilac["Écritures tardives"]
        assert ok is True
        assert sev == "INFO"
        assert "export groupé" in detail


# ---------------------------------------------------------------------------
# Tests unitaires par contrôle (DataFrames synthétiques)
# ---------------------------------------------------------------------------

class TestEquilibre:

    def test_fec_equilibre(self):
        df = _fec_minimal()
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, _, _ = res["Équilibre Débit/Crédit"]
        assert ok is True

    def test_fec_desequilibre(self):
        df = _fec_minimal()
        df.loc[0, "Debit"] += 500.0  # crée un écart
        df["Solde"] = df["Debit"] - df["Credit"]
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Équilibre Débit/Crédit"]
        assert ok is False
        assert sev == "BLOQUANT"
        assert "écart=" in detail


class TestLignesZero:

    def test_avec_lignes_zero(self):
        df = _fec_minimal().copy()
        df.loc[len(df)] = {
            "JournalCode": "OD", "JournalLib": "OD", "EcritureNum": "Z1",
            "EcritureDate": "20250601", "CompteNum": "999999", "CompteLib": "Test",
            "CompAuxNum": "", "CompAuxLib": "", "PieceRef": "", "PieceDate": "20250601",
            "EcritureLib": "zéro", "Debit": 0.0, "Credit": 0.0,
            "EcritureLet": "", "DateLet": "", "ValidDate": "20260110",
            "Montantdevise": 0.0, "Idevise": "EUR", "Solde": 0.0,
        }
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, _ = res["Lignes à zéro"]
        assert ok is False
        assert "1" in detail


class TestCoherenceDates:

    def test_date_hors_exercice(self):
        df = _fec_minimal().copy()
        df.loc[0, "EcritureDate"] = "20240101"  # exercice N-1
        df["Solde"] = df["Debit"] - df["Credit"]
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Cohérence des dates"]
        assert ok is False
        assert sev == "WARNING"

    def test_date_invalide(self):
        df = _fec_minimal().copy()
        df.loc[0, "EcritureDate"] = "invalide"
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, _ = res["Cohérence des dates"]
        assert ok is False
        assert "non parsable" in detail


class TestEcrituresDimanche:

    def test_ecriture_dimanche(self):
        df = _fec_minimal().copy()
        df.loc[0, "EcritureDate"] = "20250105"  # dimanche 5 jan 2025
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Écritures un dimanche"]
        assert ok is False
        assert sev == "WARNING"
        assert "1" in detail

    def test_aucun_dimanche(self):
        df = _fec_minimal()
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, _, _ = res["Écritures un dimanche"]
        assert ok is True


class TestMontantsRonds:

    def test_montant_rond_seuil(self):
        df = _fec_minimal().copy()
        df.loc[0, "Credit"] = 15000.0  # entier >= 10k
        df.loc[1, "Debit"] = 15000.0
        df["Solde"] = df["Debit"] - df["Credit"]
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Montants ronds"]
        assert ok is False
        assert sev == "INFO"

    def test_montant_avec_centimes_non_signale(self):
        df = _fec_minimal().copy()
        df.loc[0, "Credit"] = 15000.50  # non entier
        df.loc[1, "Debit"] = 15000.50
        df["Solde"] = df["Debit"] - df["Credit"]
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, _, _ = res["Montants ronds"]
        assert ok is True


class TestDoublons:

    def test_doublon_detecte(self):
        df = _fec_minimal()
        df_double = pd.concat([df, df.iloc[[0, 1]]], ignore_index=True)
        df_double["Solde"] = df_double["Debit"] - df_double["Credit"]
        res = _resultats_dict(run_all(df_double, DATE_CLOTURE))
        ok, detail, sev = res["Doublons potentiels"]
        assert ok is False
        assert sev == "WARNING"

    def test_pas_de_doublon(self):
        df = _fec_minimal()
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, _, _ = res["Doublons potentiels"]
        assert ok is True


class TestBenford:

    def test_distribution_anormale(self):
        """Distribution concentrée sur le chiffre 9 → MAD très élevé."""
        df = _fec_minimal().copy()
        df["Debit"] = [9000.0, 0.0, 9000.0, 0.0]
        df["Credit"] = [0.0, 9000.0, 0.0, 9000.0]
        df["Solde"] = df["Debit"] - df["Credit"]
        res = _resultats_dict(run_all(df, "31/12/2025", seuil_benford_mad=0.015))
        ok, detail, sev = res["Benford (1er chiffre)"]
        assert ok is False
        assert sev == "INFO"

    def test_champs_mad_chi2_dans_detail(self, fec_gilac):
        res = _resultats_dict(run_all(fec_gilac, DATE_CLOTURE))
        _, detail, _ = res["Benford (1er chiffre)"]
        assert "MAD=" in detail
        assert "χ²=" in detail
        assert "n=" in detail


class TestEcrituresTardives:

    def test_ecriture_tardive_reelle(self):
        df = _fec_minimal().copy()
        df.loc[0, "ValidDate"] = "20260301"  # après 31/01/2026
        # Les autres restent avant le seuil pour éviter export groupé
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Écritures tardives"]
        assert ok is False
        assert sev == "WARNING"

    def test_validdate_vide(self):
        df = _fec_minimal().copy()
        df["ValidDate"] = ""
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Écritures tardives"]
        assert ok is True
        assert sev == "INFO"
        assert "absent" in detail

    def test_export_groupe_non_signifiant(self):
        """Toutes les ValidDate identiques et tardives = export groupé → ok=True."""
        df = _fec_minimal().copy()
        df["ValidDate"] = "20260311"  # date unique tardive
        res = _resultats_dict(run_all(df, DATE_CLOTURE))
        ok, detail, sev = res["Écritures tardives"]
        assert ok is True
        assert "export groupé" in detail
