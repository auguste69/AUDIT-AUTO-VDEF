"""
Test end-to-end du rapprochement de comptes N/N-1 (P6) dans run_pipeline.

Scénario synthétique : le compte bancaire "512003" (N-1) devient "512300"
en N (même libellé). Avec --rapprochements-auto le solde N-1 suit le compte
renuméroté ; sans validation (mode non interactif), AUCUNE fusion n'est
appliquée — la confirmation explicite est obligatoire.
"""

from pathlib import Path

import pytest
from openpyxl import Workbook

from main import run_pipeline

# FEC synthétique équilibré : capital 100 000 au crédit, banque au débit.
# 18 colonnes obligatoires, séparateur TAB, dates dans l'exercice 2025.
_COLONNES = (
    "JournalCode\tJournalLib\tEcritureNum\tEcritureDate\tCompteNum\t"
    "CompteLib\tCompAuxNum\tCompAuxLib\tPieceRef\tPieceDate\tEcritureLib\t"
    "Debit\tCredit\tEcritureLet\tDateLet\tValidDate\tMontantdevise\tIdevise"
)
_LIGNES = [
    "BQ\tBanque\tB1\t20250115\t512300\tBANQUE CREDIT AGRICOLE\t\t\tP1\t"
    "20250115\tApport\t100000,00\t0,00\t\t\t20250131\t0,00\tEUR",
    "OD\tDivers\tO1\t20250115\t101000\tCAPITAL SOCIAL\t\t\tP1\t"
    "20250115\tApport\t0,00\t100000,00\t\t\t20250131\t0,00\tEUR",
]


@pytest.fixture()
def fec_path(tmp_path) -> Path:
    chemin = tmp_path / "TEST_2025_FEC.txt"
    chemin.write_text("\n".join([_COLONNES] + _LIGNES), encoding="utf-8")
    return chemin


@pytest.fixture()
def balance_n1_path(tmp_path) -> Path:
    """Balance N-1 Excel simple : 512003 (ancien numéro) + 101000."""
    chemin = tmp_path / "balance_n1.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["CompteNum", "CompteLib", "Solde"])
    ws.append(["101000", "CAPITAL SOCIAL", -100000.0])
    ws.append(["512003", "BANQUE CREDIT AGRICOLE", 100000.0])
    wb.save(chemin)
    return chemin


def _run(fec_path, balance_n1_path, tmp_path, **kwargs) -> dict:
    return run_pipeline(
        fec_path=str(fec_path),
        client="TESTCLIENT",
        date_cloture="31/12/2025",
        n1_fm=str(balance_n1_path),
        templates_dir=None,
        output_dir=str(tmp_path / "out"),
        **kwargs,
    )


def test_rapprochement_auto_fusionne(fec_path, balance_n1_path, tmp_path):
    """--rapprochements-auto : la proposition 512003 → 512300 est appliquée,
    le solde N-1 suit le compte renuméroté."""
    resultats = _run(fec_path, balance_n1_path, tmp_path,
                     rapprochements_auto=True)

    proposes = resultats["rapprochements_proposes"]
    appliques = resultats["rapprochements_appliques"]
    assert len(proposes) == 1
    assert (proposes[0].compte_n1, proposes[0].compte_n) == ("512003", "512300")
    assert appliques == proposes

    bm = resultats["balance_mappee"]
    ligne = bm[bm["CompteNum"] == "512300"].iloc[0]
    assert ligne["Solde_N1_KE"] == pytest.approx(100.0)
    assert float(ligne["Var_KE"]) == pytest.approx(0.0)
    assert resultats["fm_path"].exists()


def test_sans_validation_aucune_fusion(fec_path, balance_n1_path, tmp_path):
    """Mode par défaut non interactif (pas de TTY, pas d'auto, pas de liste
    validée) : la proposition existe mais AUCUNE fusion n'est appliquée."""
    resultats = _run(fec_path, balance_n1_path, tmp_path)

    assert len(resultats["rapprochements_proposes"]) == 1
    assert resultats["rapprochements_appliques"] == []

    bm = resultats["balance_mappee"]
    ligne = bm[bm["CompteNum"] == "512300"].iloc[0]
    assert ligne["Solde_N1_KE"] == pytest.approx(0.0)  # pas de fusion


def test_liste_validee_streamlit(fec_path, balance_n1_path, tmp_path):
    """rapprochements_valides (paires validées dans l'UI) : seules les
    paires fournies sont fusionnées."""
    resultats = _run(fec_path, balance_n1_path, tmp_path,
                     rapprochements_valides=[("512003", "512300")])

    assert len(resultats["rapprochements_appliques"]) == 1
    bm = resultats["balance_mappee"]
    ligne = bm[bm["CompteNum"] == "512300"].iloc[0]
    assert ligne["Solde_N1_KE"] == pytest.approx(100.0)


def test_liste_validee_vide_aucune_fusion(fec_path, balance_n1_path, tmp_path):
    """rapprochements_valides=[] (l'utilisateur a tout ignoré) : aucune
    fusion malgré la proposition."""
    resultats = _run(fec_path, balance_n1_path, tmp_path,
                     rapprochements_valides=[])

    assert len(resultats["rapprochements_proposes"]) == 1
    assert resultats["rapprochements_appliques"] == []
