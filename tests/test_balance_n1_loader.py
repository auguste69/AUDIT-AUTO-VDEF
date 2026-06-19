"""
Tests du chargeur unifié de balance N-1 (src/parsers/balance_n1_loader.py).

Fichiers synthétiques générés dans tmp_path, plus le FM GILAC réel
("data/FM GILAC.xlsx") pour les tests d'intégration légers.
"""

from pathlib import Path

import openpyxl
import pytest

from src.parsers.balance_n1_loader import _extraire_soldes_from_fm, load_balance_n1
from src.parsers.mapping_parser import detect_balance_sheet

# Chemin du FM GILAC réel (référence de format, pas de données en dur)
_FM_GILAC = Path(__file__).parent.parent / "data" / "FM GILAC.xlsx"

# Les 18 colonnes obligatoires du FEC
_COLONNES_FEC = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib", "Debit", "Credit",
    "EcritureLet", "DateLet", "ValidDate", "Montantdevise", "Idevise",
]


def _ecrire_fec_synthetique(chemin: Path) -> None:
    """Écrit un FEC N-1 synthétique minimal (TAB, UTF-8, montants à virgule)."""
    lignes = ["\t".join(_COLONNES_FEC)]
    ecritures = [
        ("AC", "601000", "ACHATS MP", "1000,00", "0,00"),
        ("AC", "401000", "FOURNISSEURS", "0,00", "1000,00"),
        ("VE", "411000", "CLIENTS", "2500,00", "0,00"),
        ("VE", "701000", "VENTES PF", "0,00", "2500,00"),
        ("VE", "411000", "CLIENTS", "500,00", "0,00"),
        ("VE", "701000", "VENTES PF", "0,00", "500,00"),
    ]
    for i, (journal, compte, lib, debit, credit) in enumerate(ecritures, start=1):
        lignes.append("\t".join([
            journal, f"Journal {journal}", str(i), "20241231",
            compte, lib, "", "",
            f"P{i}", "20241231", f"Ecriture {i}", debit, credit,
            "", "", "20241231", "", "",
        ]))
    chemin.write_text("\n".join(lignes), encoding="utf-8")


def _ecrire_balance_simple(chemin: Path) -> None:
    """Écrit une balance Excel simple synthétique (soldes en €)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["CompteNum", "CompteLib", "Solde"])
    ws.append(["101000", "CAPITAL SOCIAL", -50000.0])
    ws.append(["411000", "CLIENTS", 12500.0])
    ws.append(["701000", "VENTES", 37500.0])
    wb.save(chemin)
    wb.close()


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_load_fec_txt_synthetique(tmp_path):
    """.txt FEC N-1 synthétique → balance agrégée en K€, mapping None."""
    fec_path = tmp_path / "fec_n1.txt"
    _ecrire_fec_synthetique(fec_path)

    balance, mapping = load_balance_n1(fec_path)

    assert mapping is None
    assert len(balance) == 4
    # Agrégation : 411000 = 2500 + 500 = 3000 € → 3.0 K€
    assert balance["411000"]["solde_ke"] == pytest.approx(3.0)
    assert balance["701000"]["solde_ke"] == pytest.approx(-3.0)
    assert balance["601000"]["solde_ke"] == pytest.approx(1.0)
    assert balance["401000"]["solde_ke"] == pytest.approx(-1.0)
    assert balance["411000"]["libelle"] == "CLIENTS"


@pytest.mark.skipif(not _FM_GILAC.exists(), reason="FM GILAC absent de data/")
def test_load_fm_xlsx_reel():
    """.xlsx FM réel → balance + mapping non vides, soldes en K€."""
    balance, mapping = load_balance_n1(_FM_GILAC)

    assert mapping is not None and len(mapping) > 0
    assert len(balance) > 0
    # Soldes en K€ : floats, ordres de grandeur raisonnables (pas des €)
    for infos in balance.values():
        assert isinstance(infos["solde_ke"], float)
        assert abs(infos["solde_ke"]) < 1_000_000  # en K€, pas en €
    # Le mapping FM contient les clés attendues
    exemple = next(iter(mapping.values()))
    assert "cycle" in exemple
    assert "compta" in exemple


def test_load_balance_xlsx_simple(tmp_path):
    """.xlsx balance simple synthétique → balance en K€, mapping None."""
    xlsx_path = tmp_path / "balance_n1.xlsx"
    _ecrire_balance_simple(xlsx_path)

    balance, mapping = load_balance_n1(xlsx_path)

    assert mapping is None
    assert len(balance) == 3
    # Conversion € → K€
    assert balance["101000"]["solde_ke"] == pytest.approx(-50.0)
    assert balance["411000"]["solde_ke"] == pytest.approx(12.5)
    assert balance["701000"]["solde_ke"] == pytest.approx(37.5)
    assert balance["101000"]["libelle"] == "CAPITAL SOCIAL"


def test_extension_inconnue_csv(tmp_path):
    """Extension non supportée (.csv) → ValueError explicite."""
    csv_path = tmp_path / "balance_n1.csv"
    csv_path.write_text("CompteNum;CompteLib;Solde\n101000;CAPITAL;-50000\n",
                        encoding="utf-8")

    with pytest.raises(ValueError, match="Format N-1 non reconnu"):
        load_balance_n1(csv_path)


@pytest.mark.skipif(not _FM_GILAC.exists(), reason="FM GILAC absent de data/")
def test_coherence_load_vs_extraction_directe():
    """load_balance_n1 (FM) == _extraire_soldes_from_fm appelé directement."""
    balance_via_loader, _ = load_balance_n1(_FM_GILAC)

    nom_feuille, mode = detect_balance_sheet(_FM_GILAC)
    assert mode == "fm"
    balance_directe = _extraire_soldes_from_fm(_FM_GILAC, nom_feuille)

    assert balance_via_loader == balance_directe
