"""
Test de régression cellulaire du FM sur données synthétiques.

Régénère le FM complet dans un répertoire temporaire à partir du FEC et de
la balance N-1 synthétiques (tests/synthetic_data.py, déterministes) et le
compare, cellule par cellule (valeurs + number_format), à la fixture de
référence tests/fixtures/FM_SYNTHETIQUE_REF.xlsx.

Toute modification du moteur ou des writers qui change une valeur, un
format ou la structure d'un onglet fait échouer ce test : si le changement
est voulu, régénérer la fixture avec scripts/generer_fixture_fm.py et
vérifier le diff onglet par onglet.

Règles de comparaison :
- Comparaison PAR NOM D'ONGLET (jamais par index) — uniquement les onglets
  présents dans la fixture ; les onglets supplémentaires du FM régénéré
  sont ignorés.
- Flottants : tolérance abs(a - b) <= 0.01.
- Pas de comparaison octet par octet.
"""

from pathlib import Path
from typing import Iterator, Tuple

import openpyxl
import pandas as pd
import pytest

from src.parsers.fec_parser import parse
from src.parsers.balance_n1_loader import load_balance_n1
from src.parsers.mapping_parser import from_pcg_config
from src.engine.balance_builder import build
from src.engine.cycle_mapper import map_cycles
from src.writers.fm_writer import write
from tests.synthetic_data import generer_balance_n1_xlsx, generer_fec_synthetique

CONFIG_DIR = Path(__file__).parent.parent / "src" / "config"
FIXTURE_FM = Path(__file__).parent / "fixtures" / "FM_SYNTHETIQUE_REF.xlsx"
PCG_PATH   = CONFIG_DIR / "mapping_pcg.yaml"

_TOLERANCE = 0.01


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def balance_mappee(tmp_path_factory) -> pd.DataFrame:
    """Construit la balance mappée synthétique (même pipeline que main.py)."""
    tmp = tmp_path_factory.mktemp("fm_regression_data")
    fec = generer_fec_synthetique(tmp / "DEMO_2025_FEC.txt")
    n1_xlsx = generer_balance_n1_xlsx(tmp / "balance_n1.xlsx")

    df = parse(fec)
    balance_n1, _ = load_balance_n1(n1_xlsx)
    pcg = from_pcg_config(PCG_PATH)

    bal = build(df, balance_n1)
    return map_cycles(bal, None, pcg)


@pytest.fixture(scope="module")
def fm_regen_path(balance_mappee, tmp_path_factory) -> Path:
    """Régénère le FM dans un tmp_path pytest avec la config PCG par défaut."""
    out = tmp_path_factory.mktemp("fm_regression")
    pcg = from_pcg_config(PCG_PATH)
    return write(balance_mappee, "DEMO", "31/12/2025", out, pcg_config=pcg)


@pytest.fixture(scope="module")
def wb_regen(fm_regen_path):
    return openpyxl.load_workbook(fm_regen_path)


@pytest.fixture(scope="module")
def wb_fixture():
    return openpyxl.load_workbook(FIXTURE_FM)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _valeurs_egales(a, b) -> bool:
    """Égalité de valeurs de cellules — tolérance 0.01 sur les flottants."""
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) \
            and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(float(a) - float(b)) <= _TOLERANCE
    return a == b


def _iter_cellules(ws_fix, ws_new) -> Iterator[Tuple[int, int]]:
    """Itère sur la plage couvrant les deux feuilles (union des dimensions)."""
    max_row = max(ws_fix.max_row, ws_new.max_row)
    max_col = max(ws_fix.max_column, ws_new.max_column)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            yield r, c


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_fixture_presente():
    """La fixture de référence existe."""
    assert FIXTURE_FM.exists(), f"Fixture absente : {FIXTURE_FM}"


def test_onglets_fixture_presents(wb_fixture, wb_regen):
    """Chaque onglet de la fixture existe dans le FM régénéré (par NOM)."""
    manquants = [n for n in wb_fixture.sheetnames if n not in wb_regen.sheetnames]
    assert not manquants, (
        f"Onglets de la fixture absents du FM régénéré : {manquants}. "
        f"Onglets régénérés : {wb_regen.sheetnames}"
    )


def test_regression_cellulaire(wb_fixture, wb_regen):
    """Toutes les cellules (valeur + number_format) sont identiques.

    Comparaison par nom d'onglet, uniquement les onglets de la fixture ;
    les onglets supplémentaires du FM régénéré sont ignorés.
    """
    ecarts = []
    for nom in wb_fixture.sheetnames:
        if nom not in wb_regen.sheetnames:
            continue  # déjà couvert par test_onglets_fixture_presents
        ws_fix = wb_fixture[nom]
        ws_new = wb_regen[nom]
        for r, c in _iter_cellules(ws_fix, ws_new):
            cell_fix = ws_fix.cell(row=r, column=c)
            cell_new = ws_new.cell(row=r, column=c)
            if not _valeurs_egales(cell_fix.value, cell_new.value):
                ecarts.append(
                    f"{nom}!{cell_fix.coordinate} : valeur "
                    f"{cell_fix.value!r} ≠ {cell_new.value!r}"
                )
            elif cell_fix.number_format != cell_new.number_format:
                ecarts.append(
                    f"{nom}!{cell_fix.coordinate} : format "
                    f"{cell_fix.number_format!r} ≠ {cell_new.number_format!r}"
                )
            if len(ecarts) >= 50:  # limiter le volume du rapport d'échec
                break
        if len(ecarts) >= 50:
            break

    assert not ecarts, (
        f"{len(ecarts)} écart(s) cellulaire(s) (50 max affichés) :\n"
        + "\n".join(ecarts)
    )
