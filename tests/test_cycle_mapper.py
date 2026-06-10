"""
Tests pour src/engine/cycle_mapper.py et src/parsers/mapping_parser.py.
Utilise la vraie balance GILAC + FM + mapping_pcg.yaml.
"""

from pathlib import Path

import pytest
import pandas as pd

from src.parsers.fec_parser import parse
from src.parsers.mapping_parser import from_fm, from_pcg_config
from src.engine.balance_builder import build
from src.engine.cycle_mapper import map_cycles, _resoudre_cycle, _resoudre_compta

DATA_DIR   = Path(__file__).parent.parent / "data"
CONFIG_DIR = Path(__file__).parent.parent / "src" / "config"

FEC_PATH = DATA_DIR / "GILAC_2025_12_31_FEC.txt"
FM_PATH  = DATA_DIR / "FM GILAC.xlsx"
PCG_PATH = CONFIG_DIR / "mapping_pcg.yaml"

# 29 nouveaux comptes 2025 absents du FM (à mapper via PCG)
NOUVEAUX_COMPTES = [
    "106810", "231500", "401500", "401600", "409100",
    "411500", "411600", "445661", "455030", "486001",
    "512113", "512153", "512155", "604000", "607030",
    "611020", "616170", "621100", "622601", "623202",
    "623203", "637810", "641430", "649000", "671100",
    "758700", "762400", "767010", "768010",
]

TOUS_CYCLES = {"A", "V", "P", "E", "F", "S", "T",
               "I Corp", "I Incorp", "I Fi", "C Propres", "C PRC", "X"}


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def fec():
    return parse(FEC_PATH)


@pytest.fixture(scope="module")
def balance(fec):
    return build(fec)


@pytest.fixture(scope="module")
def mapping_fm():
    return from_fm(FM_PATH)


@pytest.fixture(scope="module")
def pcg():
    return from_pcg_config(PCG_PATH)


@pytest.fixture(scope="module")
def balance_mappee(balance, mapping_fm, pcg):
    return map_cycles(balance, mapping_fm, pcg)


# ---------------------------------------------------------------------------
# Tests : mapping_parser.from_fm
# ---------------------------------------------------------------------------

class TestFromFm:

    def test_nb_comptes_fm(self, mapping_fm):
        assert len(mapping_fm) == 288

    def test_structure_entree(self, mapping_fm):
        info = mapping_fm["101300"]
        assert set(info.keys()) == {
            "cycle", "ref",
            "etatfi_n", "etatfi_n1", "compta_n", "compta_n1",
            "etatfi", "compta",  # aliases historiques (= valeurs N)
        }

    def test_valeurs_compte_connu(self, mapping_fm):
        info = mapping_fm["101300"]
        assert info["cycle"]  == "C Propres"
        assert info["compta"] == "Passif"
        assert info["ref"]    == "C Propres0"

    def test_nouveaux_comptes_absents_du_fm(self, mapping_fm):
        for num in NOUVEAUX_COMPTES:
            assert num not in mapping_fm, f"{num} ne devrait pas être dans le FM"

    def test_tous_les_cycles_presents(self, mapping_fm):
        cycles_fm = {v["cycle"] for v in mapping_fm.values()}
        assert cycles_fm == TOUS_CYCLES

    def test_fichier_inexistant(self):
        with pytest.raises(FileNotFoundError):
            from_fm("/tmp/inexistant.xlsx")


# ---------------------------------------------------------------------------
# Tests : mapping_parser.from_pcg_config
# ---------------------------------------------------------------------------

class TestFromPcgConfig:

    def test_cles_presentes(self, pcg):
        assert "prefixes" in pcg
        assert "surcharges" in pcg
        assert "passif_prefixes" in pcg
        assert "ordre_cycles" in pcg
        assert "noms_cycles" in pcg

    def test_nb_prefixes(self, pcg):
        # Le YAML couvre les 7 classes avec ~277 préfixes
        assert len(pcg["prefixes"]) >= 250

    def test_classe_1_present(self, pcg):
        assert pcg["prefixes"].get("10") == "C Propres"
        assert pcg["prefixes"].get("15") == "C PRC"
        assert pcg["prefixes"].get("16") == "F"

    def test_surcharges_classe_6(self, pcg):
        # 603 → S (variation stocks), pas A comme le défaut "60"
        assert pcg["prefixes"].get("603") == "S"
        assert pcg["prefixes"].get("60")  == "A"

    def test_passif_prefixes(self, pcg):
        assert "401" in pcg["passif_prefixes"]
        assert "421" in pcg["passif_prefixes"]
        assert "444" in pcg["passif_prefixes"]

    def test_ordre_cycles(self, pcg):
        assert pcg["ordre_cycles"][0] == "C Propres"
        assert len(pcg["ordre_cycles"]) == 13

    def test_fichier_inexistant(self):
        with pytest.raises(FileNotFoundError):
            from_pcg_config("/tmp/inexistant.yaml")


# ---------------------------------------------------------------------------
# Tests : résolution interne (_resoudre_cycle, _resoudre_compta)
# ---------------------------------------------------------------------------

class TestResolutionCycle:

    def test_prefixe_3_chiffres(self, pcg):
        # 486 → A (CCA), surcharge sur "48" → T
        assert _resoudre_cycle("486001", pcg["prefixes"], pcg["surcharges"]) == "A"

    def test_prefixe_2_chiffres(self, pcg):
        # 48 → T (régularisation générique)
        assert _resoudre_cycle("488000", pcg["prefixes"], pcg["surcharges"]) == "T"

    def test_prefixe_1_chiffre_fallback(self, pcg):
        # Compte fictif 6XXXXX → classe 6 → "60" défaut A
        assert _resoudre_cycle("600000", pcg["prefixes"], pcg["surcharges"]) == "A"

    def test_surcharge_cabinet_prioritaire(self):
        prefixes  = {"60": "A", "601": "A"}
        surcharge = {"601": "S"}
        assert _resoudre_cycle("601000", prefixes, surcharge) == "S"

    def test_variation_stocks_603(self, pcg):
        # 603 → S même si "60" → A
        assert _resoudre_cycle("603100", pcg["prefixes"], pcg["surcharges"]) == "S"


class TestResolutionCompta:

    def test_classe_1_passif(self, pcg):
        assert _resoudre_compta("101000", pcg["passif_prefixes"]) == "Passif"

    def test_classe_2_actif(self, pcg):
        assert _resoudre_compta("213000", pcg["passif_prefixes"]) == "Actif"

    def test_classe_3_actif(self, pcg):
        assert _resoudre_compta("371000", pcg["passif_prefixes"]) == "Actif"

    def test_classe_5_actif(self, pcg):
        assert _resoudre_compta("512000", pcg["passif_prefixes"]) == "Actif"

    def test_classe_6_charges(self, pcg):
        assert _resoudre_compta("641000", pcg["passif_prefixes"]) == "Charges"

    def test_classe_7_produits(self, pcg):
        assert _resoudre_compta("701000", pcg["passif_prefixes"]) == "Produits"

    def test_classe_4_fournisseur_passif(self, pcg):
        # 401 → Passif
        assert _resoudre_compta("401000", pcg["passif_prefixes"]) == "Passif"

    def test_classe_4_client_actif(self, pcg):
        # 411 → Actif (pas dans passif_prefixes)
        assert _resoudre_compta("411000", pcg["passif_prefixes"]) == "Actif"

    def test_classe_4_cca_actif(self, pcg):
        # 486 → Actif
        assert _resoudre_compta("486001", pcg["passif_prefixes"]) == "Actif"


# ---------------------------------------------------------------------------
# Tests : map_cycles — bout en bout sur GILAC
# ---------------------------------------------------------------------------

class TestMapCycles:

    def test_shape(self, balance, balance_mappee):
        assert len(balance_mappee) == len(balance)
        # cycle, etatfi_n, etatfi_n1, compta_n, compta_n1, ref
        # + aliases etatfi/compta = 8 colonnes ajoutées
        assert len(balance_mappee.columns) == len(balance.columns) + 8

    def test_colonnes_ajoutees(self, balance_mappee):
        for col in ("cycle", "compta", "etatfi", "ref"):
            assert col in balance_mappee.columns

    def test_aucun_compte_sans_cycle(self, balance_mappee):
        sans_cycle = balance_mappee[balance_mappee["cycle"] == ""]
        assert len(sans_cycle) == 0, (
            f"Comptes sans cycle : {sans_cycle['CompteNum'].tolist()}"
        )

    def test_cycles_valides(self, balance_mappee):
        cycles_obtenus = set(balance_mappee["cycle"].unique())
        invalides = cycles_obtenus - TOUS_CYCLES
        assert not invalides, f"Cycles inconnus : {invalides}"

    def test_13_cycles_representes(self, balance_mappee):
        assert len(balance_mappee["cycle"].unique()) == 13

    def test_comptes_fm_prioritaires(self, balance_mappee, mapping_fm):
        """Les comptes présents dans le FM gardent exactement son cycle."""
        for num, info in list(mapping_fm.items())[:20]:
            ligne = balance_mappee[balance_mappee["CompteNum"] == num]
            if ligne.empty:
                continue  # compte FM absent de la balance 2025
            assert ligne.iloc[0]["cycle"]  == info["cycle"],  f"Cycle FM incorrect pour {num}"
            assert ligne.iloc[0]["compta"] == info["compta"], f"Compta FM incorrect pour {num}"

    def test_nouveaux_comptes_via_pcg(self, balance_mappee):
        """Les 29 nouveaux comptes (absents du FM) ont un cycle du PCG."""
        presents = balance_mappee[balance_mappee["CompteNum"].isin(NOUVEAUX_COMPTES)]
        assert len(presents) > 0, "Aucun nouveau compte trouvé dans la balance"
        assert (presents["cycle"] != "").all(), "Nouveau compte sans cycle PCG"
        # etatfi doit être vide (pas connu sans FM)
        assert (presents["etatfi"] == "").all(), "Nouveau compte ne devrait pas avoir d'etatfi"

    def test_ref_format(self, balance_mappee):
        """ref = cycle + '0' pour les comptes PCG, valeur FM pour les comptes FM."""
        pcg_lignes = balance_mappee[balance_mappee["etatfi"] == ""]
        for _, row in pcg_lignes.iterrows():
            assert row["ref"] == f"{row['cycle']}0", (
                f"ref incorrect pour {row['CompteNum']} : attendu '{row['cycle']}0', "
                f"obtenu '{row['ref']}'"
            )

    def test_compte_486_cca_dans_A(self, balance_mappee):
        """486001 (CCA) doit être en cycle A, compta Actif."""
        ligne = balance_mappee[balance_mappee["CompteNum"] == "486001"]
        assert not ligne.empty
        assert ligne.iloc[0]["cycle"]  == "A"
        assert ligne.iloc[0]["compta"] == "Actif"

    def test_compte_401_fournisseur_passif(self, balance_mappee):
        ligne = balance_mappee[balance_mappee["CompteNum"] == "401000"]
        assert not ligne.empty
        assert ligne.iloc[0]["cycle"]  == "A"
        assert ligne.iloc[0]["compta"] == "Passif"

    def test_sans_fm_tout_pcg(self, balance, pcg):
        """Sans FM, tous les comptes sont mappés via le PCG."""
        result = map_cycles(balance, None, pcg)
        assert (result["cycle"] != "").all()
        assert (result["etatfi"] == "").all()

    def test_immuabilite_balance_source(self, balance, mapping_fm, pcg):
        """map_cycles ne modifie pas le DataFrame source."""
        cols_avant = list(balance.columns)
        map_cycles(balance, mapping_fm, pcg)
        assert list(balance.columns) == cols_avant
