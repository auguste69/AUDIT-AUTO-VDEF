"""Tests de structure de la section liasse_fiscale du mapping_pcg.yaml.

Vérifie UNIQUEMENT la structure (clés et types), pas le contenu :
les listes de préfixes sont remplies par les prompts 4, 9 et 10.
"""

from pathlib import Path

import pytest
import yaml

_YAML_PATH = Path(__file__).resolve().parent.parent / "src" / "config" / "mapping_pcg.yaml"


@pytest.fixture(scope="module")
def config() -> dict:
    """Charge le mapping_pcg.yaml complet."""
    with _YAML_PATH.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def test_section_liasse_fiscale_existe(config):
    """La section liasse_fiscale est présente à la racine du YAML."""
    assert "liasse_fiscale" in config


def test_sous_sections_top_level(config):
    """Les 5 sous-sections attendues sont présentes."""
    attendu = {"bilan", "treso", "aace", "ebit", "pl_detaille"}
    assert set(config["liasse_fiscale"].keys()) == attendu


def test_bilan_structure(config):
    """La sous-section bilan contient actif, passif et comptes_bascule."""
    bilan = config["liasse_fiscale"]["bilan"]
    assert "actif" in bilan
    assert "passif" in bilan
    assert "comptes_bascule" in bilan


def _verifier_feuilles_sont_des_listes(noeud, chemin: str) -> None:
    """Parcours récursif : toute valeur feuille (non dict) doit être une list."""
    if isinstance(noeud, dict):
        for cle, valeur in noeud.items():
            _verifier_feuilles_sont_des_listes(valeur, f"{chemin}.{cle}")
    else:
        assert isinstance(noeud, list), (
            f"Valeur feuille non-list dans liasse_fiscale : {chemin} "
            f"({type(noeud).__name__})"
        )


def test_listes_feuilles_sont_des_listes(config):
    """Toute valeur feuille de liasse_fiscale est une liste (éventuellement vide)."""
    _verifier_feuilles_sont_des_listes(config["liasse_fiscale"], "liasse_fiscale")
