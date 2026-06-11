"""
Sous-package des writers d'onglets de feuilles maîtresses.

Chaque module écrit UN onglet du FM à partir d'un objet calculé par
src/engine/financial_engine.py — écriture pure, aucun calcul métier.
"""

from src.writers.fm.actif_detaille import ecrire_actif_detaille_tab
from src.writers.fm.ebit import ecrire_ebit_tab
from src.writers.fm.passif_detaille import ecrire_passif_detaille_tab
from src.writers.fm.pl_detaille import ecrire_pl_detaille_tab

__all__ = [
    "ecrire_actif_detaille_tab",
    "ecrire_ebit_tab",
    "ecrire_passif_detaille_tab",
    "ecrire_pl_detaille_tab",
]
