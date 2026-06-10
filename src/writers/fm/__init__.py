"""
Sous-package des writers d'onglets de feuilles maîtresses.

Chaque module écrit UN onglet du FM à partir d'un objet calculé par
src/engine/financial_engine.py — écriture pure, aucun calcul métier.
"""

from src.writers.fm.ebit import ecrire_ebit_tab

__all__ = ["ecrire_ebit_tab"]
