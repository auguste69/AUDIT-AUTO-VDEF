"""
Écriture de l'onglet Passif détaillé (cerfa 2051) du FM.

Écriture pure : toutes les valeurs proviennent d'un PassifDetaille calculé
par src.engine.financial_engine.calculer_passif_detaille. Le rendu est
strictement identique à celui de l'Actif détaillé (sections, sous-totaux,
total) — il est délégué à ecrire_etat_detaille_tab du module
actif_detaille (même mise en page cerfa 2050/2051).
"""

import logging

from openpyxl.worksheet.worksheet import Worksheet

from src.models.financial_statements import PassifDetaille
from src.writers.fm.actif_detaille import ecrire_etat_detaille_tab

logger = logging.getLogger(__name__)


def ecrire_passif_detaille_tab(ws: Worksheet, passif: PassifDetaille,
                               date_n: str, date_n1: str,
                               client: str) -> None:
    """Écrit l'onglet Passif détaillé (cerfa 2051).

    Paramètres
    ----------
    ws : Worksheet
        Feuille openpyxl cible (déjà créée et nommée "Passif détaillé").
    passif : PassifDetaille
        Résultat de financial_engine.calculer_passif_detaille (postes
        créditeurs déjà présentés en positif).
    date_n : str
        Date de clôture N au format 'JJ/MM/AAAA' (string, jamais datetime).
    date_n1 : str
        Date de clôture N-1 au format 'JJ/MM/AAAA'.
    client : str
        Nom du client.
    """
    ecrire_etat_detaille_tab(ws, passif, "Passif détaillé",
                             date_n, date_n1, client)
