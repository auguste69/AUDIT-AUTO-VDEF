"""
Écriture de l'onglet Actif détaillé (cerfa 2050) du FM.

Écriture pure : toutes les valeurs proviennent d'un ActifDetaille calculé
par src.engine.financial_engine.calculer_actif_detaille. Seules les
variations de présentation (Var. K€ et Var. %) sont dérivées ici.

Structure (alignée sur l'onglet Actif détaillé du FM de référence) :
ligne 1 lien retour Sommaire, lignes 3-4 titre + client, ligne 8 en-têtes,
données à partir de la ligne 10 : postes par section, sous-total par
section, agrégats (Actif immobilisé net, Actif circulant) et TOTAL ACTIF.
"""

import logging
from typing import Tuple, Union

from openpyxl.worksheet.worksheet import Worksheet

from src.models.financial_statements import ActifDetaille, PassifDetaille
from src.writers.styles import (
    FONT_META, FONT_NORMAL, NUM_KE, NUM_PCT,
    remove_gridlines, set_col_widths, write_data_row, write_header_row,
    write_title_block, write_total_row,
)

logger = logging.getLogger(__name__)

# Largeurs de colonnes (mêmes proportions que l'onglet EBIT)
_WIDTHS_DETAILLE = {"A": 6, "B": 56, "C": 4, "D": 14, "E": 14,
                    "F": 12, "G": 10}


def _var(valeur_n: float, valeur_n1: float) -> Tuple[float, Union[float, str]]:
    """Variations de présentation : Var. K€ et Var. % ('n/a' si N-1 ≈ 0)."""
    var_ke = round(valeur_n - valeur_n1, 3)
    if abs(valeur_n1) >= 0.001:
        var_pct: Union[float, str] = round(var_ke / abs(valeur_n1), 4)
    else:
        var_pct = "n/a"
    return var_ke, var_pct


def _ecrire_poste(ws: Worksheet, row: int, libelle: str,
                  valeur_n: float, valeur_n1: float) -> None:
    """Écrit une ligne de poste (données)."""
    var_ke, var_pct = _var(valeur_n, valeur_n1)
    write_data_row(ws, row, [
        (2, libelle,   None,   FONT_NORMAL),
        (4, valeur_n,  NUM_KE, FONT_NORMAL),
        (5, valeur_n1, NUM_KE, FONT_NORMAL),
        (6, var_ke,    NUM_KE, FONT_NORMAL),
        (7, var_pct,
           NUM_PCT if not isinstance(var_pct, str) else None,
           FONT_NORMAL),
    ])


def _ecrire_total(ws: Worksheet, row: int, libelle: str,
                  valeur_n: float, valeur_n1: float) -> None:
    """Écrit une ligne de sous-total / agrégat / total (bold, trait dessus)."""
    var_ke, var_pct = _var(valeur_n, valeur_n1)
    write_total_row(ws, row, [
        (2, libelle),
        (4, valeur_n,  NUM_KE),
        (5, valeur_n1, NUM_KE),
        (6, var_ke,    NUM_KE),
        (7, var_pct, NUM_PCT if not isinstance(var_pct, str) else None),
    ])


def ecrire_etat_detaille_tab(ws: Worksheet,
                             etat: Union[ActifDetaille, PassifDetaille],
                             titre: str, date_n: str, date_n1: str,
                             client: str) -> None:
    """Écrit un onglet d'état détaillé (Actif ou Passif — même rendu).

    Paramètres
    ----------
    ws : Worksheet
        Feuille openpyxl cible (déjà créée et nommée).
    etat : ActifDetaille ou PassifDetaille
        Résultat de financial_engine.calculer_actif_detaille /
        calculer_passif_detaille (sections, agrégats, total).
    titre : str
        Titre de l'onglet ("Actif détaillé" / "Passif détaillé").
    date_n : str
        Date de clôture N au format 'JJ/MM/AAAA' (string, jamais datetime).
    date_n1 : str
        Date de clôture N-1 au format 'JJ/MM/AAAA'.
    client : str
        Nom du client.
    """
    remove_gridlines(ws)
    set_col_widths(ws, _WIDTHS_DETAILLE)

    ws.cell(row=1, column=1, value="Retour sommaire").font = FONT_META

    write_title_block(ws, row_titre=3, row_client=4,
                      titre=titre, client=client)

    write_header_row(ws, 8, [
        (2, "En milliers d'€uros"),
        (4, date_n), (5, date_n1),
        (6, "Var. K€"), (7, "Var. %"),
    ])

    # Agrégats indexés par section après laquelle ils s'insèrent
    agregats_apres: dict = {}
    for agregat in etat.agregats:
        agregats_apres.setdefault(agregat.apres_section, []).append(agregat)

    row = 10
    for section in etat.sections:
        for poste in section.postes:
            _ecrire_poste(ws, row, poste.libelle,
                          poste.valeur_n, poste.valeur_n1)
            row += 1
        if section.sous_total is not None:
            _ecrire_total(ws, row, section.sous_total.libelle,
                          section.sous_total.valeur_n,
                          section.sous_total.valeur_n1)
            row += 1
        row += 1  # ligne vide entre sections

        for agregat in agregats_apres.get(section.cle, []):
            _ecrire_total(ws, row, agregat.libelle,
                          agregat.valeur_n, agregat.valeur_n1)
            row += 2  # agrégat + ligne vide

    _ecrire_total(ws, row, etat.total.libelle,
                  etat.total.valeur_n, etat.total.valeur_n1)

    logger.info(
        "%s : onglet généré (total N=%.0f K€, N-1=%.0f K€)",
        titre, etat.total.valeur_n, etat.total.valeur_n1,
    )


def ecrire_actif_detaille_tab(ws: Worksheet, actif: ActifDetaille,
                              date_n: str, date_n1: str,
                              client: str) -> None:
    """Écrit l'onglet Actif détaillé (cerfa 2050).

    Paramètres
    ----------
    ws : Worksheet
        Feuille openpyxl cible (déjà créée et nommée "Actif détaillé").
    actif : ActifDetaille
        Résultat de financial_engine.calculer_actif_detaille.
    date_n : str
        Date de clôture N au format 'JJ/MM/AAAA'.
    date_n1 : str
        Date de clôture N-1 au format 'JJ/MM/AAAA'.
    client : str
        Nom du client.
    """
    ecrire_etat_detaille_tab(ws, actif, "Actif détaillé",
                             date_n, date_n1, client)
