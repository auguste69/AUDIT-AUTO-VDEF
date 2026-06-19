"""
Écriture de l'onglet EBIT (résultat d'exploitation) du FM.

Écriture pure : toutes les valeurs proviennent d'un EbitSynthetique calculé
par src.engine.financial_engine.calculer_ebit. Comme dans les autres onglets
(Bilan, Tréso), seules les variations de présentation (Var. K€ = N − N-1 et
Var. % associée) sont dérivées ici, ligne par ligne.

Structure (alignée sur l'onglet EBIT du FM de référence et le cerfa 2052) :
ligne 1 lien retour Sommaire, lignes 3-4 titre + client, ligne 8 en-têtes,
données à partir de la ligne 10, sous-totaux CA / Production / totaux et
ligne finale EBIT.
"""

import logging
from typing import List, Tuple, Union

from openpyxl.worksheet.worksheet import Worksheet

from src.models.financial_statements import EbitSynthetique
from src.writers.styles import (
    FONT_META, FONT_NORMAL, FONT_SECTION, NUM_KE, NUM_PCT,
    remove_gridlines, set_col_widths, write_data_row, write_header_row,
    write_title_block, write_total_row,
)

logger = logging.getLogger(__name__)

# Largeurs de colonnes (mêmes proportions que l'onglet Tréso)
_WIDTHS_EBIT = {"A": 6, "B": 52, "C": 4, "D": 14, "E": 14, "F": 12, "G": 10}

# Lignes de l'onglet : (type, libellé, clé dans EbitSynthetique.as_dict()).
# Types : "s" = label de section, "p" = poste, "t" = sous-total/total,
#         "b" = ligne vide.
_LIGNES_EBIT: List[Tuple[str, str, str]] = [
    ("s", "PRODUITS D'EXPLOITATION", ""),
    ("p", "Ventes de marchandises",                          "ventes_marchandises"),
    ("p", "Production vendue (Biens et Services)",           "production_vendue"),
    ("t", "= Chiffre d'affaires",                            "ca"),
    ("b", "", ""),
    ("p", "Production stockée",                              "production_stockee"),
    ("p", "Production immobilisée",                          "production_immobilisee"),
    ("p", "Subventions d'exploitation",                      "subventions_exploitation"),
    ("p", "Reprises amort./prov., transferts de charges",    "reprises_transferts"),
    ("p", "Autres produits",                                 "autres_produits"),
    ("t", "= Total produits d'exploitation",                 "total_produits"),
    ("b", "", ""),
    ("s", "CHARGES D'EXPLOITATION", ""),
    ("p", "Achats de marchandises",                          "achats_marchandises"),
    ("p", "Variation de stocks de marchandises",             "variation_stocks_marchandises"),
    ("p", "Achats de matières premières et autres appro.",   "achats_matieres_premieres"),
    ("p", "Variation de stocks de matières",                 "variation_stocks_matieres"),
    ("p", "Autres achats et charges externes",               "autres_charges_externes"),
    ("p", "Impôts, taxes et versements assimilés",           "impots_taxes"),
    ("p", "Salaires et traitements",                         "salaires_traitements"),
    ("p", "Charges sociales",                                "charges_sociales"),
    ("p", "Dotations aux amortissements sur immobilisations", "dotations_amortissements"),
    ("p", "Dotations aux dépréciations sur immobilisations", "dotations_dep_immobilisations"),
    ("p", "Dotations aux dépréciations sur actif circulant", "dotations_dep_actif_circulant"),
    ("p", "Dotations aux provisions",                        "dotations_provisions"),
    ("p", "Autres charges",                                  "autres_charges"),
    ("t", "= Total charges d'exploitation",                  "total_charges"),
    ("b", "", ""),
    ("t", "= RÉSULTAT D'EXPLOITATION (EBIT)",                "ebit"),
]


def _var(valeur_n: float, valeur_n1: float) -> Tuple[float, Union[float, str]]:
    """Variations de présentation : Var. K€ et Var. % ('n/a' si N-1 ≈ 0)."""
    var_ke = round(valeur_n - valeur_n1, 3)
    if abs(valeur_n1) >= 0.001:
        var_pct: Union[float, str] = round(var_ke / abs(valeur_n1), 4)
    else:
        var_pct = "n/a"
    return var_ke, var_pct


def ecrire_ebit_tab(ws: Worksheet, ebit: EbitSynthetique,
                    date_n: str, date_n1: str, client: str) -> None:
    """Écrit l'onglet EBIT — compte de résultat d'exploitation (cerfa 2052).

    Paramètres
    ----------
    ws : Worksheet
        Feuille openpyxl cible (déjà créée et nommée "EBIT").
    ebit : EbitSynthetique
        Résultat de financial_engine.calculer_ebit (postes et agrégats,
        produits et charges déjà présentés en positif).
    date_n : str
        Date de clôture N au format 'JJ/MM/AAAA' (string, jamais datetime).
    date_n1 : str
        Date de clôture N-1 au format 'JJ/MM/AAAA'.
    client : str
        Nom du client.
    """
    remove_gridlines(ws)
    set_col_widths(ws, _WIDTHS_EBIT)

    ws.cell(row=1, column=1, value="Retour sommaire").font = FONT_META

    write_title_block(ws, row_titre=3, row_client=4,
                      titre="EBIT", client=client)

    write_header_row(ws, 8, [
        (2, "En milliers d'€uros"),
        (4, date_n), (5, date_n1),
        (6, "Var. K€"), (7, "Var. %"),
    ])

    valeurs = ebit.as_dict()
    row = 10
    for type_ligne, libelle, cle in _LIGNES_EBIT:
        if type_ligne == "b":
            row += 1
            continue
        if type_ligne == "s":
            ws.cell(row=row, column=2, value=libelle).font = FONT_SECTION
            row += 1
            continue

        valeur_n, valeur_n1 = valeurs[cle]
        var_ke, var_pct = _var(valeur_n, valeur_n1)
        fmt_pct = NUM_PCT if not isinstance(var_pct, str) else None
        if type_ligne == "p":
            write_data_row(ws, row, [
                (2, libelle,   None,    FONT_NORMAL),
                (4, valeur_n,  NUM_KE,  FONT_NORMAL),
                (5, valeur_n1, NUM_KE,  FONT_NORMAL),
                (6, var_ke,    NUM_KE,  FONT_NORMAL),
                (7, var_pct,   fmt_pct, FONT_NORMAL),
            ])
        else:  # "t" — sous-total ou total
            write_total_row(ws, row, [
                (2, libelle),
                (4, valeur_n,  NUM_KE),
                (5, valeur_n1, NUM_KE),
                (6, var_ke,    NUM_KE),
                (7, var_pct,   fmt_pct),
            ])
        row += 1

    logger.info(
        "EBIT : onglet généré (CA=%.0f, EBIT=%.0f K€)",
        ebit.ca.valeur_n, ebit.ebit.valeur_n,
    )
