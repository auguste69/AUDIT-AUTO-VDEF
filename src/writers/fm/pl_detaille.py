"""
Écriture de l'onglet P&L détaillé (compte de résultat, cerfa 2052/2053).

Écriture pure : toutes les valeurs proviennent d'un PlDetaille calculé par
src.engine.financial_engine.calculer_pl_detaille. Convention liasse :
produits en positif, charges en NÉGATIF (les charges EBIT, présentées en
positif dans l'onglet EBIT, sont négativées ici), résultats = sommes
algébriques. Le résultat net est égal à −(somme des classes 6 et 7).

Structure : ligne 1 lien retour Sommaire, lignes 3-4 titre + client,
ligne 8 en-têtes, données à partir de la ligne 10.
"""

import logging
from typing import List, Tuple, Union

from openpyxl.worksheet.worksheet import Worksheet

from src.models.financial_statements import PlDetaille
from src.writers.styles import (
    FONT_META, FONT_NORMAL, FONT_SECTION, NUM_KE, NUM_PCT,
    remove_gridlines, set_col_widths, write_data_row, write_header_row,
    write_title_block, write_total_row,
)

logger = logging.getLogger(__name__)

# Largeurs de colonnes (mêmes proportions que l'onglet EBIT)
_WIDTHS_PL = {"A": 6, "B": 56, "C": 4, "D": 14, "E": 14, "F": 12, "G": 10}

# Lignes de l'onglet : (type, libellé, clé, signe).
# Types : "s" = label de section, "p" = poste, "t" = sous-total/résultat,
#         "b" = ligne vide. Clés préfixées "ebit:" → EbitSynthetique.as_dict()
# (charges EBIT × -1 pour l'affichage liasse), sinon PlDetaille.as_dict().
_LIGNES_PL: List[Tuple[str, str, str, int]] = [
    ("s", "PRODUITS D'EXPLOITATION", "", 1),
    ("p", "Ventes de marchandises",                        "ebit:ventes_marchandises", 1),
    ("p", "Production vendue (Biens et Services)",         "ebit:production_vendue", 1),
    ("t", "= Chiffre d'affaires",                          "ebit:ca", 1),
    ("p", "Production stockée",                            "ebit:production_stockee", 1),
    ("p", "Production immobilisée",                        "ebit:production_immobilisee", 1),
    ("p", "Subventions d'exploitation",                    "ebit:subventions_exploitation", 1),
    ("p", "Reprises amort./prov., transferts de charges",  "ebit:reprises_transferts", 1),
    ("p", "Autres produits",                               "ebit:autres_produits", 1),
    ("p", "Autres produits d'exploitation (divers)",       "produits_expl_divers", 1),
    ("t", "= Total produits d'exploitation",               "produits_exploitation", 1),
    ("b", "", "", 1),
    ("s", "CHARGES D'EXPLOITATION", "", 1),
    ("p", "Achats de marchandises",                        "ebit:achats_marchandises", -1),
    ("p", "Variation de stocks de marchandises",           "ebit:variation_stocks_marchandises", -1),
    ("p", "Achats de matières premières et autres appro.", "ebit:achats_matieres_premieres", -1),
    ("p", "Variation de stocks de matières",               "ebit:variation_stocks_matieres", -1),
    ("p", "Autres achats et charges externes",             "ebit:autres_charges_externes", -1),
    ("p", "Impôts, taxes et versements assimilés",         "ebit:impots_taxes", -1),
    ("p", "Salaires et traitements",                       "ebit:salaires_traitements", -1),
    ("p", "Charges sociales",                              "ebit:charges_sociales", -1),
    ("p", "Dotations aux amortissements sur immobilisations", "ebit:dotations_amortissements", -1),
    ("p", "Dotations aux dépréciations sur immobilisations",  "ebit:dotations_dep_immobilisations", -1),
    ("p", "Dotations aux dépréciations sur actif circulant",  "ebit:dotations_dep_actif_circulant", -1),
    ("p", "Dotations aux provisions",                      "ebit:dotations_provisions", -1),
    ("p", "Autres charges",                                "ebit:autres_charges", -1),
    ("p", "Autres charges d'exploitation (divers)",        "charges_expl_divers", 1),
    ("t", "= Total charges d'exploitation",                "charges_exploitation", 1),
    ("b", "", "", 1),
    ("t", "= RÉSULTAT D'EXPLOITATION",                     "resultat_exploitation", 1),
    ("b", "", "", 1),
    ("p", "Quotes-parts de résultat sur opérations faites en commun", "quotes_parts", 1),
    ("b", "", "", 1),
    ("s", "RÉSULTAT FINANCIER", "", 1),
    ("p", "Produits financiers",                           "produits_financiers", 1),
    ("p", "Charges financières",                           "charges_financieres", 1),
    ("t", "= Résultat financier",                          "resultat_financier", 1),
    ("b", "", "", 1),
    ("t", "= RÉSULTAT COURANT AVANT IMPÔTS",               "resultat_courant", 1),
    ("b", "", "", 1),
    ("s", "RÉSULTAT EXCEPTIONNEL", "", 1),
    ("p", "Produits exceptionnels",                        "produits_exceptionnels", 1),
    ("p", "Charges exceptionnelles",                       "charges_exceptionnelles", 1),
    ("t", "= Résultat exceptionnel",                       "resultat_exceptionnel", 1),
    ("b", "", "", 1),
    ("p", "Participation des salariés aux résultats",      "participation_salaries", 1),
    ("p", "Impôts sur les bénéfices",                      "impots_benefices", 1),
    ("b", "", "", 1),
    ("t", "= RÉSULTAT NET COMPTABLE",                      "resultat_net", 1),
]


def _var(valeur_n: float, valeur_n1: float) -> Tuple[float, Union[float, str]]:
    """Variations de présentation : Var. K€ et Var. % ('n/a' si N-1 ≈ 0)."""
    var_ke = round(valeur_n - valeur_n1, 3)
    if abs(valeur_n1) >= 0.001:
        var_pct: Union[float, str] = round(var_ke / abs(valeur_n1), 4)
    else:
        var_pct = "n/a"
    return var_ke, var_pct


def ecrire_pl_detaille_tab(ws: Worksheet, pl: PlDetaille,
                           date_n: str, date_n1: str, client: str) -> None:
    """Écrit l'onglet P&L détaillé — compte de résultat (cerfa 2052/2053).

    Paramètres
    ----------
    ws : Worksheet
        Feuille openpyxl cible (déjà créée et nommée "P&L détaillé").
    pl : PlDetaille
        Résultat de financial_engine.calculer_pl_detaille (postes hors
        exploitation signés, détail exploitation via pl.ebit).
    date_n : str
        Date de clôture N au format 'JJ/MM/AAAA' (string, jamais datetime).
    date_n1 : str
        Date de clôture N-1 au format 'JJ/MM/AAAA'.
    client : str
        Nom du client.
    """
    remove_gridlines(ws)
    set_col_widths(ws, _WIDTHS_PL)

    ws.cell(row=1, column=1, value="Retour sommaire").font = FONT_META

    write_title_block(ws, row_titre=3, row_client=4,
                      titre="Compte de résultat détaillé", client=client)

    write_header_row(ws, 8, [
        (2, "En milliers d'€uros"),
        (4, date_n), (5, date_n1),
        (6, "Var. K€"), (7, "Var. %"),
    ])

    valeurs = pl.as_dict()
    valeurs_ebit = pl.ebit.as_dict()

    def _valeur(cle: str, signe: int) -> Tuple[float, float]:
        source = valeurs_ebit if cle.startswith("ebit:") else valeurs
        valeur_n, valeur_n1 = source[cle.replace("ebit:", "")]
        return round(valeur_n * signe, 3), round(valeur_n1 * signe, 3)

    row = 10
    for type_ligne, libelle, cle, signe in _LIGNES_PL:
        if type_ligne == "b":
            row += 1
            continue
        if type_ligne == "s":
            ws.cell(row=row, column=2, value=libelle).font = FONT_SECTION
            row += 1
            continue

        valeur_n, valeur_n1 = _valeur(cle, signe)
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
        else:  # "t" — sous-total ou résultat
            write_total_row(ws, row, [
                (2, libelle),
                (4, valeur_n,  NUM_KE),
                (5, valeur_n1, NUM_KE),
                (6, var_ke,    NUM_KE),
                (7, var_pct,   fmt_pct),
            ])
        row += 1

    logger.info(
        "P&L détaillé : onglet généré (résultat net N=%.0f K€)",
        pl.resultat_net.valeur_n,
    )
