"""
Chargeur unifié de la balance N-1.

Centralise la logique d'extraction de la balance N-1 (et du mapping FM
associé) auparavant dupliquée entre main.py et app.py.

Sources supportées :
  - FEC N-1 (.txt)             → balance agrégée, pas de mapping FM
  - FM existant (.xlsx)        → soldes N du FM (déjà en K€) + mapping FM
  - Balance Excel simple (.xlsx) → balance convertie en K€, pas de mapping FM

Ce module réutilise les fonctions publiques de
``src.parsers.mapping_parser`` (from_fec_n1, detect_balance_sheet,
from_balance_excel, from_fm) sans en modifier le contrat.
"""

import logging
from pathlib import Path
from typing import Optional, Tuple, Union

import openpyxl

from src.parsers.mapping_parser import (
    BalanceN1Dict,
    MappingCompte,
    detect_balance_sheet,
    from_balance_excel,
    from_fec_n1,
    from_fm,
)

logger = logging.getLogger(__name__)


def _extraire_soldes_from_fm(
    path: Union[str, Path], nom_feuille: str
) -> BalanceN1Dict:
    """
    Extrait les soldes N-1 depuis la feuille balance d'un FM existant.

    Structure attendue (headers en ligne 8, données à partir de la ligne 10) :
      - colonne B (index 1) : CompteNum (int ou float)
      - colonne C (index 2) : CompteLib
      - colonne D (index 3) : Solde N en K€ (déjà en K€, pas de conversion)

    Paramètres
    ----------
    path : str | Path
        Chemin vers le fichier FM (.xlsx).
    nom_feuille : str
        Nom de la feuille balance (détectée via detect_balance_sheet).

    Retourne
    --------
    dict {compte_num_str: {"libelle": str, "solde_ke": float}}
    """
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb[nom_feuille]
    balance: BalanceN1Dict = {}
    for row in ws.iter_rows(min_row=10, values_only=True):
        if row[1] is None:
            continue
        try:
            num = str(int(float(row[1])))
        except (ValueError, TypeError):
            continue
        solde = float(row[3]) if row[3] is not None else 0.0
        balance[num] = {
            "libelle": str(row[2]) if row[2] else "",
            "solde_ke": solde,
        }
    wb.close()
    return balance


def load_balance_n1(
    path: Union[str, Path]
) -> Tuple[BalanceN1Dict, Optional[MappingCompte]]:
    """
    Charge la balance N-1 depuis n'importe quelle source supportée.

    Détection par extension :
      - ``.txt``  : FEC N-1 — balance agrégée via from_fec_n1, mapping None
      - ``.xlsx`` : detect_balance_sheet décide entre FM complet
        (soldes N du FM + mapping via from_fm) et balance Excel simple
        (from_balance_excel, mapping None)
      - autre     : ValueError explicite

    Paramètres
    ----------
    path : str | Path
        Chemin vers le fichier N-1.

    Retourne
    --------
    tuple (balance_n1, mapping_fm)
        balance_n1 : dict {compte_num_str: {"libelle": str, "solde_ke": float}}
        mapping_fm : dict de mapping FM, ou None si la source n'est pas un FM.

    Lève
    ----
    ValueError
        Si l'extension du fichier n'est pas reconnue.
    """
    chemin = Path(path)
    suffix = chemin.suffix.lower()

    if suffix == ".txt":
        # Source FEC N-1 : pas de mapping FM, uniquement balance agrégée
        logger.info("Source N-1 : FEC — %s", path)
        balance_n1 = from_fec_n1(chemin)
        logger.info(
            "Balance N-1 : %d comptes chargés depuis le FEC N-1", len(balance_n1)
        )
        return balance_n1, None

    if suffix == ".xlsx":
        nom_feuille, mode = detect_balance_sheet(chemin)

        if mode == "fm":
            logger.info("Source N-1 : FM existant — %s", path)
            mapping_fm = from_fm(chemin)
            balance_n1 = _extraire_soldes_from_fm(chemin, nom_feuille)
            logger.info(
                "Balance N-1 : %d comptes chargés depuis le FM", len(balance_n1)
            )
            return balance_n1, mapping_fm

        # Balance Excel simple : pas de mapping FM
        logger.info("Source N-1 : Balance Excel simple — %s", path)
        balance_n1 = from_balance_excel(chemin)
        logger.info(
            "Balance N-1 : %d comptes chargés depuis la balance Excel",
            len(balance_n1),
        )
        return balance_n1, None

    raise ValueError(
        f"Format N-1 non reconnu pour '{path}'. "
        "Formats acceptés : .txt (FEC), .xlsx (FM ou balance simple)."
    )
