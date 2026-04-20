"""
Ventilation des comptes de la balance par cycle d'audit.

Logique de résolution (ordre de priorité décroissante) :
  1. Mapping FM existant du client          (priorité absolue)
  2. Surcharges cabinet du YAML
  3. Préfixe 3 chiffres dans le YAML
  4. Préfixe 2 chiffres dans le YAML
  5. Classe 1 chiffre dans le YAML          (fallback garanti)

Classification Actif/Passif/Charges/Produits :
  Classe 1        → Passif
  Classe 2, 3, 5  → Actif
  Classe 6        → Charges
  Classe 7        → Produits
  Classe 4        → préfixes dans classification_bilan.passif → Passif, sinon Actif
"""

import logging
from typing import Optional

import pandas as pd

from src.parsers.mapping_parser import MappingCompte

logger = logging.getLogger(__name__)

# Classification par classe (1 chiffre)
_COMPTA_PAR_CLASSE = {
    "1": "Passif",
    "2": "Actif",
    "3": "Actif",
    "5": "Actif",
    "6": "Charges",
    "7": "Produits",
}


def _resoudre_cycle(
    compte_num: str,
    prefixes_pcg: dict,
    surcharges: dict,
) -> Optional[str]:
    """
    Résout le cycle d'un compte par préfixe (3 → 2 → 1 chiffre).
    Cherche d'abord dans les surcharges cabinet, puis dans le PCG.
    """
    for longueur in (3, 2, 1):
        prefixe = compte_num[:longueur]
        if prefixe in surcharges:
            return surcharges[prefixe]
        if prefixe in prefixes_pcg:
            return prefixes_pcg[prefixe]
    return None


def _resoudre_compta(
    compte_num: str,
    passif_prefixes: set,
) -> str:
    """
    Détermine la classification bilan (Actif/Passif/Charges/Produits).
    """
    if not compte_num:
        return ""
    classe = compte_num[0]

    if classe in _COMPTA_PAR_CLASSE:
        return _COMPTA_PAR_CLASSE[classe]

    if classe == "4":
        # Tester 3 chiffres puis 2 chiffres
        for longueur in (3, 2):
            if compte_num[:longueur] in passif_prefixes:
                return "Passif"
        return "Actif"

    return ""


def map_cycles(
    balance: pd.DataFrame,
    mapping_fm: Optional[MappingCompte],
    pcg_config: dict,
) -> pd.DataFrame:
    """
    Enrichit la balance avec les colonnes de mapping cycle.

    Paramètres
    ----------
    balance : pd.DataFrame
        Produit par balance_builder.build() — doit contenir CompteNum.
    mapping_fm : dict ou None
        Mapping du FM existant, produit par mapping_parser.from_fm().
        Peut être None si aucun FM disponible.
    pcg_config : dict
        Config PCG, produit par mapping_parser.from_pcg_config().

    Retourne
    --------
    pd.DataFrame
        Copie de la balance avec 4 colonnes supplémentaires :
        cycle, compta, etatfi, ref.
    """
    if mapping_fm is None:
        mapping_fm = {}

    prefixes_pcg  = pcg_config["prefixes"]
    surcharges     = pcg_config["surcharges"]
    passif_pref    = pcg_config["passif_prefixes"]

    df = balance.copy()

    cycles:  list = []
    comptas: list = []
    etatfis: list = []
    refs:    list = []

    nb_fm  = 0
    nb_pcg = 0
    nb_inconnu = 0

    for compte_num in df["CompteNum"]:
        num = str(compte_num).strip()

        # --- Priorité 1 : FM existant ---
        if num in mapping_fm:
            info = mapping_fm[num]
            cycles.append(info["cycle"])
            comptas.append(info["compta"])
            etatfis.append(info["etatfi"])
            refs.append(info["ref"])
            nb_fm += 1
            continue

        # --- Priorités 2-5 : YAML (surcharges + PCG par préfixe) ---
        cycle_pcg = _resoudre_cycle(num, prefixes_pcg, surcharges)

        if cycle_pcg:
            compta = _resoudre_compta(num, passif_pref)
            ref    = f"{cycle_pcg}0"
            cycles.append(cycle_pcg)
            comptas.append(compta)
            etatfis.append("")
            refs.append(ref)
            nb_pcg += 1
        else:
            # Ne devrait jamais arriver avec un YAML exhaustif
            logger.warning("Compte %s sans cycle trouvé — vérifie le mapping_pcg.yaml", num)
            cycles.append("")
            comptas.append("")
            etatfis.append("")
            refs.append("")
            nb_inconnu += 1

    df["cycle"]  = cycles
    df["compta"] = comptas
    df["etatfi"] = etatfis
    df["ref"]    = refs

    logger.info(
        "Mapping : %d comptes FM, %d comptes PCG, %d inconnus (total %d)",
        nb_fm, nb_pcg, nb_inconnu, len(df),
    )

    return df
