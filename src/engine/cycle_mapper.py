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
        Copie de la balance avec les colonnes supplémentaires :
        cycle, etatfi_n, etatfi_n1, compta_n, compta_n1, ref,
        plus les aliases historiques etatfi (= etatfi_n) et
        compta (= compta_n) pour les consommateurs existants.

    Notes
    -----
    Pour les comptes absents du mapping FM (nouveaux comptes résolus par
    le fallback PCG), les valeurs N-1 (etatfi_n1, compta_n1) restent VIDES :
    le compte n'existait pas dans l'exercice N-1, il n'a donc pas de
    classification N-1 — c'est le choix le plus cohérent (pas d'invention
    d'historique).
    """
    if mapping_fm is None:
        mapping_fm = {}

    prefixes_pcg  = pcg_config["prefixes"]
    surcharges     = pcg_config["surcharges"]
    passif_pref    = pcg_config["passif_prefixes"]

    df = balance.copy()

    cycles:     list = []
    comptas_n:  list = []
    comptas_n1: list = []
    etatfis_n:  list = []
    etatfis_n1: list = []
    refs:       list = []

    nb_fm  = 0
    nb_pcg = 0
    nb_inconnu = 0

    for compte_num in df["CompteNum"]:
        num = str(compte_num).strip()

        # --- Priorité 1 : FM existant ---
        if num in mapping_fm:
            info = mapping_fm[num]
            # .get avec fallback sur les aliases : tolère les mappings
            # construits avant l'introduction des clés _n/_n1
            etatfi_n  = info.get("etatfi_n",  info.get("etatfi", ""))
            etatfi_n1 = info.get("etatfi_n1", "")
            compta_n  = info.get("compta_n",  info.get("compta", ""))
            compta_n1 = info.get("compta_n1", "")
            cycles.append(info["cycle"])
            comptas_n.append(compta_n)
            comptas_n1.append(compta_n1)
            etatfis_n.append(etatfi_n)
            etatfis_n1.append(etatfi_n1)
            refs.append(info["ref"])
            nb_fm += 1
            continue

        # --- Priorités 2-5 : YAML (surcharges + PCG par préfixe) ---
        cycle_pcg = _resoudre_cycle(num, prefixes_pcg, surcharges)

        if cycle_pcg:
            compta = _resoudre_compta(num, passif_pref)
            ref    = f"{cycle_pcg}0"
            cycles.append(cycle_pcg)
            comptas_n.append(compta)
            comptas_n1.append("")   # nouveau compte : pas d'historique N-1
            etatfis_n.append("")
            etatfis_n1.append("")
            refs.append(ref)
            nb_pcg += 1
        else:
            # Ne devrait jamais arriver avec un YAML exhaustif
            logger.warning("Compte %s sans cycle trouvé — vérifie le mapping_pcg.yaml", num)
            cycles.append("")
            comptas_n.append("")
            comptas_n1.append("")
            etatfis_n.append("")
            etatfis_n1.append("")
            refs.append("")
            nb_inconnu += 1

    df["cycle"]     = cycles
    df["etatfi_n"]  = etatfis_n
    df["etatfi_n1"] = etatfis_n1
    df["compta_n"]  = comptas_n
    df["compta_n1"] = comptas_n1
    df["ref"]       = refs
    # Aliases historiques (consommateurs existants : app.py, template_writer…)
    df["etatfi"] = df["etatfi_n"]
    df["compta"] = df["compta_n"]

    logger.info(
        "Mapping : %d comptes FM, %d comptes PCG, %d inconnus (total %d)",
        nb_fm, nb_pcg, nb_inconnu, len(df),
    )

    return df
