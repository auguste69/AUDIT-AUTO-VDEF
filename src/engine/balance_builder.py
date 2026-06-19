"""
Construction de la balance générale à partir du FEC.

Agrège les écritures par compte, calcule les variations vs N-1
et convertit les montants en K€.
"""

import logging
from typing import Dict, Any, Optional, Union

import pandas as pd

logger = logging.getLogger(__name__)

# Type du dict balance N-1 : {compte_num: {"libelle": str, "solde_ke": float}}
BalanceN1 = Dict[str, Dict[str, Any]]


def build(df_fec: pd.DataFrame, balance_n1: Optional[BalanceN1] = None) -> pd.DataFrame:
    """
    Construit la balance générale comparative N vs N-1.

    Paramètres
    ----------
    df_fec : pd.DataFrame
        DataFrame produit par fec_parser.parse().
    balance_n1 : dict, optionnel
        Dict {compte_num: {"libelle": str, "solde_ke": float}}.
        Si None ou vide, les colonnes N-1 seront à 0 et Var_PCT à 'n/a'.

    Retourne
    --------
    pd.DataFrame
        Colonnes : CompteNum, CompteLib, Debit, Credit, Solde,
                   Solde_KE, Solde_N1_KE, Var_KE, Var_PCT.
        Trié par CompteNum. Les comptes présents en N-1 mais soldés en N
        (absents du FEC) figurent dans la balance avec des montants N à 0,
        pour que la comparaison N vs N-1 soit exhaustive.

    Lève
    ----
    ValueError
        Si la somme des soldes dépasse 0,01 € (FEC déséquilibré).
    """
    if balance_n1 is None:
        balance_n1 = {}

    # --- Agrégation par compte ---
    # Clé d'agrégation : CompteNum SEUL. Un même compte peut porter
    # plusieurs libellés dans le FEC (renommage en cours d'exercice) :
    # agréger sur (CompteNum, CompteLib) créerait des lignes en doublon
    # et fausserait les totaux N-1. On retient le premier libellé.
    balance = (
        df_fec
        .groupby("CompteNum", as_index=False)
        .agg(
            CompteLib=("CompteLib", "first"),
            Debit=("Debit", "sum"),
            Credit=("Credit", "sum"),
            Solde=("Solde", "sum"),
        )
        .sort_values("CompteNum")
        .reset_index(drop=True)
    )

    # --- Contrôle d'équilibre ---
    somme_soldes = balance["Solde"].sum()
    if abs(somme_soldes) > 0.01:
        raise ValueError(
            f"FEC déséquilibré : somme des soldes = {somme_soldes:.2f} € "
            f"(seuil = 0,01 €)"
        )
    logger.info(
        "Balance construite : %d comptes, somme des soldes = %.4f €",
        len(balance), somme_soldes,
    )

    # --- Conversion en K€ ---
    balance["Solde_KE"] = (balance["Solde"] / 1000).round(3)

    # --- Soldes N-1 ---
    balance["Solde_N1_KE"] = balance["CompteNum"].map(
        lambda num: balance_n1.get(str(num), {}).get("solde_ke", 0.0)
    )

    # --- Comptes N-1 soldés en N (absents du FEC) ---
    # Sans ces lignes, la somme des soldes N-1 ne reboucle pas à 0 et les
    # variations des comptes soldés (emprunt remboursé, stock liquidé…)
    # disparaissent de la revue analytique.
    nums_n = set(balance["CompteNum"].astype(str))
    orphelins = [
        {
            "CompteNum": str(num),
            "CompteLib": str((info or {}).get("libelle", "")),
            "Debit": 0.0, "Credit": 0.0, "Solde": 0.0, "Solde_KE": 0.0,
            "Solde_N1_KE": float((info or {}).get("solde_ke", 0.0)),
        }
        for num, info in balance_n1.items()
        if str(num) not in nums_n
        and abs(float((info or {}).get("solde_ke", 0.0))) >= 0.0005
    ]
    if orphelins:
        logger.info(
            "Balance comparative : %d compte(s) N-1 soldé(s) en N ajouté(s) "
            "(somme N-1 = %.3f K€)",
            len(orphelins), sum(o["Solde_N1_KE"] for o in orphelins),
        )
        balance = (
            pd.concat([balance, pd.DataFrame(orphelins)], ignore_index=True)
            .sort_values("CompteNum")
            .reset_index(drop=True)
        )

    # --- Variations ---
    balance["Var_KE"] = (balance["Solde_KE"] - balance["Solde_N1_KE"]).round(3)

    def _var_pct(row: pd.Series) -> Union[float, str]:
        n1 = row["Solde_N1_KE"]
        if abs(n1) < 0.001:
            return "n/a"
        return round((row["Var_KE"] / abs(n1)), 4)

    balance["Var_PCT"] = balance.apply(_var_pct, axis=1)

    return balance[
        ["CompteNum", "CompteLib", "Debit", "Credit", "Solde",
         "Solde_KE", "Solde_N1_KE", "Var_KE", "Var_PCT"]
    ]
