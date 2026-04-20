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
        Trié par CompteNum.

    Lève
    ----
    ValueError
        Si la somme des soldes dépasse 0,01 € (FEC déséquilibré).
    """
    if balance_n1 is None:
        balance_n1 = {}

    # --- Agrégation par compte ---
    balance = (
        df_fec
        .groupby(["CompteNum", "CompteLib"], as_index=False)
        .agg(
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
