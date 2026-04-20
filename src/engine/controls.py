"""
Contrôles d'intégrité du FEC.

Chaque contrôle retourne un tuple :
    (nom: str, ok: bool, detail: str, severity: str)

Severity :
    - BLOQUANT : écart comptable, impossible de continuer
    - WARNING  : anomalie à examiner
    - INFO     : point d'attention, information
"""

import logging
import math
from collections import Counter
from typing import List, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

# Colonnes obligatoires FEC (art. L.47 A LPF)
COLONNES_OBLIGATOIRES = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib", "Debit", "Credit",
    "EcritureLet", "DateLet", "ValidDate", "Montantdevise", "Idevise",
]

# Distribution de Benford : probabilité attendue pour chaque premier chiffre 1-9
BENFORD_ATTENDU = {d: math.log10(1.0 + 1.0 / d) for d in range(1, 10)}

ResultatControle = Tuple[str, bool, str, str]


def run_all(
    df: pd.DataFrame,
    date_cloture: str,
    seuil_montant_rond: float = 10_000.0,
    seuil_benford_mad: float = 0.015,
) -> List[ResultatControle]:
    """
    Exécute tous les contrôles d'intégrité sur le FEC.

    Paramètres
    ----------
    df : pd.DataFrame
        DataFrame produit par fec_parser.parse().
    date_cloture : str
        Date de clôture de l'exercice au format 'JJ/MM/AAAA' (ex: '31/12/2025').
    seuil_montant_rond : float
        Seuil en € au-dessus duquel un montant rond est signalé (défaut : 10 000).
    seuil_benford_mad : float
        Seuil du MAD Benford au-dessus duquel la distribution est anormale (défaut : 0.015).

    Retourne
    --------
    list[tuple[str, bool, str, str]]
        Liste de (nom_controle, ok, detail, severity).
    """
    date_cloture_ts = pd.to_datetime(date_cloture, format="%d/%m/%Y")
    resultats: List[ResultatControle] = []

    controles = [
        _ctrl_colonnes_obligatoires,
        _ctrl_equilibre,
        _ctrl_lignes_zero,
        _ctrl_coherence_dates,
        _ctrl_ecritures_dimanche,
        _ctrl_montants_ronds,
        _ctrl_doublons,
        _ctrl_benford,
        _ctrl_ecritures_tardives,
    ]

    kwargs = {
        "date_cloture_ts": date_cloture_ts,
        "seuil_montant_rond": seuil_montant_rond,
        "seuil_benford_mad": seuil_benford_mad,
    }

    for controle in controles:
        try:
            res = controle(df, **kwargs)
            resultats.append(res)
            statut = "OK" if res[1] else "KO"
            logger.info("[%s] %s — %s", statut, res[0], res[2])
        except Exception as exc:  # noqa: BLE001
            logger.error("Erreur dans le contrôle %s : %s", controle.__name__, exc)
            resultats.append((controle.__name__, False, f"Erreur inattendue : {exc}", "BLOQUANT"))

    return resultats


# =============================================================================
# Contrôles individuels
# =============================================================================

def _ctrl_colonnes_obligatoires(df: pd.DataFrame, **_) -> ResultatControle:
    """Vérifie que les 18 colonnes obligatoires sont présentes."""
    manquantes = [c for c in COLONNES_OBLIGATOIRES if c not in df.columns]
    if manquantes:
        return (
            "Colonnes obligatoires",
            False,
            f"{len(manquantes)} colonne(s) manquante(s) : {', '.join(manquantes)}",
            "BLOQUANT",
        )
    return ("Colonnes obligatoires", True, "18 colonnes obligatoires présentes", "BLOQUANT")


def _ctrl_equilibre(df: pd.DataFrame, **_) -> ResultatControle:
    """Vérifie que la somme des Débits égale la somme des Crédits."""
    total_debit = df["Debit"].sum()
    total_credit = df["Credit"].sum()
    ecart = abs(total_debit - total_credit)
    ok = bool(ecart <= 0.01)
    detail = (
        f"Débit={total_debit:,.2f} € / Crédit={total_credit:,.2f} € — écart={ecart:.4f} €"
    )
    return ("Équilibre Débit/Crédit", ok, detail, "BLOQUANT")


def _ctrl_lignes_zero(df: pd.DataFrame, **_) -> ResultatControle:
    """Signale les lignes où Débit=0 ET Crédit=0."""
    masque = (df["Debit"] == 0.0) & (df["Credit"] == 0.0)
    nb = masque.sum()
    ok = bool(nb == 0)
    detail = f"{nb} ligne(s) avec Débit=0 et Crédit=0" if nb else "Aucune ligne à zéro"
    return ("Lignes à zéro", ok, detail, "WARNING")


def _ctrl_coherence_dates(
    df: pd.DataFrame,
    date_cloture_ts: pd.Timestamp,
    **_,
) -> ResultatControle:
    """Vérifie que toutes les EcritureDates sont dans l'exercice (01/01/N — date clôture)."""
    dates = pd.to_datetime(df["EcritureDate"], format="%Y%m%d", errors="coerce")
    hors_plage_nat = dates.isna().sum()
    debut_exercice = pd.Timestamp(year=date_cloture_ts.year, month=1, day=1)
    hors_plage = ((dates < debut_exercice) | (dates > date_cloture_ts)).sum()
    total_anomalies = int(hors_plage_nat) + int(hors_plage)
    ok = bool(total_anomalies == 0)
    if ok:
        detail = (
            f"Toutes les dates sont dans l'exercice "
            f"({debut_exercice.strftime('%d/%m/%Y')} — {date_cloture_ts.strftime('%d/%m/%Y')})"
        )
    else:
        detail = (
            f"{total_anomalies} date(s) hors exercice "
            f"(dont {hors_plage_nat} non parsable(s))"
        )
    return ("Cohérence des dates", ok, detail, "WARNING")


def _ctrl_ecritures_dimanche(
    df: pd.DataFrame,
    **_,
) -> ResultatControle:
    """Signale les écritures dont EcritureDate tombe un dimanche."""
    dates = pd.to_datetime(df["EcritureDate"], format="%Y%m%d", errors="coerce")
    dimanches = (dates.dt.dayofweek == 6).sum()
    ok = bool(dimanches == 0)
    detail = (
        f"{dimanches} écriture(s) passée(s) un dimanche"
        if dimanches
        else "Aucune écriture un dimanche"
    )
    return ("Écritures un dimanche", ok, detail, "WARNING")


def _ctrl_montants_ronds(
    df: pd.DataFrame,
    seuil_montant_rond: float = 10_000.0,
    **_,
) -> ResultatControle:
    """Signale les montants ronds (sans centimes) supérieurs au seuil."""
    montants = pd.concat([
        df.loc[df["Debit"] >= seuil_montant_rond, "Debit"],
        df.loc[df["Credit"] >= seuil_montant_rond, "Credit"],
    ])
    ronds = montants[montants % 1.0 == 0.0]
    nb = len(ronds)
    ok = bool(nb == 0)
    seuil_ke = seuil_montant_rond / 1000
    detail = (
        f"{nb} montant(s) rond(s) ≥ {seuil_ke:.0f} K€ (à documenter)"
        if nb
        else f"Aucun montant rond ≥ {seuil_ke:.0f} K€"
    )
    return ("Montants ronds", ok, detail, "INFO")


def _ctrl_doublons(df: pd.DataFrame, **_) -> ResultatControle:
    """Détecte les doublons potentiels : même journal, numéro, compte, montants et libellé."""
    cles = ["JournalCode", "EcritureNum", "CompteNum", "Debit", "Credit", "EcritureLib"]
    masque = df.duplicated(subset=cles, keep=False)
    nb_lignes = int(masque.sum())
    # Nombre de groupes distincts en doublon
    nb_groupes = df[masque].groupby(cles).ngroups if nb_lignes else 0
    ok = bool(nb_lignes == 0)
    detail = (
        f"{nb_lignes} ligne(s) en doublon potentiel ({nb_groupes} groupe(s) distincts)"
        if nb_lignes
        else "Aucun doublon potentiel détecté"
    )
    return ("Doublons potentiels", ok, detail, "WARNING")


def _ctrl_benford(
    df: pd.DataFrame,
    seuil_benford_mad: float = 0.015,
    **_,
) -> ResultatControle:
    """
    Test de Benford sur le premier chiffre des montants (Débit et Crédit > 0).

    Calcule le MAD (Mean Absolute Deviation) entre distribution observée et théorique,
    ainsi que le chi² (sans correction de Yates).
    """
    # Extraire tous les montants positifs (Débit + Crédit)
    vals = pd.concat([
        df.loc[df["Debit"] > 0, "Debit"],
        df.loc[df["Credit"] > 0, "Credit"],
    ]).values

    if len(vals) == 0:
        return ("Benford", True, "Aucun montant à analyser", "INFO")

    # Premier chiffre significatif
    premiers: List[int] = []
    for v in vals:
        s = f"{abs(v):.10f}".lstrip("0").replace(".", "")
        if s:
            d = int(s[0])
            if 1 <= d <= 9:
                premiers.append(d)

    total = len(premiers)
    if total == 0:
        return ("Benford", True, "Aucun premier chiffre exploitable", "INFO")

    obs_count = Counter(premiers)
    obs_freq = {d: obs_count.get(d, 0) / total for d in range(1, 10)}

    # MAD
    mad = sum(abs(obs_freq[d] - BENFORD_ATTENDU[d]) for d in range(1, 10)) / 9

    # Chi²
    chi2 = sum(
        (obs_count.get(d, 0) - total * BENFORD_ATTENDU[d]) ** 2
        / (total * BENFORD_ATTENDU[d])
        for d in range(1, 10)
    )

    ok = bool(mad <= seuil_benford_mad)
    distribution = "  ".join(
        f"{d}:{obs_freq[d]:.3f}(att:{BENFORD_ATTENDU[d]:.3f})"
        for d in range(1, 10)
    )
    detail = (
        f"MAD={mad:.4f} (seuil={seuil_benford_mad}) — χ²={chi2:.2f} — "
        f"n={total:,} montants\n  {distribution}"
    )
    return ("Benford (1er chiffre)", ok, detail, "INFO")


def _ctrl_ecritures_tardives(
    df: pd.DataFrame,
    date_cloture_ts: pd.Timestamp,
    **_,
) -> ResultatControle:
    """
    Signale les écritures validées après le 31/01/N+1 (si ValidDate disponible).

    Note : si toutes les ValidDate sont identiques (export groupé), le contrôle
    retourne un avertissement spécifique au lieu d'un faux positif massif.
    """
    valides = df["ValidDate"].str.strip().replace("", pd.NaT)
    dates_valid = pd.to_datetime(valides, format="%Y%m%d", errors="coerce")
    nb_disponibles = dates_valid.notna().sum()

    if nb_disponibles == 0:
        return ("Écritures tardives", True, "Champ ValidDate absent ou vide", "INFO")

    seuil = pd.Timestamp(year=date_cloture_ts.year + 1, month=1, day=31)

    # Détecter un export groupé (toutes les dates identiques)
    dates_uniques = dates_valid.dropna().unique()
    if len(dates_uniques) == 1:
        date_unique = pd.Timestamp(dates_uniques[0])
        if date_unique > seuil:
            return (
                "Écritures tardives",
                True,
                f"ValidDate unique = {date_unique.strftime('%d/%m/%Y')} "
                f"(export groupé après clôture — non signifiant)",
                "INFO",
            )

    tardives = int((dates_valid > seuil).sum())
    ok = bool(tardives == 0)
    detail = (
        f"{tardives} écriture(s) avec ValidDate > {seuil.strftime('%d/%m/%Y')} "
        f"(sur {nb_disponibles} ValidDate renseignées)"
        if tardives
        else f"Aucune écriture tardive (seuil : {seuil.strftime('%d/%m/%Y')})"
    )
    return ("Écritures tardives", ok, detail, "WARNING")
