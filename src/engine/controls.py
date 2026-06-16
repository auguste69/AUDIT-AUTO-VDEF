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
from typing import List, Optional, Tuple

import pandas as pd

# Import engine → engine uniquement (jamais src/writers/)
from src.engine.financial_engine import (
    calculer_actif_detaille,
    calculer_bilan,
    calculer_passif_detaille,
    calculer_pl_detaille,
    calculer_treso,
)

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
    balance_mappee: Optional[pd.DataFrame] = None,
    liasse_config: Optional[dict] = None,
    seuil_equilibre_bilan_ke: float = 1.0,
    bilan_non_bloquant: bool = False,
) -> List[ResultatControle]:
    """
    Exécute tous les contrôles d'intégrité sur le FEC.

    Si balance_mappee ET liasse_config sont fournis, les contrôles
    financiers supplémentaires sont exécutés (équilibre du bilan AC-1,
    cohérence du résultat, cohérence des états détaillés, cohérence du
    résultat net P&L). Sinon, seuls les 9 contrôles FEC historiques
    sont exécutés (rétrocompatibilité totale).

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
    balance_mappee : pd.DataFrame, optionnel
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE) — active
        les contrôles financiers AC-1 et cohérence du résultat.
    liasse_config : dict, optionnel
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_equilibre_bilan_ke : float
        Seuil de tolérance de l'écart Actif/Passif en K€ (défaut : 1.0 K€).
    bilan_non_bloquant : bool
        Si True, le contrôle AC-1 est rétrogradé en WARNING (jamais bloquant).

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

    # Contrôles financiers (AC-1 + cohérence résultat) — uniquement si la
    # balance mappée et la config liasse sont disponibles. Si l'un des deux
    # est None, le comportement historique (9 contrôles) est inchangé.
    if balance_mappee is not None and liasse_config is not None:
        resultats.extend(
            run_controles_financiers(
                balance_mappee,
                liasse_config,
                seuil_equilibre_bilan_ke=seuil_equilibre_bilan_ke,
                bilan_non_bloquant=bilan_non_bloquant,
            )
        )

    return resultats


def run_controles_financiers(
    balance_mappee: pd.DataFrame,
    liasse_config: dict,
    seuil_equilibre_bilan_ke: float = 1.0,
    bilan_non_bloquant: bool = False,
) -> List[ResultatControle]:
    """
    Exécute les contrôles financiers sur la balance mappée :
    AC-1 (équilibre du bilan) et cohérence du résultat.

    Appelable séparément de run_all() : dans le pipeline, les 9 contrôles
    FEC tournent avant la construction de la balance, alors que ces deux
    contrôles nécessitent la balance mappée (disponible après map_cycles).

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE).
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_equilibre_bilan_ke : float
        Seuil de tolérance de l'écart Actif/Passif en K€ (défaut : 1.0 K€).
    bilan_non_bloquant : bool
        Si True, AC-1 est rétrogradé en WARNING.

    Retourne
    --------
    list[tuple[str, bool, str, str]]
        Liste de (nom_controle, ok, detail, severity) — 5 éléments
        (AC-1, cohérence du résultat, cohérence des états détaillés,
        cohérence du résultat net P&L, cohérence Tréso).
    """
    resultats: List[ResultatControle] = []

    controles = [
        (_ctrl_bilan_equilibre, "WARNING" if bilan_non_bloquant else "BLOQUANT"),
        (_ctrl_coherence_resultat, "WARNING"),
        (_ctrl_coherence_actif_detaille, "WARNING"),
        (_ctrl_coherence_pl_resultat, "WARNING"),
        (_ctrl_coherence_treso, "WARNING"),
    ]
    for controle, severity_erreur in controles:
        try:
            res = controle(
                balance_mappee,
                liasse_config,
                seuil_equilibre_bilan_ke=seuil_equilibre_bilan_ke,
                bilan_non_bloquant=bilan_non_bloquant,
            )
            resultats.append(res)
            statut = "OK" if res[1] else "KO"
            logger.info("[%s] %s — %s", statut, res[0], res[2].split("\n")[0])
        except Exception as exc:  # noqa: BLE001
            logger.error("Erreur dans le contrôle %s : %s", controle.__name__, exc)
            resultats.append((
                controle.__name__,
                False,
                f"Erreur inattendue lors du contrôle financier : {exc}",
                severity_erreur,
            ))

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

    # Premier chiffre significatif. On retire en tête les zéros ET le point
    # décimal : sans cela, un montant < 0,10 € (ex. 0.05 → "0.0500000000")
    # garde un "0" après la virgule et serait écarté à tort, alors que son
    # premier chiffre significatif est 5.
    premiers: List[int] = []
    for v in vals:
        s = f"{abs(v):.10f}".lstrip("0.").replace(".", "")
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


# =============================================================================
# Contrôles financiers (sur la balance mappée)
# =============================================================================

def _ctrl_bilan_equilibre(
    balance_mappee: pd.DataFrame,
    liasse_config: dict,
    seuil_equilibre_bilan_ke: float = 1.0,
    bilan_non_bloquant: bool = False,
    **_,
) -> ResultatControle:
    """
    Contrôle AC-1 : équilibre du bilan (Total Actif = Total Passif).

    BLOQUANT uniquement sur l'écart N : |Total Actif N − Total Passif N|
    doit être ≤ seuil_equilibre_bilan_ke (défaut : 1.0 K€).

    L'écart N-1 est contrôlé en WARNING uniquement (jamais bloquant) :
    les données N-1 proviennent du FM historique du client et peuvent
    porter un écart non corrigeable.

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE).
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_equilibre_bilan_ke : float
        Seuil de tolérance de l'écart en K€ (défaut : 1.0 K€).
    bilan_non_bloquant : bool
        Si True, la sévérité du contrôle devient WARNING.
    """
    bilan = calculer_bilan(balance_mappee, liasse_config)
    actif_n, actif_n1 = bilan.total_actif.as_tuple()
    passif_n, passif_n1 = bilan.total_passif.as_tuple()

    ecart_n = abs(actif_n - passif_n)
    ecart_n1 = abs(actif_n1 - passif_n1)

    ok = bool(ecart_n <= seuil_equilibre_bilan_ke)
    severity = "WARNING" if bilan_non_bloquant else "BLOQUANT"

    detail = (
        f"Bilan N : Actif={actif_n:,.1f} K€ / Passif={passif_n:,.1f} K€ — "
        f"écart={ecart_n:.3f} K€ (seuil : {seuil_equilibre_bilan_ke:.1f} K€)"
    )
    if not ok:
        detail += (
            " — le bilan de l'exercice N est déséquilibré : vérifier le FEC "
            "et le mapping des comptes (mapping_pcg.yaml)."
        )

    if ecart_n1 > seuil_equilibre_bilan_ke:
        msg_n1 = (
            f"WARNING N-1 : écart Actif/Passif = {ecart_n1:.3f} K€ "
            f"(Actif={actif_n1:,.1f} K€ / Passif={passif_n1:,.1f} K€) — "
            f"non bloquant, données issues du FM historique du client."
        )
        detail += "\n  " + msg_n1
        logger.warning("Équilibre du bilan (AC-1) — %s", msg_n1)

    return ("Équilibre du bilan (AC-1)", ok, detail, severity)


def _ctrl_coherence_resultat(
    balance_mappee: pd.DataFrame,
    liasse_config: dict,
    seuil_coherence_resultat_ke: float = 1.0,
    **_,
) -> ResultatControle:
    """
    Contrôle de cohérence du résultat (WARNING, jamais bloquant).

    Vérifie que le résultat comptable déduit de la balance
    (− somme des soldes des classes 6 et 7) est cohérent avec le poste
    résultat du bilan (comptes 12x + résultat en cours classes 6/7),
    avec une tolérance de 1 K€.

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE).
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_coherence_resultat_ke : float
        Tolérance de l'écart en K€ (défaut : 1.0 K€).
    """
    # Résultat déduit de la balance : − (somme classes 6 et 7)
    masque_67 = balance_mappee["CompteNum"].astype(str).str.startswith(("6", "7"))
    resultat_balance = -float(balance_mappee.loc[masque_67, "Solde_KE"].sum())

    # Poste résultat du bilan : compte 12x + résultat en cours (classes 6/7)
    bilan = calculer_bilan(balance_mappee, liasse_config)
    resultat_bilan = (
        bilan.postes["resultat"].valeur_n
        + bilan.postes["resultat_encours"].valeur_n
    )

    ecart = abs(resultat_balance - resultat_bilan)
    ok = bool(ecart <= seuil_coherence_resultat_ke)

    detail = (
        f"Résultat déduit de la balance (− classes 6/7) = {resultat_balance:,.1f} K€ / "
        f"poste résultat du bilan (12x + résultat en cours) = {resultat_bilan:,.1f} K€ — "
        f"écart={ecart:.3f} K€ (tolérance : {seuil_coherence_resultat_ke:.1f} K€)"
    )
    if not ok:
        detail += (
            " — incohérence : le résultat de l'exercice ne correspond pas "
            "entre la balance et le bilan (résultat déjà affecté en 12x ou "
            "mapping incomplet)."
        )

    return ("Cohérence du résultat", ok, detail, "WARNING")


def _ctrl_coherence_actif_detaille(
    balance_mappee: pd.DataFrame,
    liasse_config: dict,
    seuil_coherence_detaille_ke: float = 1.0,
    **_,
) -> ResultatControle:
    """
    Contrôle de cohérence des états détaillés (WARNING, jamais bloquant).

    Vérifie que les totaux de l'Actif détaillé et du Passif détaillé
    (cerfa 2050/2051) coïncident avec les totaux du bilan synthétique,
    avec une tolérance de 1 K€ : un écart signale une partition incomplète
    des structures actif/passif détaillés du mapping_pcg.yaml.

    Si les structures sont absentes de la config (YAML antérieur au
    prompt 10), le contrôle est marqué non exécuté (INFO, ok=True).

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE).
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_coherence_detaille_ke : float
        Tolérance de l'écart en K€ (défaut : 1.0 K€).
    """
    nom = "Cohérence des états détaillés"
    if not (liasse_config.get("actif_detaille_structure") or {}).get("sections") \
            or not (liasse_config.get("passif_detaille_structure") or {}).get("sections"):
        return (
            nom, True,
            "Contrôle non exécuté : structures actif/passif détaillés "
            "absentes de la config PCG (mapping_pcg.yaml).",
            "INFO",
        )

    bilan = calculer_bilan(balance_mappee, liasse_config)
    actif = calculer_actif_detaille(balance_mappee, liasse_config)
    passif = calculer_passif_detaille(balance_mappee, liasse_config)

    ecart_actif = abs(actif.total.valeur_n - bilan.total_actif.valeur_n)
    ecart_passif = abs(passif.total.valeur_n - bilan.total_passif.valeur_n)
    ok = bool(ecart_actif <= seuil_coherence_detaille_ke
              and ecart_passif <= seuil_coherence_detaille_ke)

    detail = (
        f"Actif détaillé N = {actif.total.valeur_n:,.1f} K€ / bilan = "
        f"{bilan.total_actif.valeur_n:,.1f} K€ (écart={ecart_actif:.3f} K€) — "
        f"Passif détaillé N = {passif.total.valeur_n:,.1f} K€ / bilan = "
        f"{bilan.total_passif.valeur_n:,.1f} K€ (écart={ecart_passif:.3f} K€) "
        f"— tolérance : {seuil_coherence_detaille_ke:.1f} K€"
    )
    if not ok:
        detail += (
            " — incohérence : la partition des états détaillés ne couvre pas "
            "les mêmes comptes que le bilan synthétique (vérifier "
            "actif_detaille_structure / passif_detaille_structure dans "
            "mapping_pcg.yaml)."
        )

    return (nom, ok, detail, "WARNING")


def _ctrl_coherence_pl_resultat(
    balance_mappee: pd.DataFrame,
    liasse_config: dict,
    seuil_coherence_resultat_ke: float = 1.0,
    **_,
) -> ResultatControle:
    """
    Contrôle de cohérence du résultat net P&L (WARNING, jamais bloquant).

    Vérifie que le résultat net du P&L détaillé (cerfa 2052/2053) coïncide
    avec le poste résultat du bilan (comptes 12x + résultat en cours
    classes 6/7), avec une tolérance de 1 K€.

    Si les sections ebit/pl_detaille sont absentes de la config (YAML
    antérieur aux prompts 9-10), le contrôle est marqué non exécuté
    (INFO, ok=True).

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE).
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_coherence_resultat_ke : float
        Tolérance de l'écart en K€ (défaut : 1.0 K€).
    """
    nom = "Cohérence du résultat net (P&L)"
    if not liasse_config.get("ebit") or not liasse_config.get("pl_detaille"):
        return (
            nom, True,
            "Contrôle non exécuté : sections ebit/pl_detaille absentes de "
            "la config PCG (mapping_pcg.yaml).",
            "INFO",
        )

    pl = calculer_pl_detaille(balance_mappee, liasse_config)

    bilan = calculer_bilan(balance_mappee, liasse_config)
    resultat_bilan = (
        bilan.postes["resultat"].valeur_n
        + bilan.postes["resultat_encours"].valeur_n
    )

    ecart = abs(pl.resultat_net.valeur_n - resultat_bilan)
    ok = bool(ecart <= seuil_coherence_resultat_ke)

    detail = (
        f"Résultat net du P&L détaillé = {pl.resultat_net.valeur_n:,.1f} K€ / "
        f"poste résultat du bilan (12x + résultat en cours) = "
        f"{resultat_bilan:,.1f} K€ — écart={ecart:.3f} K€ "
        f"(tolérance : {seuil_coherence_resultat_ke:.1f} K€)"
    )
    if not ok:
        detail += (
            " — incohérence : le résultat net du compte de résultat ne "
            "correspond pas au résultat porté au bilan (résultat déjà "
            "affecté en 12x ou mapping incomplet)."
        )

    return (nom, ok, detail, "WARNING")


def _ctrl_coherence_treso(
    balance_mappee: pd.DataFrame,
    liasse_config: dict,
    seuil_coherence_treso_ke: float = 1.0,
    **_,
) -> ResultatControle:
    """
    Contrôle de cohérence de la Tréso (WARNING, jamais bloquant).

    Vérifie que la trésorerie nette issue de l'approche bilancielle
    (TN = FRNG − BFR) retombe sur la trésorerie directe (classe 5,
    partitionnée par signe), avec une tolérance de 1 K€ sur N.

    Un écart signifie que des comptes ne sont capturés par aucune rubrique
    Tréso (ou par plusieurs) — voir la section liasse_fiscale.treso du
    mapping_pcg.yaml. L'écart N-1 est signalé dans le détail sans faire
    échouer le contrôle (données issues du FM historique du client).

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (colonnes CompteNum, Solde_KE, Solde_N1_KE).
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().
    seuil_coherence_treso_ke : float
        Tolérance de l'écart en K€ (défaut : 1.0 K€).
    """
    treso = calculer_treso(balance_mappee, liasse_config)
    tn_n, tn_n1 = treso.tn.as_tuple()
    verif_n, verif_n1 = treso.postes["tn_verif"].as_tuple()

    ecart_n = abs(tn_n - verif_n)
    ecart_n1 = abs(tn_n1 - verif_n1)
    ok = bool(ecart_n <= seuil_coherence_treso_ke)

    detail = (
        f"TN (FRNG − BFR) = {tn_n:,.1f} K€ / trésorerie directe "
        f"(classe 5) = {verif_n:,.1f} K€ — écart={ecart_n:.3f} K€ "
        f"(tolérance : {seuil_coherence_treso_ke:.1f} K€)"
    )
    if not ok:
        detail += (
            " — incohérence : des comptes ne sont capturés par aucune "
            "rubrique Tréso (ou par plusieurs) — vérifier la section "
            "liasse_fiscale.treso du mapping_pcg.yaml."
        )
    if ecart_n1 > seuil_coherence_treso_ke:
        msg_n1 = (
            f"WARNING N-1 : écart TN/trésorerie directe = {ecart_n1:.3f} K€ "
            f"(TN={tn_n1:,.1f} K€ / vérification={verif_n1:,.1f} K€) — "
            f"non bloquant, données issues du FM historique du client."
        )
        detail += "\n  " + msg_n1
        logger.warning("Cohérence Tréso — %s", msg_n1)

    return ("Cohérence Tréso (TN vs trésorerie directe)", ok, detail,
            "WARNING")
