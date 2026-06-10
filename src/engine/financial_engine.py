"""
Moteur de calcul des états financiers synthétiques (Bilan, Tréso, AACE).

Sépare le calcul métier de la présentation Excel : fm_writer.py consomme
les objets retournés ici (BilanSynthetique, TresoSynthetique, DataFrame
AACE) et ne fait plus que de l'écriture.

Toutes les listes de préfixes PCG proviennent de la section liasse_fiscale
du mapping_pcg.yaml (validée par liasse_fiscale_loader.load_liasse_fiscale).
Aucun préfixe n'est codé en dur : seuls les signes de présentation (× -1
pour le passif) et les conditions sur le signe du solde (">0" / "<0")
relèvent de la logique générique de ce module.

Les DataFrames reçus ne sont jamais modifiés en place.
"""

import logging
from typing import List, Tuple

import pandas as pd

from src.models.financial_statements import (
    BilanSynthetique,
    PosteComptable,
    TresoSynthetique,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers d'agrégation par préfixe (privés)
# ---------------------------------------------------------------------------

def _sommer_prefixes(balance: pd.DataFrame, prefixes: list,
                     colonne: str = "Solde_KE") -> float:
    """Somme les valeurs de colonne pour tous les comptes commençant par un des préfixes."""
    masque = balance["CompteNum"].apply(
        lambda n: any(str(n).startswith(p) for p in prefixes)
    )
    return float(balance.loc[masque, colonne].sum())


def _sommer_prefixes_si(balance: pd.DataFrame, prefixes: list,
                         colonne: str, signe_cond: str) -> float:
    """Somme avec filtre sur le signe de la valeur ('>0' ou '<0')."""
    masque = balance["CompteNum"].apply(
        lambda n: any(str(n).startswith(p) for p in prefixes)
    )
    if signe_cond == ">0":
        masque = masque & (balance[colonne] > 0)
    elif signe_cond == "<0":
        masque = masque & (balance[colonne] < 0)
    return float(balance.loc[masque, colonne].sum())


def _classer_comptes_bascule(balance: pd.DataFrame,
                             prefixes_bascule: List[str]) -> tuple:
    """
    Reclasse dynamiquement les comptes bascule (préfixes issus de
    liasse_fiscale.bilan.comptes_bascule.prefixes) selon le signe de leur
    Solde_KE / Solde_N1_KE.

    Retourne ((actif_n, actif_n1), (passif_n, passif_n1)).
    Convention passif : valeur positive (−Solde_KE pour les soldes créditeurs).
    """
    if "Solde_KE" not in balance.columns or "Solde_N1_KE" not in balance.columns:
        raise ValueError(
            "Bascule : colonnes 'Solde_KE' ou 'Solde_N1_KE' absentes de la balance — "
            "vérifier que balance_builder.build() a bien été appelé."
        )
    masque = balance["CompteNum"].apply(
        lambda n: any(str(n).startswith(p) for p in prefixes_bascule)
    )
    df_b = balance.loc[masque]

    actif_n   = round(float(df_b.loc[df_b["Solde_KE"]    > 0, "Solde_KE"].sum()),    3)
    actif_n1  = round(float(df_b.loc[df_b["Solde_N1_KE"] > 0, "Solde_N1_KE"].sum()), 3)
    passif_n  = round(-float(df_b.loc[df_b["Solde_KE"]    < 0, "Solde_KE"].sum()),   3)
    passif_n1 = round(-float(df_b.loc[df_b["Solde_N1_KE"] < 0, "Solde_N1_KE"].sum()), 3)

    logger.debug(
        "Bascule : actif N=%.1f K€, passif N=%.1f K€ (%d comptes bascule)",
        actif_n, passif_n, len(df_b),
    )
    return (actif_n, actif_n1), (passif_n, passif_n1)


# ---------------------------------------------------------------------------
# Bilan synthétique
# ---------------------------------------------------------------------------

def calculer_bilan(balance: pd.DataFrame,
                   liasse_config: dict) -> BilanSynthetique:
    """
    Calcule les postes actif et passif du bilan synthétique pour N et N-1.

    Les listes de préfixes proviennent de liasse_config["bilan"] (section
    liasse_fiscale du mapping_pcg.yaml). Les signes de présentation et les
    conditions sur le signe du solde restent dans ce code (logique générique).

    Convention :
    - Actif : valeurs nettes (brut + amort, les amort ayant des Solde_KE négatifs)
    - Passif : × -1 (comptes créditeurs → soldes négatifs → affichage positif)

    Garantie universelle : un poste résiduel catch-all absorbe les comptes
    non capturés pour assurer l'équilibre avec n'importe quel FEC.

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().

    Retourne
    --------
    BilanSynthetique
        Postes actif/passif + totaux (valeurs N et N-1 en K€).
    """
    def sp(prefixes, signe=1):
        return (
            round(_sommer_prefixes(balance, prefixes, "Solde_KE")    * signe, 3),
            round(_sommer_prefixes(balance, prefixes, "Solde_N1_KE") * signe, 3),
        )

    def sp_si(prefixes, condition, signe=1):
        return (
            round(_sommer_prefixes_si(balance, prefixes, "Solde_KE",    condition) * signe, 3),
            round(_sommer_prefixes_si(balance, prefixes, "Solde_N1_KE", condition) * signe, 3),
        )

    p_actif   = liasse_config["bilan"]["actif"]
    p_passif  = liasse_config["bilan"]["passif"]
    p_bascule = liasse_config["bilan"]["comptes_bascule"]

    # --- ACTIF ---
    # Les comptes d'amort/dép ont des Solde_KE négatifs — les additionner
    # aux comptes bruts donne directement la valeur nette.
    cap_non_appele = sp(p_actif["capital_souscrit_non_appele"])
    immo_incorp    = sp(p_actif["immo_incorp"])
    immo_corp      = sp(p_actif["immo_corp"])
    immo_fi        = sp(p_actif["immo_fi"])
    stocks         = sp(p_actif["stocks"])
    avances        = sp(p_actif["avances_acomptes_verses"])
    crean_cli      = sp(p_actif["creances_clients"])
    # Les préfixes bascule (425, 444, 451, 455, 456, 467, 478, 486) sont
    # absents des listes actif/passif — gérés par _classer_comptes_bascule
    autres_crean   = sp(p_actif["autres_creances"])
    vmp            = sp(p_actif["vmp"])
    dispo          = sp(p_actif["disponibilites"])

    # --- PASSIF (× -1 car comptes créditeurs) ---
    capital        = sp(p_passif["capital"],                  signe=-1)
    primes         = sp(p_passif["primes"],                   signe=-1)
    reserve_legale = sp(p_passif["reserve_legale"],           signe=-1)
    autres_res     = sp(p_passif["autres_reserves"],          signe=-1)
    report         = sp(p_passif["report_nouveau"],           signe=-1)
    resultat       = sp(p_passif["resultat_exercice"],        signe=-1)
    subventions    = sp(p_passif["subventions"],              signe=-1)
    prov_regl      = sp(p_passif["prov_reglementees"],        signe=-1)
    prov_risques_b = sp(p_passif["prov_risques"],             signe=-1)
    prov_charges_b = sp(p_passif["prov_charges"],             signe=-1)
    emprunts       = sp(p_passif["emprunts"],                 signe=-1)
    det_fourn_b    = sp(p_passif["dettes_fournisseurs"],      signe=-1)
    det_fisc_b     = sp(p_passif["dettes_fiscales_sociales"], signe=-1)
    autres_dettes  = sp(p_passif["autres_dettes"],            signe=-1)
    # Résultat en cours (classes 6 & 7 si exercice non clôturé / arrêté partiel)
    resultat_encours = sp(p_passif["resultat_en_cours"],      signe=-1)

    # --- Comptes bascule : reclassement dynamique selon le signe ---
    bascule_actif, bascule_passif = _classer_comptes_bascule(
        balance, p_bascule["prefixes"]
    )

    # Sous-totaux interco affichés séparément (display only — déjà inclus dans bascule_actif/passif)
    crean_interco           = sp_si(p_bascule["interco"], ">0")
    dettes_interco          = sp_si(p_bascule["interco"], "<0", signe=-1)
    bascule_reclasses_actif = (
        round(bascule_actif[0]  - crean_interco[0],  3),
        round(bascule_actif[1]  - crean_interco[1],  3),
    )
    bascule_reclasses_passif = (
        round(bascule_passif[0] - dettes_interco[0], 3),
        round(bascule_passif[1] - dettes_interco[1], 3),
    )

    # --- Totaux ---
    total_actif = (
        round(cap_non_appele[0] + immo_incorp[0] + immo_corp[0] + immo_fi[0]
              + stocks[0] + avances[0] + crean_cli[0] + autres_crean[0]
              + bascule_actif[0] + vmp[0] + dispo[0], 3),
        round(cap_non_appele[1] + immo_incorp[1] + immo_corp[1] + immo_fi[1]
              + stocks[1] + avances[1] + crean_cli[1] + autres_crean[1]
              + bascule_actif[1] + vmp[1] + dispo[1], 3),
    )
    total_passif = (
        round(capital[0] + primes[0] + reserve_legale[0] + autres_res[0]
              + report[0] + resultat[0] + subventions[0] + prov_regl[0]
              + prov_risques_b[0] + prov_charges_b[0] + emprunts[0]
              + det_fourn_b[0] + det_fisc_b[0] + autres_dettes[0]
              + bascule_passif[0] + resultat_encours[0], 3),
        round(capital[1] + primes[1] + reserve_legale[1] + autres_res[1]
              + report[1] + resultat[1] + subventions[1] + prov_regl[1]
              + prov_risques_b[1] + prov_charges_b[1] + emprunts[1]
              + det_fourn_b[1] + det_fisc_b[1] + autres_dettes[1]
              + bascule_passif[1] + resultat_encours[1], 3),
    )

    # Diagnostic : log l'écart résiduel pour traçabilité
    imbalance_n = round(total_actif[0] - total_passif[0], 3)
    if abs(imbalance_n) > 0.1:
        logger.warning(
            "Bilan : résiduel actif/passif N = %.3f K€ "
            "(comptes non capturés → vérifier mapping_pcg.yaml)", imbalance_n,
        )

    # Contrôle équilibre bilan
    for annee_lbl, idx in [("N", 0), ("N-1", 1)]:
        ecart = abs(total_actif[idx] - total_passif[idx])
        if ecart > 0.01:
            logger.warning(
                "Bilan : Total Actif ≠ Total Passif en %s "
                "(Actif=%.1f, Passif=%.1f, écart=%.3f K€)",
                annee_lbl, total_actif[idx], total_passif[idx], ecart,
            )

    valeurs: dict = dict(
        cap_non_appele=cap_non_appele,
        immo_incorp=immo_incorp, immo_corp=immo_corp, immo_fi=immo_fi,
        stocks=stocks, avances=avances, crean_cli=crean_cli,
        autres_crean=autres_crean,
        crean_interco=crean_interco,
        bascule_reclasses_actif=bascule_reclasses_actif,
        vmp=vmp, dispo=dispo,
        capital=capital, primes=primes, reserve_legale=reserve_legale,
        autres_res=autres_res, report=report, resultat=resultat,
        subventions=subventions, prov_regl=prov_regl,
        prov_risques_b=prov_risques_b, prov_charges_b=prov_charges_b,
        emprunts=emprunts, det_fourn_b=det_fourn_b, det_fisc_b=det_fisc_b,
        autres_dettes=autres_dettes,
        dettes_interco=dettes_interco,
        bascule_reclasses_passif=bascule_reclasses_passif,
        resultat_encours=resultat_encours,
    )
    return BilanSynthetique(
        postes={cle: PosteComptable(cle, v[0], v[1])
                for cle, v in valeurs.items()},
        total_actif=PosteComptable("total_actif", total_actif[0],
                                   total_actif[1]),
        total_passif=PosteComptable("total_passif", total_passif[0],
                                    total_passif[1]),
    )


# ---------------------------------------------------------------------------
# Tréso (BFR / FRNG / TN)
# ---------------------------------------------------------------------------

def calculer_treso(balance: pd.DataFrame,
                   liasse_config: dict) -> TresoSynthetique:
    """
    Calcule tous les postes BFR/FRNG/TN pour N (Solde_KE) et N-1 (Solde_N1_KE).

    Les listes de préfixes proviennent de liasse_config["treso"] (section
    liasse_fiscale du mapping_pcg.yaml). Les signes et les conditions sur le
    signe du solde restent dans ce code (logique générique).

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().

    Retourne
    --------
    TresoSynthetique
        Postes détaillés/sous-totaux + agrégats FRNG, BFR, TN.
    """
    def sp(prefixes, signe=1):
        """Somme simple × signe pour N et N-1."""
        return (
            round(_sommer_prefixes(balance, prefixes, "Solde_KE")    * signe, 3),
            round(_sommer_prefixes(balance, prefixes, "Solde_N1_KE") * signe, 3),
        )

    def sp_si(prefixes, cond, signe=1):
        """Somme conditionnelle × signe pour N et N-1."""
        return (
            round(_sommer_prefixes_si(balance, prefixes, "Solde_KE",    cond) * signe, 3),
            round(_sommer_prefixes_si(balance, prefixes, "Solde_N1_KE", cond) * signe, 3),
        )

    def add(*postes):
        """Additionne plusieurs tuples (val_n, val_n1)."""
        return (round(sum(p[0] for p in postes), 3),
                round(sum(p[1] for p in postes), 3))

    p_res   = liasse_config["treso"]["ressources_stables"]
    p_emp   = liasse_config["treso"]["emplois_stables"]
    p_ac    = liasse_config["treso"]["actif_circulant_exploitation"]
    p_pc    = liasse_config["treso"]["passif_circulant_exploitation"]
    p_treso = liasse_config["treso"]["tresorerie_directe"]

    # --- Ressources stables (× -1 car passif/créditeur) ---
    cap_propres  = sp(p_res["capitaux_propres"],       signe=-1)
    amort_dep    = sp(p_res["amort_dep"],              signe=-1)
    prov_risques = sp(p_res["prov_risques"],           signe=-1)
    dettes_mlt   = sp(p_res["dettes_financieres_mlt"], signe=-1)
    total_res    = add(cap_propres, amort_dep, prov_risques, dettes_mlt)

    # --- Emplois stables (positif car actif/débiteur) ---
    actif_immo = sp(p_emp["actif_immobilise_brut"])
    total_emp  = actif_immo

    # --- FRNG ---
    frng = (round(total_res[0] - total_emp[0], 3),
            round(total_res[1] - total_emp[1], 3))

    # --- Actif circulant d'exploitation ---
    stocks       = sp(p_ac["stocks"])
    crean_cli    = sp(p_ac["creances_clients"])
    autres_crean = sp_si(p_ac["autres_creances"], ">0")
    cca          = sp(p_ac["cca"])
    total_ac     = add(stocks, crean_cli, autres_crean, cca)

    # --- Passif circulant d'exploitation (× -1 pour affichage positif) ---
    det_fourn  = sp(p_pc["dettes_fournisseurs"], signe=-1)
    det_fisc   = sp_si(p_pc["dettes_fiscales_sociales"], "<0", signe=-1)
    autres_det = sp_si(p_pc["autres_dettes"],            "<0", signe=-1)
    pca        = sp(p_pc["pca"], signe=-1)
    total_pc   = add(det_fourn, det_fisc, autres_det, pca)

    # --- BFR ---
    bfr = (round(total_ac[0] - total_pc[0], 3),
           round(total_ac[1] - total_pc[1], 3))

    # --- TN ---
    tn = (round(frng[0] - bfr[0], 3),
          round(frng[1] - bfr[1], 3))

    # --- Vérification TN (trésorerie directe) ---
    treso_active   = sp_si(p_treso["tresorerie_active"], ">0")
    treso_pass_519 = sp(p_treso["tresorerie_passive_519"], signe=-1)
    treso_pass_512 = sp_si(p_treso["tresorerie_passive_512"], "<0", signe=-1)
    treso_passive  = (round(treso_pass_519[0] + treso_pass_512[0], 3),
                      round(treso_pass_519[1] + treso_pass_512[1], 3))
    tn_verif = (round(treso_active[0] - treso_passive[0], 3),
                round(treso_active[1] - treso_passive[1], 3))

    # Contrôle cohérence fondamentale FRNG = BFR + TN
    for annee_lbl, idx in [("N", 0), ("N-1", 1)]:
        ecart = abs(frng[idx] - (bfr[idx] + tn[idx]))
        if ecart > 0.01:
            logger.warning(
                "Tréso : FRNG ≠ BFR + TN en %s "
                "(FRNG=%.1f, BFR=%.1f, TN=%.1f, écart=%.3f K€)",
                annee_lbl, frng[idx], bfr[idx], tn[idx], ecart,
            )

    valeurs: dict = dict(
        cap_propres=cap_propres, amort_dep=amort_dep,
        prov_risques=prov_risques, dettes_mlt=dettes_mlt, total_res=total_res,
        actif_immo=actif_immo, total_emp=total_emp,
        stocks=stocks, crean_cli=crean_cli, autres_crean=autres_crean,
        cca=cca, total_ac=total_ac,
        det_fourn=det_fourn, det_fisc=det_fisc, autres_det=autres_det,
        pca=pca, total_pc=total_pc,
        treso_active=treso_active, treso_passive=treso_passive,
        tn_verif=tn_verif,
    )
    return TresoSynthetique(
        postes={cle: PosteComptable(cle, v[0], v[1])
                for cle, v in valeurs.items()},
        frng=PosteComptable("frng", frng[0], frng[1]),
        bfr=PosteComptable("bfr", bfr[0], bfr[1]),
        tn=PosteComptable("tn", tn[0], tn[1]),
    )


# ---------------------------------------------------------------------------
# AACE
# ---------------------------------------------------------------------------

def filtrer_aace(balance_mappee: pd.DataFrame,
                 liasse_config: dict) -> pd.DataFrame:
    """Retourne les lignes dont CompteNum commence par un préfixe AACE, triées.

    Les préfixes proviennent de liasse_config["aace"]["prefixes"] (section
    liasse_fiscale du mapping_pcg.yaml).

    Paramètres
    ----------
    balance_mappee : pd.DataFrame
        Balance mappée (cycle_mapper.map_cycles) — non modifiée en place.
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().

    Retourne
    --------
    pd.DataFrame
        Copie filtrée et triée par CompteNum (index réinitialisé).
    """
    prefixes = tuple(liasse_config["aace"]["prefixes"])
    masque = balance_mappee["CompteNum"].astype(str).apply(
        lambda n: n.startswith(prefixes)
    )
    return balance_mappee[masque].sort_values("CompteNum").reset_index(drop=True)
