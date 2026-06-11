"""
Moteur de calcul des états financiers synthétiques (Bilan, Tréso, EBIT, AACE).

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
    ActifDetaille,
    BilanSynthetique,
    EbitSynthetique,
    LigneDetail,
    PassifDetaille,
    PlDetaille,
    PosteComptable,
    SectionDetail,
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

    Il n'y a PAS de poste résiduel : un compte capturé par aucun préfixe
    (ou par plusieurs) déséquilibre le bilan — l'écart est loggé ici et
    contrôlé par AC-1 (controls._ctrl_bilan_equilibre, bloquant sur N).

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
    # Disponibilités : 511/513-518/53 inconditionnels + 512 si débiteur
    # (512 créditeur = banque créditrice → reclassé en emprunts au passif)
    dispo_fixe     = sp(p_actif["disponibilites"])
    dispo_cond     = sp_si(p_actif.get("disponibilites_cond") or [], ">0")
    dispo          = (round(dispo_fixe[0] + dispo_cond[0], 3),
                      round(dispo_fixe[1] + dispo_cond[1], 3))

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
    # Emprunts : 16/17/517/519 inconditionnels + 512 si créditeur
    emprunts_fixe  = sp(p_passif["emprunts"],                 signe=-1)
    emprunts_cond  = sp_si(p_passif.get("emprunts_cond") or [], "<0",
                           signe=-1)
    emprunts       = (round(emprunts_fixe[0] + emprunts_cond[0], 3),
                      round(emprunts_fixe[1] + emprunts_cond[1], 3))
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

def _resume_comptes_non_captures_treso(balance: pd.DataFrame,
                                       treso_config: dict) -> str:
    """Liste les comptes ne matchant aucun préfixe des rubriques Tréso.

    Sert au diagnostic du contrôle de cohérence TN : un compte hors de
    toutes les rubriques fausse l'égalité TN = trésorerie directe.
    """
    prefixes: set = set()
    for postes in treso_config.values():
        for liste in postes.values():
            prefixes.update(str(p) for p in liste)
    tous = tuple(prefixes)
    masque = ~balance["CompteNum"].astype(str).str.startswith(tous)
    comptes = sorted(balance.loc[masque, "CompteNum"].astype(str))
    if not comptes:
        return "aucun"
    extrait = ", ".join(comptes[:10])
    suite = f"… (+{len(comptes) - 10})" if len(comptes) > 10 else ""
    return f"{len(comptes)} compte(s) : {extrait}{suite}"


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
    # Résultat de l'exercice en cours (classes 6/7 si non clôturé) :
    # un bénéfice est une ressource, une perte la diminue.
    resultat_enc = sp(p_res["resultat_en_cours"],      signe=-1)
    amort_dep    = sp(p_res["amort_dep"],              signe=-1)
    prov_risques = sp(p_res["prov_risques"],           signe=-1)
    dettes_mlt   = sp(p_res["dettes_financieres_mlt"], signe=-1)
    total_res    = add(cap_propres, resultat_enc, amort_dep, prov_risques,
                       dettes_mlt)

    # --- Emplois stables (positif car actif/débiteur) ---
    actif_immo = sp(p_emp["actif_immobilise_brut"])
    total_emp  = actif_immo

    # --- FRNG ---
    frng = (round(total_res[0] - total_emp[0], 3),
            round(total_res[1] - total_emp[1], 3))

    # --- Actif circulant d'exploitation (comptes débiteurs uniquement) ---
    stocks       = sp(p_ac["stocks"])
    crean_cli    = sp_si(p_ac["creances_clients"], ">0")
    autres_crean = sp_si(p_ac["autres_creances"],  ">0")
    cca          = sp_si(p_ac["cca"],              ">0")
    total_ac     = add(stocks, crean_cli, autres_crean, cca)

    # --- Passif circulant d'exploitation (comptes créditeurs, × -1) ---
    det_fourn  = sp_si(p_pc["dettes_fournisseurs"],      "<0", signe=-1)
    det_fisc   = sp_si(p_pc["dettes_fiscales_sociales"], "<0", signe=-1)
    autres_det = sp_si(p_pc["autres_dettes"],            "<0", signe=-1)
    pca        = sp_si(p_pc["pca"],                      "<0", signe=-1)
    total_pc   = add(det_fourn, det_fisc, autres_det, pca)

    # --- BFR ---
    bfr = (round(total_ac[0] - total_pc[0], 3),
           round(total_ac[1] - total_pc[1], 3))

    # --- TN ---
    tn = (round(frng[0] - bfr[0], 3),
          round(frng[1] - bfr[1], 3))

    # --- Vérification TN (trésorerie directe : classe 5 par signe) ---
    treso_active  = sp_si(p_treso["tresorerie_active"],  ">0")
    treso_passive = sp_si(p_treso["tresorerie_passive"], "<0", signe=-1)
    tn_verif = (round(treso_active[0] - treso_passive[0], 3),
                round(treso_active[1] - treso_passive[1], 3))

    # Contrôle de cohérence : la TN issue de FRNG − BFR doit retomber sur
    # la trésorerie directe (classe 5). Un écart signifie que des comptes
    # ne sont capturés par aucune rubrique (ou par plusieurs).
    for annee_lbl, idx in [("N", 0), ("N-1", 1)]:
        ecart = abs(tn[idx] - tn_verif[idx])
        if ecart > 0.5:
            logger.warning(
                "Tréso : TN (FRNG − BFR) ≠ trésorerie directe en %s "
                "(TN=%.1f, vérification=%.1f, écart=%.3f K€) — comptes non "
                "capturés : %s",
                annee_lbl, tn[idx], tn_verif[idx], ecart,
                _resume_comptes_non_captures_treso(balance,
                                                   liasse_config["treso"]),
            )

    valeurs: dict = dict(
        cap_propres=cap_propres, resultat_enc=resultat_enc,
        amort_dep=amort_dep,
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
# EBIT (résultat d'exploitation)
# ---------------------------------------------------------------------------

# Clés de postes EBIT, dans l'ordre de présentation du cerfa 2052.
_POSTES_EBIT_PRODUITS: List[str] = [
    "ventes_marchandises", "production_vendue", "production_stockee",
    "production_immobilisee", "subventions_exploitation",
    "reprises_transferts", "autres_produits",
]
_POSTES_EBIT_CHARGES: List[str] = [
    "achats_marchandises", "variation_stocks_marchandises",
    "achats_matieres_premieres", "variation_stocks_matieres",
    "autres_charges_externes", "impots_taxes", "salaires_traitements",
    "charges_sociales", "dotations_amortissements",
    "dotations_dep_immobilisations", "dotations_dep_actif_circulant",
    "dotations_provisions", "autres_charges",
]


def _sommer_poste_avec_exclusions(balance: pd.DataFrame, section: dict,
                                  cle: str, colonne: str,
                                  chemin: str) -> float:
    """Somme un poste EBIT : préfixes inclus moins préfixes exclus.

    Les exclusions sont portées par la clé "<cle>_exclusions" de la section
    (sous-préfixes plus spécifiques, ex. "709 sauf 7097") — voir le
    commentaire de la section liasse_fiscale.ebit du mapping_pcg.yaml.
    """
    if cle not in section:
        raise ValueError(
            f"liasse_fiscale.ebit.{chemin} : clé obligatoire '{cle}' "
            f"absente. Clés disponibles : {sorted(section.keys())}."
        )
    prefixes = [str(p) for p in (section[cle] or [])]
    exclusions = [str(p) for p in (section.get(f"{cle}_exclusions") or [])]
    total = _sommer_prefixes(balance, prefixes, colonne) if prefixes else 0.0
    if exclusions:
        total -= _sommer_prefixes(balance, exclusions, colonne)
    return total


def calculer_ebit(balance: pd.DataFrame,
                  liasse_config: dict) -> EbitSynthetique:
    """
    Calcule le résultat d'exploitation (EBIT) pour N et N-1.

    Les listes de préfixes proviennent de liasse_config["ebit"] (section
    liasse_fiscale du mapping_pcg.yaml). Chaque poste peut porter une clé
    "<poste>_exclusions" listant les sous-préfixes à déduire (ex. 709 sauf
    7097). Seuls les signes de présentation restent dans ce code :

    - produits (classe 7, créditeurs → Solde_KE négatif) : × -1 → positif ;
    - charges (classe 6, débitrices → Solde_KE positif) : telles quelles ;
    - EBIT = total produits − total charges.

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    liasse_config : dict
        Dict contenant la clé "ebit" (section liasse_fiscale du YAML,
        sous-sections produits_exploitation et charges_exploitation).

    Retourne
    --------
    EbitSynthetique
        Postes détaillés + agrégats CA, total produits, total charges, EBIT.

    Lève
    ----
    ValueError
        Si la section ebit (ou une de ses clés obligatoires) est absente.
    """
    ebit_cfg = liasse_config.get("ebit") or {}
    p_prod = ebit_cfg.get("produits_exploitation")
    p_chrg = ebit_cfg.get("charges_exploitation")
    if not p_prod or not p_chrg:
        raise ValueError(
            "liasse_fiscale.ebit : sous-sections 'produits_exploitation' et "
            "'charges_exploitation' requises — vérifier mapping_pcg.yaml."
        )

    def sp(section: dict, cle: str, chemin: str, signe: int) -> tuple:
        """Somme un poste (avec exclusions) × signe, pour N et N-1."""
        return (
            round(_sommer_poste_avec_exclusions(
                balance, section, cle, "Solde_KE", chemin) * signe, 3),
            round(_sommer_poste_avec_exclusions(
                balance, section, cle, "Solde_N1_KE", chemin) * signe, 3),
        )

    # Produits × -1 (créditeurs → affichage positif), charges telles quelles.
    produits = {cle: sp(p_prod, cle, "produits_exploitation", -1)
                for cle in _POSTES_EBIT_PRODUITS}
    charges = {cle: sp(p_chrg, cle, "charges_exploitation", 1)
               for cle in _POSTES_EBIT_CHARGES}

    def add(valeurs: list) -> tuple:
        return (round(sum(v[0] for v in valeurs), 3),
                round(sum(v[1] for v in valeurs), 3))

    ca = add([produits["ventes_marchandises"],
              produits["production_vendue"]])
    total_produits = add(list(produits.values()))
    total_charges = add(list(charges.values()))
    ebit = (round(total_produits[0] - total_charges[0], 3),
            round(total_produits[1] - total_charges[1], 3))

    logger.info(
        "EBIT : CA N=%.1f K€, produits N=%.1f, charges N=%.1f, EBIT N=%.1f K€",
        ca[0], total_produits[0], total_charges[0], ebit[0],
    )

    postes = {**produits, **charges}
    return EbitSynthetique(
        postes={cle: PosteComptable(cle, v[0], v[1])
                for cle, v in postes.items()},
        ca=PosteComptable("ca", ca[0], ca[1]),
        total_produits=PosteComptable("total_produits", total_produits[0],
                                      total_produits[1]),
        total_charges=PosteComptable("total_charges", total_charges[0],
                                     total_charges[1]),
        ebit=PosteComptable("ebit", ebit[0], ebit[1]),
    )


# ---------------------------------------------------------------------------
# États détaillés (Actif détaillé / Passif détaillé — cerfa 2050 / 2051)
# ---------------------------------------------------------------------------

def _sommer_poste_structure(balance: pd.DataFrame, poste: dict,
                            colonne: str) -> float:
    """Somme un poste d'état détaillé : prefixes + prefixes_cond filtrés.

    Le poste peut combiner une composante inconditionnelle ("prefixes") et
    une composante conditionnée sur le signe du solde ("prefixes_cond" +
    "condition" ">0"/"<0") — reflet des comptes bascule du bilan synthétique.
    """
    total = 0.0
    prefixes = [str(p) for p in (poste.get("prefixes") or [])]
    if prefixes:
        total += _sommer_prefixes(balance, prefixes, colonne)
    prefixes_cond = [str(p) for p in (poste.get("prefixes_cond") or [])]
    if prefixes_cond:
        condition = poste.get("condition")
        if condition not in (">0", "<0"):
            raise ValueError(
                f"État détaillé : poste '{poste.get('cle')}' porte des "
                f"prefixes_cond sans condition valide (attendu '>0' ou "
                f"'<0', trouvé {condition!r})."
            )
        total += _sommer_prefixes_si(balance, prefixes_cond, colonne, condition)
    return total


def _calculer_etat_detaille(balance: pd.DataFrame, structure: dict,
                            signe: int, nom_etat: str) -> tuple:
    """Calcule un état détaillé (sections, agrégats, total) depuis sa structure.

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    structure : dict
        Structure de présentation (actif_detaille_structure ou
        passif_detaille_structure du mapping_pcg.yaml).
    signe : int
        +1 (actif) ou -1 (passif, comptes créditeurs affichés en positif).
    nom_etat : str
        Nom de l'état pour les messages d'erreur ("Actif détaillé", …).

    Retourne
    --------
    tuple (sections, agregats, total)
        Listes de SectionDetail / LigneDetail et LigneDetail total.
    """
    if not structure or not structure.get("sections"):
        raise ValueError(
            f"{nom_etat} : structure absente ou sans sections — vérifier la "
            f"section correspondante de mapping_pcg.yaml."
        )

    sections: List[SectionDetail] = []
    for sec in structure["sections"]:
        postes: List[LigneDetail] = []
        for p in sec.get("postes") or []:
            valeur_n = round(
                _sommer_poste_structure(balance, p, "Solde_KE") * signe, 3)
            valeur_n1 = round(
                _sommer_poste_structure(balance, p, "Solde_N1_KE") * signe, 3)
            postes.append(LigneDetail(p["cle"], p["libelle"],
                                      valeur_n, valeur_n1))
        libelle_st = sec.get("sous_total")
        sous_total = None
        if libelle_st:
            sous_total = LigneDetail(
                sec["cle"], libelle_st,
                round(sum(p.valeur_n for p in postes), 3),
                round(sum(p.valeur_n1 for p in postes), 3),
            )
        sections.append(SectionDetail(sec["cle"], libelle_st, postes,
                                      sous_total))

    par_cle = {s.cle: s for s in sections}

    def _somme_sections(cles: list) -> Tuple[float, float]:
        inconnues = [c for c in cles if c not in par_cle]
        if inconnues:
            raise ValueError(
                f"{nom_etat} : agrégat référençant des sections inconnues "
                f"{inconnues} — sections disponibles : {sorted(par_cle)}."
            )
        return (
            round(sum(p.valeur_n for c in cles
                      for p in par_cle[c].postes), 3),
            round(sum(p.valeur_n1 for c in cles
                      for p in par_cle[c].postes), 3),
        )

    agregats: List[LigneDetail] = []
    for ag in structure.get("agregats") or []:
        valeur_n, valeur_n1 = _somme_sections(ag["sections"])
        agregats.append(LigneDetail(ag["cle"], ag["libelle"],
                                    valeur_n, valeur_n1,
                                    apres_section=ag.get("apres_section")))

    total_n, total_n1 = _somme_sections([s.cle for s in sections])
    libelle_total = (structure.get("total") or {}).get("libelle", "TOTAL")
    total = LigneDetail("total", libelle_total, total_n, total_n1)

    logger.info("%s : total N=%.1f K€, N-1=%.1f K€", nom_etat, total_n,
                total_n1)
    return sections, agregats, total


def calculer_actif_detaille(balance: pd.DataFrame,
                            liasse_config: dict) -> ActifDetaille:
    """Calcule l'Actif détaillé (cerfa 2050) pour N et N-1.

    La structure de présentation provient de
    liasse_config["actif_detaille_structure"] (racine du mapping_pcg.yaml,
    exposée par load_liasse_fiscale). Les partitions reprennent exactement
    les listes du bilan synthétique raffinées par sous-poste : le total est
    strictement égal au Total Actif du bilan.

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().

    Retourne
    --------
    ActifDetaille
        Sections, agrégats et total (valeurs N et N-1 en K€).
    """
    sections, agregats, total = _calculer_etat_detaille(
        balance, liasse_config.get("actif_detaille_structure"),
        signe=1, nom_etat="Actif détaillé",
    )
    return ActifDetaille(sections=sections, agregats=agregats, total=total)


def calculer_passif_detaille(balance: pd.DataFrame,
                             liasse_config: dict) -> PassifDetaille:
    """Calcule le Passif détaillé (cerfa 2051) pour N et N-1.

    Même principe que calculer_actif_detaille, avec signe -1 (comptes
    créditeurs affichés en positif). Le total est strictement égal au
    Total Passif du bilan synthétique.

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    liasse_config : dict
        Section liasse_fiscale chargée par load_liasse_fiscale().

    Retourne
    --------
    PassifDetaille
        Sections, agrégats et total (valeurs N et N-1 en K€).
    """
    sections, agregats, total = _calculer_etat_detaille(
        balance, liasse_config.get("passif_detaille_structure"),
        signe=-1, nom_etat="Passif détaillé",
    )
    return PassifDetaille(sections=sections, agregats=agregats, total=total)


# ---------------------------------------------------------------------------
# P&L détaillé (cerfa 2052 / 2053)
# ---------------------------------------------------------------------------

# Postes hors exploitation, clés obligatoires de liasse_fiscale.pl_detaille
_POSTES_PL: List[str] = [
    "produits_financiers", "charges_financieres",
    "produits_exceptionnels", "charges_exceptionnelles",
    "participation_salaries", "impots_benefices", "quotes_parts",
]


def calculer_pl_detaille(balance: pd.DataFrame,
                         liasse_config: dict) -> PlDetaille:
    """Calcule le compte de résultat détaillé (cerfa 2052/2053) pour N et N-1.

    Les postes hors exploitation proviennent de
    liasse_config["pl_detaille"]. La partie exploitation est calculée en
    RÉSIDUEL : produits d'exploitation = −(classe 7) − postes hors
    exploitation de classe 7, idem pour les charges (classe 6). Le résultat
    net est donc strictement égal à −(somme des classes 6 et 7), quel que
    soit le plan de comptes du client.

    Convention de signe : tous les postes valent −solde (produits positifs,
    charges négatives) ; les résultats sont des sommes algébriques.

    Paramètres
    ----------
    balance : pd.DataFrame
        Balance avec colonnes CompteNum, Solde_KE, Solde_N1_KE.
    liasse_config : dict
        Dict contenant les clés "ebit" et "pl_detaille" (sections
        liasse_fiscale du YAML).

    Retourne
    --------
    PlDetaille
        Détail exploitation (EBIT) + postes hors exploitation + résultats.

    Lève
    ----
    ValueError
        Si la section pl_detaille (ou une clé obligatoire) est absente.
    """
    pl_cfg = liasse_config.get("pl_detaille") or {}
    manquantes = [cle for cle in _POSTES_PL if cle not in pl_cfg]
    if manquantes:
        raise ValueError(
            f"liasse_fiscale.pl_detaille : clé(s) obligatoire(s) "
            f"{manquantes} absente(s) — vérifier mapping_pcg.yaml."
        )

    ebit = calculer_ebit(balance, liasse_config)

    def sneg(prefixes: list, colonne: str) -> float:
        """−(somme des soldes) : produits positifs, charges négatives."""
        prefixes = [str(p) for p in prefixes]
        return -_sommer_prefixes(balance, prefixes, colonne) if prefixes else 0.0

    postes = {
        cle: (round(sneg(pl_cfg[cle], "Solde_KE"), 3),
              round(sneg(pl_cfg[cle], "Solde_N1_KE"), 3))
        for cle in _POSTES_PL
    }

    def _exploitation(colonne: str) -> Tuple[float, float]:
        """Produits et charges d'exploitation en résiduel des classes 7 / 6."""
        produits = -_sommer_prefixes(balance, ["7"], colonne)
        charges = -_sommer_prefixes(balance, ["6"], colonne)
        for cle in _POSTES_PL:
            for prefixe in pl_cfg[cle]:
                prefixe = str(prefixe)
                valeur = -_sommer_prefixes(balance, [prefixe], colonne)
                if prefixe.startswith("7"):
                    produits -= valeur
                else:
                    charges -= valeur
        return round(produits, 3), round(charges, 3)

    prod_n, chrg_n = _exploitation("Solde_KE")
    prod_n1, chrg_n1 = _exploitation("Solde_N1_KE")
    postes["produits_exploitation"] = (prod_n, prod_n1)
    postes["charges_exploitation"] = (chrg_n, chrg_n1)

    # Résiduels vs EBIT synthétique : comptes d'exploitation non capturés
    # par les postes EBIT (ex. frais accessoires d'achats par convention
    # cabinet) — affichés en lignes "divers" dans le writer.
    postes["produits_expl_divers"] = (
        round(prod_n - ebit.total_produits.valeur_n, 3),
        round(prod_n1 - ebit.total_produits.valeur_n1, 3),
    )
    postes["charges_expl_divers"] = (
        round(chrg_n + ebit.total_charges.valeur_n, 3),
        round(chrg_n1 + ebit.total_charges.valeur_n1, 3),
    )

    def add(*cles: str) -> Tuple[float, float]:
        return (round(sum(postes[c][0] for c in cles), 3),
                round(sum(postes[c][1] for c in cles), 3))

    r_expl = add("produits_exploitation", "charges_exploitation")
    r_fin = add("produits_financiers", "charges_financieres")
    r_courant = (round(r_expl[0] + postes["quotes_parts"][0] + r_fin[0], 3),
                 round(r_expl[1] + postes["quotes_parts"][1] + r_fin[1], 3))
    r_exc = add("produits_exceptionnels", "charges_exceptionnelles")
    r_net = (
        round(r_courant[0] + r_exc[0] + postes["participation_salaries"][0]
              + postes["impots_benefices"][0], 3),
        round(r_courant[1] + r_exc[1] + postes["participation_salaries"][1]
              + postes["impots_benefices"][1], 3),
    )

    logger.info(
        "P&L détaillé : résultat exploitation N=%.1f, financier N=%.1f, "
        "exceptionnel N=%.1f, net N=%.1f K€",
        r_expl[0], r_fin[0], r_exc[0], r_net[0],
    )

    return PlDetaille(
        ebit=ebit,
        postes={cle: PosteComptable(cle, v[0], v[1])
                for cle, v in postes.items()},
        resultat_exploitation=PosteComptable("resultat_exploitation",
                                             r_expl[0], r_expl[1]),
        resultat_financier=PosteComptable("resultat_financier",
                                          r_fin[0], r_fin[1]),
        resultat_courant=PosteComptable("resultat_courant",
                                        r_courant[0], r_courant[1]),
        resultat_exceptionnel=PosteComptable("resultat_exceptionnel",
                                             r_exc[0], r_exc[1]),
        resultat_net=PosteComptable("resultat_net", r_net[0], r_net[1]),
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
