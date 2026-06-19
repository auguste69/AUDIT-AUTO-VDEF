"""
Rapprochement des comptes changeant de numéro entre N-1 et N (P6).

Un compte peut changer de numéro entre deux exercices (ex: "512003" →
"5123001") : sans rapprochement, les deux numéros sont traités comme des
comptes distincts, créant des doublons et faussant les variations.

Algorithme en 3 phases :
1. detecter_orphelins      — comptes N-1 sans correspondance exacte en N,
                             et inversement ;
2. scorer_matching         — score composite pour chaque paire d'orphelins :
                             préfixe commun (40 %), similarité de libellé
                             (35 %), même cycle (15 %), même classification
                             bilan (10 %) ; seuil minimal pour PROPOSER ;
3. validation utilisateur  — OBLIGATOIRE : aucune fusion sans confirmation
                             explicite (CLI interactive, UI Streamlit ou
                             option --rapprochements-auto). Le seuil sert à
                             proposer, jamais à fusionner seul.

Ordonnancement (révision 09/06/2026) : le matching s'exécute AVANT
map_cycles (pipeline : parse → load_n1 → MATCHING → build → map_cycles).
Le cycle des comptes N n'est donc pas encore calculé au moment du scoring :
il est dérivé à la volée par résolution de préfixe PCG (logique de
cycle_mapper), sans modifier la balance.

Paramètres externalisés dans la section `rapprochements` du
mapping_pcg.yaml (seuil, poids, mots_vides).
"""

import logging
import re
from typing import Dict, List, Optional, Set, Tuple

from src.engine.cycle_mapper import _resoudre_compta, _resoudre_cycle
from src.models.rapprochement import Rapprochement

logger = logging.getLogger(__name__)

# Valeurs par défaut si la section rapprochements du YAML est incomplète
_SEUIL_DEFAUT = 0.5
_POIDS_DEFAUT = {"prefixe": 0.40, "libelle": 0.35,
                 "cycle": 0.15, "classification": 0.10}


# ---------------------------------------------------------------------------
# Phase 1 — Détection des orphelins
# ---------------------------------------------------------------------------

def detecter_orphelins(
    comptes_n: Dict[str, str],
    balance_n1: dict,
) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Détecte les comptes sans correspondance exacte entre N et N-1.

    Paramètres
    ----------
    comptes_n : dict {compte_num: libelle}
        Comptes de l'exercice N (extraits du FEC).
    balance_n1 : dict {compte_num: {"libelle": str, "solde_ke": float}}
        Balance N-1 (balance_n1_loader / balance_parser).

    Retourne
    --------
    tuple (orphelins_n1, orphelins_n)
        Deux dicts {compte_num: libelle} : comptes N-1 absents de N, puis
        comptes N absents de N-1. Vides → aucun rapprochement nécessaire.
    """
    nums_n = set(comptes_n)
    nums_n1 = set(balance_n1)
    orphelins_n1 = {
        num: str((balance_n1[num] or {}).get("libelle", ""))
        for num in sorted(nums_n1 - nums_n)
    }
    orphelins_n = {
        num: str(comptes_n[num]) for num in sorted(nums_n - nums_n1)
    }
    logger.info(
        "Rapprochements : %d orphelin(s) N-1, %d orphelin(s) N",
        len(orphelins_n1), len(orphelins_n),
    )
    return orphelins_n1, orphelins_n


# ---------------------------------------------------------------------------
# Phase 2 — Scoring
# ---------------------------------------------------------------------------

def _prefixe_commun(a: str, b: str) -> float:
    """Longueur du plus long préfixe commun ÷ longueur max des deux numéros."""
    if not a or not b:
        return 0.0
    longueur = 0
    for ca, cb in zip(a, b):
        if ca != cb:
            break
        longueur += 1
    return longueur / max(len(a), len(b))


def _tokens(libelle: str, mots_vides: Set[str]) -> Set[str]:
    """Tokens d'un libellé : minuscules, alphanumériques, hors mots vides."""
    return {
        t for t in re.split(r"[^0-9a-zà-ÿ]+", str(libelle).lower())
        if t and t not in mots_vides
    }


def _similarite_libelle(libelle_a: str, libelle_b: str,
                        mots_vides: Set[str]) -> float:
    """Ratio de tokens communs (insensible à la casse, hors mots vides)."""
    tokens_a = _tokens(libelle_a, mots_vides)
    tokens_b = _tokens(libelle_b, mots_vides)
    if not tokens_a or not tokens_b:
        return 0.0
    return len(tokens_a & tokens_b) / max(len(tokens_a), len(tokens_b))


def _granularite_compatible(compte_a: str, compte_b: str,
                            longueur_collectif: int) -> bool:
    """Vrai si les deux numéros sont de même granularité.

    Un compte collectif court (≤ longueur_collectif chiffres) et un compte de
    détail long ne représentent pas la même population : on interdit de les
    rapprocher (un total agrégé n'est pas un changement de numéro de compte).
    """
    a_collectif = len(str(compte_a)) <= longueur_collectif
    b_collectif = len(str(compte_b)) <= longueur_collectif
    return a_collectif == b_collectif


def _cycle_et_compta(compte_num: str, mapping_fm: Optional[dict],
                     pcg_config: dict) -> Tuple[str, str]:
    """Cycle et classification bilan d'un compte : mapping FM du client en
    priorité, sinon résolution de préfixe PCG (logique de cycle_mapper)."""
    if mapping_fm and compte_num in mapping_fm:
        info = mapping_fm[compte_num]
        cycle = str(info.get("cycle", ""))
        compta = str(info.get("compta_n", info.get("compta", "")))
        if cycle:
            return cycle, compta or _resoudre_compta(
                compte_num, pcg_config["passif_prefixes"])
    cycle = _resoudre_cycle(compte_num, pcg_config["prefixes"],
                            pcg_config["surcharges"]) or ""
    compta = _resoudre_compta(compte_num, pcg_config["passif_prefixes"])
    return cycle, compta


def scorer_matching(
    compte_n1: str, libelle_n1: str,
    compte_n: str, libelle_n: str,
    mapping_fm: Optional[dict],
    pcg_config: dict,
    config_rapprochements: Optional[dict] = None,
) -> Rapprochement:
    """Calcule le score composite d'une paire (orphelin N-1, orphelin N).

    Composantes (poids par défaut, surchargés par la section
    rapprochements.poids du YAML) :
    - préfixe commun (40 %), similarité de libellé (35 %),
    - même cycle (15 %), même classification bilan (10 %).

    Retourne
    --------
    Rapprochement
        Proposition scorée (le filtrage par seuil relève de
        proposer_rapprochements).
    """
    cfg = config_rapprochements or {}
    poids = {**_POIDS_DEFAUT, **(cfg.get("poids") or {})}
    mots_vides = {str(m).lower() for m in (cfg.get("mots_vides") or [])}

    score_prefixe = _prefixe_commun(compte_n1, compte_n)
    score_libelle = _similarite_libelle(libelle_n1, libelle_n, mots_vides)

    cycle_n1, compta_n1 = _cycle_et_compta(compte_n1, mapping_fm, pcg_config)
    cycle_n, compta_n = _cycle_et_compta(compte_n, None, pcg_config)
    meme_cycle = bool(cycle_n1) and cycle_n1 == cycle_n
    meme_classification = bool(compta_n1) and compta_n1 == compta_n

    score = (
        poids["prefixe"] * score_prefixe
        + poids["libelle"] * score_libelle
        + poids["cycle"] * (1.0 if meme_cycle else 0.0)
        + poids["classification"] * (1.0 if meme_classification else 0.0)
    )

    return Rapprochement(
        compte_n1=compte_n1, libelle_n1=libelle_n1,
        compte_n=compte_n, libelle_n=libelle_n,
        score=round(score, 4),
        score_prefixe=round(score_prefixe, 4),
        score_libelle=round(score_libelle, 4),
        meme_cycle=meme_cycle,
        meme_classification=meme_classification,
    )


def proposer_rapprochements(
    comptes_n: Dict[str, str],
    balance_n1: dict,
    mapping_fm: Optional[dict],
    pcg_config: dict,
    config_rapprochements: Optional[dict] = None,
) -> List[Rapprochement]:
    """Propose les rapprochements entre orphelins N-1 et N (score ≥ seuil).

    Affectation gloutonne un-pour-un : les paires sont triées par score
    décroissant, chaque compte (N-1 ou N) ne peut apparaître que dans une
    seule proposition.

    Paramètres
    ----------
    comptes_n : dict {compte_num: libelle}
        Comptes de l'exercice N (extraits du FEC).
    balance_n1 : dict
        Balance N-1 {compte_num: {"libelle", "solde_ke"}}.
    mapping_fm : dict ou None
        Mapping du FM existant (cycle/compta des comptes N-1).
    pcg_config : dict
        Config PCG (from_pcg_config) — résolution cycle/classification.
    config_rapprochements : dict, optionnel
        Section rapprochements du YAML (seuil, poids, mots_vides).
        Défaut : pcg_config["rapprochements"].

    Retourne
    --------
    list[Rapprochement]
        Propositions triées par score décroissant. AUCUNE fusion n'est
        appliquée ici : la validation utilisateur est obligatoire.
    """
    if config_rapprochements is None:
        config_rapprochements = pcg_config.get("rapprochements") or {}
    seuil = float(config_rapprochements.get("seuil", _SEUIL_DEFAUT))
    longueur_collectif = int(config_rapprochements.get("longueur_collectif", 3))

    orphelins_n1, orphelins_n = detecter_orphelins(comptes_n, balance_n1)
    if not orphelins_n1 or not orphelins_n:
        return []

    # Garde-fou granularité : on ne score jamais une paire collectif↔détail
    # (ex. À-NOUVEAU sur compte collectif "411" vs détail "41110000").
    paires = [
        scorer_matching(num_n1, lib_n1, num_n, lib_n,
                        mapping_fm, pcg_config, config_rapprochements)
        for num_n1, lib_n1 in orphelins_n1.items()
        for num_n, lib_n in orphelins_n.items()
        if _granularite_compatible(num_n1, num_n, longueur_collectif)
    ]
    paires = [p for p in paires if p.score >= seuil]
    paires.sort(key=lambda p: (-p.score, p.compte_n1, p.compte_n))

    # Affectation gloutonne un-pour-un
    utilises_n1: Set[str] = set()
    utilises_n: Set[str] = set()
    propositions: List[Rapprochement] = []
    for paire in paires:
        if paire.compte_n1 in utilises_n1 or paire.compte_n in utilises_n:
            continue
        propositions.append(paire)
        utilises_n1.add(paire.compte_n1)
        utilises_n.add(paire.compte_n)

    for p in propositions:
        logger.info(
            "Rapprochement proposé : %s (%s) → %s (%s) — score %.2f",
            p.compte_n1, p.libelle_n1, p.compte_n, p.libelle_n, p.score,
        )
    return propositions


# ---------------------------------------------------------------------------
# Phase 3 — Application (après validation utilisateur uniquement)
# ---------------------------------------------------------------------------

def appliquer_rapprochements(
    balance_n1: dict,
    rapprochements: List[Rapprochement],
    mapping_fm: Optional[dict] = None,
) -> Tuple[dict, Optional[dict]]:
    """Applique les rapprochements VALIDÉS : renumérote les comptes N-1.

    Pour chaque rapprochement, l'entrée balance_n1[compte_n1] est déplacée
    sous le numéro compte_n (le solde N-1 suit le compte renuméroté) ; le
    mapping FM est renuméroté de la même façon (le compte N hérite du
    cycle/classification historique du client — priorité absolue du FM).

    Les dicts reçus ne sont JAMAIS modifiés en place. Chaque fusion est
    loggée pour traçabilité.

    Paramètres
    ----------
    balance_n1 : dict
        Balance N-1 {compte_num: {"libelle", "solde_ke"}}.
    rapprochements : list[Rapprochement]
        Rapprochements CONFIRMÉS par l'utilisateur.
    mapping_fm : dict ou None
        Mapping du FM existant à renuméroter de la même façon.

    Retourne
    --------
    tuple (balance_n1, mapping_fm)
        Copies renumérotées (mapping_fm None si None en entrée).
    """
    balance = dict(balance_n1)
    mapping = dict(mapping_fm) if mapping_fm is not None else None

    for r in rapprochements:
        if r.compte_n1 not in balance:
            logger.warning(
                "Rapprochement %s → %s ignoré : compte N-1 absent de la "
                "balance", r.compte_n1, r.compte_n,
            )
            continue
        if r.compte_n in balance:
            logger.warning(
                "Rapprochement %s → %s ignoré : le compte N existe déjà "
                "dans la balance N-1 (pas un orphelin)",
                r.compte_n1, r.compte_n,
            )
            continue
        balance[r.compte_n] = balance.pop(r.compte_n1)
        if mapping is not None and r.compte_n1 in mapping:
            mapping[r.compte_n] = mapping.pop(r.compte_n1)
        logger.info(
            "Rapprochement appliqué : %s → %s (score %.2f) — solde N-1 "
            "%.1f K€ transféré",
            r.compte_n1, r.compte_n, r.score,
            float(balance[r.compte_n].get("solde_ke", 0.0)),
        )

    return balance, mapping
