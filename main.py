"""
CLI du pipeline d'audit — point d'entrée en ligne de commande.

Usage :
    python3 main.py <fec> --client NOM --date-cloture JJ/MM/AAAA [options]

Exemples :
    # Minimum (FEC seul, mapping 100% PCG automatique)
    python3 main.py data/client_FEC.txt --client ACME --date-cloture 31/12/2025

    # Avec FM N-1 et templates
    python3 main.py data/client_FEC.txt \\
        --client ACME --date-cloture 31/12/2025 \\
        --n1-fm data/FM_ACME_N-1.xlsx \\
        --templates data/templates/ \\
        --output output/
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# Chemin par défaut de la config PCG (relatif à ce fichier)
_PCG_DEFAULT = Path(__file__).parent / "src" / "config" / "mapping_pcg.yaml"


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Pipeline d'audit : FEC → FM + feuilles de travail",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("fec", help="Chemin vers le FEC .txt")
    p.add_argument("--client",       required=True, help="Nom du client (ex: ACME)")
    p.add_argument("--date-cloture", required=True,
                   help="Date de clôture JJ/MM/AAAA (ex: 31/12/2025)")
    p.add_argument("--n1-fm",        default=None,
                   help="FM existant N-1 pour le mapping et les soldes N-1 (optionnel)")
    p.add_argument("--templates",    default=None,
                   help="Dossier contenant les templates .xlsx (optionnel)")
    p.add_argument("--output",       default="output",
                   help="Dossier de sortie (défaut : output/)")
    p.add_argument("--pcg-config",   default=str(_PCG_DEFAULT),
                   help="Chemin vers mapping_pcg.yaml")
    p.add_argument("--no-templates", action="store_true",
                   help="Désactive la génération des feuilles de travail")
    p.add_argument("--bilan-non-bloquant", action="store_true",
                   help="Le contrôle d'équilibre du bilan (AC-1) devient un "
                        "WARNING au lieu d'être bloquant")
    p.add_argument("--rapprochements-auto", action="store_true",
                   help="Applique automatiquement TOUS les rapprochements de "
                        "comptes N/N-1 proposés (score ≥ seuil), sans "
                        "validation interactive. Par défaut, chaque "
                        "rapprochement est validé interactivement.")
    p.add_argument("-v", "--verbose", action="store_true",
                   help="Mode verbeux (DEBUG)")
    return p.parse_args()


def _valider_rapprochements(
    propositions: list,
    rapprochements_auto: bool = False,
    rapprochements_valides: Optional[list] = None,
    rapprochements_interactif: bool = False,
) -> list:
    """
    Sélectionne les rapprochements à appliquer parmi les propositions.

    Règle absolue (P6) : AUCUNE fusion sans confirmation explicite. Le
    seuil de score sert à proposer, jamais à fusionner seul. Les sources
    de confirmation acceptées, par priorité :
    1. rapprochements_auto=True   — l'utilisateur a explicitement demandé
       l'application de toutes les propositions (--rapprochements-auto) ;
    2. rapprochements_valides     — paires (compte_n1, compte_n) déjà
       validées en amont (UI Streamlit) ;
    3. rapprochements_interactif  — validation CLI ligne par ligne via
       input() (o/n, t = tout valider, a = tout ignorer). Sans TTY,
       aucune fusion n'est appliquée (avec avertissement).
    """
    if rapprochements_auto:
        logger.info(
            "Rapprochements : application automatique des %d proposition(s) "
            "(--rapprochements-auto)", len(propositions),
        )
        return list(propositions)

    if rapprochements_valides is not None:
        paires = {tuple(p) for p in rapprochements_valides}
        return [r for r in propositions
                if (r.compte_n1, r.compte_n) in paires]

    if not rapprochements_interactif:
        logger.warning(
            "Rapprochements : %d proposition(s) NON appliquée(s) — aucune "
            "validation fournie (utiliser --rapprochements-auto, la "
            "validation interactive CLI ou l'interface Streamlit).",
            len(propositions),
        )
        return []

    if not sys.stdin.isatty():
        logger.warning(
            "Rapprochements : validation interactive impossible (stdin "
            "n'est pas un terminal) — %d proposition(s) ignorée(s). "
            "Utiliser --rapprochements-auto pour les appliquer.",
            len(propositions),
        )
        return []

    valides: list = []
    tout_valider = False
    logger.info("Validation interactive de %d rapprochement(s) proposé(s) :",
                len(propositions))
    for i, r in enumerate(propositions, start=1):
        if tout_valider:
            valides.append(r)
            continue
        reponse = input(
            f"[{i}/{len(propositions)}] Fusionner {r.compte_n1} "
            f"({r.libelle_n1}) → {r.compte_n} ({r.libelle_n}) — "
            f"score {r.score:.2f} ? [o/N/t=tout valider/a=tout ignorer] "
        ).strip().lower()
        if reponse == "a":
            break
        if reponse == "t":
            tout_valider = True
            valides.append(r)
        elif reponse == "o":
            valides.append(r)
    return valides


def run_pipeline(
    fec_path: str,
    client: str,
    date_cloture: str,
    n1_fm: Optional[str] = None,
    templates_dir: Optional[str] = None,
    output_dir: str = "output",
    pcg_config_path: Optional[str] = None,
    bilan_non_bloquant: bool = False,
    rapprochements_auto: bool = False,
    rapprochements_valides: Optional[list] = None,
    rapprochements_interactif: bool = False,
) -> dict:
    """
    Exécute le pipeline complet et retourne un dict de résultats.

    Compatible CLI et Streamlit — aucun print(), tout passe par logging
    (sauf la validation interactive des rapprochements, qui passe par
    input() quand rapprochements_interactif=True et stdin est un TTY).

    Paramètres notables
    -------------------
    bilan_non_bloquant : bool
        Si True, le contrôle d'équilibre du bilan AC-1 est rétrogradé en
        WARNING : le pipeline continue même si le bilan N est déséquilibré.
    rapprochements_auto : bool
        Si True, TOUS les rapprochements de comptes N/N-1 proposés sont
        appliqués sans validation interactive (confirmation explicite
        donnée par l'option elle-même).
    rapprochements_valides : list, optionnel
        Liste de paires (compte_n1, compte_n) déjà validées par
        l'utilisateur (UI Streamlit) : seules les propositions
        correspondantes sont appliquées.
    rapprochements_interactif : bool
        Si True (CLI), chaque proposition est validée via input().
        Sans TTY, aucune fusion n'est appliquée (validation obligatoire).

    Retourne
    --------
    dict avec les clés :
        fec_lignes, nb_comptes, controles, balance_mappee, fm_path,
        rapprochements_proposes, rapprochements_appliques,
        zip_path (optionnel)
    """
    from src.parsers.fec_parser import parse
    from src.parsers.balance_n1_loader import load_balance_n1
    from src.parsers.mapping_parser import from_pcg_config
    from src.engine.balance_builder import build
    from src.engine.controls import run_all, run_controles_financiers
    from src.engine.cycle_mapper import map_cycles
    from src.engine.liasse_fiscale_loader import load_liasse_fiscale
    from src.writers.excel_writer import write as write_travail
    from src.writers.fm_writer import write as write_fm
    from src.writers.template_writer import write as write_tpl

    if pcg_config_path is None:
        pcg_config_path = str(_PCG_DEFAULT)

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)

    resultats: dict = {}

    # ------------------------------------------------------------------
    # 1. Config PCG
    # ------------------------------------------------------------------
    logger.info("Chargement de la config PCG : %s", pcg_config_path)
    pcg = from_pcg_config(pcg_config_path)

    # ------------------------------------------------------------------
    # 2. Lecture du FEC
    # ------------------------------------------------------------------
    logger.info("=== ÉTAPE 1/5 — Lecture du FEC ===")
    df_fec = parse(fec_path)
    resultats["fec_lignes"] = len(df_fec)
    logger.info("FEC chargé : %d écritures", len(df_fec))

    # ------------------------------------------------------------------
    # 3. Contrôles d'intégrité
    # ------------------------------------------------------------------
    logger.info("=== ÉTAPE 2/5 — Contrôles d'intégrité ===")
    controles = run_all(df_fec, date_cloture)
    resultats["controles"] = controles

    bloquants_ko = [c for c in controles if c[3] == "BLOQUANT" and not c[1]]
    for nom, ok, detail, sev in controles:
        symbole = "✓" if ok else "✗"
        logger.info("  [%s] %s (%s) — %s", symbole, nom, sev, detail.split("\n")[0])

    if bloquants_ko:
        msg = f"{len(bloquants_ko)} contrôle(s) BLOQUANT(S) échoué(s) : " \
              f"{[c[0] for c in bloquants_ko]}"
        logger.error(msg)
        raise ValueError(msg)

    # ------------------------------------------------------------------
    # 4. Mapping N-1 et balance
    # ------------------------------------------------------------------
    logger.info("=== ÉTAPE 3/5 — Balance générale ===")
    mapping_fm = None
    balance_n1 = None

    if n1_fm:
        balance_n1, mapping_fm = load_balance_n1(n1_fm)

    # --- Rapprochement des comptes N/N-1 (P6) — AVANT build/map_cycles ---
    propositions = []
    appliques = []
    if balance_n1:
        from src.engine.account_matcher import (
            appliquer_rapprochements, proposer_rapprochements,
        )
        comptes_n = dict(
            df_fec.groupby("CompteNum")["CompteLib"].first().astype(str)
        )
        propositions = proposer_rapprochements(
            comptes_n, balance_n1, mapping_fm, pcg,
        )
        if propositions:
            appliques = _valider_rapprochements(
                propositions,
                rapprochements_auto=rapprochements_auto,
                rapprochements_valides=rapprochements_valides,
                rapprochements_interactif=rapprochements_interactif,
            )
            balance_n1, mapping_fm = appliquer_rapprochements(
                balance_n1, appliques, mapping_fm,
            )
            logger.info(
                "Rapprochements : %d proposé(s), %d appliqué(s)",
                len(propositions), len(appliques),
            )
    resultats["rapprochements_proposes"] = propositions
    resultats["rapprochements_appliques"] = appliques

    balance = build(df_fec, balance_n1)
    resultats["nb_comptes"] = len(balance)
    logger.info("Balance construite : %d comptes", len(balance))

    travail_path = write_travail(df_fec, balance, balance_n1, client, date_cloture, output)
    resultats["travail_path"] = travail_path
    resultats["travail_bytes"] = travail_path.read_bytes()
    resultats["travail_nom"]   = travail_path.name
    logger.info("Fichier de travail : %s", travail_path)

    # ------------------------------------------------------------------
    # 5. Mapping cycles
    # ------------------------------------------------------------------
    logger.info("=== ÉTAPE 4/5 — Mapping des cycles ===")
    balance_mappee = map_cycles(balance, mapping_fm, pcg)
    resultats["balance_mappee"] = balance_mappee
    nb_inconnus = (balance_mappee["cycle"] == "").sum()
    if nb_inconnus:
        logger.warning("%d compte(s) sans cycle — vérifier mapping_pcg.yaml", nb_inconnus)

    cycles = balance_mappee["cycle"].value_counts().to_dict()
    for cycle, nb in sorted(cycles.items()):
        logger.info("  Cycle %-12s : %3d comptes", cycle, nb)

    # ------------------------------------------------------------------
    # 5 bis. Contrôles financiers (AC-1 + cohérence résultat)
    # ------------------------------------------------------------------
    # Choix d'implémentation (le moins invasif) : les 9 contrôles FEC
    # restent à l'étape 2 (avant la construction de la balance, comportement
    # historique inchangé). AC-1 et la cohérence du résultat nécessitent la
    # balance mappée, disponible seulement après map_cycles : ils sont donc
    # exécutés ici via un second appel ciblé (run_controles_financiers,
    # même moteur que run_all avec balance_mappee) sans ré-exécuter les
    # 9 contrôles FEC.
    liasse_config = load_liasse_fiscale(pcg)
    controles_financiers = run_controles_financiers(
        balance_mappee,
        liasse_config,
        bilan_non_bloquant=bilan_non_bloquant,
    )
    controles.extend(controles_financiers)
    resultats["controles"] = controles

    for nom, ok, detail, sev in controles_financiers:
        symbole = "✓" if ok else "✗"
        logger.info("  [%s] %s (%s) — %s", symbole, nom, sev, detail.split("\n")[0])

    bilan_bloquants_ko = [
        c for c in controles_financiers if c[3] == "BLOQUANT" and not c[1]
    ]
    if bilan_bloquants_ko:
        msg = (
            f"Bilan déséquilibré — {bilan_bloquants_ko[0][2].splitlines()[0]} "
            f"Corrigez le FEC ou le mapping, ou relancez avec "
            f"--bilan-non-bloquant pour forcer la génération."
        )
        logger.error(msg)
        raise ValueError(msg)

    # ------------------------------------------------------------------
    # 6. Feuilles maîtresses
    # ------------------------------------------------------------------
    logger.info("=== ÉTAPE 5/5 — Génération des fichiers ===")
    fm_path = write_fm(balance_mappee, client, date_cloture, output,
                       pcg_config=pcg)
    resultats["fm_path"] = fm_path
    logger.info("Feuilles maîtresses : %s", fm_path)

    # ------------------------------------------------------------------
    # 7. Templates (optionnel)
    # ------------------------------------------------------------------
    if templates_dir and Path(templates_dir).exists():
        zip_path = write_tpl(
            templates_dir=templates_dir,
            nom_client=client,
            date_cloture=date_cloture,
            output_path=output,
            mapping_templates=pcg["templates"],
            fm_path=fm_path,
            integration_templates=pcg.get("integration_templates"),
        )
        resultats["zip_path"] = zip_path
        logger.info("Feuilles de travail : %s", zip_path)
    elif templates_dir:
        logger.warning("Dossier templates introuvable : %s", templates_dir)

    logger.info("=== Pipeline terminé ===")
    logger.info("  Travail: %s", resultats.get("travail_path"))
    logger.info("  FM     : %s", resultats.get("fm_path"))
    logger.info("  ZIP    : %s", resultats.get("zip_path", "(non demandé)"))
    logger.info("  Bilan  : %d écritures / %d comptes / %d contrôles",
                resultats["fec_lignes"], resultats["nb_comptes"], len(controles))

    return resultats


def main() -> None:
    args = _parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    templates = None if args.no_templates else args.templates

    try:
        run_pipeline(
            fec_path=args.fec,
            client=args.client,
            date_cloture=args.date_cloture,
            n1_fm=args.n1_fm,
            templates_dir=templates,
            output_dir=args.output,
            pcg_config_path=args.pcg_config,
            bilan_non_bloquant=args.bilan_non_bloquant,
            rapprochements_auto=args.rapprochements_auto,
            # Validation interactive par défaut en CLI (sans TTY : aucune
            # fusion — la confirmation explicite reste obligatoire)
            rapprochements_interactif=not args.rapprochements_auto,
        )
    except ValueError as exc:
        logger.error("Erreur bloquante : %s", exc)
        sys.exit(1)
    except FileNotFoundError as exc:
        logger.error("Fichier introuvable : %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
