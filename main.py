"""
CLI du pipeline d'audit — point d'entrée en ligne de commande.

Usage :
    python3 main.py <fec> --client NOM --date-cloture JJ/MM/AAAA [options]

Exemples :
    # Minimum (FEC seul, mapping 100% PCG automatique)
    python3 main.py data/client_FEC.txt --client ACME --date-cloture 31/12/2025

    # Avec FM N-1 et templates
    python3 main.py data/GILAC_2025_12_31_FEC.txt \\
        --client GILAC --date-cloture 31/12/2025 \\
        --n1-fm data/FM\\ GILAC.xlsx \\
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
    p.add_argument("-v", "--verbose", action="store_true",
                   help="Mode verbeux (DEBUG)")
    return p.parse_args()


def run_pipeline(
    fec_path: str,
    client: str,
    date_cloture: str,
    n1_fm: Optional[str] = None,
    templates_dir: Optional[str] = None,
    output_dir: str = "output",
    pcg_config_path: Optional[str] = None,
) -> dict:
    """
    Exécute le pipeline complet et retourne un dict de résultats.

    Compatible CLI et Streamlit — aucun print(), tout passe par logging.

    Retourne
    --------
    dict avec les clés :
        fec_lignes, nb_comptes, controles, fm_path, zip_path (optionnel)
    """
    from src.parsers.fec_parser import parse
    from src.parsers.mapping_parser import from_fm, from_pcg_config
    from src.engine.balance_builder import build
    from src.engine.controls import run_all
    from src.engine.cycle_mapper import map_cycles
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
        n1_path = Path(n1_fm)
        suffix = n1_path.suffix.lower()

        if suffix == ".txt":
            # Source FEC N-1 : pas de mapping FM, uniquement balance agrégée
            logger.info("Source N-1 : FEC — %s", n1_fm)
            from src.parsers.mapping_parser import from_fec_n1
            balance_n1 = from_fec_n1(n1_fm)
            logger.info("Balance N-1 : %d comptes chargés depuis le FEC N-1", len(balance_n1))

        elif suffix == ".xlsx":
            from src.parsers.mapping_parser import detect_balance_sheet, from_balance_excel
            import openpyxl
            nom_feuille, mode = detect_balance_sheet(n1_path)

            if mode == "fm":
                logger.info("Source N-1 : FM existant — %s", n1_fm)
                mapping_fm = from_fm(n1_fm)
                # Extraire les soldes N-1 depuis la feuille détectée
                wb2 = openpyxl.load_workbook(n1_fm, read_only=True, data_only=True)
                ws = wb2[nom_feuille]
                balance_n1 = {}
                for row in ws.iter_rows(min_row=10, values_only=True):
                    if row[1] is None:
                        continue
                    try:
                        num = str(int(float(row[1])))
                    except (ValueError, TypeError):
                        continue
                    solde = float(row[3]) if row[3] is not None else 0.0
                    balance_n1[num] = {
                        "libelle": str(row[2]) if row[2] else "",
                        "solde_ke": solde,
                    }
                wb2.close()
                logger.info("Balance N-1 : %d comptes chargés depuis le FM", len(balance_n1))

            else:
                # Balance Excel simple : pas de mapping FM
                logger.info("Source N-1 : Balance Excel simple — %s", n1_fm)
                balance_n1 = from_balance_excel(n1_fm)
                logger.info("Balance N-1 : %d comptes chargés depuis la balance Excel", len(balance_n1))

        else:
            raise ValueError(
                f"Format N-1 non reconnu pour '{n1_fm}'. "
                "Formats acceptés : .txt (FEC), .xlsx (FM ou balance simple)."
            )

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
    nb_inconnus = (balance_mappee["cycle"] == "").sum()
    if nb_inconnus:
        logger.warning("%d compte(s) sans cycle — vérifier mapping_pcg.yaml", nb_inconnus)

    cycles = balance_mappee["cycle"].value_counts().to_dict()
    for cycle, nb in sorted(cycles.items()):
        logger.info("  Cycle %-12s : %3d comptes", cycle, nb)

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
            balance_mappee=balance_mappee,
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
        )
    except ValueError as exc:
        logger.error("Erreur bloquante : %s", exc)
        sys.exit(1)
    except FileNotFoundError as exc:
        logger.error("Fichier introuvable : %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()
