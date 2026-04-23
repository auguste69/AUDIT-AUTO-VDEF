#!/usr/bin/env python3
"""
Orchestrateur de prompts XML séquentiels pour audit-automation.

Exécute les prompts 3 à 12 de scripts/prompts.md les uns après les autres
via des agents Claude Code isolés. Après chaque prompt : pytest, auto-fix
en cas de régression, vérification indépendante que l'objectif est atteint.

Usage :
    python3 scripts/prompt_runner.py --prompts-file scripts/prompts.md
    python3 scripts/prompt_runner.py --prompts-file scripts/prompts.md --resume
    python3 scripts/prompt_runner.py --prompts-file scripts/prompts.md --dry-run
"""

import argparse
import json
import logging
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent.parent
STATE_FILE = ROOT / "scripts" / ".run_state.json"
LOG_FILE = ROOT / "scripts" / "run_log.md"
CLAUDE_BIN = "/opt/node22/bin/claude"

AGENT_RULES = """RÈGLES STRICTES POUR CET AGENT — NE PAS DÉVIER :
1. Implémenter ENTIÈREMENT l'objectif. Zéro TODO, zéro placeholder, zéro mock.
2. Avant de terminer, vérifier que `python3 -m pytest tests/ -q` ne régresse pas.
3. Scope : uniquement les changements demandés. Ne pas refactorer du code non lié.
4. Respecter les conventions CLAUDE.md du projet (type hints, docstrings français, pas de print()).
5. Tout numéro de compte, libellé ou montant client doit venir de données, jamais hardcodé.
6. Lire les fichiers existants (RP-2) avant de les modifier.
7. Si une ambiguïté existe, choisir l'interprétation la plus complète compatible avec les contraintes."""

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Parsing des prompts
# ---------------------------------------------------------------------------
PROMPT_BLOCK_RE = re.compile(
    r"<prompt[^>]*>(.*?)</prompt>",
    re.DOTALL,
)
TAG_RE = re.compile(r"<(\w+)>(.*?)</\1>", re.DOTALL)


def parse_prompts(path: Path) -> list[dict]:
    """Extrait les blocs <prompt>...</prompt> du fichier markdown."""
    text = path.read_text(encoding="utf-8")
    prompts = []
    for block in PROMPT_BLOCK_RE.findall(text):
        p: dict = {}
        for tag, content in TAG_RE.findall(block):
            p[tag] = content.strip()
        if "objectif" in p:
            prompts.append(p)
    return prompts


# ---------------------------------------------------------------------------
# Invocation Claude
# ---------------------------------------------------------------------------
def run_claude(
    prompt: str,
    model: str = "sonnet",
    cwd: Optional[Path] = None,
    timeout: int = 900,
    add_rules: bool = True,
) -> tuple[int, str]:
    """Lance claude -p et retourne (returncode, output)."""
    full_prompt = f"{AGENT_RULES}\n\n---\n\n{prompt}" if add_rules else prompt
    cmd = [
        CLAUDE_BIN, "-p",
        "--dangerously-skip-permissions",
        "--model", model,
        full_prompt,
    ]
    try:
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            cwd=str(cwd or ROOT),
            timeout=timeout,
        )
        return r.returncode, r.stdout + r.stderr
    except subprocess.TimeoutExpired:
        return 1, f"TIMEOUT après {timeout}s"


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------
def run_tests(cwd: Path = ROOT) -> dict:
    """Lance pytest et retourne {passed, failed, errors, output}."""
    r = subprocess.run(
        [sys.executable, "-m", "pytest", "tests/", "-q", "--tb=short", "--no-header"],
        capture_output=True,
        text=True,
        cwd=str(cwd),
        timeout=300,
    )
    output = r.stdout + r.stderr
    passed = failed = errors = 0
    for line in reversed(output.splitlines()):
        if re.search(r"\d+ passed|\d+ failed|\d+ error", line):
            if m := re.search(r"(\d+) passed", line):
                passed = int(m.group(1))
            if m := re.search(r"(\d+) failed", line):
                failed = int(m.group(1))
            if m := re.search(r"(\d+) error", line):
                errors = int(m.group(1))
            break
    return {"passed": passed, "failed": failed, "errors": errors, "output": output}


def detect_baseline(cwd: Path = ROOT) -> int:
    """Détecte le baseline pytest actuel."""
    result = run_tests(cwd)
    return result["passed"]


# ---------------------------------------------------------------------------
# Agents spécialisés
# ---------------------------------------------------------------------------
def check_conflicts(prompt: dict, idx: int) -> list[str]:
    """Agent Haiku : vérifie que les fichiers/fonctions cités dans le prompt existent."""
    tree_output = subprocess.run(
        ["find", "src", "tests", "-name", "*.py", "-not", "-path", "*/__pycache__/*"],
        capture_output=True, text=True, cwd=str(ROOT),
    ).stdout

    checker_prompt = f"""Analyse ce prompt et identifie tous les fichiers Python, modules et fonctions
qu'il suppose DÉJÀ EXISTANTS dans le projet (fichiers à modifier, imports attendus, etc.).
Vérifie leur présence dans l'arborescence ci-dessous.

Réponds UNIQUEMENT en JSON valide sur une seule ligne :
{{"conflicts": [{{"item": "chemin/ou/fonction", "reason": "explication courte"}}]}}
Si aucun conflit : {{"conflicts": []}}

Prompt à analyser :
<objectif>{prompt.get('objectif', '')}</objectif>
<fichiers_a_modifier>{prompt.get('fichiers_a_modifier', '')}</fichiers_a_modifier>
<instructions>{prompt.get('instructions', '')[:1000]}</instructions>

Arborescence actuelle :
{tree_output}"""

    rc, output = run_claude(checker_prompt, model="haiku", timeout=120, add_rules=False)
    if rc != 0:
        logger.warning("Conflict checker échoué (rc=%d) — on continue sans vérification", rc)
        return []

    try:
        json_match = re.search(r'\{.*"conflicts".*\}', output, re.DOTALL)
        if json_match:
            data = json.loads(json_match.group())
            return [c["item"] for c in data.get("conflicts", [])]
    except (json.JSONDecodeError, KeyError):
        pass
    return []


def resolve_conflicts(prompt: dict, conflicts: list[str]) -> dict:
    """Agent Resolver : adapte minimalement le prompt si des prérequis manquent."""
    conflicts_str = "\n".join(f"- {c}" for c in conflicts)
    resolver_prompt = f"""Les éléments suivants sont référencés dans le prompt mais n'existent pas encore :
{conflicts_str}

Adapte MINIMALEMENT le prompt pour qu'il puisse s'exécuter correctement.
Si des fichiers manquants doivent être créés en prérequis, indique-le dans les instructions.
Retourne le prompt modifié complet, en conservant exactement les mêmes balises XML.

Prompt original :
<objectif>{prompt.get('objectif', '')}</objectif>
<instructions>{prompt.get('instructions', '')}</instructions>
<contraintes>{prompt.get('contraintes', '')}</contraintes>
<validation>{prompt.get('validation', '')}</validation>"""

    rc, output = run_claude(resolver_prompt, model="sonnet", timeout=180, add_rules=False)
    if rc != 0:
        logger.warning("Resolver échoué — utilisation du prompt original")
        return prompt

    resolved = dict(prompt)
    for tag, content in TAG_RE.findall(output):
        if tag in resolved:
            resolved[tag] = content.strip()
    return resolved


def run_impl_agent(prompt: dict, idx: int, attempt: int, prev_feedback: str = "") -> str:
    """Agent d'implémentation principal."""
    feedback_block = ""
    if prev_feedback:
        feedback_block = f"\n\nFEEDBACK DU VÉRIFICATEUR (tentative précédente) :\n{prev_feedback}\nCorrige ces points spécifiques.\n"

    impl_prompt = f"""PROMPT {idx + 1} — IMPLÉMENTATION{feedback_block}

<objectif>
{prompt.get('objectif', '')}
</objectif>

<fichiers_a_modifier>
{prompt.get('fichiers_a_modifier', '')}
</fichiers_a_modifier>

<fichiers_a_creer>
{prompt.get('fichiers_a_creer', '')}
</fichiers_a_creer>

<instructions>
{prompt.get('instructions', '')}
</instructions>

<contraintes>
{prompt.get('contraintes', '')}
</contraintes>

<validation>
{prompt.get('validation', '')}
</validation>

Répertoire de travail : {ROOT}
Commence par lire les fichiers concernés, puis implémente entièrement l'objectif.
Vérifie avec `python3 -m pytest tests/ -q` avant de terminer."""

    rc, output = run_claude(impl_prompt, model="sonnet", cwd=ROOT, timeout=900)
    if rc != 0:
        logger.warning("ImplAgent rc=%d pour prompt %d tentative %d", rc, idx + 1, attempt + 1)
    return output


def fix_regressions(test_output: str, idx: int, fix_attempt: int) -> bool:
    """Agent Fixer : corrige les régressions pytest. Retourne True si succès."""
    fixer_prompt = f"""Les tests pytest suivants ont échoué après une implémentation (prompt {idx + 1}) :

{test_output[-3000:]}

Corrige UNIQUEMENT ces échecs sans modifier les fichiers de test eux-mêmes.
Changements chirurgicaux uniquement — ne touche pas au code non lié.
Après correction, vérifie que `python3 -m pytest tests/ -q` passe."""

    rc, output = run_claude(fixer_prompt, model="sonnet", cwd=ROOT, timeout=600)
    if rc != 0:
        logger.warning("FixerAgent rc=%d, tentative %d", rc, fix_attempt + 1)
        return False

    result = run_tests(ROOT)
    return result["failed"] == 0 and result["errors"] == 0


def verify_objective(prompt: dict, idx: int) -> tuple[bool, str]:
    """Agent Vérificateur indépendant. Retourne (pass, feedback)."""
    verif_prompt = f"""Tu es un réviseur indépendant strict. Le prompt suivant a été implémenté :

<objectif>
{prompt.get('objectif', '')}
</objectif>

<validation>
{prompt.get('validation', '')}
</validation>

<fichiers_a_modifier>
{prompt.get('fichiers_a_modifier', '')}
</fichiers_a_modifier>

<fichiers_a_creer>
{prompt.get('fichiers_a_creer', '')}
</fichiers_a_creer>

Inspecte le code du projet (lis les fichiers pertinents) et détermine si
l'objectif a été ENTIÈREMENT atteint selon les critères de validation.

Réponds par EXACTEMENT une de ces deux formes (première ligne) :
PASS: [raison en 1 ligne]
FAIL: [liste précise et courte des points manquants]

Ne modifie AUCUN fichier."""

    rc, output = run_claude(verif_prompt, model="sonnet", cwd=ROOT, timeout=300, add_rules=False)
    if rc != 0:
        return False, f"Vérificateur échoué (rc={rc})"

    first_line = output.strip().splitlines()[0] if output.strip() else ""
    if first_line.startswith("PASS"):
        return True, first_line
    return False, output.strip()[:500]


# ---------------------------------------------------------------------------
# État / log
# ---------------------------------------------------------------------------
def load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {"current_prompt": 0, "baseline": 0, "results": []}


def save_state(state: dict) -> None:
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def write_log(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"\n{message}\n")
    logger.info(message.replace("\n", " | ")[:120])


# ---------------------------------------------------------------------------
# Boucle principale par prompt
# ---------------------------------------------------------------------------
MAX_IMPL_ATTEMPTS = 3
MAX_FIX_ATTEMPTS = 3


def run_prompt(prompt: dict, idx: int, state: dict) -> bool:
    """Exécute un prompt complet. Retourne True si succès."""
    baseline = state["baseline"]
    ts_start = datetime.now().strftime("%Y-%m-%d %H:%M")

    write_log(f"\n## {ts_start} — Prompt {idx + 1}/10 : {prompt.get('objectif', '')[:80]}")

    # 1. Conflict check
    conflicts = check_conflicts(prompt, idx)
    if conflicts:
        write_log(f"**Conflits détectés** : {', '.join(conflicts)}")
        prompt = resolve_conflicts(prompt, conflicts)
        write_log("**Prompt adapté par le Resolver**")
    else:
        write_log("**Conflits** : aucun")

    # 2. Boucle implémentation
    prev_feedback = ""
    for attempt in range(MAX_IMPL_ATTEMPTS):
        write_log(f"\n### Tentative {attempt + 1}/{MAX_IMPL_ATTEMPTS}")

        # 2a. Implémentation
        impl_output = run_impl_agent(prompt, idx, attempt, prev_feedback)
        write_log(f"ImplAgent terminé ({len(impl_output)} chars de sortie)")

        # 2b. Tests
        test_result = run_tests(ROOT)
        write_log(
            f"pytest : {test_result['passed']} passés, "
            f"{test_result['failed']} échoués, "
            f"{test_result['errors']} erreurs"
        )

        # 2c. Fix regressions si besoin
        has_regression = (
            test_result["failed"] > 0
            or test_result["errors"] > 0
            or test_result["passed"] < baseline
        )
        if has_regression:
            write_log(f"**Régression détectée** (baseline={baseline}, passed={test_result['passed']})")
            fixed = False
            for fix_attempt in range(MAX_FIX_ATTEMPTS):
                write_log(f"FixerAgent tentative {fix_attempt + 1}/{MAX_FIX_ATTEMPTS}...")
                fixed = fix_regressions(test_result["output"], idx, fix_attempt)
                if fixed:
                    test_result = run_tests(ROOT)
                    write_log(f"Fix réussi : {test_result['passed']} tests passés")
                    break
                write_log("Fix échoué, nouvelle tentative...")

            if not fixed:
                write_log(f"**ÉCHEC IRRÉSOLUBLE** : régression non corrigée après {MAX_FIX_ATTEMPTS} tentatives")
                return False

        # 2d. Vérification de l'objectif
        obj_ok, feedback = verify_objective(prompt, idx)
        if obj_ok:
            write_log(f"**Vérificateur** : ✅ {feedback}")
            state["results"].append({
                "idx": idx,
                "status": "success",
                "attempts": attempt + 1,
                "tests_passed": test_result["passed"],
                "ts": ts_start,
            })
            # Mettre à jour le baseline si on a plus de tests
            if test_result["passed"] > baseline:
                state["baseline"] = test_result["passed"]
                write_log(f"Baseline mis à jour : {test_result['passed']} tests")
            write_log(f"\n**Résultat** : ✅ succès en {attempt + 1} tentative(s)\n---")
            return True
        else:
            write_log(f"**Vérificateur** : ❌ objectif non atteint\n```\n{feedback}\n```")
            prev_feedback = feedback

    write_log(f"**ÉCHEC** : objectif non atteint après {MAX_IMPL_ATTEMPTS} tentatives\n---")
    state["results"].append({
        "idx": idx,
        "status": "failed",
        "attempts": MAX_IMPL_ATTEMPTS,
        "ts": ts_start,
    })
    return False


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Orchestrateur de prompts XML séquentiels")
    p.add_argument("--prompts-file", required=True, type=Path,
                   help="Fichier .md contenant les prompts XML")
    p.add_argument("--resume", action="store_true",
                   help="Reprend depuis scripts/.run_state.json")
    p.add_argument("--baseline", type=int, default=None,
                   help="Nombre de tests de référence (défaut : auto-détecté)")
    p.add_argument("--start-from", type=int, default=None,
                   help="Forcer le démarrage depuis le prompt i (0-indexé)")
    p.add_argument("--dry-run", action="store_true",
                   help="Affiche les prompts parsés sans exécuter")
    return p.parse_args()


def main() -> None:
    args = parse_args()

    prompts_file = args.prompts_file
    if not prompts_file.exists():
        logger.error("Fichier de prompts introuvable : %s", prompts_file)
        sys.exit(1)

    prompts = parse_prompts(prompts_file)
    if not prompts:
        logger.error("Aucun prompt trouvé dans %s", prompts_file)
        sys.exit(1)

    logger.info("%d prompt(s) chargés depuis %s", len(prompts), prompts_file)

    # Dry run
    if args.dry_run:
        for i, p in enumerate(prompts):
            print(f"\n{'='*60}")
            print(f"Prompt {i+1}/{len(prompts)}")
            print(f"Objectif : {p.get('objectif', '')[:200]}")
            fichiers = p.get('fichiers_a_modifier', '') + p.get('fichiers_a_creer', '')
            print(f"Fichiers : {fichiers[:200]}")
        print(f"\n{len(prompts)} prompt(s) prêts à l'exécution.")
        return

    # Vérifier que pytest est disponible
    r = subprocess.run(
        [sys.executable, "-m", "pytest", "--version"],
        capture_output=True, text=True, cwd=str(ROOT),
    )
    if r.returncode != 0:
        logger.info("Installation de pytest...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "pytest", "--break-system-packages"],
            check=True,
        )

    # État
    if args.resume:
        state = load_state()
        logger.info("Reprise depuis le prompt %d", state["current_prompt"] + 1)
    else:
        state = {"current_prompt": 0, "baseline": 0, "results": []}

    if args.start_from is not None:
        state["current_prompt"] = args.start_from
        logger.info("Démarrage forcé depuis le prompt %d", args.start_from + 1)

    # Baseline
    if args.baseline is not None:
        state["baseline"] = args.baseline
    elif state["baseline"] == 0:
        logger.info("Détection du baseline pytest...")
        state["baseline"] = detect_baseline(ROOT)
        logger.info("Baseline : %d tests", state["baseline"])

    write_log(
        f"\n# Run du {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        f"\n**Baseline** : {state['baseline']} tests"
        f"\n**Départ depuis prompt** : {state['current_prompt'] + 1}/{len(prompts)}"
    )

    # Boucle principale
    for i in range(state["current_prompt"], len(prompts)):
        logger.info("=" * 60)
        logger.info("PROMPT %d/%d", i + 1, len(prompts))
        logger.info("=" * 60)

        success = run_prompt(prompts[i], i, state)
        if not success:
            logger.error(
                "Prompt %d échoué de façon irrésoluble. "
                "Relancez avec --resume après correction manuelle.",
                i + 1,
            )
            state["current_prompt"] = i
            save_state(state)
            write_log(f"\n**RUN INTERROMPU** au prompt {i+1}. Relancer avec --resume.\n")
            sys.exit(1)

        state["current_prompt"] = i + 1
        save_state(state)

    # Succès complet
    write_log(
        f"\n# ✅ TOUS LES PROMPTS COMPLÉTÉS"
        f"\n**Tests finaux** : {state['baseline']} passés"
        f"\n**Date** : {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
    )
    logger.info("Tous les prompts exécutés avec succès.")
    logger.info("Log complet : %s", LOG_FILE)


if __name__ == "__main__":
    main()
