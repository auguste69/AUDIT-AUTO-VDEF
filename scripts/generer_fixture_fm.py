"""
Régénère la fixture de régression cellulaire tests/fixtures/FM_SYNTHETIQUE_REF.xlsx.

À lancer UNIQUEMENT quand un changement de valeurs/formats du FM est voulu
(évolution du moteur ou des writers), après avoir vérifié le diff onglet
par onglet contre l'ancienne fixture.

Usage :
    python3 -m scripts.generer_fixture_fm
"""

import logging
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.parsers.fec_parser import parse                       # noqa: E402
from src.parsers.balance_n1_loader import load_balance_n1      # noqa: E402
from src.parsers.mapping_parser import from_pcg_config         # noqa: E402
from src.engine.balance_builder import build                   # noqa: E402
from src.engine.cycle_mapper import map_cycles                 # noqa: E402
from src.writers.fm_writer import write                        # noqa: E402
from tests.synthetic_data import (                             # noqa: E402
    generer_balance_n1_xlsx,
    generer_fec_synthetique,
)

logging.basicConfig(level=logging.INFO, format="%(levelname)-8s %(message)s")
logger = logging.getLogger(__name__)

RACINE = Path(__file__).resolve().parent.parent
FIXTURE = RACINE / "tests" / "fixtures" / "FM_SYNTHETIQUE_REF.xlsx"
PCG_PATH = RACINE / "src" / "config" / "mapping_pcg.yaml"


def main() -> None:
    """Génère le FM synthétique et l'installe comme fixture de référence."""
    with tempfile.TemporaryDirectory() as tmp_str:
        tmp = Path(tmp_str)
        fec = generer_fec_synthetique(tmp / "DEMO_2025_FEC.txt")
        n1_xlsx = generer_balance_n1_xlsx(tmp / "balance_n1.xlsx")

        df = parse(fec)
        balance_n1, _ = load_balance_n1(n1_xlsx)
        pcg = from_pcg_config(PCG_PATH)
        balance_mappee = map_cycles(build(df, balance_n1), None, pcg)

        fm_path = write(balance_mappee, "DEMO", "31/12/2025", tmp,
                        pcg_config=pcg)
        FIXTURE.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(fm_path, FIXTURE)

    logger.info("Fixture régénérée : %s", FIXTURE)


if __name__ == "__main__":
    main()
