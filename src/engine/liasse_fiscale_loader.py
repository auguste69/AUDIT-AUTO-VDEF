"""
Chargeur et validateur de la section liasse_fiscale du mapping_pcg.yaml.

La section liasse_fiscale externalise les LISTES de préfixes PCG utilisées
par les calculs des états financiers (Bilan, Tréso, AACE). Les signes de
présentation (× -1 pour le passif) et les conditions de filtrage sur le
signe du solde (">0" / "<0") restent dans le code appelant : seules les
listes de préfixes sont configurables.

Usage :
    pcg = mapping_parser.from_pcg_config("src/config/mapping_pcg.yaml")
    liasse = load_liasse_fiscale(pcg)
    prefixes_aace = liasse["aace"]["prefixes"]
"""

import logging
from typing import Dict, List

logger = logging.getLogger(__name__)

# Clés obligatoires de la section liasse_fiscale, par sous-section.
# Les sections ebit et pl_detaille ne sont pas validées ici : elles restent
# vides tant que les prompts 9 et 10 ne les ont pas remplies.
_CLES_REQUISES: Dict[str, Dict[str, List[str]]] = {
    "bilan": {
        "actif": [
            "capital_souscrit_non_appele", "immo_incorp", "immo_corp",
            "immo_fi", "stocks", "avances_acomptes_verses",
            "creances_clients", "autres_creances", "vmp",
            "disponibilites", "cca",
        ],
        "passif": [
            "capital", "primes", "reserve_legale", "autres_reserves",
            "report_nouveau", "resultat_exercice", "resultat_en_cours",
            "subventions", "prov_reglementees", "prov_risques",
            "prov_charges", "emprunts", "dettes_fournisseurs",
            "dettes_fiscales_sociales", "autres_dettes", "pca",
        ],
        "comptes_bascule": ["prefixes", "interco"],
    },
    "treso": {
        "ressources_stables": [
            "capitaux_propres", "amort_dep", "prov_risques",
            "dettes_financieres_mlt",
        ],
        "emplois_stables": ["actif_immobilise_brut"],
        "actif_circulant_exploitation": [
            "stocks", "creances_clients", "autres_creances", "cca",
        ],
        "passif_circulant_exploitation": [
            "dettes_fournisseurs", "dettes_fiscales_sociales",
            "autres_dettes", "pca",
        ],
        "tresorerie_directe": [
            "tresorerie_active", "tresorerie_passive_519",
            "tresorerie_passive_512",
        ],
    },
    "aace": {"prefixes": []},
}


def _normaliser_prefixes(valeurs: list, chemin: str) -> List[str]:
    """Convertit une liste de préfixes en liste de strings.

    Lève ValueError si la valeur n'est pas une liste.
    """
    if not isinstance(valeurs, list):
        raise ValueError(
            f"liasse_fiscale : la clé '{chemin}' doit être une liste de "
            f"préfixes, trouvé {type(valeurs).__name__}."
        )
    return [str(v) for v in valeurs]


def load_liasse_fiscale(pcg_config: dict) -> dict:
    """Extrait, valide et normalise la section liasse_fiscale d'une config PCG.

    Paramètres
    ----------
    pcg_config : dict
        Config produite par mapping_parser.from_pcg_config() — doit contenir
        la clé "liasse_fiscale".

    Retourne
    --------
    dict
        Structure {bilan: {actif: {...}, passif: {...}, comptes_bascule: {...}},
        treso: {...}, aace: {...}} où toute feuille est une liste de préfixes
        (strings).

    Lève
    ----
    ValueError
        Si la section liasse_fiscale est absente ou si une clé obligatoire
        manque (message explicite indiquant la clé fautive).
    """
    liasse = pcg_config.get("liasse_fiscale")
    if not liasse:
        raise ValueError(
            "Section 'liasse_fiscale' absente ou vide dans la config PCG — "
            "vérifier src/config/mapping_pcg.yaml."
        )

    resultat: dict = {}
    for section, sous_sections in _CLES_REQUISES.items():
        if section not in liasse:
            raise ValueError(
                f"liasse_fiscale : sous-section obligatoire '{section}' "
                f"absente. Sections disponibles : {sorted(liasse.keys())}."
            )
        contenu = liasse[section]

        # Cas plat : la sous-section est directement {cle: [prefixes]}
        # (ex. aace.prefixes) — détecté quand la spec n'a pas de sous-clés.
        resultat[section] = {}
        for cle, feuilles in sous_sections.items():
            if cle not in contenu:
                raise ValueError(
                    f"liasse_fiscale.{section} : clé obligatoire '{cle}' "
                    f"absente. Clés disponibles : {sorted(contenu.keys())}."
                )
            if feuilles:  # niveau intermédiaire (ex. bilan.actif.{postes})
                resultat[section][cle] = {}
                for feuille in feuilles:
                    if feuille not in contenu[cle]:
                        raise ValueError(
                            f"liasse_fiscale.{section}.{cle} : clé "
                            f"obligatoire '{feuille}' absente. Clés "
                            f"disponibles : {sorted(contenu[cle].keys())}."
                        )
                    resultat[section][cle][feuille] = _normaliser_prefixes(
                        contenu[cle][feuille],
                        f"{section}.{cle}.{feuille}",
                    )
            else:  # feuille directe (ex. aace.prefixes)
                resultat[section][cle] = _normaliser_prefixes(
                    contenu[cle], f"{section}.{cle}"
                )

    logger.debug(
        "liasse_fiscale chargée : sections %s", sorted(resultat.keys())
    )
    return resultat
