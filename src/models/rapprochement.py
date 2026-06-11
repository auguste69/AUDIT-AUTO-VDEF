"""
Dataclass des rapprochements de comptes N / N-1 (P6 — account_matcher).

Un Rapprochement décrit une PROPOSITION de fusion entre un compte orphelin
de l'exercice N-1 et un compte orphelin de l'exercice N (compte ayant
changé de numéro entre les deux exercices). La fusion n'est appliquée
qu'après confirmation explicite de l'utilisateur.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Rapprochement:
    """Proposition de rapprochement entre un compte N-1 et un compte N.

    Attributs
    ---------
    compte_n1 : str
        Numéro du compte orphelin de l'exercice N-1 (ex: "512003").
    libelle_n1 : str
        Libellé du compte N-1.
    compte_n : str
        Numéro du compte orphelin de l'exercice N (ex: "5123001").
    libelle_n : str
        Libellé du compte N.
    score : float
        Score composite [0, 1] (préfixe 40 % + libellé 35 % + cycle 15 %
        + classification bilan 10 %).
    score_prefixe : float
        Composante préfixe commun [0, 1] (longueur du plus long préfixe
        commun ÷ longueur max des deux numéros).
    score_libelle : float
        Composante similarité de libellé [0, 1] (ratio de tokens communs,
        insensible à la casse, hors mots vides).
    meme_cycle : bool
        True si les deux comptes relèvent du même cycle d'audit.
    meme_classification : bool
        True si même classification bilan (Actif/Passif/Charges/Produits).
    """

    compte_n1: str
    libelle_n1: str
    compte_n: str
    libelle_n: str
    score: float
    score_prefixe: float
    score_libelle: float
    meme_cycle: bool
    meme_classification: bool
