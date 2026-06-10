"""
Dataclasses des états financiers synthétiques (Bilan, Tréso, EBIT).

Ces modèles transportent les résultats du moteur de calcul
(src/engine/financial_engine.py) vers les writers Excel
(src/writers/fm_writer.py), sans couplage entre calcul et présentation.

Conventions :
- Chaque poste porte une valeur N et une valeur N-1 en K€, déjà arrondies
  (3 décimales) et signées selon la convention de présentation (passif et
  produits affichés en positif).
- ``as_dict()`` restitue le format historique ``{cle: (valeur_n, valeur_n1)}``
  consommé par les fonctions d'écriture de fm_writer — il limite les
  réécritures côté writer et garantit le zéro changement de comportement.
"""

from dataclasses import dataclass
from typing import Dict, Tuple


@dataclass(frozen=True)
class PosteComptable:
    """Poste agrégé d'un état financier.

    Attributs
    ---------
    cle : str
        Identifiant technique du poste (ex: "immo_corp", "frng").
    valeur_n : float
        Valeur de l'exercice N en K€ (signe de présentation appliqué).
    valeur_n1 : float
        Valeur de l'exercice N-1 en K€ (signe de présentation appliqué).
    """

    cle: str
    valeur_n: float
    valeur_n1: float

    def as_tuple(self) -> Tuple[float, float]:
        """Retourne (valeur_n, valeur_n1) — format historique des writers."""
        return (self.valeur_n, self.valeur_n1)


@dataclass(frozen=True)
class BilanSynthetique:
    """Bilan synthétique : postes actif/passif + totaux.

    Attributs
    ---------
    postes : Dict[str, PosteComptable]
        Postes détaillés (actif puis passif), indexés par clé technique,
        dans l'ordre de calcul. Les totaux n'y figurent pas.
    total_actif : PosteComptable
        Total de l'actif (N, N-1).
    total_passif : PosteComptable
        Total du passif (N, N-1).
    """

    postes: Dict[str, PosteComptable]
    total_actif: PosteComptable
    total_passif: PosteComptable

    def as_dict(self) -> Dict[str, Tuple[float, float]]:
        """Aplatit postes + totaux en {cle: (valeur_n, valeur_n1)}."""
        resultat = {cle: p.as_tuple() for cle, p in self.postes.items()}
        resultat["total_actif"] = self.total_actif.as_tuple()
        resultat["total_passif"] = self.total_passif.as_tuple()
        return resultat


@dataclass(frozen=True)
class EbitSynthetique:
    """Compte de résultat d'exploitation synthétique (EBIT, cerfa 2052).

    Attributs
    ---------
    postes : Dict[str, PosteComptable]
        Postes détaillés produits puis charges d'exploitation, indexés par
        clé technique (ex: "ventes_marchandises", "impots_taxes"), dans
        l'ordre de présentation. Les agrégats (ca, total_produits,
        total_charges, ebit) n'y figurent pas.
    ca : PosteComptable
        Chiffre d'affaires = ventes de marchandises + production vendue.
    total_produits : PosteComptable
        Total des produits d'exploitation (N, N-1).
    total_charges : PosteComptable
        Total des charges d'exploitation (N, N-1).
    ebit : PosteComptable
        Résultat d'exploitation = total produits − total charges.
    """

    postes: Dict[str, PosteComptable]
    ca: PosteComptable
    total_produits: PosteComptable
    total_charges: PosteComptable
    ebit: PosteComptable

    def as_dict(self) -> Dict[str, Tuple[float, float]]:
        """Aplatit postes + agrégats en {cle: (valeur_n, valeur_n1)}."""
        resultat = {cle: p.as_tuple() for cle, p in self.postes.items()}
        resultat["ca"] = self.ca.as_tuple()
        resultat["total_produits"] = self.total_produits.as_tuple()
        resultat["total_charges"] = self.total_charges.as_tuple()
        resultat["ebit"] = self.ebit.as_tuple()
        return resultat


@dataclass(frozen=True)
class TresoSynthetique:
    """Analyse de trésorerie : postes BFR/FRNG/TN + agrégats clés.

    Attributs
    ---------
    postes : Dict[str, PosteComptable]
        Postes détaillés et sous-totaux (ressources/emplois stables, actif et
        passif circulants, trésorerie directe), indexés par clé technique.
        Les trois agrégats clés (frng, bfr, tn) n'y figurent pas.
    frng : PosteComptable
        Fonds de roulement net global (N, N-1).
    bfr : PosteComptable
        Besoin en fonds de roulement (N, N-1).
    tn : PosteComptable
        Trésorerie nette = FRNG − BFR (N, N-1).
    """

    postes: Dict[str, PosteComptable]
    frng: PosteComptable
    bfr: PosteComptable
    tn: PosteComptable

    def as_dict(self) -> Dict[str, Tuple[float, float]]:
        """Aplatit postes + agrégats en {cle: (valeur_n, valeur_n1)}."""
        resultat = {cle: p.as_tuple() for cle, p in self.postes.items()}
        resultat["frng"] = self.frng.as_tuple()
        resultat["bfr"] = self.bfr.as_tuple()
        resultat["tn"] = self.tn.as_tuple()
        return resultat
