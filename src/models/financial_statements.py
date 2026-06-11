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
from typing import Dict, List, Optional, Tuple


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
class LigneDetail:
    """Ligne d'un état détaillé (poste, sous-total ou agrégat).

    Attributs
    ---------
    cle : str
        Identifiant technique de la ligne (ex: "fonds_commercial").
    libelle : str
        Libellé de présentation (ex: "Fonds commercial").
    valeur_n : float
        Valeur N en K€ (signe de présentation appliqué).
    valeur_n1 : float
        Valeur N-1 en K€ (signe de présentation appliqué).
    apres_section : Optional[str]
        Pour les agrégats uniquement : clé de la section après laquelle la
        ligne doit être insérée par le writer (None sinon).
    """

    cle: str
    libelle: str
    valeur_n: float
    valeur_n1: float
    apres_section: Optional[str] = None

    def as_tuple(self) -> Tuple[float, float]:
        """Retourne (valeur_n, valeur_n1)."""
        return (self.valeur_n, self.valeur_n1)


@dataclass(frozen=True)
class SectionDetail:
    """Section d'un état détaillé : postes + sous-total optionnel.

    Attributs
    ---------
    cle : str
        Identifiant technique de la section (ex: "immo_incorp").
    libelle : Optional[str]
        Libellé du sous-total de section (None → pas de ligne sous-total).
    postes : List[LigneDetail]
        Lignes de la section, dans l'ordre de présentation.
    sous_total : Optional[LigneDetail]
        Sous-total de la section (somme des postes), None si libelle absent.
    """

    cle: str
    libelle: Optional[str]
    postes: List[LigneDetail]
    sous_total: Optional[LigneDetail]


@dataclass(frozen=True)
class ActifDetaille:
    """Actif détaillé (format liasse fiscale 2050).

    Attributs
    ---------
    sections : List[SectionDetail]
        Sections dans l'ordre de présentation (immobilisations, stocks, …).
    agregats : List[LigneDetail]
        Agrégats inter-sections (ex: "Actif immobilisé net"), positionnés
        via leur attribut apres_section.
    total : LigneDetail
        Total de l'actif (somme de toutes les sections).
    """

    sections: List[SectionDetail]
    agregats: List[LigneDetail]
    total: LigneDetail


@dataclass(frozen=True)
class PassifDetaille:
    """Passif détaillé (format liasse fiscale 2051).

    Attributs
    ---------
    sections : List[SectionDetail]
        Sections dans l'ordre de présentation (capitaux propres, dettes, …).
    agregats : List[LigneDetail]
        Agrégats inter-sections (souvent vide côté passif).
    total : LigneDetail
        Total du passif (somme de toutes les sections).
    """

    sections: List[SectionDetail]
    agregats: List[LigneDetail]
    total: LigneDetail


@dataclass(frozen=True)
class PlDetaille:
    """Compte de résultat détaillé (format liasse fiscale 2052/2053).

    Convention de signe : tous les postes valent −solde (produits créditeurs
    → positif, charges débitrices → négatif). Les résultats intermédiaires
    sont des sommes algébriques. Le résultat net est donc strictement égal à
    −(somme des soldes des classes 6 et 7).

    Attributs
    ---------
    ebit : EbitSynthetique
        Détail exploitation (postes EBIT, charges présentées en positif —
        à négativer pour l'affichage liasse).
    postes : Dict[str, PosteComptable]
        Postes hors EBIT : produits_financiers, charges_financieres,
        produits_exceptionnels, charges_exceptionnelles,
        participation_salaries, impots_benefices, quotes_parts, et les
        résiduels d'exploitation produits_expl_divers / charges_expl_divers
        (comptes 6/7 d'exploitation non capturés par les postes EBIT),
        ainsi que produits_exploitation / charges_exploitation (totaux
        liasse, résiduels des classes 7 / 6).
    resultat_exploitation : PosteComptable
        Résultat d'exploitation liasse (produits + charges d'exploitation).
    resultat_financier : PosteComptable
        Produits financiers + charges financières.
    resultat_courant : PosteComptable
        Résultat courant avant impôts (exploitation + quotes-parts + financier).
    resultat_exceptionnel : PosteComptable
        Produits exceptionnels + charges exceptionnelles.
    resultat_net : PosteComptable
        Résultat net comptable = −(somme classes 6 et 7).
    """

    ebit: "EbitSynthetique"
    postes: Dict[str, PosteComptable]
    resultat_exploitation: PosteComptable
    resultat_financier: PosteComptable
    resultat_courant: PosteComptable
    resultat_exceptionnel: PosteComptable
    resultat_net: PosteComptable

    def as_dict(self) -> Dict[str, Tuple[float, float]]:
        """Aplatit postes + résultats en {cle: (valeur_n, valeur_n1)}."""
        resultat = {cle: p.as_tuple() for cle, p in self.postes.items()}
        resultat["resultat_exploitation"] = self.resultat_exploitation.as_tuple()
        resultat["resultat_financier"] = self.resultat_financier.as_tuple()
        resultat["resultat_courant"] = self.resultat_courant.as_tuple()
        resultat["resultat_exceptionnel"] = self.resultat_exceptionnel.as_tuple()
        resultat["resultat_net"] = self.resultat_net.as_tuple()
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
