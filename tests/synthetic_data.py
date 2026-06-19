"""
Générateur de données de test synthétiques et anonymes.

Remplace les fichiers client réels (FEC + FM) supprimés du dépôt pour des
raisons de confidentialité. Produit, de façon DÉTERMINISTE (aucun aléa) :

- un FEC synthétique au format réel : encodage CP1252, séparateur TAB,
  montants à virgule française, 18 colonnes obligatoires, équilibré au
  centime, couvrant les 13 cycles d'audit et les cas sensibles du moteur
  financier (519 créditeur, 512 débiteur, 455 c/c associé créditeur,
  486/487, 491, 409/419, résultat en cours non clôturé) ;
- une balance N-1 Excel simple (CompteNum / CompteLib / Solde en €),
  sommant à zéro, avec des comptes soldés en N (orphelins N-1).

Utilisé par tests/test_pipeline_synthetique.py et
tests/test_fm_regression_cellulaire.py.
"""

from pathlib import Path
from typing import Dict, List, Tuple

import openpyxl

# ---------------------------------------------------------------------------
# Plan de comptes synthétique (libellés accentués → exercice le CP1252)
# ---------------------------------------------------------------------------

COMPTES: Dict[str, str] = {
    "101000": "Capital social",
    "106100": "Réserve légale",
    "110000": "Report à nouveau",
    "151000": "Provisions pour risques",
    "164000": "Emprunts auprès des établissements de crédit",
    "205000": "Logiciels et licences",
    "213500": "Bâtiments industriels",
    "215400": "Matériel industriel",
    "261000": "Titres de participation",
    "275000": "Dépôts et cautionnements versés",
    "280500": "Amortissements logiciels",
    "281350": "Amortissements bâtiments",
    "281540": "Amortissements matériel",
    "310000": "Stocks de matières premières",
    "370000": "Stocks de marchandises",
    "391000": "Dépréciation stocks matières",
    "401100": "Fournisseurs",
    "409100": "Fournisseurs - avances versées",
    "411100": "Clients",
    "419100": "Clients - avances reçues",
    "421000": "Personnel - rémunérations dues",
    "431000": "Sécurité sociale",
    "445660": "TVA déductible",
    "445710": "TVA collectée",
    "444000": "État - impôt sur les bénéfices",
    "455100": "Associés - comptes courants",
    "467000": "Autres comptes débiteurs ou créditeurs",
    "486000": "Charges constatées d'avance",
    "487000": "Produits constatés d'avance",
    "491100": "Dépréciation des comptes clients",
    "512100": "Banque principale",
    "512200": "Banque secondaire",
    "519000": "Concours bancaires courants",
    "530000": "Caisse",
    "601000": "Achats de matières premières",
    "603100": "Variation des stocks de matières",
    "604000": "Achats d'études et prestations",
    "613500": "Locations mobilières",
    "622600": "Honoraires",
    "641000": "Rémunérations du personnel",
    "645000": "Charges de sécurité sociale",
    "661100": "Intérêts des emprunts",
    "671200": "Pénalités et amendes",
    "681120": "Dotations amortissements immobilisations",
    "681740": "Dotations dépréciations créances clients",
    "695000": "Impôt sur les bénéfices",
    "701000": "Ventes de produits finis",
    "706000": "Prestations de services",
    "713500": "Variation des stocks de produits",
    "758000": "Produits divers de gestion courante",
    "761000": "Produits de participations",
    "771800": "Autres produits exceptionnels",
}

# Comptes présents UNIQUEMENT en N-1 (soldés en N) — exercent l'inclusion
# des orphelins dans la balance comparative.
COMPTES_N1_SEULS: Dict[str, str] = {
    "274000": "Prêts au personnel",
    "168800": "Intérêts courus sur emprunts",
}

_COLONNES_FEC = [
    "JournalCode", "JournalLib", "EcritureNum", "EcritureDate",
    "CompteNum", "CompteLib", "CompAuxNum", "CompAuxLib",
    "PieceRef", "PieceDate", "EcritureLib", "Debit", "Credit",
    "EcritureLet", "DateLet", "ValidDate", "Montantdevise", "Idevise",
]


def _montant_fr(centimes: int) -> str:
    """Formate un montant en centimes → chaîne à virgule française."""
    return f"{centimes // 100},{centimes % 100:02d}"


def _ecritures_synthetiques() -> List[Tuple[str, str, str, str,
                                            List[Tuple[str, int, int]]]]:
    """Construit la liste des écritures (journal, num, date, libellé, lignes).

    Chaque ligne est (compte, debit_centimes, credit_centimes). Toutes les
    écritures sont équilibrées par construction.
    """
    ecritures = []

    # --- À-nouveaux au 01/01/2025 (équilibrés) -------------------------
    an = [
        ("101000",         0, 30000000),   # capital 300 K€
        ("106100",         0,  3000000),
        ("110000",         0,  4500000),
        ("151000",         0,  2000000),
        ("164000",         0, 18000000),
        ("205000",   6000000,        0),
        ("213500",  30000000,        0),
        ("215400",  16000000,        0),
        ("261000",   5000000,        0),
        ("275000",    800000,        0),
        ("280500",         0,  2400000),
        ("281350",         0,  9000000),
        ("281540",         0,  6400000),
        ("310000",   7000000,        0),
        ("370000",   3500000,        0),
        ("391000",         0,   400000),
        ("401100",         0,  5200000),
        ("411100",   7400000,        0),
        ("491100",         0,   300000),
        ("455100",         0,  6000000),   # c/c associés créditeurs 60 K€
        ("512100",  21000000,        0),
        ("530000",    100000,        0),
        ("519000",         0,  3600000),   # CBC 36 K€ (créditeur)
        ("486000",   1400000,        0),
        ("487000",         0,   600000),
        ("467000",         0,  6800000),
    ]
    ecritures.append(("AN", "AN-1", "20250101", "À-nouveaux 2025", an))

    # --- Cycle mensuel : ventes, achats, paie, encaissements -----------
    for mois in range(1, 13):
        date = f"2025{mois:02d}15"
        # Ventes : clients TTC / ventes HT + TVA collectée
        ecritures.append(("VE", f"VE-{mois}", date, "Facture ventes", [
            ("411100", 14400000, 0),
            ("701000", 0, 10000000),
            ("706000", 0,  2000000),
            ("445710", 0,  2400000),
        ]))
        # Achats : charges HT + TVA déductible / fournisseurs TTC
        ecritures.append(("AC", f"AC-{mois}", date, "Facture achats", [
            ("601000", 5000000, 0),
            ("604000",  800000, 0),
            ("613500",  500000, 0),
            ("622600",  400000, 0),
            ("445660", 1340000, 0),
            ("401100", 0, 8040000),
        ]))
        # Paie
        ecritures.append(("OD", f"PA-{mois}", date, "Paie mensuelle", [
            ("641000", 2500000, 0),
            ("645000", 1000000, 0),
            ("421000", 0, 2600000),
            ("431000", 0,  900000),
        ]))
        # Règlements (clients → banque, banque → fournisseurs et dettes)
        ecritures.append(("BQ", f"BQ-{mois}", date, "Règlements bancaires", [
            ("512100", 14000000, 0),
            ("411100", 0, 14000000),
        ]))
        ecritures.append(("BQ", f"BR-{mois}", date, "Paiements fournisseurs", [
            ("401100", 7800000, 0),
            ("421000", 2600000, 0),
            ("431000",  900000, 0),
            ("512100", 0, 11300000),
        ]))

    # --- Écritures spécifiques (cas sensibles du moteur) ---------------
    ecritures += [
        ("OD", "OD-1", "20250630", "Dotation amortissements", [
            ("681120", 2500000, 0),
            ("281350", 0, 1500000),
            ("281540", 0,  800000),
            ("280500", 0,  200000),
        ]),
        ("OD", "OD-2", "20250630", "Dépréciation clients", [
            ("681740", 200000, 0),
            ("491100", 0, 200000),
        ]),
        ("OD", "OD-3", "20250930", "Variation de stocks", [
            ("603100", 600000, 0),
            ("310000", 0, 600000),
            ("370000", 400000, 0),
            ("713500", 0, 400000),
        ]),
        ("BQ", "BQ-X1", "20250420", "Avance versée fournisseur", [
            ("409100", 900000, 0),
            ("512100", 0, 900000),
        ]),
        ("BQ", "BQ-X2", "20250505", "Avance reçue client", [
            ("512200", 1200000, 0),
            ("419100", 0, 1200000),
        ]),
        ("OD", "OD-4", "20251231", "Charges constatées d'avance", [
            ("486000", 300000, 0),
            ("613500", 0, 300000),
        ]),
        ("OD", "OD-5", "20251231", "Produits constatés d'avance", [
            ("706000", 250000, 0),
            ("487000", 0, 250000),
        ]),
        ("BQ", "BQ-X3", "20250810", "Intérêts d'emprunt et agios", [
            ("661100", 450000, 0),
            ("512100", 0, 450000),
        ]),
        ("BQ", "BQ-X4", "20250915", "Remboursement annuité emprunt", [
            ("164000", 2000000, 0),
            ("512100", 0, 2000000),
        ]),
        ("OD", "OD-6", "20251105", "Pénalité fiscale (exceptionnel)", [
            ("671200", 120000, 0),
            ("512200", 0, 120000),
        ]),
        ("BQ", "BQ-X5", "20250612", "Dividendes de participations reçus", [
            ("512200", 380000, 0),
            ("761000", 0, 380000),
        ]),
        ("OD", "OD-7", "20251230", "Produit exceptionnel divers", [
            ("467000", 500000, 0),
            ("771800", 0, 220000),
            ("758000", 0, 280000),
        ]),
        ("OD", "OD-8", "20251231", "Acompte impôt sur les bénéfices", [
            ("695000", 800000, 0),
            ("444000", 0, 800000),
        ]),
    ]
    return ecritures


def generer_fec_synthetique(path: Path) -> Path:
    """Écrit le FEC synthétique (CP1252, TAB, virgule française) et le retourne."""
    lignes = ["\t".join(_COLONNES_FEC)]
    for journal, num, date, libelle, details in _ecritures_synthetiques():
        assert sum(d for _, d, _ in details) == sum(c for _, _, c in details), \
            f"Écriture déséquilibrée : {num}"
        for compte, debit, credit in details:
            lignes.append("\t".join([
                journal, f"Journal {journal}", num, date,
                compte, COMPTES[compte], "", "",
                num, date, libelle,
                _montant_fr(debit), _montant_fr(credit),
                "", "", "20260115", "", "",
            ]))
    path = Path(path)
    path.write_text("\r\n".join(lignes) + "\r\n", encoding="cp1252")
    return path


def soldes_n1_synthetiques() -> Dict[str, int]:
    """Soldes N-1 en CENTIMES (somme = 0), comptes N + orphelins N-1."""
    soldes = {
        "101000": -30000000,
        "106100":  -3000000,
        # Report à nouveau débiteur (pertes antérieures) — équilibre la balance
        "110000":  23500000,
        "151000":  -2000000,
        "164000": -20000000,
        "205000":   6000000,
        "213500":  30000000,
        "215400":  14000000,
        "261000":   5000000,
        "275000":    800000,
        "280500":  -2200000,
        "281350":  -7500000,
        "281540":  -5600000,
        "310000":   6400000,
        "370000":   3100000,
        "391000":   -400000,
        "401100":  -4800000,
        "411100":   6900000,
        "491100":   -300000,
        "455100":  -5500000,
        "512100":   9000000,
        "530000":     80000,
        "519000":  -2800000,
        "486000":   1100000,
        "487000":   -500000,
        "467000":  -6580000,
        # Orphelins : présents en N-1, soldés en N
        "274000":   1300000,
        "168800":   -800000,
        # Résultat en cours N-1 (classes 6/7 non clôturées)
        "601000":  55000000,
        "641000":  27000000,
        "701000": -98000000,
        "695000":   1000000,
        "661100":    300000,
        "771800":   -500000,
    }
    assert sum(soldes.values()) == 0, "Balance N-1 synthétique déséquilibrée"
    return soldes


def generer_balance_n1_xlsx(path: Path) -> Path:
    """Écrit la balance N-1 synthétique (Excel simple, soldes en €)."""
    libelles = {**COMPTES, **COMPTES_N1_SEULS}
    wb = openpyxl.Workbook()
    ws = wb.active
    # Nom neutre (sans "bal") : la détection passe par le scan de contenu
    # ("CompteNum" dans les 10 premières lignes → mode balance simple).
    ws.title = "Export comptable"
    ws.append(["CompteNum", "CompteLib", "Solde"])
    for num, centimes in sorted(soldes_n1_synthetiques().items()):
        ws.append([num, libelles[num], centimes / 100.0])
    path = Path(path)
    wb.save(path)
    return path
