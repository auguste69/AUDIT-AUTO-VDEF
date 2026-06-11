# CLAUDE.md — Audit Automation Platform

## Contexte du projet

Ce projet automatise le processus d'audit d'un cabinet d'audit externe français. Le flux actuel est 100% manuel dans Excel et sujet à des erreurs de copier-coller. On remplace ce processus par un pipeline Python automatisé.

### Universalité

Le système fonctionne pour N'IMPORTE QUEL client disposant d'un FEC, quelle que soit la taille de l'entreprise, le secteur d'activité ou le logiciel comptable utilisé. Le mapping PCG dans `mapping_pcg.yaml` couvre exhaustivement TOUS les numéros de comptes possibles du Plan Comptable Général français (classes 1 à 7, ~300 préfixes). Tout numéro de compte trouvera un cycle d'audit et une classification bilan, sans intervention manuelle. Le système a été testé avec le client GILAC (270 comptes, industrie) mais est conçu pour fonctionner avec des FEC de toute taille (de la micro-entreprise aux groupes avec des milliers de comptes).

### Le flux métier (de A à Z)

1. **Réception du FEC** (.txt) — Fichier des Écritures Comptables, obligatoire en France (art. L.47 A LPF). Contient toutes les écritures comptables d'un exercice. Format : texte délimité (tab ou pipe), encodage variable (UTF-8, CP1252, Latin-1).
2. **Import et nettoyage** — Détection auto du séparateur et de l'encodage, conversion des montants (virgule française → point), ajout de la colonne Solde = Débit - Crédit.
3. **Construction de la balance générale** — Agrégation par CompteNum/CompteLib, somme des Débits, Crédits et Soldes. Contrôle : la somme de tous les soldes DOIT être = 0.
4. **Comparaison N vs N-1** — La balance N est comparée avec la balance de l'exercice précédent. Calcul des variations en valeur absolue (K€) et en pourcentage.
5. **Mapping par cycle d'audit** — Chaque compte est assigné à un cycle d'audit (A=Achats, V=Ventes, P=Personnel, etc.) selon le plan comptable général (PCG). Le mapping peut venir d'un fichier FM existant ou être déduit automatiquement par préfixe de compte.
6. **Génération des feuilles maîtresses** — Un fichier Excel avec : un sommaire, une balance comparative N vs N-1, et un onglet par cycle d'audit ventilant les comptes par Actif/Passif/Charges/Produits.
7. **Injection dans les templates de feuilles de travail** — Les templates du cabinet (un fichier .xlsx par cycle avec des sous-feuilles par procédure d'audit) sont copiés, renommés, et les placeholders (#NomClient, #DateClôture) sont remplacés. Les soldes sont injectés dans les cellules appropriées.

### Les 13 cycles d'audit

| Code | Nom | Exemples de comptes |
|------|-----|---------------------|
| C Propres | Capitaux propres | 101xxx, 106xxx, 120xxx |
| C PRC | Provisions pour risques et charges | 151xxx |
| F | Financement | 164xxx, 512xxx, 661xxx |
| I Incorp | Immobilisations incorporelles | 207xxx, 280xxx |
| I Corp | Immobilisations corporelles | 213xxx-218xxx, 281xxx |
| I Fi | Immobilisations financières | 261xxx, 274xxx, 275xxx |
| S | Stocks | 311xxx-397xxx, 713xxx |
| A | Achats (Fournisseurs) | 401xxx, 408xxx, 486xxx, 601xxx-629xxx, 681xxx |
| V | Ventes (Clients) | 411xxx-419xxx, 491xxx, 654xxx, 701xxx-709xxx |
| P | Personnel | 421xxx-438xxx, 641xxx-649xxx |
| E | État (fiscal/social) | 441xxx-449xxx, 631xxx-637xxx, 691xxx-699xxx |
| T | Autres créances/dettes | 461xxx-478xxx, 655xxx, 658xxx, 751xxx, 758xxx |
| X | Résultat exceptionnel | 671xxx-678xxx, 771xxx-778xxx |

### Les 18 colonnes standard du FEC

JournalCode, JournalLib, EcritureNum, EcritureDate (AAAAMMJJ), CompteNum, CompteLib, CompAuxNum, CompAuxLib, PieceRef, PieceDate, EcritureLib, Debit, Credit, EcritureLet, DateLet, ValidDate, Montantdevise, Idevise

### Classification Actif/Passif/Charges/Produits

- Classe 1 (10x-15x) → Passif (capitaux propres, provisions, emprunts)
- Classe 2 (20x-29x) → Actif (immobilisations)
- Classe 3 (31x-39x) → Actif (stocks)
- Classe 4 → Mixte selon le préfixe :
  - 401-408 → Passif (fournisseurs créditeurs)
  - 409 → Actif (fournisseurs débiteurs)
  - 411-418 → Actif (clients débiteurs)
  - 419 → Passif (clients créditeurs)
  - 421-449 → Passif (dettes sociales/fiscales)
  - 486 → Actif (CCA)
  - 487 → Passif (PCA)
  - 491 → Actif (provisions dépréciation clients)
- Classe 5 (50x-59x) → Actif (trésorerie)
- Classe 6 → Charges
- Classe 7 → Produits

**Attention** : certains comptes changent de sens selon leur solde. Un compte 425 (Personnel avances) est normalement Actif mais peut être Passif s'il est créditeur. Le FM existant du client fait foi pour la classification — le mapping automatique par PCG est un fallback.

---

## Architecture du projet

```
audit-automation/
├── src/
│   ├── __init__.py
│   ├── parsers/
│   │   ├── __init__.py
│   │   ├── fec_parser.py          # Lecture et nettoyage du FEC
│   │   ├── balance_parser.py      # Lecture balance N-1 (Excel, CSV, FEC)
│   │   └── mapping_parser.py      # Extraction mapping depuis FM ou config
│   ├── engine/
│   │   ├── __init__.py
│   │   ├── controls.py            # Contrôles d'intégrité du FEC
│   │   ├── balance_builder.py     # Construction balance + comparaison N/N-1
│   │   └── cycle_mapper.py        # Ventilation des comptes par cycle
│   ├── writers/
│   │   ├── __init__.py
│   │   ├── styles.py              # Styles Excel centralisés (Design A)
│   │   ├── excel_writer.py        # Fichier de travail (FEC + BG N + BG N-1)
│   │   ├── fm_writer.py           # Feuilles maîtresses
│   │   └── template_writer.py     # Injection dans templates feuilles de travail
│   └── config/
│       ├── __init__.py
│       └── mapping_pcg.yaml       # Mapping PCG → cycles (configurable)
├── tests/
│   ├── __init__.py
│   ├── test_fec_parser.py
│   ├── test_controls.py
│   ├── test_balance_builder.py
│   ├── test_cycle_mapper.py
│   └── test_writers.py
├── data/                           # Fichiers de test (ne pas commiter)
│   ├── GILAC_2025_FEC.txt
│   ├── FM_GILAC.xlsx
│   └── templates/
│       ├── 20XX_XX_YYY_A_Fournisseurs-Achats.xlsx
│       └── ...
├── app.py                          # Interface Streamlit
├── main.py                         # CLI
├── requirements.txt
├── CLAUDE.md                       # Ce fichier
└── README.md
```

### Principes architecturaux

1. **Séparation stricte** — Les parsers lisent, le moteur calcule, les writers écrivent. Aucun module ne fait deux choses.
2. **Zéro duplication** — L'app Streamlit et le CLI appellent exactement les mêmes modules. Pas de copier-coller de logique.
3. **Config externalisée** — Le mapping PCG et les seuils d'alerte sont dans des fichiers YAML, pas en dur dans le code.
4. **Testable** — Chaque module a ses propres tests unitaires qui tournent avec les vrais fichiers dans `data/`.

---

## Spécifications détaillées par module

### src/parsers/fec_parser.py

**Input** : chemin ou file-like object vers un FEC .txt
**Output** : pandas DataFrame avec colonnes typées

Fonctionnalités :
- Détection automatique de l'encodage : tester UTF-8, CP1252, Latin-1 sur le fichier ENTIER (pas juste les premiers octets — certains FEC ont des caractères spéciaux en fin de fichier)
- Détection automatique du séparateur : tab (\t), pipe (|), point-virgule (;)
- Nettoyage : strip sur toutes les colonnes string, suppression des retours chariot Windows (\r)
- Conversion : Debit et Credit en float (virgule → point), Montantdevise en float
- Ajout : colonne Solde = Debit - Credit
- Validation : vérifier que les 18 colonnes obligatoires sont présentes

### src/parsers/balance_parser.py

**Input** : fichier Excel balance N-1, ou FEC N-1, ou FM existant
**Output** : dict {compte_num: {'libelle': str, 'solde_ke': float}}

Trois modes :
1. `from_fec(path)` — lit un FEC, construit la balance, convertit en K€
2. `from_excel(path)` — lit un Excel avec colonnes CompteNum, CompteLib, Solde
3. `from_fm(path)` — extrait les soldes N depuis l'onglet "Balance N Vs N-1" d'un FM existant

### src/parsers/mapping_parser.py

**Input** : FM existant (.xlsx) ou fichier mapping_pcg.yaml
**Output** : dict {compte_num: {'cycle': str, 'etatfi': str, 'compta': str}}

Deux sources :
1. `from_fm(path)` — lit l'onglet "Balance N Vs N-1" du FM, colonnes : B=CompteNum (int/float), C=CompteLib, H=Ref cycle, J=Cycle, K=EtatFi N, M=ComptaN
2. `from_pcg_config(path)` — lit le YAML de mapping par préfixe

### src/engine/controls.py

**Input** : DataFrame FEC, date de clôture
**Output** : liste de tuples (nom_controle, ok: bool, detail: str, severity: str)

Contrôles à implémenter :
- Équilibre Débit/Crédit (BLOQUANT si écart > 0.01)
- Colonnes obligatoires présentes
- Lignes à zéro (D=0 et C=0)
- Cohérence des dates (toutes dans l'exercice)
- Écritures un dimanche (WARNING)
- Montants ronds ≥ 10K€ (INFO)
- Doublons potentiels (même journal + même numéro + même compte + mêmes montants + même libellé)
- Test de Benford sur le premier chiffre des montants (distribution attendue vs observée, calcul du chi² et du MAD)
- Écritures passées tard (après le 31/01/N+1 si ValidDate disponible)

### src/engine/balance_builder.py

**Input** : DataFrame FEC, dict balance N-1
**Output** : DataFrame avec CompteNum, CompteLib, Debit, Credit, Solde, Solde_KE, Solde_N1_KE, Var_KE, Var_PCT

Fonctionnalités :
- Agrégation par CompteNum + CompteLib
- Conversion en K€ (÷ 1000)
- Calcul des variations vs N-1
- Var_PCT = 'n/a' si abs(N-1) < 0.001 (évite division par zéro)
- Vérification : somme des soldes = 0

### src/engine/cycle_mapper.py

**Input** : DataFrame balance, dict mapping existant, config YAML
**Output** : DataFrame balance enrichi avec colonnes cycle, compta, etatfi, ref

Logique de résolution (ordre de priorité) :
1. Si le compte existe dans le mapping du FM existant du client → utiliser ce mapping (priorité absolue)
2. Sinon → surcharges_cabinet du YAML (préfixes spécifiques au cabinet)
3. Sinon → chercher par préfixe dans le YAML, du plus spécifique au plus général :
   - Préfixe 3 chiffres (ex: "486" → A, "603" → S)
   - Préfixe 2 chiffres (ex: "48" → T, "60" → A)
   - Classe 1 chiffre — fallback garanti (ex: "6" → charges, "4" → T)
4. Le YAML couvre EXHAUSTIVEMENT le PCG français (classes 1-7). TOUT numéro de compte trouvera un match — il n'y a plus de fallback "T" arbitraire.
5. Classification Actif/Passif : utiliser la section classification_bilan du YAML pour la classe 4 (mixte). Classes 1 → Passif, 2-3-5 → Actif, 6 → Charges, 7 → Produits.

Le fichier mapping_pcg.yaml est structuré par classe (classe_1, classe_2, etc.). Le parser doit fusionner toutes les classes en un seul dict flat pour la résolution.

### src/writers/styles.py

**Design A — Minimaliste**. Zéro couleur de fond, hiérarchie par typographie seule.

Styles centralisés :
- `FONT_TITLE` : Arial 16, bold
- `FONT_SUBTITLE` : Arial 11, italic, color #808080
- `FONT_HEADER` : Arial 10, bold, color #808080
- `FONT_NORMAL` : Arial 10
- `FONT_BOLD` : Arial 10, bold
- `FONT_SECTION` : Arial 9, bold, color #999999 (pour les labels ACTIF/PASSIF/etc.)
- `FONT_META` : Arial 9, color #AAAAAA (pour les colonnes secondaires : Ref, Cycle, EtatFi)
- `BORDER_BOTTOM_MED` : trait medium noir (sous les en-têtes)
- `BORDER_BOTTOM_HAIR` : trait hairline #C0C0C0 (entre les lignes de données)
- `BORDER_TOP_THIN` : trait thin noir (au-dessus des totaux)
- `NUM_KE` : format `#,##0;(#,##0);"-"` — entiers, négatifs entre parenthèses, zéros = "-"
- `NUM_PCT` : format `0%` — pourcentages sans décimales

Fonctions utilitaires :
- `remove_gridlines(ws)` — supprime le quadrillage Excel
- `set_col_widths(ws)` — largeurs standard
- `write_header_row(ws, row, headers)` — en-tête avec trait épais
- `write_data_row(ws, row, ...)` — ligne de données avec trait fin
- `write_total_row(ws, row, ...)` — ligne total avec trait au-dessus, bold

**IMPORTANT** :
- Pas de PatternFill nulle part (aucune couleur de fond)
- Pas de datetime dans les en-têtes, uniquement des strings "31/12/2025"
- Montants arrondis à l'entier (pas de décimales), format K€
- Pas de couleur sur les variations (même significatives)

### src/writers/fm_writer.py

**Output** : fichier FM_{client}_{annee}.xlsx

Onglets à générer :
1. **Sommaire** — titre, nom client, date, liens vers les onglets
2. **Balance N Vs N-1** — tous les comptes triés par numéro, colonnes : Ref, CompteNum, CompteLib, Solde N, Solde N-1, Var K€, Var %, Ref cycle, Cycle, EtatFi N, EtatFi N-1, ComptaN, ComptaN-1
3. **Un onglet par cycle** (13 onglets) — nom du format "A0", "V0", "C_Propres0", etc. Même structure que la balance mais filtré par cycle, avec sous-sections ACTIF/PASSIF/CHARGES/PRODUITS et totaux par section

### src/writers/excel_writer.py

**Output** : fichier Travail_{client}_{annee}.xlsx

Onglets :
1. **FEC** — le FEC brut complet avec en-têtes stylés
2. **Balance N** — CompteNum, CompteLib, Débit, Crédit, Solde, Solde K€
3. **Balance N-1** — CompteNum, CompteLib, Solde K€

### src/writers/template_writer.py

**Output** : ZIP avec les 13 fichiers de feuilles de travail

Fonctionnalités :
- Copie chaque template, renomme (20XX_XX_YYY → {annee}_12_{client})
- Remplace #NomClient par le nom du client
- Remplace #DateClôture par la date au format JJ/MM/AAAA
- Supprime les quadrillages
- **Injection des soldes** : dans chaque feuille Synthèse, renseigner les soldes de la feuille maîtresse correspondante dans les cellules prévues

Mapping template → cycle :
- `A_Fournisseurs` → A
- `V_Clients` → V
- `P_Personnel` → P
- `E_Etat` → E
- `F_Financement` → F
- `S_Stocks` → S
- `T_Autres` → T
- `I_Immobilisations_corp` → I Corp
- `I_Immobilisations_incorporelles` → I Incorp
- `I_Immobilisations_financie` → I Fi
- `C_Capitaux` → C Propres
- `C_Provision` → C PRC
- `X_Re_sultat` → X

---

## Structure du FM existant (FM_GILAC.xlsx) — à utiliser comme référence

L'onglet "Balance N Vs N-1" a cette structure (row 8 = headers, data à partir de row 10) :

| Colonne | Index | Contenu |
|---------|-------|---------|
| A | 0 | Ref. (vide pour les lignes de données) |
| B | 1 | CompteNum (int ou float, ex: 101300) |
| C | 2 | CompteLib (string) |
| D | 3 | Solde N en K€ (float) |
| E | 4 | Solde N-1 en K€ (float) |
| F | 5 | Variation K€ |
| G | 6 | Variation % |
| H | 7 | Ref. cycle (ex: "C Propres0", "A0") |
| I | 8 | (vide) |
| J | 9 | Cycle (ex: "C Propres", "A", "V") |
| K | 10 | EtatFi N (ex: "Capital social", "Réserves") |
| L | 11 | EtatFi N-1 |
| M | 12 | ComptaN (ex: "Passif", "Actif", "Charges") |
| N | 13 | ComptaN-1 |

Les onglets cycle (A0, V0, etc.) ont la même structure avec en plus des sous-sections ACTIF/PASSIF/CHARGES/PRODUITS (label en colonne B, pas de données dans les colonnes numériques).

---

## Données de test (client GILAC — exemple, pas une limitation)

Ces données servent à tester le pipeline. Le système fonctionne avec n'importe quel FEC français.

- **FEC 2025** : 106 272 lignes, encodage CP1252, séparateur TAB, dates du 01/01/2025 au 31/12/2025
- **Balance 2025** : 270 comptes, équilibre parfait (écart = 0.00), résultat = -2 572 709 € (-2 573 K€)
- **Balance 2024 (N-1)** : 288 comptes (extraits du FM_GILAC existant qui couvre 2024 vs 2023)
- **29 nouveaux comptes en 2025** (absents du mapping 2024) : 106810, 231500, 401500, 401600, 409100, 411500, 411600, 445661, 455030, 486001, 512113, 512153, 512155, 604000, 607030, 611020, 616170, 621100, 622601, 623202, 623203, 637810, 641430, 649000, 671100, 758700, 762400, 767010, 768010
- **13 templates de feuilles de travail** avec entre 2 et 17 onglets chacun

---

## Commandes utiles

```bash
# Installer les dépendances
pip3 install -r requirements.txt

# Lancer les tests
python3 -m pytest tests/ -v

# Lancer le CLI
python3 main.py data/GILAC_2025_FEC.txt --client GILAC --date-cloture 31/12/2025 --n1-fm data/FM_GILAC.xlsx --templates data/templates/ --output output/

# Lancer l'interface Streamlit
python3 -m streamlit run app.py
```

---

## Règles de codage

- Python 3.9+ (compatibilité Mac du stagiaire)
- Type hints sur toutes les fonctions publiques
- Docstrings en français
- Pas de print() — utiliser le module logging
- Noms de variables en français pour le domaine métier (solde_ke, compte_num, cycle_audit)
- Noms techniques en anglais (parse, build, write, config)
- Chaque fonction fait une seule chose
- Pas de try/except silencieux — logger les erreurs
- Les DataFrames ne sont jamais modifiés en place — toujours retourner une copie

---

## RÈGLES AGENTS AUTO-EXÉCUTÉS (prompt_runner)

Ces règles s'appliquent quand le projet est exécuté par `scripts/prompt_runner.py`
(agents Claude Code non-interactifs, `--dangerously-skip-permissions`).

1. **Complétude** : implémenter entièrement l'objectif du prompt. Aucun TODO, aucun placeholder, aucune implémentation partielle.
2. **Tests** : avant de terminer, s'assurer que `python3 -m pytest tests/ -q` ne régresse pas par rapport au baseline.
3. **Scope strict** : ne modifier que le code directement nécessaire à l'objectif. Aucun refactoring non demandé.
4. **Zéro hardcode client** : aucune valeur GILAC en dur (voir "Règles de codage").
5. **Vérification** : relire le prompt après implémentation pour confirmer chaque point des `<validation>`.

---

## Ordre de construction recommandé

1. `src/config/mapping_pcg.yaml` + `src/parsers/mapping_parser.py`
2. `src/parsers/fec_parser.py` + `tests/test_fec_parser.py`
3. `src/engine/balance_builder.py` + `tests/test_balance_builder.py`
4. `src/engine/controls.py` + `tests/test_controls.py`
5. `src/engine/cycle_mapper.py` + `tests/test_cycle_mapper.py`
6. `src/writers/styles.py`
7. `src/writers/excel_writer.py`
8. `src/writers/fm_writer.py`
9. `src/writers/template_writer.py` + `tests/test_writers.py`
10. `main.py` (CLI)
11. `app.py` (Streamlit)
12. Tests d'intégration end-to-end
# Audit Automation — Document de référence complet

**Date :** 26/03/2026  
**Projet :** audit-automation  
**Stack :** Python 3.9+ · Streamlit · openpyxl · pandas · PyYAML  
**Taille du code actuel :** ~2 500 lignes (hors tests, config)  
**Cabinet :** AVIZEO — Expert-comptable en Rhône-Alpes

**⚠️ RAPPEL FONDAMENTAL :** Ce projet est UNIVERSEL. Il fonctionne pour N'IMPORTE QUEL FEC de N'IMPORTE QUELLE entreprise française. GILAC est un jeu de test, pas une cible. Aucune valeur, aucun libellé, aucun numéro de compte ne doit être codé en dur pour un client spécifique. Tous les documents (FM GILAC, FEC GILAC) servent d'exemple de format/structure uniquement. Seuls les templates de feuilles de travail sont des fichiers de production.

---

# PARTIE 1 — DIAGNOSTIC ET PRIORISATION (Mission 1)

## 1.1 Tableau de diagnostic

| ID | Problème | Priorité | Faisabilité | Solution proposée | Effort |
|---|---|---|---|---|---|
| **P1** | Les feuilles maîtresses ne sont pas insérées comme onglets dans les templates | **Critique** | Modérée | Copier chaque onglet FM comme nouvelle feuille dans le template correspondant, avant la feuille "Synthèse" | 3-4 jours |
| **P3** | Les 3 livrables ne sont pas tous générés (manquent Bilan, EBIT, Actif/Passif/P&L détaillés, Tréso, AACE, et le fichier FEC+Balances) | **Critique** | Complexe | Créer `excel_writer.py` pour livrable 1, ajouter 7 onglets dans `fm_writer.py` pour livrable 2, coupler P1 pour livrable 3 | 5-7 jours |
| **P5** | L'ingestion N-1 ne détecte pas automatiquement la balance dans un classeur multi-feuilles | **Haute** | Simple | Scanner les feuilles du classeur et identifier celle contenant la balance (heuristique colonnes CompteNum/Solde) | 1 jour |
| **P6** | Doublons de comptes N vs N-1 (rapprochement de numéros de comptes qui changent entre exercices) | **Haute** | Complexe | Algorithme de matching par préfixe + libellé + validation utilisateur interactive (modal Streamlit) | 3-4 jours |
| **P4** | Design des FM et de l'interface web | **Moyenne** | Modérée | Refonte styles.py pour reproduire le format FM exemple, customisation Streamlit (CSS, logo, couleurs AVIZEO) | 2-3 jours |
| **P2** | Templates "Concessions Automobiles" | **Basse** | Simple | Toggle Streamlit + dossier `data/templates_concess/` + logique de sélection dynamique (seuls S_Stocks et A_Fournisseurs changent) | 1 jour |

---

## 1.2 Détail par problème

### P1 — Intégration des feuilles maîtresses dans les templates (CRITIQUE)

**Situation actuelle :** Le `template_writer.py` injecte les données de balance du cycle *à la fin de la feuille Synthèse existante*, en ajoutant un tableau en bas de page. Ce n'est PAS le comportement attendu.

**Comportement attendu :** Chaque feuille maîtresse générée doit être copiée comme un **nouvel onglet Excel** dans le template correspondant, positionné **juste AVANT la feuille "Synthèse"**.

**Mapping feuille → template confirmé :**

| Feuilles FM à insérer | Template cible | Feuille "Synthèse" |
|---|---|---|
| AACE, A0 | `_A_Fournisseurs-Achats.xlsx` | "A Synthèse" |
| C Propres0 | `_C_Capitaux propres.xlsx` | "C Synthèse" |
| C PRC0 | `_C_Provision pour risques et charges.xlsx` | "C Synthèse" |
| E0 | `_E_Etat.xlsx` | "E Synthèse" |
| Tréso, F0 | `_F_Financement.xlsx` | "F Synthèse" |
| I Corp0, I Incorp0 | `_I_Immobilisations corp.xlsx` | "I Synthèse" |
| *(aucune)* | `_I_Immobilisations incorporelles.xlsx` | "I Synthèse" |
| I Fi0 | `_I_Immobilisations financières.xlsx` | "I Synthèse" |
| P0 | `_P_Personnel.xlsx` | "P Synthèse" |
| S0 | `_S_Stocks.xlsx` | "S Synthèse" |
| T0 | `_T_Autres créances - dettes.xlsx` | "T Synthèse" |
| V0 | `_V_Clients-Ventes.xlsx` | "V Synthèse" |
| X0 | `_X_Résultat exceptionnel.xlsx` | "X Synthèse" |

**Note importante :** Le template `_I_Immobilisations corp.xlsx` reçoit les deux feuilles (I Corp0 ET I Incorp0) — les immos corp et incorp sont dans le même fichier template (confirmé). Le template `_I_Immobilisations incorporelles.xlsx` existe dans `data/templates/` (13 fichiers au total) mais ne reçoit AUCUNE feuille FM : il est copié/renommé avec remplacement des placeholders uniquement. Les clés du mapping `integration_templates` doivent correspondre aux noms de fichiers RÉELS de `data/templates/` (ex. "Provision" au singulier).

**Complexité technique :** openpyxl ne supporte pas nativement la copie d'onglets entre workbooks. Il faut recréer l'onglet cellule par cellule en copiant valeurs, styles, fusions et dimensions.

**⚠️ Risque images/dessins (vérifié) :** 9 des 13 templates contiennent des images ou dessins (jusqu'à 12 éléments dans `E_Etat.xlsx`). openpyxl perd les *shapes*/dessins à la simple ouverture+sauvegarde d'un fichier, et la préservation des images/graphiques connaît des régressions selon les versions. Toute évolution de `template_writer.py` doit inclure un test automatique comptant les médias (`xl/media/*`, `xl/drawings/*` dans l'archive ZIP) avant/après génération, et toute perte doit être détectée et documentée.

**Solution technique :**
1. Modifier `template_writer.py` : supprimer `_injecter_balance_cycle()` (approche incorrecte)
2. Créer `_inserer_onglet_fm(wb_template, nom_onglet, df_cycle, cycle, ...)` qui reproduit la structure d'un onglet cycle du FM et positionne la feuille avant "Synthèse" via `wb.move_sheet()`
3. Étendre le mapping pour supporter les feuilles multi-cycle (AACE + A0 dans Achats, Tréso + F0 dans Financement)

---

### P3 — Livrables / Export (CRITIQUE)

**Situation actuelle :** Le pipeline produit 2 fichiers (FM + ZIP templates). Il manque le fichier de travail et 7 onglets dans le FM.

#### Les 3 livrables attendus :

**Livrable 1 — Fichier de travail `Travail_{client}_{annee}.xlsx`**
Module : `excel_writer.py` (à créer — mentionné dans CLAUDE.md mais inexistant)

| Onglet | Contenu |
|---|---|
| FEC N | FEC brut complet (18 colonnes + Solde), headers stylés |
| Balance N | CompteNum, CompteLib, Débit, Crédit, Solde, Solde K€ |
| Balance N-1 | CompteNum, CompteLib, Solde K€ |

**Livrable 2 — FM complet `FM_{client}_{annee}.xlsx`**
Module : `fm_writer.py` (à enrichir — 7 onglets manquants)

| Onglet | Statut | Description |
|---|---|---|
| Sommaire | ✅ Existant | |
| Balance N Vs N-1 | ✅ Existant | |
| Bilan | ❌ À créer | Bilan synthétique Actif/Passif côte à côte |
| EBIT | ❌ À créer | Compte de résultat jusqu'au résultat d'exploitation |
| Actif détaillé | ❌ À créer | Détail de l'actif ligne par ligne (format liasse 2050) |
| Passif détaillé | ❌ À créer | Détail du passif ligne par ligne (format liasse 2051) |
| P&L détaillé | ❌ À créer | Compte de résultat détaillé (format liasse 2052/2053) |
| Tréso | ❌ À créer | BFR, FRNG, Trésorerie nette |
| AACE | ❌ À créer | Détail des Autres Achats et Charges Externes |
| C Propres0 … X0 | ✅ Existants | Les 13 cycles d'audit |

Ordre : Sommaire → Balance → Bilan → EBIT → Actif détaillé → Passif détaillé → P&L détaillé → Tréso → AACE → [cycles dans l'ordre canonique]

**Livrable 3 — ZIP templates `FT_{client}_{annee}.zip`**
Module : `template_writer.py` (à refactorer pour P1)
Chaque template avec ses feuilles FM insérées comme onglets AVANT la Synthèse.

---

### P5 — Ingestion fichier N-1 multi-feuilles (HAUTE)

**Problème :** L'utilisateur peut uploader un fichier "Feuilles maîtresses N-1" contenant de nombreuses feuilles (Sommaire, Balance, Bilan, EBIT, A0, V0, etc.). Le code doit identifier automatiquement la feuille contenant la Balance.

**Logique de détection (ordre de priorité) :**
1. Chercher un onglet nommé exactement "Balance N Vs N-1" → mode FM complet (extraire mapping + soldes N comme soldes N-1)
2. Chercher un onglet dont le nom contient "balance" (insensible à la casse) → mode FM probable
3. Chercher un onglet dont le nom contient "bal" → fallback
4. Scanner chaque feuille : chercher une colonne contenant "CompteNum" ou "Compte" ou "Num" dans les 10 premières lignes → mode balance simple
5. Si aucune feuille détectée → lever une erreur explicite avec la liste des onglets disponibles

**Extraction des soldes N-1 :**
- Si feuille "Balance N Vs N-1" trouvée (structure FM) : les soldes sont en colonne D (Solde N), déjà en K€
- Si feuille balance simple trouvée : les soldes sont en €, diviser par 1000
- Si FEC N-1 (.txt) : parser et agréger comme actuellement

**Module impacté :** `mapping_parser.py`

---

### P6 — Rapprochement de comptes N vs N-1 (HAUTE) — ✅ IMPLÉMENTÉ (prompt 12)

**Statut : implémenté.** Modules : `src/engine/account_matcher.py`
(detecter_orphelins, scorer_matching, proposer_rapprochements,
appliquer_rapprochements), `src/models/rapprochement.py` (dataclass
Rapprochement), section `rapprochements` du `mapping_pcg.yaml` (seuil,
poids, mots_vides). Intégration : `main.py` (validation CLI interactive par
défaut, option `--rapprochements-auto`) et `app.py` (tableau interactif
avec "Tout valider" / "Tout ignorer", choix ligne par ligne prioritaires).
Ordre pipeline : parse → load_n1 → MATCHING → build → map_cycles →
fm_writer. Le cycle/classification du compte N est dérivé à la volée par
résolution de préfixe PCG (logique de cycle_mapper). Aucune fusion sans
confirmation explicite. Tests : `tests/test_account_matcher.py`,
`tests/test_account_matcher_e2e.py`.

**Problème :** Un compte peut changer de numéro entre N-1 et N (ex: "512003" → "5123001"). Actuellement les deux sont traités comme des comptes différents, créant des doublons et faussant les variations.

**Algorithme en 3 phases :**

**Phase 1 — Détection des orphelins :**
- Comptes N-1 sans correspondance exacte en N → liste `orphelins_n1`
- Comptes N sans correspondance exacte en N-1 → liste `orphelins_n`
- Si les deux listes sont vides → aucun rapprochement nécessaire

**Phase 2 — Scoring de matching :**
Pour chaque paire (orphelin_n1, orphelin_n), score composite :
- Préfixe commun (poids 40%) : longueur du plus long préfixe commun ÷ longueur max
- Similarité libellé (poids 35%) : ratio de tokens communs (insensible casse, hors mots vides)
- Même cycle (poids 15%) : 1.0 si même cycle, 0.0 sinon
- Même classification bilan (poids 10%) : 1.0 si même Actif/Passif/Charges/Produits, 0.0 sinon
- Seuil minimal : score ≥ 0.5 pour proposer un rapprochement

**Phase 3 — Validation utilisateur (OBLIGATOIRE) :**
Aucune fusion sans confirmation explicite.
- Interface Streamlit : tableau interactif avec colonnes [Compte N-1 | Libellé N-1 | → | Compte N | Libellé N | Score | Action]
- **Affichage ligne par ligne** : chaque rapprochement proposé est affiché individuellement, l'utilisateur peut choisir ✅ Fusionner ou ❌ Ignorer pour chaque paire
- **Bouton "Tout valider"** : en bas du tableau, applique ✅ Fusionner à toutes les lignes d'un coup (raccourci quand l'utilisateur fait confiance aux scores)
- Bouton "Tout ignorer" : passe sans aucune fusion
- Les choix ligne par ligne priment sur "Tout valider" (si l'utilisateur a déjà marqué certaines lignes comme ❌ Ignorer, elles restent ignorées)
- Logger toutes les fusions pour traçabilité

**Module à créer :** `src/engine/account_matcher.py`

---

### P4 — Design / UX (MOYENNE)

**Refonte des styles FM :**
- Les dates en headers doivent être des `datetime` Excel (pas des strings)
- Les pourcentages en format `0%` (float décimal)
- Ajouter des lignes de total par section
- Colonnes EtatFi N et EtatFi N-1 distinctes (le code actuel met la même valeur)
- Reproduire les 14 colonnes A-N du FM GILAC

**Refonte interface Streamlit :**
- Header avec logo AVIZEO et titre "Audit Automation — AVIZEO"
- Palette : bleu principal `#2E75B6`, bleu foncé `#1B4F7A`, bleu clair `#D6E8F7`
- CSS custom via `st.markdown(unsafe_allow_html=True)`
- Sidebar stylée aux couleurs AVIZEO
- Favicon personnalisé

---

### P2 — Templates Concessions Automobiles (BASSE)

**Règle confirmée :** Seuls 2 templates changent quand l'entreprise est une concession auto :
- `_S_Stocks.xlsx` → remplacé par le template concession Stocks
- `_A_Fournisseurs-Achats.xlsx` → remplacé par le template concession Achats

**Templates concession — différences structurelles :**

| Template | Standard | Concession |
|---|---|---|
| **S_Stocks** | S Synthèse, S100-S710 (13 onglets) | S Synthèse, Questionnaire, S100, S200, S400, S300, S500, S610 Dép VN, S620 Dép VO, S630 Dép PR (10 onglets) |
| **A_Fournisseurs** | A Synthèse, A100-A800 (14 onglets) | A Synthèse, Questionnaire, A100-A700 + A200 Rappro constructeur, A430 Revue Primes const (13 onglets) |

**Implémentation :**
- `data/templates_concess/` avec les 2 templates concession
- Toggle Streamlit : "Concession automobile ?" (False par défaut)
- Si True → remplacer uniquement les 2 templates concernés, les 11 autres restent standards
- Le mapping cycle reste identique (A → A, S → S)

---

# PARTIE 2 — SPÉCIFICATIONS TECHNIQUES DES MODULES MANQUANTS

## 2.1 Branding — AVIZEO

**Nom :** AVIZEO  
**Baseline :** Expert-comptable & plus encore  
**Localisation :** Rhône-Alpes (Chambéry, Lyon, Villefranche-sur-Saône, Montrevel-en-Bresse)  
**Site :** https://www.avizeo.eu  

**Charte graphique :**
- Bleu principal : `#2E75B6`
- Bleu foncé : `#1B4F7A`
- Bleu clair : `#D6E8F7`
- Gris texte : `#3D3D3D`
- Gris secondaire : `#808080`
- Blanc : `#FFFFFF`
- Police : Arial

**Logo :** fichier PNG fourni, à intégrer dans le header Streamlit et en filigrane optionnel sur les exports.

---

## 2.2 Onglet "AACE" (Autres Achats et Charges Externes)

**Nature :** Détail compte par compte de tous les AACE (sous-ensemble du cycle A).  
**Filtre :** Comptes dont le préfixe commence par 606, 607 (hors marchandises pures), 608, 609, 61, 62.  
**Structure :** Identique aux onglets cycle (headers row 8, données row 10+), avec :
- Colonnes : Ref, CompteNum, CompteLib, Solde N K€, Solde N-1 K€, Var K€, Var %, Ref cycle
- Pas de sous-sections ACTIF/PASSIF — tout est "Charges"
- Convention de signe : charges en positif (comptes classe 6 débiteurs)

---

## 2.3 Onglet "Tréso" (BFR, FRNG, Trésorerie nette)

**Formules confirmées :**

```
FRNG = Ressources stables − Emplois stables

Ressources stables :
  - Capitaux propres (comptes 10x, 11x, 12x, 13x, 14x)
  - Amortissements & dépréciations (comptes 28x, 29x, 39x, 49x)
  - Provisions pour risques & charges (comptes 15x)
  - Dettes financières MLT (comptes 16x, 17x)

Emplois stables :
  - Actif immobilisé brut (comptes 20x-27x bruts, i.e. avant amortissements)

BFR = Actif circulant d'exploitation − Passif circulant d'exploitation

Actif circulant d'exploitation :
  - Stocks bruts (comptes 31x-37x)
  - Créances clients brutes (comptes 41x)
  - Autres créances d'exploitation (comptes 40x débiteurs, 44x débiteurs, 46x-47x débiteurs)
  - Charges constatées d'avance exploitation (486)

Passif circulant d'exploitation :
  - Dettes fournisseurs (comptes 40x créditeurs)
  - Dettes fiscales & sociales (comptes 42x-43x, 44x créditeurs)
  - Autres dettes d'exploitation (comptes 46x-47x créditeurs)
  - Produits constatés d'avance exploitation (487)

TN = FRNG − BFR

Vérification : TN = Disponibilités + VMP − CBC − Soldes créditeurs banques
  - Disponibilités (comptes 51x, 53x)
  - VMP (comptes 50x)
  - Concours bancaires courants (compte 519)
  - Soldes créditeurs de banques (512 si créditeur)

Cohérence fondamentale : FRNG = BFR + TN
```

**Structure de l'onglet :** Lignes agrégées par poste (pas compte par compte), colonnes N, N-1, Var K€, Var %.

---

## 2.4 Onglet "Bilan"

**Structure :** Actif à gauche (colonnes B-G), Passif à droite (colonnes I-N), côte à côte.

**Mapping Actif (basé sur le cerfa 2050) :**

| Ligne bilan | Racines de comptes |
|---|---|
| Capital souscrit non appelé | 109 |
| Immobilisations incorporelles | 201, 203, 205, 206, 207, 208, 232, 237 (brut) − 280x, 290x (amort/dép) |
| Immobilisations corporelles | 211-218, 231, 238 (brut) − 281x, 291x (amort/dép) |
| Immobilisations financières | 261-275, 276 (brut) − 296x, 297x (dép) |
| Stocks et en-cours | 31x-37x (brut) − 39x (dép) |
| Avances et acomptes versés | 4091 |
| Créances clients | 411, 413, 416, 418 (brut) − 491 (dép) |
| Autres créances | 4096-4098, 425, 441, 444, 4456, 451, 455, 456, 462, 465, 467, 478 |
| VMP | 50x − 590x |
| Disponibilités | 51x, 53x |
| Charges constatées d'avance | 486 |

**Mapping Passif (basé sur le cerfa 2051) :**

| Ligne bilan | Racines de comptes |
|---|---|
| Capital | 101, 108 |
| Primes d'émission, fusion, apport | 104 |
| Écarts de réévaluation | 105 |
| Réserve légale | 1061 |
| Réserves statutaires | 1063 |
| Réserves réglementées | 1062, 1064 (choix cabinet : 1062 "réserves indisponibles" rattaché ici par convention — à documenter dans le YAML) |
| Autres réserves | 1068 |
| Report à nouveau | 11x |
| Résultat de l'exercice | 12x |
| Subventions d'investissement | 13x |
| Provisions réglementées | 14x |
| Provisions pour risques | 151 |
| Provisions pour charges | Autres 15x |
| Emprunts établissements de crédit | 164, 512 créditeur, 517, 519 |
| Emprunts et dettes financières | 162, 165, 166, 168, 17x, 426, 45x |
| Avances reçues sur commandes | 4191 |
| Dettes fournisseurs | 401, 403, 4081, 4088 |
| Dettes fiscales et sociales | 421-427, 43x, 442, 444, 4455, 4457, 446, 447, 449, 457 |
| Dettes sur immobilisations | 404, 405, 4084 |
| Autres dettes | 4196-4198, 464, 467, 478, 509 |
| Produits constatés d'avance | 487 |

---

## 2.5 Onglet "EBIT" (Résultat d'exploitation)

**Mapping P&L (basé sur le cerfa 2052) :**

| Ligne P&L | Racines de comptes |
|---|---|
| **Produits d'exploitation** | |
| Ventes de marchandises | 707, 7097 (RRR accordés sur ventes de marchandises, en déduction) |
| Production vendue (Biens et Services) | Autres 70x (701-706, 708 y compris 7087, 709 sauf 7097) |
| **= Chiffre d'affaires** | Somme des deux ci-dessus |
| Production stockée | 713 |
| Production immobilisée | 72x |
| Subventions d'exploitation | 74x |
| Reprises amort./prov., transferts charges | 781, 791 |
| Autres produits | 75x (sauf 755) |
| **Total produits exploitation** | |
| **Charges d'exploitation** | |
| Achats de marchandises | 607, 6087, 6097 |
| Variation stocks marchandises | 6037 |
| Achats matières premières | 601, 602, 6081, 6082, 6091, 6092 |
| Variation stocks matières | 6031, 6032 |
| Autres charges externes | 604-606, 61x, 62x, 6084-6096 |
| Impôts et taxes | 63x |
| Salaires et traitements | 641, 644 |
| Charges sociales | 645-649 |
| Dotations amortissements (ligne GA) | 6811, 6812 |
| Dotations dépréciations immobilisations (ligne GB) | 6816 |
| Dotations dépréciations actif circulant (ligne GC) | 6817 |
| Dotations provisions (ligne GD) | 6815 |
| Autres charges | 65x (sauf 655) |
| **Total charges exploitation** | |
| **= Résultat d'exploitation** | Produits − Charges |

---

## 2.6 Onglets "Actif détaillé" et "Passif détaillé"

Format liasse fiscale 2050/2051, une ligne par poste, colonnes N/N-1/Var K€/Var %. Utilisent les mappings Actif et Passif du Bilan ci-dessus, avec le détail par sous-poste.

---

## 2.7 Onglet "P&L détaillé"

Format liasse fiscale 2052/2053 complet :
- Résultat d'exploitation (= EBIT ci-dessus)
- Résultat financier (produits financiers 76x − charges financières 66x)
- Résultat courant avant impôts
- Résultat exceptionnel (77x − 67x)
- Participation des salariés (691)
- Impôts sur les bénéfices (695, 698 intégration fiscale, 699 carry-back en déduction)
- **= Résultat net comptable**

---

# PARTIE 3 — VENTILATION DES TÂCHES ET SOUS-AGENTS (Mission 2)

## 3.1 Architecture en blocs

| Bloc | Tâches | Sous-agent | Modèle Claude | Justification |
|---|---|---|---|---|
| **A — Architecture & Specs** | CLAUDE.md v2, specs fonctionnelles, mapping liasse fiscale | Architecte | **Opus** | Décisions structurantes, compréhension métier profonde |
| **B — Parsers & Ingestion** | P5 (détection balance auto), nouveau `balance_parser.py` | Parseur | **Sonnet** | Code standard, patterns bien définis |
| **C — Engine / Calculs** | P6 (rapprochement comptes), mapping liasse fiscale, calculs BFR/FR/Trésorerie | Calculateur | **Opus** | Logique métier complexe, algorithme de matching |
| **D — Writers / Excel** | P1 (intégration templates), P3 livrable 1, P3 livrable 2, P4 (styles) | Rédacteur Excel | **Sonnet** | Code openpyxl répétitif, volume important |
| **E — Templates** | P1 (copie onglets entre workbooks), P2 (concessions auto) | Intégrateur | **Sonnet** | Manipulation openpyxl, patterns identifiés |
| **F — UI / Streamlit** | P4 (refonte interface), P6 (modal validation), P2 (toggle concessions), branding AVIZEO | Frontendeur | **Sonnet** | Streamlit simple, CSS custom standard |
| **G — Tests & QA** | Tests unitaires, test end-to-end, validation vs FM GILAC | Testeur | **Haiku** | Tâches simples et volumineuses |
| **H — DevOps & Packaging** | requirements.txt, Dockerfile, documentation | Ops | **Haiku** | Tâches standard |

## 3.2 Dépendances

```
A (Specs) ──→ B (Parsers) ──→ C (Engine) ──→ D (Writers) ──→ E (Templates)
                                    │                              │
                                    └──→ F (UI) ←─────────────────┘
                                                      │
                                              G (Tests) ←── tous les blocs
                                                      │
                                              H (DevOps)
```

## 3.3 Estimation globale

| Bloc | Effort | Sprint |
|---|---|---|
| A — Architecture | 1 jour | Sprint 0 |
| B — Parsers | 1-2 jours | Sprint 1 |
| C — Engine | 4-5 jours | Sprint 1-2 |
| D — Writers | 5-7 jours | Sprint 2-3 |
| E — Templates | 3-4 jours | Sprint 3 |
| F — UI | 2-3 jours | Sprint 3 |
| G — Tests | 2-3 jours | Continu |
| H — DevOps | 0.5 jour | Sprint 4 |
| **Total** | **19-26 jours** | |

---

# PARTIE 4 — RECOMMANDATION STRATÉGIQUE (Mission 3)

## Verdict : CONTINUER — ne PAS repartir de zéro

**Raisons :**

1. **Architecture solide** — Séparation parsers/engine/writers propre et respectée. Aucun module ne fait deux choses.
2. **CLAUDE.md excellent** — 350+ lignes couvrant le domaine métier, l'architecture, les specs et les conventions.
3. **Pipeline fonctionnel** — FEC parsé, balance construite, contrôles exécutés, cycles mappés, FM généré. Le socle est là.
4. **Problèmes = AJOUTS, pas corrections** — P1 est un changement de stratégie d'intégration. P3 = modules manquants. P5/P6 = améliorations. Le code existant n'est pas "cassé".
5. **Coût d'un restart** — 10+ jours juste pour revenir au point actuel (parser FEC, mapping PCG 300+ préfixes, 9 contrôles, cycle mapper).

**Il n'est PAS nécessaire de réexpliquer le processus d'audit.** Le CLAUDE.md est déjà très complet.

**Plan d'action :**
- Phase 1 : Enrichir CLAUDE.md (0.5 jour) — ajouter ce document comme addendum
- Phase 2 : Refactoring template_writer (1 jour) — supprimer l'injection en fin de Synthèse, remplacer par création d'onglets
- Phase 3 : Modules manquants (8-10 jours) — excel_writer, 7 onglets FM, account_matcher
- Phase 4 : UI et finitions (3-4 jours) — Streamlit, branding AVIZEO, toggle concessions

---

# PARTIE 5 — GUIDE DES OUTILS PAR TÂCHE (Mission 4)

## 5.1 Outil recommandé par tâche

| Tâche | Outil | Justification |
|---|---|---|
| Architecture, specs, CLAUDE.md | **Claude.ai (Opus)** | Réflexion stratégique, compréhension métier |
| Écriture de nouveaux modules Python | **Claude Code (Sonnet)** | Accès filesystem, exécution tests, itérations rapides |
| Debugging / refactoring | **Claude Code (Sonnet)** | Erreurs en temps réel, exécuter le pipeline |
| Design des styles FM | **Claude Code (Sonnet)** | Itérer sur styles.py, générer un FM test |
| Interface Streamlit | **Claude Code (Sonnet)** | Modifications CSS, components, live reload |
| Tests unitaires | **Claude Code (Haiku)** | Tests patterns, assertions simples, volume |
| Analyse du FM GILAC exemple | **Claude.ai** | Upload du fichier, analyse de structure |
| Documentation utilisateur | **Claude.ai (Sonnet)** | Rédaction, structuration |

**NON recommandés :** Claude Excel (le travail est programmatique, pas interactif), MCP Servers (pas de services externes).

## 5.2 Optimisations proposées (par priorité)

**1. Modularité — Créer les modules manquants (CRITIQUE)**
- `excel_writer.py` — décrit dans CLAUDE.md mais absent
- Extraire les fonctions de balance N-1 de `mapping_parser.py` vers `balance_parser.py`

**2. Rapidité — Optimiser le parsing FEC (HAUTE)**
- `_detecter_encodage()` lit le fichier ENTIER pour chaque encodage (jusqu'à 4 lectures de 16 Mo)
- Optimisation : lire les 10 000 premiers et derniers octets

**3. Tokens — Éviter la duplication de logique (HAUTE)**
- `app.py` contient ~80 lignes dupliquées pour reconstruire la balance
- Solution : retourner la balance mappée dans `resultats` depuis `run_pipeline`

**4. Séparation logique métier dans fm_writer.py (MOYENNE)**
- Extraire `fm_builder.py` pour préparer les DataFrames, `fm_writer.py` ne fait qu'écrire
- Avantage : réutilisabilité pour P1

**5. Configuration externalisée (MOYENNE)**
- Le mapping feuille → template devrait être dans le YAML :
```yaml
integration_templates:
  A_Fournisseurs:
    feuilles: [AACE, A0]
    position: avant_synthese
  F_Financement:
    feuilles: [Tréso, F0]
    position: avant_synthese
```

**6. Cache Streamlit (BASSE)** — `@st.cache_resource` pour les objets lourds

**7. Logging structuré (BASSE)** — JSON au lieu de texte

---

# PARTIE 6 — INFORMATIONS COMPLÉMENTAIRES

## 6.1 Déploiement

- **Phase 1 :** Local uniquement (macOS, Python 3.9+, `python3 -m streamlit run app.py`)
- **Phase 2 :** Serveur interne (Docker + reverse proxy, prévoir Dockerfile)

## 6.2 Universalité du système

Le code ne doit JAMAIS contenir :
- De numéros de comptes en dur (sauf les racines PCG dans le YAML/mapping)
- De libellés de comptes en dur
- De noms de clients en dur
- De montants ou seuils spécifiques à un client

Tout doit passer par :
- `mapping_pcg.yaml` pour les racines de comptes → cycles
- Les paramètres d'entrée (nom client, date clôture)
- Le FM N-1 du client pour le mapping historique (priorité absolue)

Le FM GILAC sert UNIQUEMENT de référence de format/mise en page, pas de source de données.

## 6.3 Questions — toutes résolues

Toutes les questions de clarification ont été répondues et intégrées dans ce document. Prêt pour implémentation.
