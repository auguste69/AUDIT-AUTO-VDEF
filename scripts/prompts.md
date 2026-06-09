# Prompts 3 à 12 — Refactoring audit-automation

**Livré le 20/04/2026. Révisé le 09/06/2026 (audit de cohérence).**

**Contexte.** Les prompts 1 et 2 ont été exécutés. Le prompt 2 a corrigé le bug double-comptage 444/467/478 dans `_calc_bilan`. Les prompts ci-dessous s'appuient sur cet état.

**Révision du 09/06/2026 — changements par rapport à la version initiale :**
1. **Prompts 5 et 6 inversés** : l'extraction du moteur financier (`financial_engine.py`) passe AVANT le contrôle AC-1, pour éviter un import temporaire engine→writers (violation de l'architecture parsers/engine/writers).
2. **AC-1 reprécisé** : bloquant sur l'exercice **N uniquement** (seuil paramétrable, défaut 1.0 K€) ; l'exercice N-1 est en WARNING. Constat vérifié : le pipeline GILAC actuel logge un écart N-1 de 2.837 K€ (données historiques du FM, non corrigeables) — un AC-1 bloquant sur N-1 rendrait la validation insatisfiable.
3. **Test de régression cellulaire (prompt 4) spécifié** : comparaison par NOM d'onglet sur un sous-ensemble stable, tolérance flottante — pour survivre aux ajouts d'onglets des prompts 9-10.
4. **Specs liasse fiscale corrigées dans CLAUDE.md** (à respecter dans les prompts 9-10) : 7087 reste en production vendue (pas en ventes de marchandises) ; 6817 = dépréciations actif circulant (ligne GC), séparé des dotations immobilisations ; IS = 695 + 698 + 699.
5. **Prompt 11 renforcé** : préservation des images/dessins des templates (9 des 13 templates en contiennent), arbitrage explicite du template `I_Immobilisations incorporelles.xlsx` (n'accueille aucune feuille FM).
6. **Bloc `<validation_cmds>` ajouté à chaque prompt** : commandes shell exécutées mécaniquement par l'orchestrateur (exit code 0 = succès, les `grep` d'absence sont préfixés par `!`). La section `<validation>` reste la check-list humaine.

**Avertissement.** Ces prompts sont figés sur mes hypothèses actuelles. Si une dérive apparaît entre deux prompts, elle ne sera pas rattrapée par cette séquence — l'orchestrateur committe après chaque prompt validé et revient au dernier commit vert en cas d'échec.

**Règles d'usage.**
- Un prompt à la fois. Ne jamais enchaîner deux prompts sans avoir vérifié le précédent.
- Tous les tests existants doivent rester verts entre chaque prompt.
- Scope strict : après chaque exécution, `git status` ne doit montrer que les fichiers listés dans `fichiers_a_modifier` ou `fichiers_a_creer` (contrôle automatisé par l'orchestrateur).

---

## Prompt 3 — Création structure YAML `liasse_fiscale` (vide)

```xml
<prompt id="3" titre="Création de la section YAML liasse_fiscale (structure vide)">

<objectif>
Ajouter la section liasse_fiscale au mapping_pcg.yaml avec structure complète mais listes vides.
Aucune migration de valeurs (prompt 4).
</objectif>

<fichiers_a_modifier>
- src/config/mapping_pcg.yaml
</fichiers_a_modifier>

<fichiers_a_creer>
- tests/test_mapping_liasse_fiscale_structure.py
</fichiers_a_creer>

<instructions>

1. LIRE mapping_pcg.yaml entièrement (RP-2). Identifier l'emplacement avant la section surcharges_cabinet.

2. AJOUTER avant surcharges_cabinet la section liasse_fiscale complète avec la structure ci-dessous. Toutes les listes restent vides — ne rien remplir.

liasse_fiscale:

  bilan:
    actif:
      capital_souscrit_non_appele: []
      immo_incorp: []
      immo_corp: []
      immo_fi: []
      stocks: []
      avances_acomptes_verses: []
      creances_clients: []
      autres_creances: []
      vmp: []
      disponibilites: []
      cca: []
    passif:
      capital: []
      primes: []
      reserve_legale: []
      autres_reserves: []
      report_nouveau: []
      resultat_exercice: []
      resultat_en_cours: []
      subventions: []
      prov_reglementees: []
      prov_risques: []
      prov_charges: []
      emprunts: []
      dettes_fournisseurs: []
      dettes_fiscales_sociales: []
      autres_dettes: []
      pca: []
    comptes_bascule:
      prefixes: []

  treso:
    ressources_stables:
      capitaux_propres: []
      amort_dep: []
      prov_risques: []
      dettes_financieres_mlt: []
    emplois_stables:
      actif_immobilise_brut: []
    actif_circulant_exploitation:
      stocks: []
      creances_clients: []
      autres_creances: []
      cca: []
    passif_circulant_exploitation:
      dettes_fournisseurs: []
      dettes_fiscales_sociales: []
      autres_dettes: []
      pca: []
    tresorerie_directe:
      tresorerie_active: []
      tresorerie_passive_519: []
      tresorerie_passive_512: []

  aace:
    prefixes: []

  ebit:
    produits_exploitation:
      ventes_marchandises: []
      production_vendue: []
      production_stockee: []
      production_immobilisee: []
      subventions_exploitation: []
      reprises_transferts: []
      autres_produits: []
    charges_exploitation:
      achats_marchandises: []
      variation_stocks_marchandises: []
      achats_matieres_premieres: []
      variation_stocks_matieres: []
      autres_charges_externes: []
      impots_taxes: []
      salaires_traitements: []
      charges_sociales: []
      dotations_amortissements: []
      dotations_dep_immobilisations: []
      dotations_dep_actif_circulant: []
      dotations_provisions: []
      autres_charges: []

  pl_detaille:
    produits_financiers: []
    charges_financieres: []
    produits_exceptionnels: []
    charges_exceptionnelles: []
    participation_salaries: []
    impots_benefices: []

3. CRÉER tests/test_mapping_liasse_fiscale_structure.py qui vérifie UNIQUEMENT la structure, pas le contenu :
   - Fixture : charge le YAML via yaml.safe_load.
   - test_section_liasse_fiscale_existe : assert 'liasse_fiscale' in config.
   - test_sous_sections_top_level : les 5 clés bilan, treso, aace, ebit, pl_detaille sont présentes.
   - test_bilan_structure : bilan contient 'actif', 'passif', 'comptes_bascule'.
   - test_listes_feuilles_sont_des_listes : parcours récursif, toute valeur feuille (non dict) est une list.
   - Ne PAS asserter que les listes sont non vides.

</instructions>

<contraintes>
- Aucune modification Python (ni src/, ni main.py, ni app.py).
- Aucune migration de valeurs — toutes les listes restent vides.
- Ne pas modifier les autres sections du YAML (classe_1 à classe_7, classification_bilan, templates, ordre_cycles, etc.).
- Le test doit passer avec des listes vides.
- Les dotations EBIT sont structurées en 4 lignes distinctes (GA/GB/GC/GD du cerfa 2052) : dotations_amortissements, dotations_dep_immobilisations, dotations_dep_actif_circulant, dotations_provisions.
</contraintes>

<validation_cmds>
python3 -c "import yaml; c = yaml.safe_load(open('src/config/mapping_pcg.yaml')); assert 'liasse_fiscale' in c; assert set(c['liasse_fiscale'].keys()) == {'bilan', 'treso', 'aace', 'ebit', 'pl_detaille'}"
python3 -m pytest tests/test_mapping_liasse_fiscale_structure.py -q
</validation_cmds>

<validation>
1. python3 -c "import yaml; c = yaml.safe_load(open('src/config/mapping_pcg.yaml')); assert 'liasse_fiscale' in c; assert set(c['liasse_fiscale'].keys()) == {'bilan', 'treso', 'aace', 'ebit', 'pl_detaille'}; print('Structure OK')"

2. python3 -m pytest tests/ -v --tb=short
   → Tous les tests précédents verts + 4 nouveaux tests verts.

3. python3 main.py data/GILAC_2025_12_31_FEC.txt --client GILAC --date-cloture 31/12/2025 --n1-fm "data/FM GILAC.xlsx" --output output/post_p3/ --no-templates
   → Pipeline fonctionne toujours. FM produit identique à celui du prompt 2.
</validation>

</prompt>
```

---

## Prompt 4 — Migration des préfixes Bilan/Tréso/AACE vers YAML

```xml
<prompt id="4" titre="Migration préfixes Bilan, Tréso, AACE de Python vers YAML">

<objectif>
Migrer tous les préfixes PCG en dur de fm_writer.py vers la section 
liasse_fiscale du YAML. Zéro changement de comportement observable — 
test de régression cellulaire obligatoire.
</objectif>

<fichiers_a_modifier>
- src/config/mapping_pcg.yaml (remplir bilan, treso, aace)
- src/parsers/mapping_parser.py (exposer liasse_fiscale)
- src/writers/fm_writer.py (consommer le mapping)
- main.py, app.py (passer pcg_config à write_fm)
</fichiers_a_modifier>

<fichiers_a_creer>
- src/engine/liasse_fiscale_loader.py
- tests/fixtures/FM_GILAC_2025_POST_FIX.xlsx (snapshot avant migration)
- tests/test_fm_regression_cellulaire.py
</fichiers_a_creer>

<instructions>

1. SNAPSHOT AVANT (obligatoire) :
   python3 main.py data/GILAC_2025_12_31_FEC.txt --client GILAC \
     --date-cloture 31/12/2025 --n1-fm "data/FM GILAC.xlsx" \
     --output output/pre_p4/ --no-templates
   cp output/pre_p4/FM_GILAC_2025.xlsx tests/fixtures/FM_GILAC_2025_POST_FIX.xlsx

2. LECTURE PRÉALABLE (RP-2) : lire _calc_bilan, _calc_treso, _filtrer_aace et la constante _PREFIXES_AACE dans fm_writer.py. ATTENTION : migrer aussi les préfixes à 1 chiffre (ex. resultat_encours = sp(["6","7"])).

3. REMPLIR la section liasse_fiscale.bilan dans mapping_pcg.yaml avec les préfixes actuels.

4. REMPLIR liasse_fiscale.treso avec les préfixes de _calc_treso.

5. REMPLIR liasse_fiscale.aace.prefixes = ["606","607","608","609","61","62"].

6. CRÉER src/engine/liasse_fiscale_loader.py avec load_liasse_fiscale().

7. MODIFIER mapping_parser.from_pcg_config pour ajouter la clé "liasse_fiscale".

8. MODIFIER fm_writer.py pour consommer le mapping.

9. ADAPTER main.py et app.py pour passer pcg à write_fm().

10. CRÉER tests/test_fm_regression_cellulaire.py. SPÉCIFICATION OBLIGATOIRE (le test doit survivre aux ajouts d'onglets des prompts 9-10) :
    - Comparer le FM régénéré à la fixture PAR NOM D'ONGLET, jamais par position/index.
    - Périmètre comparé : uniquement les onglets présents dans la fixture ("Balance N Vs N-1", "Bilan", "Tréso", "AACE" et les 13 onglets cycle). Les onglets présents dans le FM régénéré mais absents de la fixture sont IGNORÉS (pas d'assertion sur la liste complète ni sur l'ordre des onglets).
    - Comparer valeurs de cellules + number_format via openpyxl. Pour les flottants : tolérance abs(a-b) <= 0.01 (K€).
    - Ne PAS comparer les fichiers octet par octet (timestamps dans docProps/core.xml → non reproductible).
    - Régénérer le FM dans un tmp_path pytest (pas dans output/).

11. CLEANUP FINAL : vérifier que fm_writer.py ne contient plus AUCUN préfixe PCG en dur (y compris à 1 chiffre).

</instructions>

<contraintes>
- Sections YAML liasse_fiscale.ebit et pl_detaille restent vides (prompts 9-10).
- Aucun changement de comportement observable.
- Ne pas créer financial_engine.py dans ce prompt.
</contraintes>

<validation_cmds>
! grep -E '\[\s*"[0-9]{1,4}"' src/writers/fm_writer.py
python3 -m pytest tests/test_fm_regression_cellulaire.py -q
</validation_cmds>

<validation>
1. grep -E '\[\s*"[0-9]{1,4}"' src/writers/fm_writer.py → 0 résultat (préfixes 1 à 4 chiffres).
2. python3 -m pytest tests/test_fm_regression_cellulaire.py -v → PASS.
3. python3 -m pytest tests/ -v --tb=short → tous verts.
</validation>

</prompt>
```

---

## Prompt 5 — Extraction du moteur financier

```xml
<prompt id="5" titre="Extraction du moteur financier : financial_engine.py + dataclasses">

<objectif>
Séparer calcul métier et présentation Excel. Créer financial_engine.py 
et src/models/ avec les dataclasses. fm_writer.py ne fait plus que de l'écriture.
(Exécuté AVANT le contrôle AC-1 — révision 09/06/2026 — pour que le contrôle
puisse s'appuyer sur le moteur sans import engine→writers.)
</objectif>

<fichiers_a_modifier>
- src/writers/fm_writer.py (dépouillé du calcul)
</fichiers_a_modifier>

<fichiers_a_creer>
- src/engine/financial_engine.py
- src/models/__init__.py
- src/models/financial_statements.py
- tests/test_financial_engine_bilan.py
- tests/test_financial_engine_treso.py
</fichiers_a_creer>

<instructions>

1. LIRE fm_writer.py en entier. Identifier toutes les fonctions de calcul : _calc_bilan, _calc_treso, _sommer_prefixes, _sommer_prefixes_si, _classer_comptes_bascule, _filtrer_aace.

2. CRÉER src/models/__init__.py et src/models/financial_statements.py avec PosteComptable, BilanSynthetique, TresoSynthetique.

3. CRÉER src/engine/financial_engine.py avec calculer_bilan, calculer_treso, filtrer_aace. Les fonctions prennent (balance: pd.DataFrame, liasse_config: dict) — aucun préfixe en dur.

4. MODIFIER fm_writer.py : supprimer les fonctions de calcul, importer depuis financial_engine.

5. CRÉER les 2 fichiers de tests.

6. Test de régression cellulaire (prompt 4) doit continuer à passer.

</instructions>

<contraintes>
- Zéro changement de comportement observable.
- fm_writer.py cible < 700 lignes après refactor (le fichier fait ~930 lignes, ~300 lignes de calcul extraites).
- Pas encore de src/writers/fm/.
- Ne PAS toucher à controls.py dans ce prompt (le contrôle AC-1 arrive au prompt 6).
</contraintes>

<validation_cmds>
! grep -E "_sommer_prefixes|_calc_bilan|_calc_treso" src/writers/fm_writer.py
test $(wc -l < src/writers/fm_writer.py) -lt 700
python3 -m pytest tests/test_financial_engine_bilan.py tests/test_financial_engine_treso.py tests/test_fm_regression_cellulaire.py -q
</validation_cmds>

<validation>
1. grep -E "_sommer_prefixes|_calc_bilan|_calc_treso" src/writers/fm_writer.py → 0 résultat.
2. wc -l src/writers/fm_writer.py → < 700 lignes.
3. python3 -m pytest tests/ -v → tous verts.
4. python3 -m pytest tests/test_fm_regression_cellulaire.py -v → PASS.
</validation>

</prompt>
```

---

## Prompt 6 — Contrôle AC-1 bloquant

```xml
<prompt id="6" titre="Contrôle AC-1 bloquant sur équilibre du bilan (exercice N)">

<objectif>
Transformer le warning silencieux de déséquilibre bilan en contrôle bloquant 
AC-1, appuyé sur financial_engine.calculer_bilan (prompt 5).
Ajouter un contrôle de cohérence du résultat en WARNING.
</objectif>

<fichiers_a_modifier>
- src/engine/controls.py (ajouter 2 contrôles)
- main.py, app.py (option --bilan-non-bloquant)
</fichiers_a_modifier>

<fichiers_a_creer>
- tests/test_ctrl_bilan_equilibre.py
- tests/test_ctrl_coherence_resultat.py
</fichiers_a_creer>

<instructions>

1. LIRE src/engine/controls.py pour comprendre le pattern des contrôles existants.

2. AJOUTER _ctrl_bilan_equilibre (BLOQUANT) et _ctrl_coherence_resultat (WARNING).
   PÉRIMÈTRE AC-1 (décision figée, révision 09/06/2026) :
   - Le contrôle appelle financial_engine.calculer_bilan (import engine→engine, propre).
   - BLOQUANT uniquement sur l'écart |Total Actif N − Total Passif N|, seuil paramétrable seuil_equilibre_bilan_ke (défaut 1.0 K€).
   - L'écart N-1 est contrôlé en WARNING uniquement (jamais bloquant) : les données N-1 viennent
     du FM historique du client et peuvent porter un écart non corrigeable
     (constaté sur GILAC : 2.837 K€ en N-1).

3. MODIFIER run_all() pour accepter balance_mappee et liasse_config optionnels (None par défaut → AC-1 non exécuté, comportement existant inchangé).

4. MODIFIER main.py::run_pipeline pour passer balance_mappee au run_all.

5. AJOUTER option CLI --bilan-non-bloquant.

6. AJOUTER checkbox Streamlit équivalente.

7. CRÉER les 2 fichiers de tests. Inclure un test vérifiant qu'un écart N-1 > seuil ne lève PAS d'erreur (WARNING seulement) et qu'un écart N > seuil lève ValueError.

</instructions>

<contraintes>
- Le contrôle AC-1 est BLOQUANT par défaut (sur N uniquement).
- Ne pas modifier les 9 contrôles existants.
- Message d'erreur en français, compréhensible par un comptable.
- controls.py ne doit importer AUCUN module de src/writers/ (architecture : le moteur calcule, les writers écrivent).
</contraintes>

<validation_cmds>
! grep -E "from src\.writers|import src\.writers" src/engine/controls.py
python3 -m pytest tests/test_ctrl_bilan_equilibre.py tests/test_ctrl_coherence_resultat.py -q
</validation_cmds>

<validation>
1. python3 -m pytest tests/ -v --tb=short → tous verts.
2. Pipeline GILAC termine sans erreur (l'écart N-1 de 2.837 K€ produit un WARNING, pas une erreur), log contient "[OK] Équilibre du bilan".
3. FEC déséquilibré en N → ValueError "Bilan déséquilibré".
</validation>

</prompt>
```

---

## Prompt 7 — Déduplication extraction N-1

```xml
<prompt id="7" titre="Déduplication extraction N-1 entre main.py et app.py">

<objectif>
Éliminer la duplication de la logique d'extraction balance N-1.
Exposer balance_mappee dans les résultats.
</objectif>

<fichiers_a_modifier>
- main.py
- app.py
</fichiers_a_modifier>

<fichiers_a_creer>
- src/parsers/balance_n1_loader.py
- tests/test_balance_n1_loader.py
</fichiers_a_creer>

<instructions>

1. LIRE main.py et app.py. Identifier la logique dupliquée d'extraction N-1 : dans chacun des deux fichiers, le bloc qui distingue suffix == ".txt" (FEC N-1) / suffix == ".xlsx" (detect_balance_sheet puis mode "fm" avec boucle iter_rows(min_row=10) sur la feuille balance, ou balance Excel simple). NE PAS se fier à des numéros de ligne — chercher le motif iter_rows(min_row=10.

2. CRÉER src/parsers/balance_n1_loader.py avec load_balance_n1() et _extraire_soldes_from_fm().

3. MODIFIER main.py : remplacer la logique if/elif/else par load_balance_n1(). Ajouter balance_mappee aux résultats.

4. MODIFIER app.py : utiliser load_balance_n1() et resultats["balance_mappee"].

5. CRÉER tests/test_balance_n1_loader.py avec 5 tests.

</instructions>

<contraintes>
- Ne pas modifier le contrat public de mapping_parser.
- Refactor sémantiquement neutre.
</contraintes>

<validation_cmds>
! grep -E "iter_rows\(min_row=10" main.py app.py
python3 -m pytest tests/test_balance_n1_loader.py -q
</validation_cmds>

<validation>
1. grep -E "iter_rows\(min_row=10" main.py app.py → 0 résultat.
2. python3 -m pytest tests/ -v → tous verts.
</validation>

</prompt>
```

---

## Prompt 8 — Correction EtatFi / Compta N/N-1

```xml
<prompt id="8" titre="Correction double stockage EtatFi et ComptaN/N-1">

<objectif>
Corriger la perte d'information sur les reclassements inter-exercices.
Les 4 colonnes EtatFi N, EtatFi N-1, ComptaN, ComptaN-1 doivent être distinctes.
</objectif>

<fichiers_a_modifier>
- src/parsers/mapping_parser.py (lire 4 colonnes)
- src/engine/cycle_mapper.py (propager 4 champs)
- src/writers/fm_writer.py (écrire 4 colonnes distinctes)
</fichiers_a_modifier>

<fichiers_a_creer>
- tests/test_etatfi_n1_distinct.py
</fichiers_a_creer>

<instructions>

1. LIRE mapping_parser.from_fm, cycle_mapper.map_cycles, fm_writer._ecrire_balance_tab.

2. MODIFIER mapping_parser.from_fm : lire row[10], [11], [12], [13] → etatfi_n, etatfi_n1, compta_n, compta_n1. Garder aliases etatfi/compta.

3. MODIFIER cycle_mapper.map_cycles : propager les 4 champs.

4. MODIFIER fm_writer._ecrire_balance_tab : écrire les 4 colonnes distinctes.

5. CRÉER tests/test_etatfi_n1_distinct.py.

6. METTRE À JOUR la fixture FM_GILAC_2025_POST_FIX.xlsx (régénération explicite + commit de la nouvelle fixture, car les colonnes K-N du FM changent légitimement).

</instructions>

<contraintes>
- Aliases etatfi/compta conservés.
- La fixture de régression doit être mise à jour explicitement.
</contraintes>

<validation_cmds>
python3 -m pytest tests/test_etatfi_n1_distinct.py tests/test_fm_regression_cellulaire.py -q
</validation_cmds>

<validation>
1. Au moins 1 compte avec etatfi_n != etatfi_n1 dans GILAC (vérifié sur les données : le FM GILAC source contient 22 comptes avec EtatFi N ≠ EtatFi N-1 et 23 avec Compta N ≠ N-1, ex. 413000, 421000, 425000).
2. python3 -m pytest tests/ -v → tous verts.
</validation>

</prompt>
```

---

## Prompt 9 — Implémentation EBIT

```xml
<prompt id="9" titre="Implémentation de l'onglet EBIT">

<objectif>
Ajouter l'onglet EBIT au FM. Valeur cible GILAC 2025 : CA ≈ 19 126,6 K€ (tolérance 1), EBIT ≈ 2 678 K€ (tolérance 5).
</objectif>

<fichiers_a_modifier>
- src/config/mapping_pcg.yaml (remplir liasse_fiscale.ebit)
- src/engine/financial_engine.py (calculer_ebit)
- src/models/financial_statements.py (EbitSynthetique)
- src/writers/fm_writer.py (orchestration)
- tests/test_writers.py (ONGLETS_ATTENDUS passe à 19)
</fichiers_a_modifier>

<fichiers_a_creer>
- src/writers/fm/__init__.py
- src/writers/fm/ebit.py
- tests/test_ebit_gilac.py
</fichiers_a_creer>

<instructions>

1. LIRE data/FM GILAC.xlsx onglet EBIT via openpyxl pour noter la structure exacte.

2. REMPLIR liasse_fiscale.ebit dans mapping_pcg.yaml en suivant le mapping CORRIGÉ de CLAUDE.md (§2.5, révision 09/06/2026) :
   - ventes_marchandises : 707, 7097 en déduction. NE PAS y mettre 7087 (bonifications clients = produits des activités annexes → production_vendue).
   - production_vendue : 701-706, 708 (y compris 7087), 709 sauf 7097.
   - dotations en 4 lignes distinctes (cerfa 2052) : dotations_amortissements = 6811+6812 (GA) ; dotations_dep_immobilisations = 6816 (GB) ; dotations_dep_actif_circulant = 6817 (GC) ; dotations_provisions = 6815 (GD).

3. AJOUTER EbitSynthetique dans financial_statements.py.

4. AJOUTER calculer_ebit dans financial_engine.py avec gestion des exclusions (7097 hors production_vendue, 755 hors autres_produits, 655 hors autres_charges).

5. CRÉER src/writers/fm/__init__.py et src/writers/fm/ebit.py < 200 lignes.

6. MODIFIER fm_writer.py : créer onglet EBIT entre Bilan et Tréso.

7. MODIFIER tests/test_writers.py : ONGLETS_ATTENDUS à 19.

8. CRÉER tests/test_ebit_gilac.py avec assertions sur CA et EBIT.

</instructions>

<contraintes>
- Tolérance EBIT : 5 K€.
- src/writers/fm/ebit.py < 200 lignes.
- Pas de logique de calcul dans le writer.
- Le test de régression cellulaire (prompt 4) compare par nom d'onglet : l'ajout de l'onglet EBIT ne doit PAS le casser. Si le test échoue à cause du nouvel onglet, c'est le test de régression qui est mal écrit — le corriger, pas le contourner.
</contraintes>

<validation_cmds>
python3 -m pytest tests/test_ebit_gilac.py tests/test_fm_regression_cellulaire.py -q
test $(wc -l < src/writers/fm/ebit.py) -lt 200
</validation_cmds>

<validation>
1. python3 -m pytest tests/test_ebit_gilac.py -v → PASS.
2. python3 -m pytest tests/ -v → tous verts, 19 onglets attendus.
</validation>

</prompt>
```

---

## Prompt 10 — Actif détaillé + Passif détaillé + P&L détaillé

```xml
<prompt id="10" titre="Implémentation Actif détaillé, Passif détaillé, P&L détaillé">

<objectif>
Ajouter les 3 onglets cerfa restants (2050, 2051, 2052-2053).
FM atteint 22 onglets. Cible Résultat net GILAC 2025 ≈ 2 572,7 K€ ± 5.
</objectif>

<fichiers_a_modifier>
- src/config/mapping_pcg.yaml (pl_detaille + structures détaillées)
- src/engine/financial_engine.py (3 fonctions)
- src/models/financial_statements.py (3 dataclasses)
- src/writers/fm_writer.py (orchestration)
- src/engine/controls.py (2 contrôles)
- tests/test_writers.py (ONGLETS_ATTENDUS passe à 22)
</fichiers_a_modifier>

<fichiers_a_creer>
- src/writers/fm/pl_detaille.py
- src/writers/fm/actif_detaille.py
- src/writers/fm/passif_detaille.py
- tests/test_pl_detaille_gilac.py
- tests/test_actif_detaille_gilac.py
- tests/test_passif_detaille_gilac.py
</fichiers_a_creer>

<instructions>

1. LIRE data/FM GILAC.xlsx onglets cerfa. Noter la structure.

2. REMPLIR liasse_fiscale.pl_detaille dans le YAML, avec le mapping CORRIGÉ de CLAUDE.md (§2.7, révision 09/06/2026) :
   - impots_benefices : 695 + 698 (intégration fiscale) + 699 (produit de carry-back, en déduction). NE PAS omettre 698.
   - participation_salaries : 691.

3. AJOUTER actif_detaille_structure et passif_detaille_structure dans le YAML. Pour les réserves : 1061 → réserve légale ; 1063 → statutaires ; 1062+1064 → réglementées (1062 "réserves indisponibles" rattaché par convention cabinet — commenter dans le YAML) ; 1068 → autres.

4. AJOUTER 3 dataclasses dans financial_statements.py.

5. AJOUTER calculer_actif_detaille, calculer_passif_detaille, calculer_pl_detaille dans financial_engine.py.

6. CRÉER les 3 writers < 200 lignes chacun.

7. MODIFIER fm_writer.py : ordre final Sommaire → Balance → Bilan → EBIT → Actif détaillé → Passif détaillé → P&L détaillé → Tréso → AACE → 13 cycles.

8. AJOUTER _ctrl_coherence_actif_detaille et _ctrl_coherence_pl_resultat dans controls.py.

9. MODIFIER tests/test_writers.py : ONGLETS_ATTENDUS à 22.

10. CRÉER les 3 tests GILAC.

</instructions>

<contraintes>
- Tolérance Résultat net : 5 K€.
- Chaque writer < 200 lignes.
- Comme au prompt 9 : l'ajout d'onglets ne doit pas casser le test de régression cellulaire (comparaison par nom).
</contraintes>

<validation_cmds>
python3 -m pytest tests/test_pl_detaille_gilac.py tests/test_actif_detaille_gilac.py tests/test_passif_detaille_gilac.py tests/test_fm_regression_cellulaire.py -q
</validation_cmds>

<validation>
1. python3 -m pytest tests/test_pl_detaille_gilac.py tests/test_actif_detaille_gilac.py tests/test_passif_detaille_gilac.py -v → PASS.
2. FM produit a 22 onglets.
3. python3 -m pytest tests/ -v → tous verts.
</validation>

</prompt>
```

---

## Prompt 11 — P1 : Insertion onglets FM dans templates

```xml
<prompt id="11" titre="P1 — Insertion feuilles FM comme onglets dans les templates">

<objectif>
Supprimer _injecter_balance_cycle et implémenter l'insertion des feuilles FM
comme onglets Excel dans chaque template, avant la feuille Synthèse,
SANS perdre les images/dessins existants des templates.
</objectif>

<fichiers_a_modifier>
- src/config/mapping_pcg.yaml (integration_templates)
- src/writers/template_writer.py (refactor)
- main.py, app.py (passer fm_path)
- tests/test_template_writer.py (nouveaux chiffres)
</fichiers_a_modifier>

<fichiers_a_creer>
- src/writers/worksheet_copy.py
- tests/test_template_insertion.py
</fichiers_a_creer>

<instructions>

1. LIRE template_writer.py entièrement. Identifier _injecter_balance_cycle.

2. AJOUTER integration_templates dans mapping_pcg.yaml. ARBITRAGES FIGÉS (révision 09/06/2026) :
   - Les clés correspondent aux noms de fichiers RÉELS de data/templates/ (13 fichiers) — attention : "C_Provision pour risques et charges" (singulier).
   - Le template I_Immobilisations corp reçoit I Corp0 ET I Incorp0 (confirmé cabinet).
   - Le template I_Immobilisations incorporelles ne reçoit AUCUNE feuille FM (copie + renommage + placeholders uniquement).
   - A_Fournisseurs-Achats reçoit [AACE, A0] ; F_Financement reçoit [Tréso, F0].

3. CRÉER src/writers/worksheet_copy.py avec copy_worksheet() : copie valeurs + styles + merged_cells + dimensions. Détails techniques :
   - openpyxl ne copie PAS d'onglet entre workbooks (Workbook.copy_worksheet est intra-classeur uniquement).
   - Styles : copier attribut par attribut avec copy.copy() (font, border, fill, number_format, alignment, protection) — ne PAS recopier cell._style (index vers les tables de styles du classeur source).
   - Insertion positionnée : préférer wb.create_sheet(title, index=wb.sheetnames.index(nom_synthese)) à move_sheet.

4. REFACTOR template_writer.py : supprimer _injecter_balance_cycle, créer _inserer_onglets_fm().

5. PRÉSERVATION DES IMAGES (obligatoire — 9 des 13 templates contiennent des images/dessins, jusqu'à 12 dans E_Etat.xlsx ; openpyxl peut les perdre à l'ouverture+sauvegarde) :
   - Dans tests/test_template_insertion.py, AJOUTER un test qui compte les entrées xl/media/* et xl/drawings/* (via zipfile) dans chaque template source et dans le fichier généré correspondant, et asserte qu'aucun média n'est perdu.
   - Si openpyxl perd des médias : conserver l'approche actuelle de template_writer (manipulation ZIP/XML directe ou keep_links/keep_vba selon ce que fait déjà le code) plutôt que load_workbook+save naïf. Lire comment le code actuel copie les templates avant de choisir.

6. ADAPTER main.py et app.py pour passer fm_path.

7. RÉÉCRIRE tests/test_template_writer.py pour les nouveaux comptes d'onglets.

8. CRÉER tests/test_template_insertion.py avec 4 tests (+ le test médias du point 5).

</instructions>

<contraintes>
- _remplacer_placeholders intact.
- Ne pas modifier les templates sources sur disque.
- Aucune perte d'images/dessins dans les fichiers générés (test automatique obligatoire).
</contraintes>

<validation_cmds>
! grep -n "_injecter_balance_cycle" src/writers/template_writer.py
python3 -m pytest tests/test_template_insertion.py tests/test_template_writer.py -q
</validation_cmds>

<validation>
1. python3 -m pytest tests/ -v → tous verts.
2. Template A contient AACE et A0 avant "A Synthèse".
3. Les fichiers générés contiennent autant d'entrées xl/media/* que leurs templates sources.
</validation>

</prompt>
```

---

## Prompt 12 — Account matcher (P6)

```xml
<prompt id="12" titre="P6 — Account matcher : rapprochement N/N-1 avec validation">

<objectif>
Implémenter le rapprochement des comptes changeant de numéro entre N-1 et N.
Algorithme : détection orphelins, scoring composite, validation utilisateur.
</objectif>

<fichiers_a_modifier>
- src/config/mapping_pcg.yaml (section rapprochements)
- main.py (options CLI)
- app.py (UI validation)
- CLAUDE.md (statut P6 → implémenté)
</fichiers_a_modifier>

<fichiers_a_creer>
- src/engine/account_matcher.py
- src/models/rapprochement.py
- tests/test_account_matcher.py
- tests/test_account_matcher_e2e.py
</fichiers_a_creer>

<instructions>

1. LIRE CLAUDE.md section P6 pour l'algorithme détaillé.

2. AJOUTER section rapprochements dans mapping_pcg.yaml avec seuil, poids, mots_vides.

3. CRÉER src/models/rapprochement.py avec dataclass Rapprochement.

4. CRÉER src/engine/account_matcher.py avec detecter_orphelins(), scorer_matching(), proposer_rapprochements(), appliquer_rapprochements().
   NOTE D'ORDONNANCEMENT (révision 09/06/2026) : le matching s'exécute AVANT map_cycles
   (ordre pipeline : parse → load_n1 → MATCHING → build → map_cycles → fm_writer).
   Le cycle des comptes N n'est donc PAS encore calculé au moment du scoring.
   Pour les composantes "même cycle" (15%) et "même classification bilan" (10%) :
   dériver le cycle et la classification du compte N à la volée par résolution
   de préfixe PCG (réutiliser la logique de cycle_mapper / le YAML), sans modifier la balance.

5. INTÉGRER dans main.py : --rapprochements-auto, validation CLI interactive par défaut.

6. INTÉGRER dans app.py : tableau interactif avec "Tout valider" / "Tout ignorer".
   Pièges Streamlit connus : "Tout valider" doit passer par un callback on_click
   (modifier st.session_state AVANT le rerun) — modifier l'état d'un widget déjà
   instancié lève StreamlitAPIException. Les choix ligne par ligne déjà faits
   priment sur "Tout valider".

7. CRÉER tests/test_account_matcher.py (8 tests unitaires).

8. CRÉER tests/test_account_matcher_e2e.py.

9. METTRE À JOUR CLAUDE.md : P6 → implémenté.

</instructions>

<contraintes>
- Validation utilisateur obligatoire par défaut.
- Paramètres externalisés YAML.
- Ordre pipeline : parse → load_n1 → MATCHING → build → map_cycles → fm_writer.
- Aucune fusion automatique sans confirmation explicite (le seuil 0.5 sert à PROPOSER, jamais à fusionner seul).
</contraintes>

<validation_cmds>
python3 -m pytest tests/test_account_matcher.py tests/test_account_matcher_e2e.py -q
</validation_cmds>

<validation>
1. python3 -m pytest tests/test_account_matcher.py -v → PASS.
2. python3 -m pytest tests/test_account_matcher_e2e.py -v → PASS.
3. python3 -m pytest tests/ -v → tous verts (total > 170 tests).
</validation>

</prompt>
```
