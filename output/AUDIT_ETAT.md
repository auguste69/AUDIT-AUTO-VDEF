# Rapport d'audit — État du projet audit-automation-2

**Date :** 19/04/2026  
**Projet :** audit-automation-2

---

## 1. Fichiers attendus

- `src/engine/financial_engine.py` (fichier) : ✗ absent
- `src/engine/account_matcher.py` (fichier) : ✗ absent
- `src/engine/liasse_fiscale_loader.py` (fichier) : ✗ absent
- `src/models` (dossier) : ✗ absent
- `src/writers/fm` (dossier) : ✗ absent
- `src/writers/worksheet_copy.py` (fichier) : ✗ absent
- `src/parsers/balance_n1_loader.py` (fichier) : ✗ absent
- `tests/fixtures` (dossier) : ✗ absent

---

## 2. Livrables

### 2a. Pipeline GILAC

- Pipeline : ✓ terminé (106272 écritures, 270 comptes, 9 contrôles)

### 2b. Onglets du FM produit

Onglets produits (18) : Sommaire, Balance N Vs N-1, Bilan, Tréso, AACE, C Propres0, C PRC0, F0, I Incorp0, I Corp0, I Fi0, S0, A0, V0, P0, E0, T0, X0

- Manquants : EBIT, Actif détaillé, Passif détaillé, P&L détaillé
- Présents non prévus : aucun

### 2c. Contenu du ZIP templates

Fichiers dans le ZIP (13) :

- `2025_12_GILAC_A_Fournisseurs-Achats.xlsx` : 14 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_C_Capitaux propres.xlsx` : 4 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_C_Provision pour risques et charges.xlsx` : 3 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_E_Etat.xlsx` : 8 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_F_Financement.xlsx` : 10 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_I_Immobilisations corp.xlsx` : 8 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_I_Immobilisations financières.xlsx` : 3 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_I_Immobilisations incorporelles.xlsx` : 2 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_P_Personnel.xlsx` : 10 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_S_Stocks.xlsx` : 13 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_T_Autres créances - dettes.xlsx` : 1 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_V_Clients-Ventes.xlsx` : 18 onglets — FM trouvés : aucun — ✗ P1 absent
- `2025_12_GILAC_X_Résultat exceptionnel.xlsx` : 2 onglets — FM trouvés : aucun — ✗ P1 absent

---

## 3. Dette technique

### 3a. Préfixes PCG en dur dans `src/` (pattern `["NNN"]`)

- `src/writers/fm_writer.py:181` → `cap_non_appele = sp(["109"])`
- `src/writers/fm_writer.py:182` → `immo_incorp    = sp(["201","203","205","206","207","208","232","237",`
- `src/writers/fm_writer.py:184` → `immo_corp      = sp(["211","212","213","214","215","218","231","238",`
- `src/writers/fm_writer.py:186` → `immo_fi        = sp(["261","266","267","268","271","272","273","274","275","276"`
- `src/writers/fm_writer.py:188` → `stocks         = sp(["31","32","33","34","35","37","39"])`
- `src/writers/fm_writer.py:189` → `avances        = sp(["4091"])`
- `src/writers/fm_writer.py:190` → `crean_cli      = sp(["411","413","416","418","491"])`
- `src/writers/fm_writer.py:191` → `autres_crean   = sp(["4096","4097","4098","425","441","444","4456",`
- `src/writers/fm_writer.py:193` → `crean_interco  = sp_si(["451","455"], ">0")`
- `src/writers/fm_writer.py:194` → `vmp            = sp(["50","590"])`
- `src/writers/fm_writer.py:195` → `dispo          = sp(["51","53"])`
- `src/writers/fm_writer.py:196` → `cca            = sp(["486"])`
- `src/writers/fm_writer.py:199` → `capital        = sp(["101","108"],                         signe=-1)`
- `src/writers/fm_writer.py:200` → `primes         = sp(["104"],                               signe=-1)`
- `src/writers/fm_writer.py:201` → `reserve_legale = sp(["1061"],                              signe=-1)`
- `src/writers/fm_writer.py:202` → `autres_res     = sp(["1062","1063","1064","1068"],          signe=-1)`
- `src/writers/fm_writer.py:203` → `report         = sp(["11"],                                signe=-1)`
- `src/writers/fm_writer.py:204` → `resultat       = sp(["12"],                                signe=-1)`
- `src/writers/fm_writer.py:205` → `subventions    = sp(["13"],                                signe=-1)`
- `src/writers/fm_writer.py:206` → `prov_regl      = sp(["14"],                                signe=-1)`
- `src/writers/fm_writer.py:207` → `prov_risques_b = sp(["151"],                               signe=-1)`
- `src/writers/fm_writer.py:208` → `prov_charges_b = sp(["152","153","154","155","156","157","158"], signe=-1)`
- `src/writers/fm_writer.py:209` → `emprunts       = sp(["16","17","519"],                     signe=-1)`
- `src/writers/fm_writer.py:210` → `det_fourn_b    = sp(["401","403","4081","4088"],            signe=-1)`
- `src/writers/fm_writer.py:211` → `det_fisc_b     = sp(["421","422","423","424","427","428","43","442","444",`
- `src/writers/fm_writer.py:213` → `autres_dettes  = sp(["4191","404","405","4084","4196","4197","4198",`
- `src/writers/fm_writer.py:215` → `pca_b          = sp(["487"],                               signe=-1)`
- `src/writers/fm_writer.py:216` → `dettes_interco = sp_si(["451","455"], "<0",                signe=-1)`
- `src/writers/fm_writer.py:468` → `cap_propres  = sp(["10", "11", "12", "13", "14"], signe=-1)`
- `src/writers/fm_writer.py:469` → `amort_dep    = sp(["28", "29", "39", "49"],        signe=-1)`
- `src/writers/fm_writer.py:470` → `prov_risques = sp(["15"],                           signe=-1)`
- `src/writers/fm_writer.py:471` → `dettes_mlt   = sp(["16", "17"],                    signe=-1)`
- `src/writers/fm_writer.py:475` → `actif_immo = sp(["20", "21", "22", "23", "24", "25", "26", "27"])`
- `src/writers/fm_writer.py:483` → `stocks       = sp(["31", "32", "33", "34", "35", "37"])`
- `src/writers/fm_writer.py:484` → `crean_cli    = sp(["411", "413", "416", "418"])`
- `src/writers/fm_writer.py:485` → `autres_crean = sp_si(["40", "44", "46", "47"], ">0")`
- `src/writers/fm_writer.py:486` → `cca          = sp(["486"])`
- `src/writers/fm_writer.py:490` → `det_fourn  = sp(["401", "403", "408"], signe=-1)`
- `src/writers/fm_writer.py:491` → `det_fisc   = sp_si(["42", "43", "44"], "<0", signe=-1)`
- `src/writers/fm_writer.py:492` → `autres_det = sp_si(["46", "47"],       "<0", signe=-1)`
- `src/writers/fm_writer.py:493` → `pca        = sp(["487"], signe=-1)`
- `src/writers/fm_writer.py:505` → `treso_active   = sp_si(["50", "51", "53"], ">0")`
- `src/writers/fm_writer.py:506` → `treso_pass_519 = sp(["519"], signe=-1)`
- `src/writers/fm_writer.py:507` → `treso_pass_512 = sp_si(["512"], "<0", signe=-1)`

### 3b. Section `liasse_fiscale` dans `mapping_pcg.yaml`

✗ Absente

### 3c. Fonctions / méthodes clés dans `template_writer.py`

- `_injecter_balance_cycle` : ✓ présent
- `move_sheet` : ✗ absent
- `_inserer_onglet_fm` : ✗ absent

### 3d. Duplication extraction N-1 (`iter_rows(min_row=10`)

- `main.py` : 1 occurrence(s)
- `app.py` : 1 occurrence(s)

---

## 4. Cibles chiffrées GILAC

- Résultat net théorique (−cl.7 − cl.6) = **2572.709 K€** (attendu 2572.7 K€ ± 0.1) — ✓
- Actif brut (classes 2+3+5) = **8790.531 K€** (attendu 8790.5 K€ ± 0.1) — ✓
- Passif brut signé (classes 1+4) = **-6217.821 K€** (attendu -6217.8 K€ ± 0.1) — ✓

### 4b. Résiduel bilan (WARNING logger `fm_writer`)

- `WARNING src.writers.fm_writer: Bilan : résiduel actif/passif N = 760.725 K€ (comptes non capturés → 'Autres (résiduel)')`
  → Valeur extraite : 760.725 K€ (attendu 760.725 K€ ± 0.1) — ✓

---

## 5. Tests

**Résultat** : 151 passé(s), 0 échoué(s), 0 en erreur

Tous les tests passent ✓

<details><summary>Sortie pytest complète (fin)</summary>

```
tests/test_template_writer.py::TestErreurs::test_dossier_inexistant PASSED [ 79%]
tests/test_template_writer.py::TestErreurs::test_dossier_vide PASSED     [ 80%]
tests/test_template_writer.py::TestErreurs::test_sans_mapping_traite_tout PASSED [ 80%]
tests/test_writers.py::TestFichierGenere::test_fichier_existe PASSED     [ 81%]
tests/test_writers.py::TestFichierGenere::test_extension_xlsx PASSED     [ 82%]
tests/test_writers.py::TestFichierGenere::test_nom_fichier PASSED        [ 82%]
tests/test_writers.py::TestOnglets::test_nombre_onglets PASSED           [ 83%]
tests/test_writers.py::TestOnglets::test_noms_onglets PASSED             [ 84%]
tests/test_writers.py::TestOnglets::test_ordre_cycles_canonique PASSED   [ 84%]
tests/test_writers.py::TestBalanceTab::test_headers_row8 PASSED          [ 85%]
tests/test_writers.py::TestBalanceTab::test_titre_row3 PASSED            [ 86%]
tests/test_writers.py::TestBalanceTab::test_client_row4 PASSED           [ 86%]
tests/test_writers.py::TestBalanceTab::test_nombre_lignes_donnees PASSED [ 87%]
tests/test_writers.py::TestBalanceTab::test_comptes_tries_par_numero PASSED [ 88%]
tests/test_writers.py::TestBalanceTab::test_solde_brut_pas_inverse PASSED [ 88%]
tests/test_writers.py::TestBalanceTab::test_cycle_colonne_j PASSED       [ 89%]
tests/test_writers.py::TestConventionSigne::test_passif_signe_inverse PASSED [ 90%]
tests/test_writers.py::TestConventionSigne::test_charges_signe_identique PASSED [ 90%]
tests/test_writers.py::TestConventionSigne::test_actif_signe_identique PASSED [ 91%]
tests/test_writers.py::TestConventionSigne::test_produits_signe_inverse PASSED [ 92%]
tests/test_writers.py::TestConventionSigne::test_capital_passif_positif PASSED [ 92%]
tests/test_writers.py::TestOngletsCycle::test_titre_cycle PASSED         [ 93%]
tests/test_writers.py::TestOngletsCycle::test_client_dans_cycle PASSED   [ 94%]
tests/test_writers.py::TestOngletsCycle::test_headers_row8_cycle PASSED  [ 94%]
tests/test_writers.py::TestOngletsCycle::test_sections_presentes_A0 PASSED [ 95%]
tests/test_writers.py::TestOngletsCycle::test_sections_presentes_V0 PASSED [ 96%]
tests/test_writers.py::TestOngletsCycle::test_nb_comptes_par_cycle PASSED [ 96%]
tests/test_writers.py::TestOngletsCycle::test_var_pct_na_pour_comptes_nouveaux PASSED [ 97%]
tests/test_writers.py::TestOngletsCycle::test_retour_sommaire_row1 PASSED [ 98%]
tests/test_writers.py::TestStyles::test_import_styles PASSED             [ 98%]
tests/test_writers.py::TestStyles::test_remove_gridlines PASSED          [ 99%]
tests/test_writers.py::TestStyles::test_write_header_row PASSED          [100%]

=============================== warnings summary ===============================
tests/test_balance_builder.py::test_colonnes_presentes
tests/test_controls.py::TestGilac::test_run_all_retourne_9_controles
tests/test_cycle_mapper.py::TestMapCycles::test_shape
tests/test_fec_parser.py::test_colonnes_obligatoires_presentes
tests/test_writers.py::TestFichierGenere::test_fichier_existe
  /home/user/audit-automation-2/src/parsers/fec_parser.py:101: Pandas4Warning: For backward compatibility, 'str' dtypes are included by select_dtypes when 'object' dtype is specified. This behavior is deprecated and will be removed in a future version. Explicitly pass 'str' to `include` to select them, or to `exclude` to remove them and silence this warning.
  See https://pandas.pydata.org/docs/user_guide/migration-3-strings.html#string-migration-select-dtypes for details on how to write code that works with pandas 2 and 3.
    colonnes_str = df.select_dtypes("object").columns

tests/test_fec_parser.py::test_pas_de_retour_chariot
  /home/user/audit-automation-2/tests/test_fec_parser.py:85: Pandas4Warning: For backward compatibility, 'str' dtypes are included by select_dtypes when 'object' dtype is specified. This behavior is deprecated and will be removed in a future version. Explicitly pass 'str' to `include` to select them, or to `exclude` to remove them and silence this warning.
  See https://pandas.pydata.org/docs/user_guide/migration-3-strings.html#string-migration-select-dtypes for details on how to write code that works with pandas 2 and 3.
    for col in fec.select_dtypes("object").columns:

-- Docs: https://docs.pytest.org/en/stable/how-to/capture-warnings.html
======================= 151 passed, 6 warnings in 49.81s =======================
```
</details>