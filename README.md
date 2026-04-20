# Audit Automation Platform

Plateforme web locale pour automatiser le processus d'audit : du FEC aux feuilles de travail.

## Installation (5 minutes)

### 1. Installer Python

Si Python n'est pas installé sur ton poste :
- Va sur https://www.python.org/downloads/
- Télécharge la dernière version (3.12+)
- **IMPORTANT** : coche "Add Python to PATH" pendant l'installation

Pour vérifier que Python est installé, ouvre un terminal (PowerShell sur Windows, Terminal sur Mac) :
```
python --version
```

### 2. Installer les dépendances

Ouvre un terminal dans le dossier du projet et lance :
```
pip install -r requirements.txt
```

### 3. Lancer la plateforme

```
streamlit run app.py
```

La plateforme s'ouvre dans ton navigateur à l'adresse http://localhost:8501

## Utilisation

### Étapes :

1. **Sidebar gauche** : Renseigne le nom du client et la date de clôture
2. **FEC N** : Dépose le fichier FEC de l'exercice courant (.txt)
3. **Balance N-1** : Trois options au choix :
   - Déposer un FEC N-1 (le script calcule la balance)
   - Déposer une balance N-1 en Excel
   - Extraire depuis un fichier FM existant
4. **FM existant** (optionnel) : Dépose un fichier FM d'un exercice précédent pour récupérer le mapping des comptes par cycle
5. **Templates** (optionnel) : Dépose les templates de feuilles de travail du cabinet
6. Clique sur **Lancer le traitement**

### Fichiers générés :

| Fichier | Contenu |
|---------|---------|
| `Travail_CLIENT_ANNEE.xlsx` | FEC brut + Balance N + Balance N-1 |
| `FM_CLIENT_ANNEE.xlsx` | Sommaire + Balance comparative + 13 feuilles maîtresses par cycle |
| `FdT_CLIENT_ANNEE.zip` | Templates remplis avec le nom du client et la date de clôture |

## Structure du dossier

```
audit-automation/
├── app.py                  ← Interface web Streamlit
├── audit_automation.py     ← Script en ligne de commande (alternative)
├── requirements.txt        ← Dépendances Python
├── README.md              ← Ce fichier
└── templates/             ← Tes templates de feuilles de travail (à créer)
    ├── 20XX_XX_YYY_A_Fournisseurs-Achats.xlsx
    ├── 20XX_XX_YYY_V_Clients-Ventes.xlsx
    └── ...
```

## Utilisation en ligne de commande (alternative)

Si tu préfères ne pas utiliser l'interface web :

```bash
python audit_automation.py mon_fec.txt \
  --client NOMCLIENT \
  --date-cloture 31/12/2025 \
  --n1-fm ancien_FM.xlsx \
  --templates ./templates/ \
  --output ./output
```

## Notes

- Les montants sont affichés en milliers d'euros (K€), arrondis à l'entier
- Les négatifs sont entre parenthèses : (513)
- Le script détecte automatiquement l'encodage et le séparateur du FEC
- Les nouveaux comptes (absents du mapping) sont auto-assignés par logique PCG
- Aucune donnée ne quitte ton poste — tout tourne en local
