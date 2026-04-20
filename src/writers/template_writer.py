"""
Injection des données client dans les templates de feuilles de travail.

Pour chaque template .xlsx dans le dossier source :
  1. Copie le fichier
  2. Renomme : 20XX_XX_YYY_  →  {annee}_12_{client}_
  3. Remplace dans toutes les cellules de toutes les feuilles :
       #NomClient    → nom du client
       #DateClôture  → date au format JJ/MM/AAAA
       #Date         → idem
  4. Supprime les quadrillages de toutes les feuilles
  5. Injecte les données de balance du cycle dans la feuille Synthèse (si balance_mappee fourni)
  6. Compresse l'ensemble dans un ZIP

Les templates sans cycle reconnu dans la config sont ignorés avec un avertissement.
"""

import io
import logging
import zipfile
from pathlib import Path
from typing import Dict, Optional, Union

import openpyxl
import pandas as pd

from src.writers.styles import (
    remove_gridlines, write_header_row, write_data_row, write_total_row,
    write_section_label, FONT_BOLD, FONT_META, FONT_NORMAL, NUM_KE, NUM_PCT,
)

logger = logging.getLogger(__name__)

# Préfixe commun à tous les noms de fichiers template
_PREFIXE_TEMPLATE = "20XX_XX_YYY_"

# Placeholders reconnus et leur remplacement cible
_PLACEHOLDERS = ("#NomClient", "#DateClôture", "#Date", "#date de clôture")


def _prefixe_sortie(client: str, date_cloture: str) -> str:
    """Construit le préfixe de sortie : '{annee}_12_{client}_'."""
    annee = pd.to_datetime(date_cloture, format="%d/%m/%Y").year
    return f"{annee}_12_{client}_"


def _normaliser(s: str) -> str:
    """
    Normalise pour la comparaison : sans accents, sans séparateurs, minuscules.
    Ex: 'X_Re_sultat' et 'X_Résultat exceptionnel' donnent tous deux 'xresultat…'.
    """
    import unicodedata
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")  # strip accents
    # Supprimer tous les séparateurs (espaces, underscores, tirets)
    for sep in (" ", "_", "-"):
        s = s.replace(sep, "")
    return s.lower()


def _detecter_cycle(nom_fichier: str, mapping_templates: Dict[str, str]) -> Optional[str]:
    """
    Retrouve le cycle d'un template à partir de son nom de fichier.

    La comparaison est normalisée (espaces → _, sans accents, minuscules)
    pour gérer les écarts entre les clés YAML et les noms de fichiers réels.
    """
    partie_norm = _normaliser(
        nom_fichier.replace(_PREFIXE_TEMPLATE, "").replace(".xlsx", "")
    )
    for cle, cycle in mapping_templates.items():
        if _normaliser(cle) in partie_norm:
            return cycle
    return None


def _remplacer_placeholders(ws, nom_client: str, date_cloture: str) -> int:
    """
    Remplace les placeholders dans toutes les cellules de la feuille.
    Retourne le nombre de remplacements effectués.
    """
    nb = 0
    date_fmt = pd.to_datetime(date_cloture, format="%d/%m/%Y").strftime("%d/%m/%Y")

    substitutions = {
        "#NomClient":       nom_client,
        "#DateClôture":     date_fmt,
        "#Date":            date_fmt,
        "#date de clôture": date_fmt,
    }

    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                valeur = cell.value
                for placeholder, remplacement in substitutions.items():
                    if placeholder in valeur:
                        valeur = valeur.replace(placeholder, remplacement)
                        nb += 1
                if valeur != cell.value:
                    cell.value = valeur

    return nb


def _injecter_balance_cycle(
    ws,
    df_cycle: "pd.DataFrame",
    cycle: str,
    date_n: str,
    date_n1: str,
) -> None:
    """
    Ajoute un tableau de balance du cycle à la fin de la feuille Synthèse.

    Le tableau est inséré après le dernier contenu existant, séparé par une ligne vide.
    Il comprend un titre, un en-tête, et les comptes organisés par section
    (ACTIF / PASSIF / CHARGES / PRODUITS) avec totaux par section.

    Paramètres
    ----------
    ws : Worksheet openpyxl (mode écriture)
    df_cycle : pd.DataFrame
        Sous-ensemble de balance_mappee pour ce cycle — colonnes requises :
        CompteNum, CompteLib, Solde_KE, Solde_N1_KE, Var_KE, Var_PCT, compta.
    cycle : str
        Code du cycle (ex: "F", "A", "C Propres").
    date_n : str
        Date de clôture N au format 'JJ/MM/AAAA'.
    date_n1 : str
        Date de clôture N-1 au format 'JJ/MM/AAAA'.
    """
    if df_cycle.empty:
        logger.warning(
            "Cycle '%s' : aucun compte à injecter dans la feuille Synthèse", cycle
        )
        return

    # Trouver la dernière ligne vraiment occupée (non vide)
    derniere_ligne = 1
    for row_idx in range(ws.max_row, 0, -1):
        row_vals = []
        for c in range(1, 10):
            try:
                row_vals.append(ws.cell(row=row_idx, column=c).value)
            except AttributeError:
                pass
        if any(v is not None for v in row_vals):
            derniere_ligne = row_idx
            break

    # Démarrer 2 lignes après le dernier contenu
    row_debut = derniere_ligne + 2

    # Estimer les lignes nécessaires et désfusionner la zone d'injection
    # (les templates ont des plages fusionnées qui peuvent s'étendre au-delà du contenu visible)
    rows_needed = 3 + len(df_cycle) + 12
    to_unmerge = [
        str(mr) for mr in list(ws.merged_cells.ranges)
        if mr.max_row >= row_debut and mr.min_row <= row_debut + rows_needed
    ]
    for rng in to_unmerge:
        ws.unmerge_cells(rng)

    # Titre du bloc injecté
    c_titre = ws.cell(row=row_debut, column=2)
    c_titre.value = f"Balance du cycle {cycle} — extrait de la feuille maîtresse"
    c_titre.font = FONT_BOLD
    row_debut += 1

    # Ligne d'en-têtes
    headers = [
        (2, "Compte"),
        (3, "Libellé"),
        (4, date_n),
        (5, date_n1),
        (6, "Var. K€"),
        (7, "Var. %"),
    ]
    write_header_row(ws, row_debut, headers)
    row_debut += 1

    # Signe de présentation par section (Passif et Produits affichés en positif)
    _SIGNE = {"Actif": 1, "Passif": -1, "Charges": 1, "Produits": -1}
    _ORDRE_SECTIONS = ["Actif", "Passif", "Charges", "Produits"]

    current_row = row_debut
    for section in _ORDRE_SECTIONS:
        masque = df_cycle["compta"].str.capitalize() == section
        df_section = df_cycle[masque]
        if df_section.empty:
            continue

        signe = _SIGNE[section]

        write_section_label(ws, current_row, section.upper())
        current_row += 1

        total_n = 0.0
        total_n1 = 0.0

        for _, r in df_section.iterrows():
            val_n = round(float(r["Solde_KE"]) * signe, 3)
            val_n1 = round(float(r["Solde_N1_KE"]) * signe, 3)
            var_ke = round(val_n - val_n1, 3)

            if abs(val_n1) >= 0.001:
                var_pct: Union[float, str] = round(var_ke / abs(val_n1), 4)
            else:
                var_pct = "n/a"

            cells = [
                (2, str(r["CompteNum"]), None, FONT_META),
                (3, r["CompteLib"], None, FONT_NORMAL),
                (4, val_n, NUM_KE, FONT_NORMAL),
                (5, val_n1, NUM_KE, FONT_NORMAL),
                (6, var_ke, NUM_KE, FONT_NORMAL),
                (7, var_pct,
                   NUM_PCT if not isinstance(var_pct, str) else None,
                   FONT_NORMAL),
            ]
            write_data_row(ws, current_row, cells)
            current_row += 1

            total_n += val_n
            total_n1 += val_n1

        # Ligne total de section
        var_total = round(total_n - total_n1, 3)
        write_total_row(ws, current_row, [
            (3, f"Total {section}",   None),
            (4, round(total_n,  3),   NUM_KE),
            (5, round(total_n1, 3),   NUM_KE),
            (6, var_total,            NUM_KE),
        ])
        current_row += 2  # ligne vide après chaque section

    logger.info(
        "Cycle '%s' : %d compte(s) injectés dans la feuille Synthèse",
        cycle, len(df_cycle),
    )


def _traiter_template(
    chemin_source: Path,
    nom_client: str,
    date_cloture: str,
    prefixe_sortie: str,
    df_cycle: Optional["pd.DataFrame"] = None,
    cycle: Optional[str] = None,
    date_n1: Optional[str] = None,
) -> Optional[tuple]:
    """
    Charge, modifie et retourne (nom_fichier_sortie, bytes_contenu) pour un template.
    Retourne None si le fichier ne peut pas être traité.

    Paramètres supplémentaires
    --------------------------
    df_cycle : pd.DataFrame ou None
        Données de balance pour le cycle de ce template.
    cycle : str ou None
        Code du cycle détecté (ex: 'F', 'A').
    date_n1 : str ou None
        Date N-1 au format 'JJ/MM/AAAA'.
    """
    wb = openpyxl.load_workbook(chemin_source)

    total_remplacements = 0
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        remove_gridlines(ws)
        total_remplacements += _remplacer_placeholders(ws, nom_client, date_cloture)

    # Injection de la balance dans la feuille Synthèse
    if df_cycle is not None and cycle is not None:
        nom_synth = next(
            (s for s in wb.sheetnames
             if "synthèse" in s.lower() or "synth" in s.lower()),
            None,
        )
        if nom_synth:
            date_fmt = pd.to_datetime(date_cloture, format="%d/%m/%Y").strftime("%d/%m/%Y")
            _injecter_balance_cycle(
                ws=wb[nom_synth],
                df_cycle=df_cycle,
                cycle=cycle,
                date_n=date_fmt,
                date_n1=date_n1 or "",
            )
        else:
            logger.warning(
                "Template '%s' : aucune feuille Synthèse trouvée — injection ignorée",
                chemin_source.name,
            )

    nom_sortie = chemin_source.name.replace(_PREFIXE_TEMPLATE, prefixe_sortie)

    # Sérialiser en mémoire
    buffer = io.BytesIO()
    wb.save(buffer)
    contenu = buffer.getvalue()

    logger.info(
        "Template '%s' → '%s' (%d feuilles, %d remplacements)",
        chemin_source.name, nom_sortie, len(wb.sheetnames), total_remplacements,
    )
    return nom_sortie, contenu


def write(
    templates_dir: Union[str, Path],
    nom_client: str,
    date_cloture: str,
    output_path: Union[str, Path],
    mapping_templates: Optional[Dict[str, str]] = None,
    balance_mappee: Optional["pd.DataFrame"] = None,
) -> Path:
    """
    Traite tous les templates et génère un ZIP de feuilles de travail.

    Paramètres
    ----------
    templates_dir : str ou Path
        Dossier contenant les fichiers 20XX_XX_YYY_*.xlsx.
    nom_client : str
        Nom du client (ex: "GILAC").
    date_cloture : str
        Date de clôture au format 'JJ/MM/AAAA'.
    output_path : str ou Path
        Dossier de sortie où sera créé le ZIP.
    mapping_templates : dict, optionnel
        {cle_fichier: cycle} depuis mapping_pcg.yaml (section 'templates').
        Si None, tous les .xlsx du dossier sont traités sans filtrage par cycle.
    balance_mappee : pd.DataFrame, optionnel
        DataFrame produit par cycle_mapper.map_cycles(). Si fourni, les données
        du cycle correspondant à chaque template sont injectées dans la feuille
        Synthèse de ce template.

    Retourne
    --------
    Path
        Chemin vers le fichier ZIP généré.
    """
    templates_dir = Path(templates_dir)
    output_dir    = Path(output_path)
    output_dir.mkdir(parents=True, exist_ok=True)

    if not templates_dir.exists():
        raise FileNotFoundError(f"Dossier templates introuvable : {templates_dir}")

    templates = sorted(templates_dir.glob("*.xlsx"))
    if not templates:
        raise ValueError(f"Aucun template .xlsx trouvé dans {templates_dir}")

    annee   = pd.to_datetime(date_cloture, format="%d/%m/%Y").year
    prefixe = _prefixe_sortie(nom_client, date_cloture)
    nom_zip = output_dir / f"FT_{nom_client}_{annee}.zip"
    date_n1 = f"31/12/{annee - 1}"

    nb_traites = 0
    nb_ignores = 0

    with zipfile.ZipFile(nom_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for chemin in templates:
            cycle = None
            df_cycle = None

            # Détection du cycle (nécessaire pour le filtrage ET l'injection)
            if mapping_templates:
                cycle = _detecter_cycle(chemin.name, mapping_templates)
                if cycle is None:
                    logger.warning(
                        "Template '%s' ignoré : aucun cycle détecté dans le mapping",
                        chemin.name,
                    )
                    nb_ignores += 1
                    continue

            # Préparer le sous-dataframe du cycle si la balance est fournie
            if balance_mappee is not None and cycle is not None:
                df_cycle = balance_mappee[
                    balance_mappee["cycle"] == cycle
                ].copy()
                if df_cycle.empty:
                    logger.warning(
                        "Template '%s' : cycle '%s' absent de la balance — "
                        "injection ignorée",
                        chemin.name, cycle,
                    )
                    df_cycle = None

            resultat = _traiter_template(
                chemin_source=chemin,
                nom_client=nom_client,
                date_cloture=date_cloture,
                prefixe_sortie=prefixe,
                df_cycle=df_cycle,
                cycle=cycle,
                date_n1=date_n1,
            )
            if resultat is None:
                nb_ignores += 1
                continue

            nom_sortie, contenu = resultat
            zf.writestr(nom_sortie, contenu)
            nb_traites += 1

    logger.info(
        "ZIP généré : %s (%d templates traités, %d ignorés)",
        nom_zip, nb_traites, nb_ignores,
    )
    return nom_zip
