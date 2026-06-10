"""
Interface Streamlit du pipeline d'audit.

Lance avec : python3 -m streamlit run app.py
"""

import io
import logging
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Config de la page
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Audit Automation",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

_PCG_DEFAULT = Path(__file__).parent / "src" / "config" / "mapping_pcg.yaml"
_TEMPLATES_DEFAULT = Path(__file__).parent / "data" / "templates"

logging.basicConfig(level=logging.INFO)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sev_emoji(sev: str) -> str:
    return {"BLOQUANT": "🔴", "WARNING": "🟠", "INFO": "🔵"}.get(sev, "⚪")


def _ok_emoji(ok: bool) -> str:
    return "✅" if ok else "❌"


@st.cache_data(show_spinner=False)
def _run_pipeline_cached(
    fec_bytes: bytes,
    fec_nom: str,
    client: str,
    date_cloture: str,
    n1_bytes: Optional[bytes],
    n1_nom: Optional[str],
    templates_bytes: Optional[dict],  # {nom: bytes}
    pcg_path: str,
    bilan_non_bloquant: bool = False,
) -> dict:
    """
    Exécute le pipeline dans un dossier temporaire et retourne les résultats
    (chemins, bytes des fichiers, contrôles, stats).
    """
    from main import run_pipeline

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Écrire le FEC
        fec_path = tmp / fec_nom
        fec_path.write_bytes(fec_bytes)

        # N-1 (FM, FEC ou balance simple) — conserver l'extension d'origine
        # pour que run_pipeline puisse auto-détecter le format
        n1_path = None
        if n1_bytes and n1_nom:
            n1_path = tmp / n1_nom
            n1_path.write_bytes(n1_bytes)

        # Templates
        tpl_dir = None
        if templates_bytes:
            tpl_dir = tmp / "templates"
            tpl_dir.mkdir()
            for nom, data in templates_bytes.items():
                (tpl_dir / nom).write_bytes(data)

        output_dir = tmp / "output"

        resultats = run_pipeline(
            fec_path=str(fec_path),
            client=client,
            date_cloture=date_cloture,
            n1_fm=str(n1_path) if n1_path else None,
            templates_dir=str(tpl_dir) if tpl_dir else None,
            output_dir=str(output_dir),
            pcg_config_path=pcg_path,
            bilan_non_bloquant=bilan_non_bloquant,
        )

        # Lire les fichiers générés en mémoire avant la fin du contexte temp
        if "fm_path" in resultats:
            resultats["fm_bytes"] = Path(resultats["fm_path"]).read_bytes()
            resultats["fm_nom"]   = Path(resultats["fm_path"]).name
        if "zip_path" in resultats:
            resultats["zip_bytes"] = Path(resultats["zip_path"]).read_bytes()
            resultats["zip_nom"]   = Path(resultats["zip_path"]).name

        return resultats


# ---------------------------------------------------------------------------
# Sidebar — Paramètres
# ---------------------------------------------------------------------------

with st.sidebar:
    st.title("📋 Audit Automation")
    st.caption("Pipeline FEC → FM + Feuilles de travail")
    st.divider()

    st.subheader("1 · FEC (obligatoire)")
    fec_upload = st.file_uploader(
        "Fichier des Écritures Comptables (.txt)",
        type=["txt"],
        help="Encodage auto-détecté (UTF-8, CP1252, Latin-1). Séparateur TAB/pipe/; détecté.",
    )

    st.subheader("2 · Paramètres client")
    client = st.text_input("Nom du client", value="", placeholder="ex: ACME")
    date_cloture = st.text_input(
        "Date de clôture", value="31/12/2025", placeholder="JJ/MM/AAAA"
    )

    st.subheader("3 · N-1 (obligatoire)")
    st.caption("Formats acceptés : FM N-1 (.xlsx), FEC N-1 (.txt), Balance N-1 (.xlsx)")
    n1_upload = st.file_uploader(
        "Fichier N-1",
        type=["xlsx", "txt"],
        help=(
            "FM N-1 : fichier Excel avec onglet 'Balance N Vs N-1'.\n"
            "FEC N-1 : FEC brut de l'exercice précédent (.txt).\n"
            "Balance N-1 : Excel simple avec colonnes CompteNum, CompteLib, Solde (en €)."
        ),
    )

    st.subheader("4 · Templates (optionnel)")
    use_bundled_tpl = st.checkbox(
        "Utiliser les templates du cabinet",
        value=_TEMPLATES_DEFAULT.exists(),
        help=f"Templates dans {_TEMPLATES_DEFAULT}",
    )
    tpl_uploads = None
    if not use_bundled_tpl:
        tpl_uploads = st.file_uploader(
            "Templates .xlsx (multi-sélection)",
            type=["xlsx"],
            accept_multiple_files=True,
        )

    st.subheader("5 · Options")
    bilan_non_bloquant = st.checkbox(
        "Bilan non bloquant",
        value=False,
        help=(
            "Si coché, le contrôle d'équilibre du bilan (AC-1) devient un "
            "WARNING : le pipeline continue même si Total Actif ≠ Total Passif "
            "sur l'exercice N."
        ),
    )

    st.divider()
    run_btn = st.button("🚀 Lancer le pipeline", type="primary", use_container_width=True)

# ---------------------------------------------------------------------------
# Invalidation du cache session si les inputs changent
# ---------------------------------------------------------------------------
_current_input_key = (
    fec_upload.name if fec_upload else None,
    fec_upload.size if fec_upload else None,
    client.strip(),
    date_cloture,
    n1_upload.name if n1_upload else None,
    n1_upload.size if n1_upload else None,
    use_bundled_tpl,
    tuple(sorted(f.name for f in tpl_uploads)) if tpl_uploads else None,
    bilan_non_bloquant,
)

if st.session_state.get("_input_key") != _current_input_key:
    # Inputs ont changé — effacer les résultats stockés
    st.session_state.pop("resultats", None)
    st.session_state["_input_key"] = _current_input_key

# ---------------------------------------------------------------------------
# Zone principale
# ---------------------------------------------------------------------------

st.title("Audit Automation — Pipeline FEC")

# Lancer le pipeline uniquement quand le bouton est cliqué
if run_btn:
    # Validation des entrées
    if not fec_upload:
        st.error("Veuillez fournir un fichier FEC.")
        st.stop()
    if not client.strip():
        st.error("Le nom du client est obligatoire.")
        st.stop()
    try:
        pd.to_datetime(date_cloture, format="%d/%m/%Y")
    except ValueError:
        st.error(f"Date de clôture invalide : '{date_cloture}'. Format attendu : JJ/MM/AAAA")
        st.stop()
    if not n1_upload:
        st.error("Le N-1 est obligatoire. Veuillez fournir un FM N-1, un FEC N-1 ou une balance N-1.")
        st.stop()

    # Préparer les templates
    templates_bytes = None
    if use_bundled_tpl and _TEMPLATES_DEFAULT.exists():
        templates_bytes = {
            f.name: f.read_bytes()
            for f in _TEMPLATES_DEFAULT.glob("*.xlsx")
        }
    elif tpl_uploads:
        templates_bytes = {f.name: f.read() for f in tpl_uploads}

    # Exécution du pipeline
    with st.spinner("Pipeline en cours…"):
        try:
            resultats = _run_pipeline_cached(
                fec_bytes=fec_upload.read(),
                fec_nom=fec_upload.name,
                client=client.strip(),
                date_cloture=date_cloture,
                n1_bytes=n1_upload.read(),
                n1_nom=n1_upload.name,
                templates_bytes=templates_bytes,
                pcg_path=str(_PCG_DEFAULT),
                bilan_non_bloquant=bilan_non_bloquant,
            )
        except ValueError as exc:
            st.error(f"**Erreur bloquante** : {exc}")
            st.stop()
        except Exception as exc:
            st.exception(exc)
            st.stop()

    # Stocker les résultats en session_state pour persister entre les re-runs
    st.session_state["resultats"] = resultats

# Afficher les résultats s'ils existent en session_state
if "resultats" not in st.session_state:
    st.info(
        "**Comment utiliser :**\n"
        "1. Uploadez le FEC dans la barre latérale\n"
        "2. Saisissez le nom client et la date de clôture\n"
        "3. Uploadez le N-1 (FM .xlsx, FEC .txt ou balance .xlsx) — obligatoire\n"
        "4. (Optionnel) Utilisez les templates du cabinet pour les feuilles de travail\n"
        "5. Cliquez sur **Lancer le pipeline**"
    )
    st.stop()

resultats = st.session_state["resultats"]

# ---------------------------------------------------------------------------
# Résultats — KPIs
# ---------------------------------------------------------------------------
st.success("Pipeline terminé avec succès !")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Écritures FEC", f"{resultats['fec_lignes']:,}")
col2.metric("Comptes distincts", resultats["nb_comptes"])
nb_ok  = sum(1 for _, ok, _, _ in resultats["controles"] if ok)
nb_ko  = len(resultats["controles"]) - nb_ok
col3.metric("Contrôles OK", nb_ok, delta=f"-{nb_ko} anomalie(s)" if nb_ko else None,
            delta_color="inverse")
col4.metric("Client", client.strip())

st.divider()

# ---------------------------------------------------------------------------
# Résultats — Onglets
# ---------------------------------------------------------------------------
tab_ctrl, tab_balance, tab_cycles, tab_dl = st.tabs(
    ["🔍 Contrôles", "📊 Balance", "🗂️ Cycles", "📥 Téléchargements"]
)

# --- Onglet Contrôles ---
with tab_ctrl:
    st.subheader("Résultats des contrôles d'intégrité")
    for nom, ok, detail, sev in resultats["controles"]:
        with st.expander(f"{_ok_emoji(ok)} {_sev_emoji(sev)} **{nom}** — {detail.split(chr(10))[0]}"):
            st.text(detail)
            st.caption(f"Sévérité : {sev}")

# --- Onglet Balance ---
with tab_balance:
    st.subheader("Balance générale N vs N-1")

    # Balance mappée exposée directement par run_pipeline (pas de
    # reconstruction : la logique N-1 est centralisée dans
    # src/parsers/balance_n1_loader.py)
    bm = resultats["balance_mappee"]

    # Filtres
    col_f1, col_f2 = st.columns([2, 1])
    with col_f1:
        filtre_compte = st.text_input("Filtrer par numéro/libellé", placeholder="ex: 401 ou FOURNISSEUR")
    with col_f2:
        filtre_cycle = st.selectbox("Cycle", ["Tous"] + sorted(bm["cycle"].unique().tolist()))

    df_affiche = bm.copy()
    if filtre_compte:
        masque = (
            df_affiche["CompteNum"].str.contains(filtre_compte, case=False, na=False) |
            df_affiche["CompteLib"].str.contains(filtre_compte, case=False, na=False)
        )
        df_affiche = df_affiche[masque]
    if filtre_cycle != "Tous":
        df_affiche = df_affiche[df_affiche["cycle"] == filtre_cycle]

    cols_affiche = ["CompteNum", "CompteLib", "Solde_KE", "Solde_N1_KE", "Var_KE", "Var_PCT", "cycle", "compta"]
    st.dataframe(
        df_affiche[cols_affiche].rename(columns={
            "Solde_KE": "N (K€)", "Solde_N1_KE": "N-1 (K€)",
            "Var_KE": "Var K€", "Var_PCT": "Var %",
        }),
        use_container_width=True,
        height=450,
    )
    st.caption(f"{len(df_affiche)} compte(s) affichés sur {len(bm)}")

# --- Onglet Cycles ---
with tab_cycles:
    st.subheader("Répartition par cycle d'audit")

    resume = (
        bm.groupby("cycle")
        .agg(
            Comptes=("CompteNum", "count"),
            Solde_N=("Solde_KE", "sum"),
            Solde_N1=("Solde_N1_KE", "sum"),
        )
        .reset_index()
        .rename(columns={"cycle": "Cycle", "Solde_N": "N (K€)", "Solde_N1": "N-1 (K€)"})
    )
    resume["Var (K€)"] = (resume["N (K€)"] - resume["N-1 (K€)"]).round(1)
    resume["N (K€)"]   = resume["N (K€)"].round(1)
    resume["N-1 (K€)"] = resume["N-1 (K€)"].round(1)

    ORDRE = ["C Propres","C PRC","F","I Incorp","I Corp","I Fi","S","A","V","P","E","T","X"]
    resume["_ord"] = resume["Cycle"].map(lambda c: ORDRE.index(c) if c in ORDRE else 99)
    resume = resume.sort_values("_ord").drop(columns="_ord")

    st.dataframe(resume, use_container_width=True, hide_index=True)

    # Sections par cycle sélectionné
    cycle_sel = st.selectbox("Détail du cycle", ORDRE, key="cycle_detail")
    df_sel = bm[bm["cycle"] == cycle_sel][["CompteNum","CompteLib","compta","Solde_KE","Solde_N1_KE","Var_KE"]]
    if df_sel.empty:
        st.info(f"Aucun compte dans le cycle {cycle_sel} pour ce FEC.")
    else:
        for section in ["Actif","Passif","Charges","Produits"]:
            df_s = df_sel[df_sel["compta"].str.capitalize() == section]
            if df_s.empty:
                continue
            st.markdown(f"**{section.upper()}** — {len(df_s)} compte(s), total N = **{df_s['Solde_KE'].sum():.0f} K€**")
            st.dataframe(df_s.rename(columns={"Solde_KE":"N (K€)","Solde_N1_KE":"N-1 (K€)","Var_KE":"Var K€"}),
                         use_container_width=True, hide_index=True, height=200)

# --- Onglet Téléchargements ---
with tab_dl:
    st.subheader("Fichiers générés")

    if "travail_bytes" in resultats:
        st.download_button(
            label=f"📥 Fichier de travail — {resultats['travail_nom']}",
            data=resultats["travail_bytes"],
            file_name=resultats["travail_nom"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning("Fichier de travail non généré.")

    if "fm_bytes" in resultats:
        st.download_button(
            label=f"📥 Feuilles maîtresses — {resultats['fm_nom']}",
            data=resultats["fm_bytes"],
            file_name=resultats["fm_nom"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning("Feuilles maîtresses non générées.")

    if "zip_bytes" in resultats:
        st.download_button(
            label=f"📥 Feuilles de travail — {resultats['zip_nom']}",
            data=resultats["zip_bytes"],
            file_name=resultats["zip_nom"],
            mime="application/zip",
            use_container_width=True,
        )
        # Liste des fichiers dans le ZIP
        with zipfile.ZipFile(io.BytesIO(resultats["zip_bytes"])) as zf:
            st.caption(f"{len(zf.namelist())} fichiers dans le ZIP :")
            for nom in sorted(zf.namelist()):
                st.caption(f"  • {nom}")
    else:
        st.info("Feuilles de travail non demandées (aucun dossier templates fourni).")
