from pathlib import Path
import streamlit as st
import pandas as pd
from utils import (
    load_raw_export, pareto_level1, pareto_level2,
    plot_pareto_chart, kpi_summary, detect_group_columns, COL_CATEGORY, COL_ASSET,
)

# Chemin absolu basé sur l'emplacement réel de app.py (évite les erreurs
# de chemin relatif selon le répertoire de travail sur Streamlit Cloud)
APP_DIR = Path(__file__).resolve().parent
LOGO_PATH = APP_DIR / "versigent_logo.png"

st.set_page_config(
    page_title="Maintenance Dashboard - Aptiv",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "📊",
    layout="wide",
)

# ---------------------------------------------------------------
# En-tête avec logo Versigent
# ---------------------------------------------------------------
col_logo, col_title = st.columns([1, 5])
with col_logo:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=160)
    else:
        st.warning(f"Logo introuvable : {LOGO_PATH.name} (vérifier qu'il est bien à la racine du repo, au même niveau que app.py).")
with col_title:
    st.title("Maintenance Downtime Dashboard")
    st.caption("Aptiv Oujda — Analyse Pareto Cutting / Maintenance (Niveau 1 & Niveau 2)")

st.divider()

# ---------------------------------------------------------------
# Import du fichier
# ---------------------------------------------------------------
uploaded_file = st.sidebar.file_uploader(
    "Importer l'export CMMS (RadGridExport.xls / Cutting_Analysis.xlsx / .csv)",
    type=["xls", "xlsx", "csv"],
)
st.sidebar.markdown("---")
st.sidebar.caption(
    "L'app calcule automatiquement :\n"
    "- **Pareto Niveau 1** : downtime par catégorie (Sub Description)\n"
    "- **Pareto Niveau 2** : downtime par équipement, **séparé par catégorie**"
)

if uploaded_file is None:
    st.info("⬅️ Importez un fichier d'export (RadGridExport ou Cutting Analysis) pour lancer l'analyse.")
    st.stop()

try:
    df = load_raw_export(uploaded_file)
except Exception as e:
    st.error(f"Erreur de lecture du fichier : {e}")
    st.stop()

if df.empty:
    st.warning("Aucune donnée exploitable après nettoyage (vérifier les colonnes du fichier).")
    st.stop()

# ---------------------------------------------------------------
# Colonne de regroupement — détectée automatiquement dans le fichier,
# jamais figée sur un nom de colonne fixe (le fichier change chaque semaine).
# ---------------------------------------------------------------
group_candidates = detect_group_columns(df)
if not group_candidates:
    st.error("Aucune colonne catégorielle exploitable (2 à 30 valeurs distinctes) n'a été trouvée pour regrouper le Pareto.")
    st.stop()

default_idx = group_candidates.index(COL_CATEGORY) if COL_CATEGORY in group_candidates else 0
group_col = st.sidebar.selectbox(
    "Colonne de regroupement des diagrammes Pareto",
    group_candidates,
    index=default_idx,
    help="Chaque valeur distincte de cette colonne génère son propre diagramme Pareto Niveau 2. "
         "Change cette colonne si le fichier de la semaine utilise un découpage différent "
         "(ex: Sub Description, Position, Trade...).",
)

# ---------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------
kpi = kpi_summary(df, group_col=group_col)
k1, k2, k3, k4 = st.columns(4)
k1.metric("Downtime total (min)", f"{kpi['total_min']:.0f}")
k2.metric("Catégorie la + impactante", kpi["top_categorie"], f"{kpi['top_categorie_min']:.0f} min")
k3.metric("Équipement le + impactant", kpi["top_asset"], f"{kpi['top_asset_min']:.0f} min")
k4.metric("Nb. catégories", kpi["nb_categories"])

st.divider()

# ---------------------------------------------------------------
# Onglets
# ---------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 Pareto Niveau 1 (par catégorie)", "🔍 Pareto Niveau 2 (par équipement)", "🧾 Données"])

with tab1:
    st.subheader(f"Pareto Niveau 1 — Downtime par « {group_col} »")
    p1 = pareto_level1(df, group_col=group_col)
    st.plotly_chart(
        plot_pareto_chart(p1, group_col, f"Downtime total par {group_col}"),
        use_container_width=True,
    )
    st.dataframe(p1, use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Pareto Niveau 2 — un graphique séparé par groupe")
    st.caption(f"Un diagramme par valeur distincte de « {group_col} » — leur nombre s'adapte automatiquement au fichier importé.")

    groups = pareto_level1(df, group_col=group_col)[group_col].tolist()

    top_n = st.slider(
        "Nombre d'équipements affichés par groupe (Top N, le reste est regroupé en 'Autres')",
        min_value=5, max_value=25, value=10, step=1,
    )

    view_mode = st.radio("Mode d'affichage", ["Grille (tous les groupes)", "Un groupe à la fois"], horizontal=True)

    if view_mode == "Un groupe à la fois":
        cat_choice = st.selectbox("Choisir un groupe", groups)
        p2 = pareto_level2(df, group_col, cat_choice, top_n=top_n)
        st.plotly_chart(
            plot_pareto_chart(p2, COL_ASSET, f"Pareto Niveau 2 — {cat_choice}"),
            use_container_width=True,
        )
        st.dataframe(p2, use_container_width=True, hide_index=True)
    else:
        # Pleine largeur, un groupe par ligne : évite le chevauchement
        # d'étiquettes qu'on avait en grille 2 colonnes avec des groupes
        # qui comptent 60+ équipements.
        for cat in groups:
            p2 = pareto_level2(df, group_col, cat, top_n=top_n)
            if p2.empty:
                continue
            st.plotly_chart(
                plot_pareto_chart(p2, COL_ASSET, f"Pareto Niveau 2 — {cat}"),
                use_container_width=True,
            )

with tab3:
    st.subheader("Données nettoyées")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button(
        "Télécharger les données nettoyées (CSV)",
        df.to_csv(index=False).encode("utf-8"),
        file_name="donnees_nettoyees.csv",
        mime="text/csv",
    )
