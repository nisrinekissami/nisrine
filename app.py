"""
app.py
------
Dashboard de diagnostic automatisé des temps d'arrêt (downtime)
Projet PFA - Génie Mécatronique - Cas Aptiv

Lancer avec :
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO

from utils import (
    load_data, standardize_columns, clean_data,
    compute_pareto, compute_pareto_level2, compute_paynter,
    compute_alerts, predict_next_week, summary_kpis,
    generate_text_summary, build_action_plan_template,
    format_excel_sheet, auto_detect_mapping, REQUIRED_COLUMNS,
)

st.set_page_config(page_title="Dashboard Downtime - Aptiv", layout="wide", page_icon="🛠️")


def fig_to_png_bytes(fig):
    """Convertit un graphique Plotly en image PNG téléchargeable (pour rapport/PPT)."""
    try:
        return fig.to_image(format="png", scale=2, width=1100, height=550)
    except Exception:
        return None

# ---------------------------------------------------------------------------
# EN-TÊTE
# ---------------------------------------------------------------------------
st.title("🛠️ Dashboard de diagnostic automatisé des temps d'arrêt")
st.caption(
    "Projet PFA — Génie Mécatronique | Remplace le calcul manuel Pareto / 5 Why / "
    "Paynter Chart réalisé aujourd'hui dans Excel."
)

with st.expander("ℹ️ Comment utiliser ce dashboard (à lire la première fois)"):
    st.markdown("""
    1. Charge un fichier Excel/CSV de pannes dans la barre latérale (ou utilise le
       **jeu de données d'exemple** pour tester tout de suite).
    2. Si les noms de colonnes de ton fichier sont différents, fais la correspondance
       dans la section **"Correspondance des colonnes"**.
    3. Les graphiques (Pareto, Paynter) et les alertes se génèrent automatiquement.
    4. Tu peux exporter les tableaux calculés en Excel via les boutons de téléchargement.
    """)

# ---------------------------------------------------------------------------
# BARRE LATÉRALE : CHARGEMENT DES DONNÉES
# ---------------------------------------------------------------------------
st.sidebar.header("1. Données")

use_sample = st.sidebar.checkbox("Utiliser le jeu de données d'exemple (Aptiv)", value=True)

uploaded_file = None
if not use_sample:
    uploaded_file = st.sidebar.file_uploader(
        "Charger un fichier de pannes (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "xlsm", "csv", "txt", "tsv"],
    )

raw_df = None
if use_sample:
    raw_df = pd.read_excel("donnees_pannes_exemple.xlsx")
    st.sidebar.success("Jeu de données d'exemple chargé (basé sur les causes réelles Aptiv).")
elif uploaded_file is not None:
    try:
        raw_df = load_data(uploaded_file)
        st.sidebar.success(f"Fichier chargé : {len(raw_df)} lignes, {len(raw_df.columns)} colonnes.")
    except Exception as e:
        st.sidebar.error(
            f"Erreur de lecture du fichier : {e}\n\n"
            "Astuce : si le fichier vient d'un très vieil export (.xls binaire), "
            "vérifie que le paquet 'xlrd' est bien installé (voir requirements.txt)."
        )

if raw_df is None:
    st.info("👈 Charge un fichier de données ou coche la case d'exemple pour démarrer.")
    st.stop()

# ---------------------------------------------------------------------------
# CORRESPONDANCE DES COLONNES (rend l'outil compatible avec n'importe quel export Aptiv)
# ---------------------------------------------------------------------------
st.sidebar.header("2. Correspondance des colonnes")
st.sidebar.caption(
    "Détection automatique (fonctionne quelle que soit la langue ou les noms de "
    "colonnes de ton fichier) — corrige si besoin."
)
cols = list(raw_df.columns)

# Détection automatique multilingue basée sur les mots-clés ET le contenu réel
# des colonnes (dates, nombres/durées, texte catégoriel) — ne dépend plus
# uniquement de noms français comme "date", "machine", "cause", "duree".
auto_mapping = auto_detect_mapping(raw_df)

def _default_index(field):
    guessed = auto_mapping.get(field)
    return cols.index(guessed) if guessed in cols else 0

mapping_ui = {}
mapping_ui["Date"] = st.sidebar.selectbox("Colonne Date", cols, index=_default_index("Date"))
mapping_ui["Machine"] = st.sidebar.selectbox("Colonne Machine / Équipement", cols, index=_default_index("Machine"))
mapping_ui["Cause"] = st.sidebar.selectbox("Colonne Cause / Défaut", cols, index=_default_index("Cause"))
mapping_ui["Duree_min"] = st.sidebar.selectbox("Colonne Durée", cols, index=_default_index("Duree_min"))

duree_unit = st.sidebar.selectbox(
    "Unité de la colonne Durée",
    ["Auto-détection", "Minutes", "Heures", "Secondes"],
    index=0,
    help="Si la colonne contient déjà du texte formaté (ex: '000 Days 00 Hrs 02 Mins', "
         "'1:30:00'), l'unité est extraite automatiquement quel que soit ce réglage.",
)
duree_unit_map = {"Auto-détection": "auto", "Minutes": "minutes", "Heures": "heures", "Secondes": "secondes"}

# Inverse le mapping : {nom_colonne_fichier: nom_standard}
inverse_mapping = {v: k for k, v in mapping_ui.items()}

try:
    df = standardize_columns(raw_df, inverse_mapping)
    df = clean_data(df, duree_unit_hint=duree_unit_map[duree_unit])
except Exception as e:
    st.error(f"Erreur de préparation des données : {e}")
    st.stop()

if df.empty:
    st.error(
        "Aucune ligne exploitable après nettoyage. Vérifie la correspondance des "
        "colonnes Date/Durée ci-contre (le format n'a peut-être pas été reconnu)."
    )
    st.stop()

if df.attrs.get("lignes_supprimees", 0) > 0:
    st.sidebar.warning(f"{df.attrs['lignes_supprimees']} ligne(s) invalide(s) ignorée(s) (date ou durée manquante).")

# ---------------------------------------------------------------------------
# FILTRES
# ---------------------------------------------------------------------------
st.sidebar.header("3. Filtres")
machines = sorted(df["Machine"].unique().tolist())
selected_machines = st.sidebar.multiselect("Filtrer par machine", machines, default=machines)
df = df[df["Machine"].isin(selected_machines)]

date_min, date_max = df["Date"].min(), df["Date"].max()
if pd.notna(date_min) and pd.notna(date_max):
    date_range = st.sidebar.date_input("Période", value=(date_min.date(), date_max.date()))
    if len(date_range) == 2:
        df = df[(df["Date"] >= pd.to_datetime(date_range[0])) & (df["Date"] <= pd.to_datetime(date_range[1]))]

if df.empty:
    st.warning("Aucune donnée pour ces filtres.")
    st.stop()

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
kpis = summary_kpis(df)
c1, c2, c3, c4 = st.columns(4)
c1.metric("Nombre de pannes", kpis["nb_evenements"])
c2.metric("Downtime total", f"{kpis['duree_totale_h']} h")
c3.metric("Machine la plus critique", kpis["machine_top"])
c4.metric("Cause la plus fréquente", kpis["cause_top"])

pareto1_preview = compute_pareto(df, group_col="Cause")
resume_text = generate_text_summary(df, pareto1_preview, group_col="Cause")
st.info(f"📝 **Résumé automatique** (copiable pour ton rapport) :\n\n{resume_text}")

st.divider()

# ---------------------------------------------------------------------------
# PARETO NIVEAU 1
# ---------------------------------------------------------------------------
st.header("📊 Pareto Niveau 1 — Causes de panne")

pareto1 = pareto1_preview

fig1 = go.Figure()
fig1.add_bar(x=pareto1["Cause"], y=pareto1["Duree_totale_min"], name="Durée (min)", marker_color="#1f77b4")
fig1.add_trace(go.Scatter(
    x=pareto1["Cause"], y=pareto1["Cumul_%"], name="Cumul %",
    yaxis="y2", mode="lines+markers", line=dict(color="red")
))
fig1.add_hline(y=80, line_dash="dash", line_color="gray", yref="y2")
fig1.update_layout(
    yaxis=dict(title="Durée totale (min)"),
    yaxis2=dict(title="Cumul %", overlaying="y", side="right", range=[0, 100]),
    xaxis=dict(tickangle=-45),
    legend=dict(orientation="h", y=1.1),
    height=450,
)
st.plotly_chart(fig1, use_container_width=True)

with st.expander("Voir le tableau Pareto niveau 1"):
    st.dataframe(pareto1, use_container_width=True)

# Export
col_exp1, col_exp2 = st.columns(2)
buf1 = BytesIO()
pareto1.to_excel(buf1, index=False)
col_exp1.download_button("⬇️ Télécharger le tableau (Excel)", buf1.getvalue(),
                          file_name="pareto_niveau1.xlsx", key="dl_pareto1_xlsx")

png1 = fig_to_png_bytes(fig1)
if png1:
    col_exp2.download_button("🖼️ Télécharger le graphique (image PNG)", png1,
                              file_name="pareto_niveau1.png", mime="image/png", key="dl_pareto1_png")

# ---------------------------------------------------------------------------
# PARETO NIVEAU 2
# ---------------------------------------------------------------------------
st.header("🔍 Pareto Niveau 2 — Détail par machine pour une cause")

top_cause = st.selectbox("Choisir la cause à décomposer", pareto1["Cause"].tolist())
pareto2 = compute_pareto_level2(df, top_value=top_cause, level1_col="Cause", level2_col="Machine")

if not pareto2.empty:
    fig2 = px.bar(pareto2, x="Machine", y="Duree_totale_min",
                   title=f"Détail machine pour la cause : {top_cause}")
    st.plotly_chart(fig2, use_container_width=True)
    st.dataframe(pareto2, use_container_width=True)

    col_exp3, col_exp4 = st.columns(2)
    buf2 = BytesIO()
    pareto2.to_excel(buf2, index=False)
    col_exp3.download_button("⬇️ Télécharger le tableau (Excel)", buf2.getvalue(),
                              file_name="pareto_niveau2.xlsx", key="dl_pareto2_xlsx")

    png2 = fig_to_png_bytes(fig2)
    if png2:
        col_exp4.download_button("🖼️ Télécharger le graphique (image PNG)", png2,
                                  file_name="pareto_niveau2.png", mime="image/png", key="dl_pareto2_png")
else:
    st.info("Pas assez de données pour cette cause.")

st.divider()

# ---------------------------------------------------------------------------
# PAYNTER CHART
# ---------------------------------------------------------------------------
st.header("📈 Paynter Chart — Évolution hebdomadaire des causes")

paynter = compute_paynter(df, group_col="Cause")

if not paynter.empty and paynter.shape[1] > 0:
    fig3 = go.Figure()
    for cause in paynter.index:
        fig3.add_trace(go.Scatter(
            x=paynter.columns, y=paynter.loc[cause], mode="lines+markers", name=cause
        ))
    fig3.update_layout(xaxis_title="Semaine", yaxis_title="Durée (min)", height=450)
    st.plotly_chart(fig3, use_container_width=True)

    with st.expander("Voir le tableau Paynter"):
        st.dataframe(paynter, use_container_width=True)

    col_exp5, col_exp6 = st.columns(2)
    buf3 = BytesIO()
    paynter.to_excel(buf3)
    col_exp5.download_button("⬇️ Télécharger le tableau (Excel)", buf3.getvalue(),
                              file_name="paynter_chart.xlsx", key="dl_paynter_xlsx")

    png3 = fig_to_png_bytes(fig3)
    if png3:
        col_exp6.download_button("🖼️ Télécharger le graphique (image PNG)", png3,
                                  file_name="paynter_chart.png", mime="image/png", key="dl_paynter_png")
else:
    st.info("Pas assez de semaines différentes dans les données pour tracer un Paynter chart.")

st.divider()

# ---------------------------------------------------------------------------
# ALERTES
# ---------------------------------------------------------------------------
st.header("🚨 Alertes machines critiques")

threshold = st.slider("Seuil d'alerte (minutes cumulées sur la période)", min_value=10, max_value=500, value=100, step=10)
alerts = compute_alerts(df, threshold_min=threshold, group_col="Machine")

if not alerts.empty:
    st.error(f"{len(alerts)} machine(s) dépassent le seuil de {threshold} min :")
    st.dataframe(alerts, use_container_width=True)
else:
    st.success("Aucune machine ne dépasse le seuil actuel.")

st.divider()

# ---------------------------------------------------------------------------
# PRÉDICTION SIMPLE
# ---------------------------------------------------------------------------
st.header("🔮 Prédiction — downtime probable la semaine prochaine")

if not paynter.empty:
    cause_pred = st.selectbox("Choisir la cause à prédire", paynter.index.tolist(), key="pred")
    result = predict_next_week(paynter, cause_pred)
    if result["prediction"] is not None:
        st.metric(f"Prédiction pour « {cause_pred} »", f"{result['prediction']} min",
                   help=f"Méthode utilisée : {result['methode']}")
    else:
        st.info("Pas assez de données pour prédire.")
else:
    st.info("Pas assez de données pour une prédiction.")

st.divider()

# ---------------------------------------------------------------------------
# EXPORT COMPLET — RAPPORT EXCEL PROFESSIONNEL (le livrable final pour Aptiv)
# ---------------------------------------------------------------------------
st.header("📦 Rapport Excel complet — prêt pour l'équipe maintenance")
st.caption(
    "Un seul fichier Excel, mis en forme et prêt à partager : synthèse, Pareto niveau 1 & 2, "
    "Paynter chart, alertes, et un plan d'action pré-rempli avec les causes prioritaires — "
    "au même format que le suivi 4Q déjà utilisé chez Aptiv. L'équipe maintenance n'a plus qu'à "
    "compléter les colonnes d'actions correctives."
)

action_plan = build_action_plan_template(pareto1, group_col="Cause", top_n=5)

buf_all = BytesIO()
with pd.ExcelWriter(buf_all, engine="openpyxl") as writer:
    # Feuille 1 : synthèse lisible en 30 secondes par un manager
    synthese = pd.DataFrame({
        "Indicateur": ["Période analysée", "Nombre de pannes", "Downtime total (h)",
                       "Machine la plus critique", "Cause la plus fréquente", "Résumé"],
        "Valeur": [
            f"{kpis['periode_debut'].date() if pd.notna(kpis['periode_debut']) else 'N/A'} au "
            f"{kpis['periode_fin'].date() if pd.notna(kpis['periode_fin']) else 'N/A'}",
            kpis["nb_evenements"], kpis["duree_totale_h"],
            kpis["machine_top"], kpis["cause_top"], resume_text,
        ],
    })
    synthese.to_excel(writer, sheet_name="Synthèse", index=False)
    format_excel_sheet(writer.sheets["Synthèse"], n_cols=2)

    pareto1.to_excel(writer, sheet_name="Pareto niveau 1", index=False)
    format_excel_sheet(writer.sheets["Pareto niveau 1"], n_cols=len(pareto1.columns))

    if not pareto2.empty:
        pareto2.to_excel(writer, sheet_name="Pareto niveau 2", index=False)
        format_excel_sheet(writer.sheets["Pareto niveau 2"], n_cols=len(pareto2.columns))

    if not paynter.empty:
        paynter.reset_index().to_excel(writer, sheet_name="Paynter chart", index=False)
        format_excel_sheet(writer.sheets["Paynter chart"], n_cols=len(paynter.columns) + 1)

    if not alerts.empty:
        alerts.to_excel(writer, sheet_name="Alertes", index=False)
        format_excel_sheet(writer.sheets["Alertes"], n_cols=len(alerts.columns))

    if not action_plan.empty:
        action_plan.to_excel(writer, sheet_name="Plan d'action", index=False)
        format_excel_sheet(writer.sheets["Plan d'action"], n_cols=len(action_plan.columns), header_color="C0392B")

st.download_button(
    "📦 Télécharger le rapport Excel complet (prêt à partager)",
    buf_all.getvalue(),
    file_name=f"Rapport_Downtime_Aptiv_{pd.Timestamp.today().date()}.xlsx",
    key="dl_all",
)

with st.expander("🗂️ Aperçu du plan d'action pré-rempli inclus dans le rapport"):
    st.dataframe(action_plan, use_container_width=True)
    st.caption(
        "Ces lignes sont générées automatiquement à partir des causes les plus impactantes du "
        "Pareto. L'ingénieur maintenance n'a plus qu'à compléter la cause racine détaillée, "
        "l'action corrective, le responsable et l'échéance — au lieu de partir d'une page blanche."
    )

st.divider()
st.caption(
    "Développé dans le cadre d'un Projet de Fin d'Année — Génie Mécatronique d'Automobile. "
    "Basé sur les données réelles de suivi downtime Aptiv (Kaizen Event 4Q)."
)
