"""
utils.py
Fonctions de chargement des exports CMMS (RadGridExport / Cutting Analysis)
et de calcul des Pareto Niveau 1 (par catégorie / Sub Description)
et Niveau 2 (par équipement / Asset Description, à l'intérieur de chaque catégorie).
"""

import pandas as pd
import numpy as np
import plotly.graph_objects as go

# Colonnes attendues dans l'export brut (RadGridExport / onglet "Data")
COL_CATEGORY = "Sub Description"      # ex: Outils, Cutting Machine, Kit Seal, Press
COL_ASSET = "Asset Description"        # ex: V827G, M02, Kit Seal 24646
COL_DURATION = "Total Duration (MM)"   # durée d'arrêt en minutes

REQUIRED_COLS = [COL_CATEGORY, COL_ASSET, COL_DURATION]

PARETO_THRESHOLD = 80  # seuil (%) séparant causes vitales / causes secondaires


def load_raw_export(uploaded_file):
    """Charge un export .xls / .xlsx / .csv (RadGridExport ou Cutting_Analysis 'Data')
    et retourne un DataFrame nettoyé avec les colonnes nécessaires au Pareto."""
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    elif name.endswith(".xls"):
        # Les .xls exportés par RadGrid nécessitent le moteur xlrd
        df = pd.read_excel(uploaded_file, engine="xlrd")
    else:
        # .xlsx : on essaie d'abord l'onglet "Data" (Cutting_Analysis), sinon la 1ère feuille
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet = "Data" if "Data" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet)

    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Colonnes manquantes dans le fichier : {missing}. "
            f"Colonnes trouvées : {list(df.columns)}"
        )

    df = df[REQUIRED_COLS + [c for c in df.columns if c not in REQUIRED_COLS]].copy()

    # Nettoyage
    df[COL_CATEGORY] = df[COL_CATEGORY].astype(str).str.strip()
    df[COL_ASSET] = df[COL_ASSET].astype(str).str.strip()
    df[COL_DURATION] = pd.to_numeric(df[COL_DURATION], errors="coerce")

    df = df.dropna(subset=[COL_DURATION])
    df = df[~df[COL_CATEGORY].isin(["", "nan", "None", "(blank)"])]
    df = df[df[COL_DURATION] > 0]

    return df.reset_index(drop=True)


def detect_group_columns(df, max_unique=30):
    """Détecte automatiquement les colonnes catégorielles exploitables comme
    'colonne de regroupement' pour les diagrammes Pareto (2 à max_unique
    valeurs distinctes). Ne suppose AUCUN nom de colonne fixe : s'adapte à
    n'importe quel fichier (Sub Description, Position, Trade, Fault Code...)."""
    candidates = []
    for col in df.columns:
        if col in (COL_ASSET, COL_DURATION):
            continue
        if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col]):
            n = df[col].nunique(dropna=True)
            if 2 <= n <= max_unique:
                candidates.append((col, n))
    # tri par nombre de valeurs croissant (les regroupements les plus "macro" en premier)
    candidates.sort(key=lambda x: x[1])
    return [c for c, n in candidates]


def compute_pareto(df, group_col, value_col=COL_DURATION, top_n=None):
    """Calcule un tableau de Pareto générique : somme, tri décroissant,
    cumul, % cumulé, et classe (Vitale / Secondaire) selon le seuil à 80%.

    Si top_n est fourni et qu'il y a plus de catégories que top_n, les
    éléments au-delà sont regroupés dans une ligne "Autres" — indispensable
    quand un groupe (ex: "Outils") contient 60+ équipements, pour garder
    un graphique lisible."""
    agg = (
        df.groupby(group_col)[value_col]
        .sum()
        .sort_values(ascending=False)
        .reset_index()
    )
    agg.columns = [group_col, "Duration"]

    if top_n is not None and len(agg) > top_n:
        head = agg.iloc[:top_n].copy()
        others_sum = agg.iloc[top_n:]["Duration"].sum()
        others_row = pd.DataFrame({group_col: ["Autres"], "Duration": [others_sum]})
        agg = pd.concat([head, others_row], ignore_index=True)

    total = agg["Duration"].sum()
    agg["Cumulative"] = agg["Duration"].cumsum()
    agg["Cumulative %"] = (agg["Cumulative"] / total * 100).round(1)
    agg["% Individuel"] = (agg["Duration"] / total * 100).round(1)
    agg["Classe"] = np.where(agg["Cumulative %"] <= PARETO_THRESHOLD, "Vitale (≤80%)", "Secondaire (>80%)")
    return agg


def pareto_level1(df, group_col=COL_CATEGORY, top_n=None):
    """Pareto Niveau 1 : downtime total par groupe (colonne choisie dynamiquement)."""
    return compute_pareto(df, group_col, top_n=top_n)


def pareto_level2(df, group_col, group_value, top_n=10):
    """Pareto Niveau 2 : downtime par équipement (Asset Description),
    calculé UNIQUEMENT à l'intérieur d'un groupe donné (group_col/group_value
    quelconques — pas seulement 'Sub Description').
    top_n=10 par défaut pour rester lisible (un groupe peut contenir
    60+ équipements)."""
    sub = df[df[group_col] == group_value]
    return compute_pareto(sub, COL_ASSET, top_n=top_n)


def plot_pareto_chart(pareto_df, group_col, title):
    """Construit un graphique Pareto (barres + courbe cumulée) avec
    seuil à 80% matérialisé par une ligne pointillée, et un code couleur
    Vitale / Secondaire sur les barres."""
    colors = ["#D9534F" if c == "Vitale (≤80%)" else "#B0B0B0" for c in pareto_df["Classe"]]

    fig = go.Figure()

    fig.add_bar(
        x=pareto_df[group_col],
        y=pareto_df["Duration"],
        name="Downtime (min)",
        marker_color=colors,
        text=pareto_df["Duration"],
        textposition="outside",
        yaxis="y1",
    )

    fig.add_trace(go.Scatter(
        x=pareto_df[group_col],
        y=pareto_df["Cumulative %"],
        name="% Cumulé",
        mode="lines+markers",
        line=dict(color="#1F3864", width=2),
        yaxis="y2",
    ))

    # Ligne de seuil à 80%
    fig.add_shape(
        type="line",
        x0=-0.5, x1=len(pareto_df) - 0.5,
        y0=PARETO_THRESHOLD, y1=PARETO_THRESHOLD,
        xref="x", yref="y2",
        line=dict(color="#E67E22", width=1.5, dash="dash"),
    )
    fig.add_annotation(
        x=len(pareto_df) - 1, y=PARETO_THRESHOLD,
        yref="y2", text="Seuil 80%", showarrow=False,
        font=dict(color="#E67E22", size=11), yshift=10,
    )

    # "Case" (cadre rouge) autour des barres vitales (cumul ≤ 80%),
    # comme le cadre tracé à la main sur le Pareto Excel de référence.
    nb_vitales = int((pareto_df["Classe"] == "Vitale (≤80%)").sum())
    if nb_vitales > 0:
        fig.add_shape(
            type="rect",
            x0=-0.5, x1=nb_vitales - 0.5,
            y0=0, y1=pareto_df["Duration"].max() * 1.15,
            xref="x", yref="y1",
            line=dict(color="#C0392B", width=3),
            fillcolor="rgba(0,0,0,0)",
        )

    fig.update_layout(
        title=title,
        xaxis=dict(title=None, tickangle=-45, tickfont=dict(size=11)),
        yaxis=dict(title="Downtime (min)", side="left"),
        yaxis2=dict(title="% Cumulé", overlaying="y", side="right", range=[0, 110]),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
        margin=dict(t=60, b=120),
        height=480,
        plot_bgcolor="white",
    )
    return fig


def kpi_summary(df, group_col=None):
    """Quelques indicateurs clés pour le bandeau du dashboard."""
    if group_col is None:
        candidates = detect_group_columns(df)
        group_col = candidates[0] if candidates else COL_CATEGORY
    total = df[COL_DURATION].sum()
    top_cat = df.groupby(group_col)[COL_DURATION].sum().idxmax()
    top_cat_val = df.groupby(group_col)[COL_DURATION].sum().max()
    top_asset = df.groupby(COL_ASSET)[COL_DURATION].sum().idxmax()
    top_asset_val = df.groupby(COL_ASSET)[COL_DURATION].sum().max()
    nb_categories = df[group_col].nunique()
    return {
        "total_min": total,
        "top_categorie": top_cat,
        "top_categorie_min": top_cat_val,
        "top_asset": top_asset,
        "top_asset_min": top_asset_val,
        "nb_categories": nb_categories,
    }
