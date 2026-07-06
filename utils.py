"""
utils.py
---------
Fonctions de calcul pour le dashboard de réduction du downtime.
Toutes les fonctions prennent un DataFrame pandas "standardisé" avec au minimum
les colonnes : Date, Machine, Cause, Duree_min
(la colonne Responsable est optionnelle)

Ce module est volontairement séparé de app.py pour que la logique de calcul
soit testable indépendamment de l'interface Streamlit.
"""

import re
import io
import unicodedata

import pandas as pd
import numpy as np

REQUIRED_COLUMNS = ["Date", "Machine", "Cause", "Duree_min"]


# ---------------------------------------------------------------------------
# CHARGEMENT DE FICHIER : .xlsx / .xls (ancien format) / .csv
# ---------------------------------------------------------------------------

def _read_excel_any_engine(file, header, sheet_name=0):
    """
    Essaie de lire un fichier Excel quel que soit son format réel :
    - .xlsx / .xlsm modernes -> openpyxl
    - .xls binaire (ancien format Excel 97-2003, souvent généré par des
      systèmes industriels type RadGrid) -> xlrd
    - Certains exports ".xls" sont en réalité du HTML ou du texte tabulé
      -> on retente une lecture HTML en dernier recours.
    Le nom de fichier n'est PAS fiable (beaucoup d'exports industriels
    s'appellent ".xls" alors qu'ils contiennent un vrai xlsx, ou l'inverse),
    donc on essaie plusieurs moteurs dans l'ordre plutôt que de se fier
    uniquement à l'extension.
    """
    errors = []
    for engine in ("openpyxl", "xlrd", "calamine"):
        try:
            file.seek(0)
        except Exception:
            pass
        try:
            return pd.read_excel(file, header=header, sheet_name=sheet_name, engine=engine)
        except Exception as e:
            errors.append(f"{engine}: {e}")

    # Dernier recours : certains "exports .xls" sont en fait du HTML (table web)
    try:
        file.seek(0)
    except Exception:
        pass
    try:
        tables = pd.read_html(file)
        if tables:
            return tables[0]
    except Exception as e:
        errors.append(f"html: {e}")

    raise ValueError(
        "Impossible de lire ce fichier comme un fichier Excel valide "
        "(essayé openpyxl / xlrd / calamine / html). Détails : " + " | ".join(errors)
    )


def detect_header_row(raw_bytes_or_file, max_scan_rows=15, engine=None, sheet_name=0) -> int:
    """
    Détecte automatiquement la ligne d'en-tête réelle d'un fichier Excel.
    De nombreux exports industriels (RadGrid, GMAO, ERP...) ont des lignes
    de titre, des lignes vides ou des sous-en-têtes avant la vraie ligne de
    colonnes. On scanne les premières lignes et on choisit celle qui a le
    plus de cellules non vides ET dont les valeurs ressemblent à du texte
    (typique d'un en-tête), plutôt que des nombres/dates (typique de données).
    """
    try:
        raw_bytes_or_file.seek(0)
    except Exception:
        pass
    preview = pd.read_excel(raw_bytes_or_file, header=None, nrows=max_scan_rows,
                             engine=engine, sheet_name=sheet_name)

    best_row, best_score = 0, -1
    for i in range(len(preview)):
        row = preview.iloc[i]
        non_null = row.notna().sum()
        if non_null == 0:
            continue
        text_like = sum(isinstance(v, str) for v in row if pd.notna(v))
        score = non_null + text_like  # favorise les lignes pleines ET textuelles
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def load_data(file) -> pd.DataFrame:
    """
    Charge un fichier Excel (.xlsx, .xls ancien format, .xlsm) ou CSV/TSV et
    retourne un DataFrame brut, quelle que soit la langue ou la mise en page
    des en-têtes. Détecte automatiquement la ligne d'en-tête et nettoie les
    noms de colonnes (espaces superflus, retours à la ligne).
    """
    name = getattr(file, "name", str(file))
    lower = name.lower()

    if lower.endswith(".csv") or lower.endswith(".txt") or lower.endswith(".tsv"):
        try:
            file.seek(0)
        except Exception:
            pass
        raw = file.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8", errors="ignore")
        sample = raw[:5000].decode("utf-8", errors="ignore")
        sep = max([",", ";", "\t"], key=sample.count)
        df = pd.read_csv(io.BytesIO(raw), sep=sep, engine="python")
    else:
        header_row = detect_header_row(file)
        try:
            file.seek(0)
        except Exception:
            pass
        df = _read_excel_any_engine(file, header=header_row)

    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
    return df


# ---------------------------------------------------------------------------
# DÉTECTION AUTOMATIQUE DES COLONNES, QUELLE QUE SOIT LA LANGUE
# ---------------------------------------------------------------------------
# Mots-clés multilingues (FR/EN/ES/DE/PT + variantes industrielles courantes,
# ex : exports "RadGrid" / GMAO en anglais). La détection ne se limite PAS à
# ces mots-clés : si aucun mot-clé ne correspond, on se rabat sur le
# CONTENU réel de la colonne (type de donnée) pour deviner quand même.
_KEYWORDS = {
    "Date": [
        "date", "fecha", "datum", "data", "heure", "time", "started", "start",
        "debut", "début", "créé", "cree", "created", "attended", "reported",
        "panne", "occurrence", "horodatage", "timestamp",
    ],
    "Machine": [
        "machine", "asset", "equipement", "équipement", "equipment", "poste",
        "station", "ligne", "line", "maquina", "máquina", "gerat", "gerät",
        "device", "outil", "tool", "resource", "ressource", "unit", "unite",
        "unité", "matricule", "id equipement",
    ],
    "Cause": [
        "cause", "raison", "reason", "fault", "defaut", "défaut", "defect",
        "motivo", "grund", "panne", "problem", "problème", "code", "task",
        "issue", "symptome", "symptôme", "nok", "root cause", "failure",
        "description",
    ],
    "Duree_min": [
        "duree", "durée", "duration", "temps", "time", "minutes", "min",
        "tiempo", "dauer", "downtime", "lost", "arret", "arrêt", "elapsed",
    ],
    "Responsable": [
        "responsable", "technicien", "technician", "operateur", "opérateur",
        "assigned", "assigne", "worker", "trade", "employee", "engineer",
        "ingenieur", "ingénieur", "owner",
    ],
}


def _normalize(text: str) -> str:
    """Minuscule + suppression des accents, pour comparer sans se soucier de la langue exacte."""
    text = str(text).lower().strip()
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def _keyword_score(col_name: str, field: str) -> float:
    norm_col = _normalize(col_name)
    score = 0.0
    for kw in _KEYWORDS[field]:
        kw_norm = _normalize(kw)
        if kw_norm == norm_col:
            score += 3.0
        elif kw_norm in norm_col:
            score += 1.5
    return score


def _datetime_content_score(series: pd.Series, sample_size=60) -> float:
    """Fraction de valeurs de la colonne interprétables comme des dates/heures."""
    sample = series.dropna()
    if sample.empty:
        return 0.0
    sample = sample.sample(min(sample_size, len(sample)), random_state=0)
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        parsed = pd.to_datetime(sample, errors="coerce")
    return parsed.notna().mean()


def _numeric_duration_content_score(series: pd.Series) -> float:
    """
    Score la plausibilité qu'une colonne représente une DURÉE :
    - valeurs numériques positives, pas trop grandes (< 100 000)
    - ou textes de type durée (ex: "1:30", "0 Days 02 Hrs 15 Mins")
    """
    sample = series.dropna()
    if sample.empty:
        return 0.0
    sample = sample.sample(min(200, len(sample)), random_state=0)

    numeric = pd.to_numeric(sample, errors="coerce")
    numeric_frac = numeric.notna().mean()
    plausible_range = ((numeric.fillna(-1) >= 0) & (numeric.fillna(1e9) < 100000)).mean()

    duration_like = sample.astype(str).str.contains(
        r"\d+\s*(?:day|jour|hr|hour|heure|h\b|min|mn|:)", case=False, regex=True, na=False
    ).mean()

    return max(numeric_frac * plausible_range, duration_like * 0.9)


def _categorical_content_score(series: pd.Series) -> float:
    """Favorise les colonnes texte avec une cardinalité raisonnable (ni ID unique, ni constante)."""
    sample = series.dropna()
    n = len(sample)
    if n == 0:
        return 0.0
    nunique = sample.astype(str).nunique()
    ratio = nunique / n
    if ratio <= 0.02 or ratio >= 0.98:
        return 0.2  # quasi constante ou quasi unique (probablement un ID) -> peu probable
    return 1.0


def auto_detect_mapping(df: pd.DataFrame) -> dict:
    """
    Devine automatiquement, pour un DataFrame de n'importe quelle langue et
    n'importe quels noms de colonnes, quelle colonne correspond à Date,
    Machine, Cause, Duree_min (et Responsable si présent).

    Combine deux signaux :
      1. Mots-clés multilingues sur le NOM de la colonne (rapide, mais
         inutile si le nom est ambigu ou dans une langue non prévue).
      2. Analyse du CONTENU réel de la colonne (type de donnée détecté :
         dates, nombres/durées, texte catégoriel) — fonctionne même si les
         noms de colonnes sont totalement différents de ce qui était prévu.

    Retourne un dict {champ_standard: nom_de_colonne_source_ou_None}.
    """
    cols = list(df.columns)
    scores = {field: {} for field in _KEYWORDS}

    for col in cols:
        series = df[col]
        kw_date = _keyword_score(col, "Date")
        kw_machine = _keyword_score(col, "Machine")
        kw_cause = _keyword_score(col, "Cause")
        kw_duree = _keyword_score(col, "Duree_min")
        kw_resp = _keyword_score(col, "Responsable")

        content_date = _datetime_content_score(series)
        content_duree = _numeric_duration_content_score(series)
        content_cat = _categorical_content_score(series)

        scores["Date"][col] = kw_date * 2 + content_date * 3
        scores["Duree_min"][col] = kw_duree * 2 + content_duree * 3
        # Machine/Cause/Responsable sont tous des colonnes texte catégorielles :
        # on s'appuie surtout sur les mots-clés, le contenu sert juste de filtre
        # (évite de choisir une colonne quasi-constante ou quasi-unique).
        scores["Machine"][col] = kw_machine * 3 + content_cat * 0.5
        scores["Cause"][col] = kw_cause * 3 + content_cat * 0.5
        scores["Responsable"][col] = kw_resp * 3 + content_cat * 0.3

    # Assignation gloutonne : on assigne d'abord le champ le plus "sûr"
    # (score maximal le plus élevé), puis on retire la colonne utilisée
    # pour ne pas l'assigner deux fois.
    mapping = {}
    remaining_cols = set(cols)
    fields_by_priority = sorted(
        scores.keys(),
        key=lambda f: max(scores[f].values()) if scores[f] else 0,
        reverse=True,
    )
    for field in fields_by_priority:
        candidates = {c: s for c, s in scores[field].items() if c in remaining_cols}
        if not candidates:
            mapping[field] = None
            continue
        best_col = max(candidates, key=candidates.get)
        best_score = candidates[best_col]
        if best_score <= 0:
            mapping[field] = None
        else:
            mapping[field] = best_col
            remaining_cols.discard(best_col)

    return mapping


def standardize_columns(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    Renomme les colonnes du fichier utilisateur vers le schéma standard
    attendu par les fonctions de calcul, à partir d'un mapping choisi
    dans l'interface (ex: {"Date panne": "Date", "Équipement": "Machine", ...}).
    """
    df = df.rename(columns=mapping)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Colonnes manquantes après mapping : {missing}. "
            f"Merci de vérifier la correspondance des colonnes."
        )
    return df


def parse_duration_to_minutes(series: pd.Series, unit_hint: str = "auto") -> pd.Series:
    """
    Convertit une colonne de durée EXPRIMÉE SOUS N'IMPORTE QUEL FORMAT COURANT
    en minutes numériques, de façon exacte (pas d'approximation) :
      - déjà numérique (minutes, heures ou secondes selon unit_hint)
      - texte type horloge "HH:MM" ou "HH:MM:SS"
      - texte type "000 Days 00 Hrs 02 Mins" / "2j 3h 15min" / "1 hr 30 min"
      - objets datetime.time ou pandas.Timedelta

    unit_hint : "auto" (détection), "minutes", "heures", "secondes".
    """
    s = series.copy()

    # Cas 1 : déjà numérique
    numeric = pd.to_numeric(s, errors="coerce")
    frac_numeric = numeric.notna().mean() if len(s) else 0

    if frac_numeric > 0.8:
        if unit_hint == "heures":
            return numeric * 60
        if unit_hint == "secondes":
            return numeric / 60
        return numeric  # "minutes" ou "auto" -> on suppose déjà en minutes

    # Cas 2 : pandas.Timedelta directement
    try:
        as_td = pd.to_timedelta(s, errors="coerce")
        if as_td.notna().mean() > 0.8:
            return as_td.dt.total_seconds() / 60
    except Exception:
        pass

    # Cas 3 : texte à parser avec des motifs "jours/heures/minutes/secondes"
    # multilingues, ou horloge HH:MM(:SS)
    def _parse_one(val):
        if pd.isna(val):
            return np.nan
        text = str(val).strip().lower()
        if text in ("", "nan", "none"):
            return np.nan

        # Horloge HH:MM:SS ou HH:MM
        clock = re.match(r"^(\d+):(\d{1,2})(?::(\d{1,2}))?$", text)
        if clock:
            h, m, sec = clock.groups()
            total = int(h) * 60 + int(m) + (int(sec) / 60 if sec else 0)
            return total

        # Motifs "X days/jours Y hrs/heures Z mins/minutes"
        days = re.search(r"(\d+)\s*(?:d|j|day|jour)", text)
        hours = re.search(r"(\d+)\s*(?:h|hr|hrs|heure)", text)
        mins = re.search(r"(\d+)\s*(?:m|min|mn|minute)", text)
        secs = re.search(r"(\d+)\s*(?:s|sec|seconde)", text)

        if days or hours or mins or secs:
            total = 0.0
            if days:
                total += int(days.group(1)) * 24 * 60
            if hours:
                total += int(hours.group(1)) * 60
            if mins:
                total += int(mins.group(1))
            if secs:
                total += int(secs.group(1)) / 60
            return total

        # Dernier recours : un seul nombre présent dans le texte
        num = re.search(r"[\d.,]+", text)
        if num:
            try:
                return float(num.group(0).replace(",", "."))
            except ValueError:
                return np.nan
        return np.nan

    return s.apply(_parse_one)


def clean_data(df: pd.DataFrame, duree_unit_hint: str = "auto") -> pd.DataFrame:
    """Nettoie les types (dates, durées numériques) et retire les lignes invalides.

    Fonctionne quel que soit le format d'origine de la durée (nombre, texte
    horloge, texte "jours/heures/minutes" en n'importe quelle langue) grâce à
    parse_duration_to_minutes, et quelle que soit la langue/format de la date
    grâce à pandas.to_datetime (détection automatique + dayfirst en repli).
    """
    df = df.copy()

    date_parsed = pd.to_datetime(df["Date"], errors="coerce")
    # Si beaucoup de dates échouent, on retente avec jour/mois inversés
    # (formats européens vs américains) et on garde le meilleur résultat.
    if date_parsed.notna().mean() < 0.5:
        alt = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
        if alt.notna().mean() > date_parsed.notna().mean():
            date_parsed = alt
    df["Date"] = date_parsed

    df["Duree_min"] = parse_duration_to_minutes(df["Duree_min"], unit_hint=duree_unit_hint)
    df["Machine"] = df["Machine"].astype(str).str.strip()
    df["Cause"] = df["Cause"].astype(str).str.strip()

    n_before = len(df)
    df = df.dropna(subset=["Date", "Duree_min"])
    df = df[df["Duree_min"] >= 0]
    n_after = len(df)

    df.attrs["lignes_supprimees"] = n_before - n_after
    return df


def compute_pareto(df: pd.DataFrame, group_col: str = "Cause") -> pd.DataFrame:
    """Pareto niveau 1 : classement des causes (ou machines) par durée totale cumulée."""
    grouped = df.groupby(group_col)["Duree_min"].sum().sort_values(ascending=False)
    pct = (grouped / grouped.sum()) * 100
    cumul = pct.cumsum()
    result = pd.DataFrame({
        "Duree_totale_min": grouped,
        "Pourcentage_%": pct.round(1),
        "Cumul_%": cumul.round(1),
    })
    result.index.name = group_col
    return result.reset_index()


def compute_pareto_level2(df: pd.DataFrame, top_value: str, level1_col: str = "Cause",
                           level2_col: str = "Machine") -> pd.DataFrame:
    """Pareto niveau 2 : détail par machine (ou autre dimension) pour UNE cause donnée."""
    subset = df[df[level1_col] == top_value]
    if subset.empty:
        return pd.DataFrame(columns=[level2_col, "Duree_totale_min", "Pourcentage_%", "Cumul_%"])
    return compute_pareto(subset, group_col=level2_col)


def compute_paynter(df: pd.DataFrame, group_col: str = "Cause") -> pd.DataFrame:
    """Tableau Paynter : durée totale par cause et par semaine ISO."""
    df = df.copy()
    df["Semaine"] = df["Date"].dt.isocalendar().year.astype(str) + "-S" + \
        df["Date"].dt.isocalendar().week.astype(str).str.zfill(2)
    paynter = df.pivot_table(
        index=group_col, columns="Semaine", values="Duree_min",
        aggfunc="sum", fill_value=0
    )
    # Trie les semaines chronologiquement
    paynter = paynter[sorted(paynter.columns)]
    return paynter


def compute_alerts(df: pd.DataFrame, threshold_min: float, group_col: str = "Machine") -> pd.DataFrame:
    """Retourne les machines (ou causes) dont le cumul de downtime dépasse le seuil."""
    totals = df.groupby(group_col)["Duree_min"].sum().sort_values(ascending=False)
    alerts = totals[totals > threshold_min]
    return alerts.reset_index().rename(columns={"Duree_min": "Duree_totale_min"})


def predict_next_week(paynter: pd.DataFrame, row_name: str) -> dict:
    """
    Prédiction simple de la durée de panne de la semaine suivante pour une cause donnée,
    par régression linéaire sur l'historique des semaines. Si moins de 3 points de
    données, retombe sur une moyenne mobile simple.
    """
    if row_name not in paynter.index:
        return {"prediction": None, "methode": "aucune donnée"}

    y = paynter.loc[row_name].values.astype(float)
    n = len(y)

    if n < 3:
        pred = float(np.mean(y)) if n > 0 else 0.0
        return {"prediction": round(pred, 1), "methode": "moyenne simple (peu de données)"}

    X = np.arange(n).reshape(-1, 1)
    try:
        from sklearn.linear_model import LinearRegression
        model = LinearRegression().fit(X, y)
        pred = model.predict([[n]])[0]
        pred = max(0.0, float(pred))  # une durée ne peut pas être négative
        return {"prediction": round(pred, 1), "methode": "régression linéaire"}
    except ImportError:
        # Fallback si scikit-learn n'est pas installé : moyenne mobile sur les 4 dernières semaines
        pred = float(np.mean(y[-4:]))
        return {"prediction": round(pred, 1), "methode": "moyenne mobile (scikit-learn indisponible)"}


def generate_text_summary(df: pd.DataFrame, pareto1: pd.DataFrame, group_col: str = "Cause") -> str:
    """
    Génère un court résumé en français, directement réutilisable dans un rapport
    ou une présentation, à partir des résultats calculés.
    """
    if df.empty or pareto1.empty:
        return "Pas assez de données pour générer un résumé."

    nb_events = len(df)
    total_h = round(df["Duree_min"].sum() / 60, 1)
    top_row = pareto1.iloc[0]
    top_name = top_row[group_col]
    top_pct = top_row["Pourcentage_%"]
    top_2_cumul = pareto1.iloc[min(1, len(pareto1) - 1)]["Cumul_%"]

    nb_causes_80 = (pareto1["Cumul_%"] <= 80).sum() + 1
    nb_causes_80 = min(nb_causes_80, len(pareto1))

    resume = (
        f"Sur la période analysée, {nb_events} événements de panne ont été enregistrés, "
        f"représentant un total de {total_h} heures d'arrêt. "
        f"La cause principale est « {top_name} », responsable à elle seule de {top_pct}% "
        f"du downtime total. "
        f"Conformément à la règle de Pareto (80/20), {nb_causes_80} cause(s) sur {len(pareto1)} "
        f"expliquent déjà {top_2_cumul}% du problème. "
        f"Il est donc recommandé de concentrer les actions correctives en priorité sur « {top_name} » "
        f"avant de traiter les causes secondaires."
    )
    return resume


def build_action_plan_template(pareto1: pd.DataFrame, group_col: str = "Cause", top_n: int = 5) -> pd.DataFrame:
    """
    Génère un plan d'action pré-rempli avec les causes prioritaires (issues du Pareto),
    dans le même format que le tableau "Action Plan" utilisé en interne chez Aptiv
    (Primary Reason Code / Secondary Reason Code + Root Cause / Corrective Action / Effect).
    Les colonnes à remplir manuellement par l'équipe maintenance sont laissées vides.
    """
    if pareto1.empty:
        return pd.DataFrame()

    top = pareto1.head(top_n).copy()
    plan = pd.DataFrame({
        "Priorité": range(1, len(top) + 1),
        "Primary Reason Code": top[group_col],
        "Root Cause Effect (%)": top["Pourcentage_%"],
        "Secondary Reason Code / Root Cause (à compléter)": "",
        "Permanent Corrective Action (à compléter)": "",
        "Responsable (à compléter)": "",
        "Date échéance (à compléter)": "",
        "Statut (à compléter)": "À traiter",
    })
    return plan


def format_excel_sheet(worksheet, n_cols, header_color="1F4E78"):
    """
    Applique une mise en forme professionnelle à une feuille Excel :
    en-tête coloré et en gras, texte blanc, colonnes ajustées à la largeur du contenu,
    première ligne figée. Rend le fichier exporté directement présentable en réunion.
    """
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx in range(1, n_cols + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Ajuste automatiquement la largeur des colonnes selon leur contenu
    for col_idx in range(1, n_cols + 1):
        letter = get_column_letter(col_idx)
        max_len = 10
        for row in worksheet.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
        worksheet.column_dimensions[letter].width = min(max_len + 3, 45)

    worksheet.freeze_panes = "A2"


def summary_kpis(df: pd.DataFrame) -> dict:
    """Indicateurs clés affichés en haut du dashboard."""
    return {
        "nb_evenements": len(df),
        "duree_totale_min": round(df["Duree_min"].sum(), 1),
        "duree_totale_h": round(df["Duree_min"].sum() / 60, 1),
        "machine_top": df.groupby("Machine")["Duree_min"].sum().idxmax() if not df.empty else "N/A",
        "cause_top": df.groupby("Cause")["Duree_min"].sum().idxmax() if not df.empty else "N/A",
        "periode_debut": df["Date"].min(),
        "periode_fin": df["Date"].max(),
    }
