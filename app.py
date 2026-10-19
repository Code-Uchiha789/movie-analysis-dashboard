"""
🎬 MOVIE ANALYSIS DASHBOARD  (Power BI -> Streamlit)

Run:  python -m streamlit run app.py

DATA FILE PATHS
---------------
Edit DATA_CANDIDATES below (or set DATA_PATH) to point at your file.
Relative paths are resolved next to this app.py. The first file that exists is used.
"""

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------- #
# CONFIG – change file paths here
# --------------------------------------------------------------------------- #
DATA_PATH = None  # e.g. r"C:\Users\Kshitij\Desktop\mymoviedb_cleaned.xlsx"
DATA_CANDIDATES = ["mymoviedb_cleaned.xlsx", "mymoviedb.csv", "data/mymoviedb_cleaned.xlsx", "data/mymoviedb.csv"]

# Colours (matched to the Power BI theme described in the brief)
BG = "#0A0F1F"
CYAN, ORANGE, GREEN, PURPLE, BLUE = "#19D3E6", "#FF8C1A", "#22E05A", "#9B5DE5", "#2E9BFF"
GRID = "rgba(255,255,255,0.12)"

# Canonical column names and the (normalised) spellings we accept for each
COLUMN_ALIASES = {
    "Release_Date": ["release_date", "releasedate", "date"],
    "Title": ["title", "movie", "movie_title", "name"],
    "Overview": ["overview", "description"],
    "Popularity": ["popularity"],
    "Vote_Count": ["vote_count", "votecount", "votes"],
    "Vote_Average": ["vote_average", "voteaverage", "rating", "average_rating"],
    "Original_Language": ["original_language", "originallanguage", "language"],
    "Genre": ["genre", "genres"],
}
REQUIRED = ["Release_Date", "Title", "Popularity", "Vote_Count", "Vote_Average", "Original_Language", "Genre"]

st.set_page_config(page_title="Movie Analysis Dashboard", page_icon="🎬", layout="wide")

# --------------------------------------------------------------------------- #
# CSS – dark Power-BI-like look, compact KPI cards, responsive grid
# --------------------------------------------------------------------------- #
st.markdown(
    f"""
<style>
.stApp {{ background-color:{BG}; color:#fff; }}
header[data-testid="stHeader"], footer, #MainMenu {{ display:none; }}
.block-container {{ padding:1.6rem 2.2rem 1.5rem 2.2rem; max-width:100%; }}
div[data-testid="stVerticalBlock"] {{ gap:1.1rem; }}
h1.dash-title {{ font-size:2.5rem; font-weight:800; letter-spacing:.5px; margin:0; padding:0.2rem 0 0 0; color:#fff; line-height:1.15; }}
/* filters */
div[data-testid="stSelectbox"] label p, div[data-testid="stDateInput"] label p {{ color:#fff; font-size:1.05rem; font-weight:600; }}
div[data-baseweb="select"] > div, div[data-testid="stDateInput"] input {{
  background-color:#141B2D; color:#fff; border:1px solid #3a3f4b; min-height:34px; }}
/* KPI cards */
.kpi-grid {{ display:grid; grid-template-columns:repeat(6,1fr); gap:18px; margin-bottom:0.6rem; }}
.kpi {{ border:1px solid #4a4f5c; border-radius:4px; padding:12px 8px; text-align:center; background:rgba(255,255,255,.02); }}
.kpi .v {{ font-size:1.7rem; font-weight:700; line-height:1.1; color:#fff; }}
.kpi .l {{ font-size:.78rem; color:#d0d4dc; margin-top:2px; }}
@media (max-width:1000px) {{ .kpi-grid {{ grid-template-columns:repeat(3,1fr); }} }}
@media (max-width:520px)  {{ .kpi-grid {{ grid-template-columns:repeat(2,1fr); }} }}
.chart-title {{ font-size:1.3rem; font-weight:700; color:#fff; margin:14px 0 4px 4px; }}
div[data-testid="stSlider"] label p {{ color:#fff; font-size:.72rem; }}
</style>
""",
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def compact(v, decimals=1):
    """10234 -> '10K', 1834 -> '1.8K', 2.1e6 -> '2.1M'."""
    if v is None or pd.isna(v):
        return "–"
    a = abs(v)
    for div, suffix in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if a >= div:
            s = f"{v / div:.{decimals}f}"
            if "." in s:
                s = s.rstrip("0").rstrip(".")
            return s + suffix
    return f"{v:,.0f}" if float(v).is_integer() else f"{v:.{decimals}f}"


def _norm(name):
    return "".join(c if c.isalnum() else "_" for c in str(name).strip().lower()).strip("_")


def standardise_columns(df):
    """Rename whatever spellings the file uses to the canonical column names."""
    lookup = {a: canon for canon, aliases in COLUMN_ALIASES.items() for a in aliases}
    renamed = {}
    for col in df.columns:
        canon = lookup.get(_norm(col))
        if canon and canon not in renamed.values():
            renamed[col] = canon
    return df.rename(columns=renamed)


def find_data_file():
    base = Path(__file__).resolve().parent
    candidates = [DATA_PATH] if DATA_PATH else DATA_CANDIDATES
    for c in candidates:
        p = Path(c)
        p = p if p.is_absolute() else base / p
        if p.exists():
            return p
    return None


def read_raw(path):
    """Return {sheet_name: DataFrame}. CSVs become a single 'csv' sheet."""
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, sheet_name=None)  # needs openpyxl
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            return {"csv": pd.read_csv(path, encoding=enc)}
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not decode the CSV file.")


@st.cache_data(show_spinner="Loading data…")
def load_data(path_str):
    """
    Returns (movies, genres, meta).
      movies : ONE row per distinct movie (Title + Release_Date)  -> KPIs, years, languages, scatter
      genres : ONE row per (movie, genre)                         -> genre count / genre rating
    Cached copies are returned by Streamlit, so filtering never mutates the source data.
    """
    sheets = {n: standardise_columns(d) for n, d in read_raw(Path(path_str)).items()}

    # Pick the movie-level sheet (has Popularity + Vote_Average) and an optional genre-only sheet
    movie_sheet = next((n for n, d in sheets.items() if {"Popularity", "Vote_Average"} <= set(d.columns)), None)
    if movie_sheet is None:
        raise KeyError(f"No sheet contains Popularity and Vote_Average. Columns found: "
                       f"{ {n: list(d.columns) for n, d in sheets.items()} }")
    raw = sheets[movie_sheet].copy()
    genre_sheet = next((n for n, d in sheets.items()
                        if n != movie_sheet and "genre" in n.lower() and {"Title", "Genre"} <= set(d.columns)), None)
    raw_genre = sheets[genre_sheet].copy() if genre_sheet else raw.copy()

    # Required columns check (Genre may come from the genre sheet)
    have = set(raw.columns) | (set(raw_genre.columns) if genre_sheet else set())
    missing = [c for c in REQUIRED if c not in have]
    if missing:
        raise KeyError(f"Missing required column(s): {missing}. Columns found: {list(raw.columns)}")

    def prep(df):
        df = df.copy()
        df["Title"] = df["Title"].astype(str).str.strip()
        df = df[df["Title"].ne("") & df["Title"].str.lower().ne("nan")]
        df["Release_Date"] = pd.to_datetime(df["Release_Date"], errors="coerce")
        # Unique movie identifier: Title + Release_Date (no ID column exists)
        df["Movie_ID"] = df["Title"] + "||" + df["Release_Date"].dt.strftime("%Y-%m-%d").fillna("NaT")
        return df

    raw, raw_genre = prep(raw), prep(raw_genre)
    for c in ("Popularity", "Vote_Count", "Vote_Average"):
        raw[c] = pd.to_numeric(raw[c], errors="coerce")
    raw["Original_Language"] = raw["Original_Language"].fillna("Unknown").astype(str).str.strip()

    # --- movie-level table: repeated genre rows carry identical attributes -> keep one row per movie
    movies = raw.drop_duplicates("Movie_ID").drop(columns=[c for c in ("Genre",) if c in raw.columns]).copy()
    year = movies["Release_Date"].dt.year
    movies["Year"] = year
    ystr = year.astype("Int64").astype(str).replace("<NA>", "n/a")
    movies["Movie_Label"] = movies["Title"] + " (" + ystr + ")"

    # --- genre-level table: split any "Action, Drama" cells, one row per (movie, genre), de-duplicated
    g = raw_genre[["Movie_ID", "Genre"]].dropna(subset=["Genre"]).copy()
    delimited = bool(g["Genre"].astype(str).str.contains(",").any())
    g["Genre"] = g["Genre"].astype(str).str.split(",")
    g = g.explode("Genre")
    g["Genre"] = g["Genre"].str.strip()
    g = g[g["Genre"].ne("") & g["Genre"].str.lower().ne("nan")].drop_duplicates(["Movie_ID", "Genre"])
    g = g[g["Movie_ID"].isin(movies["Movie_ID"])]

    n_rows = len(raw)
    meta = dict(
        file=Path(path_str).name, sheet=movie_sheet, genre_sheet=genre_sheet,
        rows=n_rows, movies=len(movies), genre_rows=len(g), delimited=delimited,
        multi_row=n_rows > len(movies),
    )
    return movies, g, meta


# --------------------------------------------------------------------------- #
# Chart helpers
# --------------------------------------------------------------------------- #
def style(fig, height, xtitle=None, ytitle=None):
    fig.update_layout(
        height=height, margin=dict(l=8, r=8, t=6, b=28),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#fff", size=11), showlegend=False, hoverlabel=dict(font_size=12),
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, title=xtitle, automargin=True)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, title=ytitle, automargin=True)
    return fig


def hbar(df, xcol, ycol, color, fmt, xtitle, hover_fmt=",.0f", height=330, label_max=None):
    """Horizontal bar chart; df must already be sorted descending.
    label_max: shorten axis labels longer than this to 'Chinese (non-stan...' (full name stays on hover)."""
    fig = go.Figure(go.Bar(
        x=df[xcol], y=df[ycol].astype(str), orientation="h", marker_color=color,
        text=[fmt(v) for v in df[xcol]], textposition="outside", cliponaxis=False,
        textfont=dict(color="#fff", size=11),
        hovertemplate=f"%{{y}}<br>{xtitle}: %{{x:{hover_fmt}}}<extra></extra>",
    ))
    style(fig, height, xtitle=xtitle)
    fig.update_yaxes(type="category", autorange="reversed", showgrid=False)
    if label_max:
        full = df[ycol].astype(str).tolist()
        short = [s if len(s) <= label_max + 3 else s[:label_max] + "..." for s in full]
        fig.update_yaxes(tickmode="array", tickvals=full, ticktext=short)
    fig.update_xaxes(range=[0, float(df[xcol].max()) * 1.2 if len(df) else 1], showgrid=True)
    return fig


def show(fig):
    st.plotly_chart(fig, config={"displayModeBar": False})


def scroll_chart(df, xcol, ycol, color, fmt, xtitle, hover_fmt=",.0f", box=330, label_max=None):
    """Bar chart inside a fixed-height scrollable box (height grows with number of bars)."""
    h = max(box - 10, 26 * len(df) + 40)
    with st.container(height=box, border=False):
        show(hbar(df, xcol, ycol, color, fmt, xtitle, hover_fmt, height=h, label_max=label_max))


# --------------------------------------------------------------------------- #
# Load data (with friendly errors)
# --------------------------------------------------------------------------- #
path = find_data_file()
if path is None:
    st.error("Data file not found. Put `mymoviedb_cleaned.xlsx` or `mymoviedb.csv` next to app.py, "
             "or set `DATA_PATH` / `DATA_CANDIDATES` at the top of app.py.")
    st.stop()
try:
    movies_all, genres_all, meta = load_data(str(path))
except KeyError as e:
    st.error(f"Column problem: {e}")
    st.stop()
except Exception as e:  # noqa: BLE001
    st.error(f"Could not read `{path.name}`: {e}")
    st.stop()

# --------------------------------------------------------------------------- #
# Header + filters
# --------------------------------------------------------------------------- #
h_title, f_genre, f_lang, f_date = st.columns([3.6, 1, 1, 1.3], gap="medium")
h_title.markdown('<h1 class="dash-title">🎬 MOVIE ANALYSIS DASHBOARD</h1>', unsafe_allow_html=True)

genre_opts = ["All"] + sorted(genres_all["Genre"].unique())
lang_opts = ["All"] + sorted(movies_all["Original_Language"].unique())
sel_genre = f_genre.selectbox("Genre", genre_opts)
sel_lang = f_lang.selectbox("Language", lang_opts)

valid_dates = movies_all["Release_Date"].dropna()
if valid_dates.empty:
    d_min = d_max = pd.Timestamp.today().date()
else:
    d_min, d_max = valid_dates.min().date(), valid_dates.max().date()
picked = f_date.date_input("Release Date", value=(d_min, d_max), min_value=d_min, max_value=d_max)
if isinstance(picked, (tuple, list)) and len(picked) == 2:
    d_from, d_to = picked
elif isinstance(picked, (tuple, list)) and len(picked) == 1:  # user is mid-selection
    d_from, d_to = picked[0], d_max
else:
    d_from, d_to = picked, d_max

# --------------------------------------------------------------------------- #
# Apply filters to COPIES (originals stay untouched)
# --------------------------------------------------------------------------- #
movies = movies_all
if sel_lang != "All":
    movies = movies[movies["Original_Language"] == sel_lang]
if (d_from, d_to) != (d_min, d_max):  # only exclude undated movies when the range is narrowed
    movies = movies[movies["Release_Date"].between(pd.Timestamp(d_from), pd.Timestamp(d_to))]

genres = genres_all[genres_all["Movie_ID"].isin(movies["Movie_ID"])]
if sel_genre != "All":
    # like a Power BI slicer on the genre table: only the chosen genre remains
    genres = genres[genres["Genre"] == sel_genre]
    movies = movies[movies["Movie_ID"].isin(genres["Movie_ID"])]

# --------------------------------------------------------------------------- #
# KPI cards (movie-level data -> no double counting of votes / ratings)
# --------------------------------------------------------------------------- #
kpis = [
    ("Total Movie", compact(movies["Movie_ID"].nunique(), 0)),
    ("Average Rating", f"{movies['Vote_Average'].mean():.2f}" if len(movies) else "–"),
    ("Average Popularity", f"{movies['Popularity'].mean():.2f}" if len(movies) else "–"),
    ("Total Votes", compact(movies["Vote_Count"].sum(), 0)),
    ("Total Languages", str(movies["Original_Language"].nunique())),
    ("Total Genres", str(genres["Genre"].nunique())),
]
st.markdown(
    '<div class="kpi-grid">' + "".join(f'<div class="kpi"><div class="v">{v}</div><div class="l">{l}</div></div>'
                                       for l, v in kpis) + "</div>",
    unsafe_allow_html=True,
)

if movies.empty:
    st.warning("No movies match the current filters.")
    st.stop()

# --------------------------------------------------------------------------- #
# Charts: 3 columns x 2 rows
# --------------------------------------------------------------------------- #
row1 = st.columns(3, gap="large")
row2 = st.columns(3, gap="large")

# Chart 1 – Total movies per year (distinct movies, orange area/line)
with row1[0]:
    st.markdown('<div class="chart-title">Total_Movie Per Year</div>', unsafe_allow_html=True)
    if movies["Year"].dropna().empty:
        st.info("No valid release dates.")
    else:
        # (year-range slider removed; the Release Date filter at the top controls the range)
        per_year = (movies.dropna(subset=["Year"]).assign(Year=lambda d: d["Year"].astype(int))
                    .groupby("Year")["Movie_ID"].nunique().reset_index(name="Movies"))
        fig = go.Figure(go.Scatter(
            x=per_year["Year"], y=per_year["Movies"], mode="lines", line=dict(color=ORANGE, width=2),
            fill="tozeroy", fillcolor="rgba(255,140,26,0.25)",
            hovertemplate="Year %{x}<br>Movies: %{y:,}<extra></extra>"))
        style(fig, 335, xtitle="Year", ytitle="Total Movie")
        show(fig)

# Chart 2 – Top 10 most popular movies
with row1[1]:
    st.markdown('<div class="chart-title">Top 10 Most Popular Movies</div>', unsafe_allow_html=True)
    top = movies.dropna(subset=["Popularity"]).nlargest(10, "Popularity").copy()
    # Same title + same year but different dates would collide as chart categories -> use full date
    dup = top["Movie_Label"].duplicated(keep=False)
    top.loc[dup, "Movie_Label"] = top["Title"] + " (" + top["Release_Date"].dt.strftime("%Y-%m-%d") + ")"
    show(hbar(top, "Popularity", "Movie_Label", ORANGE, compact, "Popularity", hover_fmt=",.1f", height=335))

# Chart 3 – Movies by genre (distinct movies per genre)
with row1[2]:
    st.markdown('<div class="chart-title">Movie by Genre</div>', unsafe_allow_html=True)
    gc = genres.groupby("Genre")["Movie_ID"].nunique().sort_values(ascending=False).reset_index(name="Movies")
    scroll_chart(gc, "Movies", "Genre", CYAN, lambda v: f"{v:,.0f}", "Movies")

# Chart 4 – Average rating by genre (genre rows de-duplicated per movie, ratings from movie-level table)
with row2[0]:
    st.markdown('<div class="chart-title">Average Rating by Genre</div>', unsafe_allow_html=True)
    gr = (genres.merge(movies[["Movie_ID", "Vote_Average"]], on="Movie_ID")
          .groupby("Genre")["Vote_Average"].mean().sort_values(ascending=False).reset_index(name="Avg"))
    scroll_chart(gr, "Avg", "Genre", GREEN, lambda v: f"{v:.2f}", "Average Rating", hover_fmt=".2f")

# Chart 5 – Movies by language
with row2[1]:
    st.markdown('<div class="chart-title">Movie by Language</div>', unsafe_allow_html=True)
    lc = (movies.groupby("Original_Language")["Movie_ID"].nunique()
          .sort_values(ascending=False).reset_index(name="Movies"))
    scroll_chart(lc, "Movies", "Original_Language", PURPLE, lambda v: f"{v:,.0f}", "Movies", label_max=17)

# Chart 6 – Popularity vs average rating (one point per movie)
with row2[2]:
    st.markdown('<div class="chart-title">Popularity vs Average Rating</div>', unsafe_allow_html=True)
    sc = movies.dropna(subset=["Popularity", "Vote_Average"])
    fig = go.Figure(go.Scatter(
        x=sc["Popularity"], y=sc["Vote_Average"], mode="markers",
        marker=dict(color=BLUE, size=6, opacity=0.65),
        customdata=np.column_stack([
            sc["Title"], sc["Release_Date"].dt.strftime("%Y-%m-%d").fillna("n/a"), sc["Vote_Count"].fillna(0)]),
        hovertemplate=("<b>%{customdata[0]}</b><br>Release: %{customdata[1]}<br>Popularity: %{x:,.1f}"
                       "<br>Avg rating: %{y:.1f}<br>Votes: %{customdata[2]:,.0f}<extra></extra>")))
    style(fig, 335, xtitle="Popularity", ytitle="Vote_Average")
    show(fig)