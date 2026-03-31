import streamlit as st
import pandas as pd
import plotly.express as px
import os
from typing import Optional, Dict

# ─────────────────────────────────────────────────────────────────────────────
# Page Configuration
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Jobs Status Dashboard 2026",
    page_icon="📊",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# Local Excel file path
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Jobs_Status_Report_2026.xlsx",
)

# ─────────────────────────────────────────────────────────────────────────────
# Global Styling
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* ── Base font ── */
  html, body, [class*="css"] {
      font-family: "72", "SAP 72", Arial, Helvetica, sans-serif;
  }

  /* ── Dashboard header banner ── */
  .dash-header {
      background: linear-gradient(135deg, #0057b8 0%, #003d82 100%);
      border-radius: 12px;
      padding: 22px 32px;
      margin-bottom: 24px;
      color: #fff;
  }
  .dash-header h1 { margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -.3px; }
  .dash-header p  { margin: 6px 0 0; font-size: 13px; opacity: .82; }

  /* ── Section titles ── */
  .section-title {
      font-size: 13px; font-weight: 700; color: #32363a;
      text-transform: uppercase; letter-spacing: .07em;
      border-left: 4px solid #0070f2; padding-left: 10px;
      margin: 32px 0 14px;
  }

  /* ── KPI cards ── */
  .kpi-card {
      background: #ffffff;
      border-radius: 12px;
      padding: 18px 22px;
      box-shadow: 0 2px 12px rgba(0,0,0,.09);
      border-top: 5px solid var(--accent, #0070f2);
      height: 100%;
  }
  .kpi-label {
      font-size: 10px; color: #6a6d70; font-weight: 700;
      text-transform: uppercase; letter-spacing: .07em;
  }
  .kpi-value {
      font-size: 30px; font-weight: 800; color: #1a1a2e;
      margin: 8px 0 4px; line-height: 1;
  }
  .kpi-sub { font-size: 12px; color: #6a6d70; }

  /* ── Status badge pills (used inside kpi-sub) ── */
  .pill {
      display: inline-block; padding: 2px 10px; border-radius: 20px;
      font-size: 11px; font-weight: 700; margin-top: 6px;
  }
  .pill-pass   { background:#e6f4ea; color:#1e7e34; }
  .pill-prog   { background:#fff3cd; color:#856404; }
  .pill-open   { background:#fde8e8; color:#bb0000; }

  /* ── Dark sidebar ── */
  section[data-testid="stSidebar"] { background: #1b2838 !important; }
  section[data-testid="stSidebar"] * { color: #cdd8e5 !important; }
  section[data-testid="stSidebar"] h1,
  section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3 { color: #ffffff !important; }
  /* keep dropdown text readable on light background */
  section[data-testid="stSidebar"] [data-baseweb="select"] span,
  section[data-testid="stSidebar"] [data-baseweb="select"] div,
  section[data-testid="stSidebar"] [data-baseweb="select"] input,
  section[data-testid="stSidebar"] [data-baseweb="popover"] * { color: #1a1a2e !important; }

  /* ── Hide Streamlit chrome ── */
  #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Guard: Excel file must exist
# ─────────────────────────────────────────────────────────────────────────────
if not os.path.exists(EXCEL_PATH):
    st.error(f"Excel file not found: `{EXCEL_PATH}`")
    st.info("Place `Jobs_Status_Report_2026.xlsx` in the same folder as `script.py` and restart.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Data Loading (cached; cache busted when file mtime changes)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading data…")
def load_all_sheets(mtime: float) -> Dict[str, pd.DataFrame]:
    """Read COM and SPC sheets. Treat 'Initial Takts' as strings."""
    xl = pd.ExcelFile(EXCEL_PATH)
    sheets: Dict[str, pd.DataFrame] = {}
    for sheet in ("COM", "SPC"):
        if sheet not in xl.sheet_names:
            continue
        df = xl.parse(sheet, dtype={"Initial Takts": str})
        # Normalise 'Initial Takts' — strip whitespace, fill blanks
        df["Initial Takts"] = df["Initial Takts"].fillna("Unknown").str.strip()
        # Drop rows that are entirely blank
        df.dropna(how="all", inplace=True)
        # Normalise Status so matching is case/space insensitive
        if "Status" in df.columns:
            df["Status"] = df["Status"].fillna("Unknown").str.strip()
        sheets[sheet] = df
    return sheets


all_sheets = load_all_sheets(os.path.getmtime(EXCEL_PATH))


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar — Sheet Selection (default: SPC)
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Jobs Status 2026")
    st.markdown("---")
    sheet_choice = st.selectbox(
        "Select Project",
        options=["SPC", "COM"],   # SPC first → default index 0
        index=0,
    )
    st.markdown("---")

if sheet_choice not in all_sheets:
    st.error(f"Sheet **{sheet_choice}** not found in the workbook.")
    st.stop()

df_raw = all_sheets[sheet_choice].copy()


# ─────────────────────────────────────────────────────────────────────────────
# Column Detection — tolerates minor name variations
# ─────────────────────────────────────────────────────────────────────────────
col_map: Dict[str, str] = {c.lower().strip(): c for c in df_raw.columns}


def find_col(candidates: list) -> Optional[str]:
    """Return the first matching actual column name, or None."""
    for c in candidates:
        if c in col_map:
            return col_map[c]
    return None


COL_TAKT     = find_col(["initial takts", "initial takt", "takt"])
COL_MODULE   = find_col(["module", "modules"])
COL_CATEGORY = find_col(["category", "category of issues", "issue category", "categories"])
COL_STATUS   = find_col(["status"])

missing = [name for name, col in [
    ("Initial Takts", COL_TAKT),
    ("Module",        COL_MODULE),
    ("Category",      COL_CATEGORY),
] if col is None]

if missing:
    st.error(f"Missing required columns: {', '.join(missing)}. Found: {list(df_raw.columns)}")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# Initial Takt Filter — main page top
# ─────────────────────────────────────────────────────────────────────────────
all_takts = sorted(df_raw[COL_TAKT].dropna().unique().tolist())
valid_takts = [t for t in all_takts if t != "Unknown"]

st.markdown(f"### 📋 {sheet_choice} Dashboard")
selected_takts = st.multiselect(
    label="Initial Takt(s)",
    options=all_takts,
    default=valid_takts[-1:] if valid_takts else [],
    label_visibility="collapsed",
)

if not selected_takts:
    st.info("Select at least one Initial Takt above to load the dashboard.")
    st.stop()

# Apply Takt filter
df = df_raw[df_raw[COL_TAKT].isin(selected_takts)].copy()

if df.empty:
    st.warning("No data found for the selected Takt(s).")
    st.stop()

st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# Summary Calculations
# ─────────────────────────────────────────────────────────────────────────────
total_jobs     = len(df)
unique_modules = df[COL_MODULE].nunique()
unique_takts   = df[COL_TAKT].nunique()

# Top category
category_counts = df[COL_CATEGORY].value_counts()
top_category  = category_counts.idxmax() if not category_counts.empty else "N/A"
top_cat_count = int(category_counts.max()) if not category_counts.empty else 0

# Status counts (with graceful fallback when column is absent)
if COL_STATUS:
    status_series = df[COL_STATUS].str.lower().str.strip()
    passed_count  = int((status_series == "passed").sum())
    prog_count    = int((status_series == "in progress").sum())
    open_count    = int((status_series == "open/new").sum())
else:
    passed_count = prog_count = open_count = 0


# ─────────────────────────────────────────────────────────────────────────────
# Dashboard Header Banner
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="dash-header">
  <h1>📊 Summary of Report &mdash; {sheet_choice}</h1>
  <p>Takt(s): <strong>{', '.join(selected_takts)}</strong>
     &nbsp;&middot;&nbsp; {total_jobs:,} total jobs
     &nbsp;&middot;&nbsp; {unique_modules} modules</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# KPI Cards Row
# ─────────────────────────────────────────────────────────────────────────────
def kpi_card(col, label: str, value, sub: str, accent: str = "#0070f2") -> None:
    """Render a styled KPI card inside a Streamlit column."""
    col.markdown(f"""
    <div class="kpi-card" style="--accent:{accent};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


c1, c2, c3, c4, c5 = st.columns(5)

kpi_card(c1, "Total Jobs",
         f"{total_jobs:,}",
         f"across {unique_modules} modules",
         "#0070f2")

kpi_card(c2, "Highest Category Issue",
         top_category,
         f"{top_cat_count} occurrences",
         "#bb0000")

kpi_card(c3, "Passed",
         f"{passed_count:,}",
         f'<span class="pill pill-pass">✔ Passed</span>',
         "#1e7e34")

kpi_card(c4, "In Progress",
         f"{prog_count:,}",
         f'<span class="pill pill-prog">⏳ In Progress</span>',
         "#f0a500")

kpi_card(c5, "Open / New",
         f"{open_count:,}",
         f'<span class="pill pill-open">⚠ Open / New</span>',
         "#bb0000")

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Module Distribution — count of category entries per module
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Module Distribution (Category Count per Module)</div>",
            unsafe_allow_html=True)

# Group by module and count how many category entries exist (i.e. rows)
module_cat_counts = (
    df.groupby(COL_MODULE)[COL_CATEGORY]
    .count()
    .reset_index()
    .rename(columns={COL_MODULE: "Module", COL_CATEGORY: "Count"})
    .sort_values("Count", ascending=False)
)

fig_module = px.bar(
    module_cat_counts,
    x="Module", y="Count", text="Count",
    color="Count",
    color_continuous_scale=["#c6e0f5", "#0057b8"],
    template="plotly_white",
    title=f"Category Count per Module — {sheet_choice} | Takt(s): {', '.join(selected_takts)}",
)
fig_module.update_traces(textposition="outside", marker_line_width=0)
fig_module.update_layout(
    coloraxis_showscale=False,
    xaxis_title="Module", yaxis_title="Count",
    title_font_size=14,
    margin=dict(t=50, b=40),
    plot_bgcolor="#f9fbfd",
    paper_bgcolor="#f9fbfd",
)
st.plotly_chart(fig_module, use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# Per-Module Category Charts — 2-column grid, sorted by module count
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Category of Issues by Module</div>",
            unsafe_allow_html=True)

modules = module_cat_counts["Module"].tolist()   # already sorted by count desc

for i in range(0, len(modules), 2):
    grid_cols = st.columns(2)
    for j, gcol in enumerate(grid_cols):
        if i + j >= len(modules):
            break
        module = modules[i + j]
        df_mod = df[df[COL_MODULE] == module]

        # Count each category within this module; sort ascending for horizontal bar
        cat_data = (
            df_mod[COL_CATEGORY].value_counts()
            .reset_index()
            .rename(columns={"index": "Category", COL_CATEGORY: "Count"})
            .sort_values("Count", ascending=True)
        )
        # Handle pandas v1/v2 column naming
        if "count" in cat_data.columns:
            cat_data.rename(columns={"count": "Count"}, inplace=True)
        if cat_data.columns.tolist() == [COL_CATEGORY, "Count"]:
            cat_data.rename(columns={COL_CATEGORY: "Category"}, inplace=True)
        # Final safety rename
        cat_data.columns = ["Category", "Count"]

        fig_cat = px.bar(
            cat_data,
            x="Count", y="Category",
            orientation="h", text="Count",
            color="Count",
            color_continuous_scale=["#fde8c8", "#e9730c"],
            template="plotly_white",
            title=f"📦 {module}  ({len(df_mod)} jobs)",
        )
        fig_cat.update_traces(textposition="outside", marker_line_width=0)
        fig_cat.update_layout(
            coloraxis_showscale=False,
            yaxis_title="", xaxis_title="Count",
            title_font_size=13,
            margin=dict(t=44, b=10, l=10, r=30),
            height=max(260, len(cat_data) * 42 + 90),
            plot_bgcolor="#fdfaf6",
            paper_bgcolor="#fdfaf6",
        )
        gcol.plotly_chart(fig_cat, use_container_width=True, key=f"cat_{sheet_choice}_{module}")


# ─────────────────────────────────────────────────────────────────────────────
# Data Preview
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='section-title'>Data Preview</div>", unsafe_allow_html=True)

with st.expander("Show filtered raw data", expanded=False):
    st.dataframe(df.reset_index(drop=True), use_container_width=True, height=360)
    st.caption(f"{len(df):,} rows × {len(df.columns)} columns")
