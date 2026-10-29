"""
CVC Analytics — Streamlit App
Searches and analyses CVC company information from CVC_Company_Info.xlsx
"""

import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="CVC Analytics",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp {
    background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
    min-height: 100vh;
}

[data-testid="stSidebar"] {
    background: rgba(255,255,255,0.04);
    border-right: 1px solid rgba(255,255,255,0.08);
    backdrop-filter: blur(12px);
}

/* Sidebar — all text white */
[data-testid="stSidebar"] * {
    color: #ffffff !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div {
    color: #ffffff !important;
}
/* Radio button labels */
[data-testid="stSidebar"] [data-testid="stRadio"] label,
[data-testid="stSidebar"] [data-testid="stRadio"] span {
    color: #ffffff !important;
}
/* Multiselect & text input labels */
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] .stMultiSelect label {
    color: #ffffff !important;
    font-weight: 500;
}
/* Sidebar markdown bold text */
[data-testid="stSidebar"] strong {
    color: #ffffff !important;
}
/* Selected multiselect tags keep their purple tint but stay readable */
[data-testid="stSidebar"] [data-baseweb="tag"] span {
    color: #ffffff !important;
}

.hero-title {
    font-size: 2.6rem;
    font-weight: 700;
    background: linear-gradient(90deg, #a78bfa, #60a5fa, #34d399);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 0.2rem;
}
.hero-sub {
    color: rgba(255,255,255,0.55);
    font-size: 1rem;
    margin-bottom: 1.5rem;
}

.metric-card {
    background: rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.1);
    border-radius: 16px;
    padding: 1.2rem 1.4rem;
    backdrop-filter: blur(8px);
    transition: border-color 0.2s;
}
.metric-card:hover { border-color: rgba(167,139,250,0.5); }
.metric-value {
    font-size: 2rem;
    font-weight: 700;
    color: #a78bfa;
}
.metric-label {
    font-size: 0.8rem;
    color: rgba(255,255,255,0.5);
    text-transform: uppercase;
    letter-spacing: 0.06em;
}

.company-card {
    background: rgba(255,255,255,0.05);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 0.9rem;
    transition: border-color 0.2s, transform 0.15s;
}
.company-card:hover {
    border-color: rgba(96,165,250,0.5);
    transform: translateY(-2px);
}
.company-name {
    font-size: 1.05rem;
    font-weight: 600;
    color: #60a5fa;
    margin-bottom: 0.4rem;
}
.company-meta span {
    display: inline-block;
    background: rgba(167,139,250,0.15);
    border: 1px solid rgba(167,139,250,0.25);
    border-radius: 999px;
    padding: 2px 10px;
    font-size: 0.73rem;
    color: #c4b5fd;
    margin-right: 6px;
    margin-bottom: 4px;
}
.company-overview {
    color: rgba(255,255,255,0.65);
    font-size: 0.85rem;
    line-height: 1.55;
    margin-top: 0.5rem;
}
.badge-yes {
    background: rgba(52,211,153,0.15) !important;
    border-color: rgba(52,211,153,0.3) !important;
    color: #6ee7b7 !important;
}
.badge-no {
    background: rgba(248,113,113,0.12) !important;
    border-color: rgba(248,113,113,0.25) !important;
    color: #fca5a5 !important;
}

.section-header {
    font-size: 1.1rem;
    font-weight: 600;
    color: rgba(255,255,255,0.85);
    border-left: 3px solid #a78bfa;
    padding-left: 0.6rem;
    margin-top: 1.2rem;
    margin-bottom: 0.8rem;
}

/* Main area labels */
label { color: rgba(255,255,255,0.85) !important; }

[data-baseweb="tab"] {
    color: rgba(255,255,255,0.55) !important;
    font-weight: 500;
}
[aria-selected="true"] {
    color: #a78bfa !important;
    border-bottom-color: #a78bfa !important;
}

hr { border-color: rgba(255,255,255,0.08) !important; }
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(167,139,250,0.4); border-radius: 99px; }
</style>
""", unsafe_allow_html=True)

# ── Data loading ───────────────────────────────────────────────────────────────
DATA_PATH = Path(__file__).parent / "CVC_Company_Info.xlsx"

@st.cache_data
def load_data() -> pd.DataFrame:
    df = pd.read_excel(DATA_PATH, sheet_name="Sheet1")
    df.columns = [
        "Company Name", "Business Overview", "Sector",
        "Stage", "Location", "Invest in GPs", "LP in Secondary Funds",
    ]
    df = df.dropna(subset=["Company Name"])
    df["Company Name"] = df["Company Name"].astype(str).str.strip()
    for col in ["Business Overview", "Sector", "Stage", "Location",
                "Invest in GPs", "LP in Secondary Funds"]:
        df[col] = df[col].fillna("Not disclosed").astype(str).str.strip()
    df["City"] = df["Location"].apply(lambda x: x.split(",")[0].split(";")[0].strip())
    return df


def badge(text: str, css_extra: str = "") -> str:
    return f'<span class="{css_extra}">{text}</span>'


def yes_no_badge(text: str) -> str:
    lower = text.lower()
    if lower.startswith("yes"):
        return badge(text, "badge-yes")
    elif lower.startswith("no"):
        return badge(text, "badge-no")
    return badge(text)


def render_company_card(row: pd.Series) -> None:
    overview = row["Business Overview"]
    if len(overview) > 320:
        overview = overview[:320] + "…"
    sector_trunc = row["Sector"][:55] + ("…" if len(row["Sector"]) > 55 else "")
    stage_trunc  = row["Stage"][:45]  + ("…" if len(row["Stage"])  > 45 else "")
    loc_trunc    = row["Location"][:40]
    gp_badge = yes_no_badge(row["Invest in GPs"][:30])
    lp_badge = yes_no_badge(row["LP in Secondary Funds"][:30])
    meta_html = (
        f'<span>{sector_trunc}</span>'
        f'<span>{stage_trunc}</span>'
        f'<span>📍 {loc_trunc}</span>'
        f'<span>GP: {gp_badge}</span>'
        f'<span>LP: {lp_badge}</span>'
    )
    st.markdown(f"""
    <div class="company-card">
        <div class="company-name">🏢 {row['Company Name']}</div>
        <div class="company-meta">{meta_html}</div>
        <div class="company-overview">{overview}</div>
    </div>
    """, unsafe_allow_html=True)


CHART_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter", color="rgba(255,255,255,0.7)", size=12),
    margin=dict(l=10, r=10, t=40, b=10),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
)
PALETTE_D = px.colors.qualitative.Pastel


def sidebar_filters(df: pd.DataFrame):
    st.sidebar.markdown("## 🔍 Filters")
    query = st.sidebar.text_input("Search companies", placeholder="Keyword…")

    raw_sectors = df["Sector"].unique().tolist()
    sector_tokens: set = set()
    for s in raw_sectors:
        for tok in re.split(r",\s*|;\s*", s):
            t = tok.strip()
            if t and t.lower() not in ("not publicly disclosed", "not disclosed", "varies"):
                sector_tokens.add(t.title())
    sector_list = sorted(sector_tokens)
    selected_sectors = st.sidebar.multiselect("Sector", sector_list)

    stages_raw = sorted(df["Stage"].unique().tolist())
    selected_stages = st.sidebar.multiselect("Investment Stage", stages_raw)

    cities = sorted(df["City"].unique().tolist())
    selected_cities = st.sidebar.multiselect("City / Region", cities)

    st.sidebar.markdown("---")
    st.sidebar.markdown("**Invest in GPs?**")
    invest_gp = st.sidebar.radio(
        "Invest in GPs", ["All", "Yes", "No"], horizontal=True, label_visibility="collapsed"
    )
    st.sidebar.markdown("**LP in Secondary Funds?**")
    lp_secondary = st.sidebar.radio(
        "LP Secondary", ["All", "Yes", "No"], horizontal=True, label_visibility="collapsed"
    )
    return query, selected_sectors, selected_stages, selected_cities, invest_gp, lp_secondary


def apply_filters(df, query, selected_sectors, selected_stages, selected_cities,
                  invest_gp, lp_secondary):
    mask = pd.Series(True, index=df.index)
    if query:
        q = query.lower()
        mask &= (
            df["Company Name"].str.lower().str.contains(q, na=False)
            | df["Business Overview"].str.lower().str.contains(q, na=False)
            | df["Sector"].str.lower().str.contains(q, na=False)
        )
    if selected_sectors:
        sec_mask = pd.Series(False, index=df.index)
        for s in selected_sectors:
            sec_mask |= df["Sector"].str.contains(re.escape(s), case=False, na=False)
        mask &= sec_mask
    if selected_stages:
        mask &= df["Stage"].isin(selected_stages)
    if selected_cities:
        mask &= df["City"].isin(selected_cities)
    if invest_gp != "All":
        mask &= df["Invest in GPs"].str.lower().str.startswith(invest_gp.lower())
    if lp_secondary != "All":
        mask &= df["LP in Secondary Funds"].str.lower().str.startswith(lp_secondary.lower())
    return df[mask].reset_index(drop=True)


def render_analytics(df: pd.DataFrame):
    st.markdown('<div class="section-header">Sector Distribution</div>', unsafe_allow_html=True)
    sector_tokens = []
    for s in df["Sector"]:
        for tok in re.split(r",\s*|;\s*", s):
            t = tok.strip()
            if t and t.lower() not in ("not publicly disclosed", "not disclosed"):
                sector_tokens.append(t.title())
    sc = pd.Series(sector_tokens).value_counts().head(18).reset_index()
    sc.columns = ["Sector", "Count"]
    fig_sector = px.bar(
        sc, x="Count", y="Sector", orientation="h",
        color="Count", color_continuous_scale="Purples",
        title="Top 18 Sector Keywords",
    )
    fig_sector.update_layout(**CHART_LAYOUT, coloraxis_showscale=False, height=480)
    fig_sector.update_traces(marker_line_width=0)
    st.plotly_chart(fig_sector, width="stretch")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="section-header">Investment Stage Breakdown</div>', unsafe_allow_html=True)
        stage_counts = df["Stage"].value_counts().reset_index()
        stage_counts.columns = ["Stage", "Count"]
        stage_counts["Stage Short"] = stage_counts["Stage"].apply(
            lambda x: x[:38] + "…" if len(x) > 38 else x
        )
        fig_stage = px.pie(
            stage_counts.head(12), values="Count", names="Stage Short",
            color_discrete_sequence=PALETTE_D, hole=0.45,
        )
        fig_stage.update_layout(**CHART_LAYOUT, height=380)
        fig_stage.update_traces(textposition="inside", textinfo="percent+label",
                                 insidetextfont=dict(size=9))
        st.plotly_chart(fig_stage, width="stretch")

    with col2:
        st.markdown('<div class="section-header">Geographic Distribution (Top Cities)</div>', unsafe_allow_html=True)
        city_counts = df["City"].value_counts().head(15).reset_index()
        city_counts.columns = ["City", "Count"]
        fig_city = px.bar(
            city_counts, x="City", y="Count",
            color="Count", color_continuous_scale="Teal",
        )
        fig_city.update_layout(**CHART_LAYOUT, coloraxis_showscale=False, height=380,
                                xaxis_tickangle=-35)
        fig_city.update_traces(marker_line_width=0)
        st.plotly_chart(fig_city, width="stretch")

    st.markdown('<div class="section-header">GP & LP Secondary Activity</div>', unsafe_allow_html=True)
    col3, col4 = st.columns(2)

    def flag_bar(series, title):
        def simplify(val):
            v = val.lower()
            if v.startswith("yes"): return "Yes"
            elif v.startswith("no"): return "No"
            else: return "Not disclosed"
        counts = series.apply(simplify).value_counts().reset_index()
        counts.columns = ["Status", "Count"]
        color_map = {"Yes": "#34d399", "No": "#f87171", "Not disclosed": "#94a3b8"}
        fig = px.bar(counts, x="Status", y="Count",
                     color="Status", color_discrete_map=color_map, title=title)
        fig.update_layout(**CHART_LAYOUT, showlegend=False, height=300)
        fig.update_traces(marker_line_width=0)
        return fig

    with col3:
        st.plotly_chart(flag_bar(df["Invest in GPs"], "Invest in GPs"), width="stretch")
    with col4:
        st.plotly_chart(flag_bar(df["LP in Secondary Funds"], "LP in Secondary Funds"), width="stretch")


def main():
    df = load_data()
    query, sel_sectors, sel_stages, sel_cities, invest_gp, lp_secondary = sidebar_filters(df)
    filtered = apply_filters(df, query, sel_sectors, sel_stages, sel_cities, invest_gp, lp_secondary)

    st.markdown(
        '<div class="hero-title">🏦 CVC Analytics Dashboard</div>'
        '<div class="hero-sub">Corporate Venture Capital — Company Intelligence & Search</div>',
        unsafe_allow_html=True,
    )

    total = len(df)
    shown = len(filtered)
    yes_gp = (df["Invest in GPs"].str.lower().str.startswith("yes")).sum()
    yes_lp = (df["LP in Secondary Funds"].str.lower().str.startswith("yes")).sum()
    unique_cities = df["City"].nunique()

    k1, k2, k3, k4, k5 = st.columns(5)
    for col, val, lbl in [
        (k1, total, "Total CVCs"),
        (k2, shown, "Matching Results"),
        (k3, yes_gp, "Invest in GPs"),
        (k4, yes_lp, "LP in Secondary"),
        (k5, unique_cities, "Unique Locations"),
    ]:
        col.markdown(
            f'<div class="metric-card"><div class="metric-value">{val}</div>'
            f'<div class="metric-label">{lbl}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    tab_search, tab_analytics, tab_table = st.tabs(["🔎 Search Results", "📊 Analytics", "📋 Data Table"])

    with tab_search:
        if shown == 0:
            st.info("No companies match the current filters. Try adjusting the sidebar.")
        else:
            st.markdown(
                f'<div class="section-header">Showing {shown} of {total} CVC companies</div>',
                unsafe_allow_html=True,
            )
            PAGE_SIZE = 20
            total_pages = max(1, (shown + PAGE_SIZE - 1) // PAGE_SIZE)
            page = st.number_input(
                f"Page (1–{total_pages})", min_value=1, max_value=total_pages, value=1, step=1,
            )
            start = (page - 1) * PAGE_SIZE
            for _, row in filtered.iloc[start: start + PAGE_SIZE].iterrows():
                render_company_card(row)

    with tab_analytics:
        render_analytics(filtered if shown < total else df)

    with tab_table:
        st.markdown('<div class="section-header">Full Data Table</div>', unsafe_allow_html=True)
        display_cols = [
            "Company Name", "Sector", "Stage", "Location",
            "Invest in GPs", "LP in Secondary Funds",
        ]
        st.dataframe(filtered[display_cols], width="stretch", hide_index=True, height=520)
        csv = filtered.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download CSV", data=csv, file_name="cvc_filtered.csv", mime="text/csv",
        )


if __name__ == "__main__":
    main()
