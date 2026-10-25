from __future__ import annotations

from pathlib import Path
import io
import math
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -------------------------------------------------------------------
# Excelerate Opportunity Intelligence | Streamlit dashboard
# Designed for descriptive analysis, with explicit data-quality guardrails.
# -------------------------------------------------------------------

st.set_page_config(
    page_title="Excelerate | Opportunity Intelligence",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#152B3C"
INK = "#223746"
MUTED = "#667784"
TEAL = "#137C78"
TEAL_LIGHT = "#E5F3F1"
BLUE = "#477A9E"
GOLD = "#C58A35"
RED = "#B75B55"
BORDER = "#E1E8EC"
CANVAS = "#F5F7F8"
WHITE = "#FFFFFF"

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
      html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; color: {INK}; }}
      .stApp {{ background: {CANVAS}; }}
      [data-testid="stHeader"] {{ background: rgba(245,247,248,0.94); }}
      [data-testid="stSidebar"] {{ background: #EDF2F4; border-right: 1px solid {BORDER}; }}
      [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {{ color: {NAVY}; }}
      .block-container {{ padding-top: 1.65rem; padding-bottom: 2.5rem; max-width: 1500px; }}
      h1, h2, h3, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3,
      [data-testid="stMarkdownContainer"] h1,
      [data-testid="stMarkdownContainer"] h2,
      [data-testid="stMarkdownContainer"] h3 {{
        font-family: 'Manrope', sans-serif;
        letter-spacing: -0.035em;
        color: {NAVY} !important;
      }}
      h1 {{ font-size: 2.05rem !important; font-weight: 800 !important; }}
      h2 {{ font-size: 1.28rem !important; font-weight: 750 !important; }}
      h3 {{ font-size: 1.03rem !important; font-weight: 700 !important; }}
      p, li {{ color: {INK}; }}
      .eyebrow {{ color: {TEAL}; text-transform: uppercase; letter-spacing: .13em; font-size: .72rem; font-weight: 700; margin-bottom: .45rem; }}
      .hero {{ padding: 1.45rem 1.65rem; background: {WHITE}; border: 1px solid {BORDER}; border-radius: 16px; margin-bottom: 1.1rem; }}
      .hero-title {{ font-family: 'Manrope', sans-serif; color: {NAVY}; font-size: 1.9rem; line-height: 1.15; font-weight: 800; letter-spacing: -.04em; margin: 0 0 .45rem 0; }}
      .hero-sub {{ color: {MUTED} !important; opacity: 1 !important; font-size: .96rem; margin: 0; max-width: 900px; }}
      .section-head {{ display:flex; align-items:center; justify-content:space-between; margin: .4rem 0 .25rem 0; }}
      .section-note {{ color: {MUTED} !important; opacity: 1 !important; font-size: .82rem; margin-top: -.2rem; margin-bottom: .7rem; }}
      .metric-card {{ background: {WHITE}; border: 1px solid {BORDER}; border-radius: 14px; padding: 1rem 1.1rem; min-height: 116px; }}
      .metric-label {{ color: {MUTED}; font-size: .78rem; font-weight: 600; margin-bottom: .45rem; }}
      .metric-value {{ color: {NAVY}; font-family: 'Manrope', sans-serif; font-size: 1.8rem; font-weight: 800; line-height: 1.1; }}
      .metric-foot {{ color: {MUTED}; font-size: .74rem; margin-top: .45rem; }}
      .insight {{ border-left: 3px solid {TEAL}; background: {TEAL_LIGHT}; padding: .85rem 1rem; border-radius: 0 10px 10px 0; color: {INK}; font-size: .9rem; }}
      .warning {{ border-left: 3px solid {GOLD}; background: #FBF5EA; padding: .85rem 1rem; border-radius: 0 10px 10px 0; color: {INK}; font-size: .88rem; }}
      .small-muted {{ color: {MUTED}; font-size: .78rem; }}
      div[data-testid="stTabs"] button {{ font-weight: 650; }}
      div[data-testid="stDataFrame"] {{ border: 1px solid {BORDER}; border-radius: 10px; overflow: hidden; }}
      .stButton > button {{ border-radius: 9px; border: 1px solid {BORDER}; font-weight: 650; }}
      hr {{ border-color: {BORDER}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

PLOT_LAYOUT = dict(
    paper_bgcolor=WHITE,
    plot_bgcolor=WHITE,
    font=dict(family="DM Sans, sans-serif", color=INK, size=12),
    margin=dict(l=12, r=18, t=24, b=12),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    hoverlabel=dict(bgcolor=WHITE, bordercolor=BORDER, font_color=INK),
)
COLOR_SEQUENCE = [TEAL, BLUE, GOLD, "#7A8C9A", "#86B7B3", "#A8B7C1", RED]


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes | None, filename: str | None) -> pd.DataFrame:
    """Read an uploaded workbook or the local default workbook."""
    if file_bytes is not None:
        source = io.BytesIO(file_bytes)
    else:
        candidates = [
            Path(__file__).with_name("excelerate_cleaned.xlsx"),
            Path(__file__).with_name("excelerate_for_eda.xlsx"),
            Path.cwd() / "excelerate_cleaned.xlsx",
            Path.cwd() / "excelerate_for_eda.xlsx",
        ]
        source = next((p for p in candidates if p.exists()), None)
        if source is None:
            raise FileNotFoundError(
                "No default Excel file found. Place excelerate_cleaned.xlsx beside app.py "
                "or upload an Excel workbook in the sidebar."
            )
    frame = pd.read_excel(source, sheet_name=0)
    frame.columns = frame.columns.astype(str).str.strip()
    # The original workbook used an unnamed column "0" for fee.
    if "fee" not in frame.columns and "0" in frame.columns:
        frame = frame.rename(columns={"0": "fee"})
    return frame


def boolify(series: pd.Series) -> pd.Series:
    """Normalize booleans represented as native booleans, numbers, or strings."""
    true_values = {"true", "1", "yes", "y", "t"}
    false_values = {"false", "0", "no", "n", "f"}
    def convert(value):
        if pd.isna(value):
            return pd.NA
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)) and value in (0, 1):
            return bool(value)
        val = str(value).strip().lower()
        if val in true_values:
            return True
        if val in false_values:
            return False
        return pd.NA
    return series.map(convert).astype("boolean")


def prepare_data(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    for col in ["created_at", "modified_at", "last_date_to_apply"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", utc=True).dt.tz_convert(None)
    for col in ["fee", "microscholarship", "duration_hours"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "is_auto_approve" in df.columns:
        df["is_auto_approve"] = boolify(df["is_auto_approve"])
    if "is_archived" in df.columns:
        df["is_archived"] = boolify(df["is_archived"])

    for col in ["category", "currency_type", "location", "name", "role", "role_responsibility"]:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()
            df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "null": pd.NA})

    if "category" not in df.columns:
        df["category"] = pd.NA
    df["category_display"] = df["category"].fillna("Unknown")

    if "location" not in df.columns:
        df["location"] = pd.NA
    df["location_display"] = df["location"].fillna("Unknown")

    if "fee" not in df.columns:
        df["fee"] = pd.NA
    df["fee_status"] = "Review"
    df.loc[df["fee"].isna(), "fee_status"] = "Missing"
    df.loc[df["fee"].eq(0), "fee_status"] = "Free (zero fee)"
    df.loc[df["fee"].gt(0), "fee_status"] = "Paid (positive fee)"
    if "opportunity_id" not in df.columns:
        df["opportunity_id"] = pd.Series(range(1, len(df) + 1), index=df.index).astype(str)
    df["opportunity_id"] = df["opportunity_id"].astype("string")
    return df


def distinct_count(frame: pd.DataFrame) -> int:
    if "opportunity_id" in frame.columns:
        return int(frame["opportunity_id"].nunique(dropna=True))
    return int(len(frame))


def fmt_int(value: int | float) -> str:
    if pd.isna(value):
        return "—"
    return f"{value:,.0f}"


def fmt_pct(value: float) -> str:
    return "—" if pd.isna(value) else f"{value:.1%}"


def metric_card(label: str, value: str, foot: str) -> None:
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div><div class="metric-foot">{foot}</div></div>',
        unsafe_allow_html=True,
    )


def style_fig(fig, height=340):
    fig.update_layout(**PLOT_LAYOUT, height=height)
    fig.update_xaxes(showgrid=False, zeroline=False, linecolor=BORDER, tickfont=dict(color=MUTED))
    fig.update_yaxes(gridcolor="#EEF2F4", zeroline=False, linecolor=BORDER, tickfont=dict(color=MUTED))
    return fig


def count_by(frame, col, label="Opportunities"):
    result = (frame.groupby(col, dropna=False)["opportunity_id"].nunique()
              .sort_values(ascending=False).rename(label).reset_index())
    return result


# ----------------------------- Sidebar ---------------------------------
st.sidebar.markdown('<div class="eyebrow">EXCELERATE ANALYTICS</div>', unsafe_allow_html=True)
st.sidebar.title("Dashboard controls")
st.sidebar.caption("Filter the opportunity portfolio. All charts and KPIs respond to these selections.")

uploaded = st.sidebar.file_uploader("Use a different Excel file", type=["xlsx", "xls"])
if uploaded is not None:
    raw_df = load_excel(uploaded.getvalue(), uploaded.name)
    source_label = uploaded.name
else:
    try:
        raw_df = load_excel(None, None)
        source_label = "excelerate_cleaned.xlsx" if Path(__file__).with_name("excelerate_cleaned.xlsx").exists() else "local Excel workbook"
    except Exception as exc:
        st.error(str(exc))
        st.stop()

df_all = prepare_data(raw_df)
if df_all.empty:
    st.warning("The workbook contains no rows.")
    st.stop()

st.sidebar.markdown("---")
category_values = sorted(df_all["category_display"].dropna().unique().tolist())
selected_categories = st.sidebar.multiselect("Opportunity category", category_values, default=category_values)
date_min = df_all["created_at"].min() if "created_at" in df_all.columns else pd.NaT
date_max = df_all["created_at"].max() if "created_at" in df_all.columns else pd.NaT
date_enabled = "created_at" in df_all.columns and pd.notna(date_min) and pd.notna(date_max)
if date_enabled:
    start_date, end_date = st.sidebar.date_input(
        "Created date range",
        value=(date_min.date(), date_max.date()),
        min_value=date_min.date(),
        max_value=date_max.date(),
    )
else:
    start_date, end_date = None, None

currency_values = sorted(df_all["currency_type"].dropna().astype(str).unique().tolist()) if "currency_type" in df_all else []
selected_currencies = st.sidebar.multiselect("Currency", currency_values, default=currency_values) if currency_values else []
fee_options = ["Free (zero fee)", "Paid (positive fee)", "Missing", "Review"]
selected_fee_status = st.sidebar.multiselect("Fee status", fee_options, default=fee_options)
approval_options = ["Auto-approved", "Not auto-approved", "Unknown"]
selected_approval = st.sidebar.multiselect("Approval setting", approval_options, default=approval_options)
location_values = sorted(df_all["location_display"].dropna().unique().tolist())
selected_locations = st.sidebar.multiselect("Location / work mode", location_values, default=location_values)

if st.sidebar.button("Reset filters", use_container_width=True):
    st.rerun()

filtered = df_all[df_all["category_display"].isin(selected_categories)].copy()
if date_enabled and start_date and end_date:
    filtered = filtered[
        filtered["created_at"].notna()
        & (filtered["created_at"].dt.date >= start_date)
        & (filtered["created_at"].dt.date <= end_date)
    ]
if selected_currencies and "currency_type" in filtered.columns:
    filtered = filtered[filtered["currency_type"].fillna("Unknown").astype(str).isin(selected_currencies)]
if selected_fee_status:
    filtered = filtered[filtered["fee_status"].isin(selected_fee_status)]
if selected_locations:
    filtered = filtered[filtered["location_display"].isin(selected_locations)]
if "is_auto_approve" in filtered.columns:
    def approval_label(v):
        if pd.isna(v): return "Unknown"
        return "Auto-approved" if bool(v) else "Not auto-approved"
    filtered["approval_display"] = filtered["is_auto_approve"].map(approval_label)
    filtered = filtered[filtered["approval_display"].isin(selected_approval)]

# ----------------------------- Header ----------------------------------
st.markdown(
    """
    <div class="hero">
      <div class="eyebrow">OPPORTUNITY INTELLIGENCE · PORTFOLIO OVERVIEW</div>
      <div class="hero-title">Excelerate opportunity dashboard</div>
      <p class="hero-sub">A focused view of opportunity mix, publishing activity, fee structure and approval settings — built for clear decisions, not chart volume.</p>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption(f"Source: {source_label}  ·  Workbook rows: {len(df_all):,}  ·  Filtered records: {len(filtered):,}")

if filtered.empty:
    st.warning("No records match the selected filters. Broaden one or more filters in the sidebar.")
    st.stop()

total = distinct_count(filtered)
known_approval = filtered["is_auto_approve"].dropna() if "is_auto_approve" in filtered else pd.Series(dtype="boolean")
auto_n = int((known_approval == True).sum()) if len(known_approval) else 0
auto_rate = float((known_approval == True).mean()) if len(known_approval) else float("nan")
free_n = int(filtered["fee_status"].eq("Free (zero fee)").sum())
paid_n = int(filtered["fee_status"].eq("Paid (positive fee)").sum())
fee_known = int(filtered["fee_status"].isin(["Free (zero fee)", "Paid (positive fee)"]).sum())
free_share = free_n / fee_known if fee_known else float("nan")

k1, k2, k3, k4 = st.columns(4, gap="medium")
with k1: metric_card("TOTAL OPPORTUNITIES", fmt_int(total), "Distinct opportunity IDs in current view")
with k2: metric_card("FREE-FEE SHARE", fmt_pct(free_share), f"{free_n:,} zero-fee of {fee_known:,} known fees")
with k3: metric_card("PAID OPPORTUNITIES", fmt_int(paid_n), "Positive fee; amount not summed across currencies")
with k4: metric_card("AUTO-APPROVAL RATE", fmt_pct(auto_rate), f"{auto_n:,} auto-approved of {len(known_approval):,} known settings")

st.markdown("")

tab_overview, tab_fees, tab_quality, tab_data = st.tabs(
    ["Overview", "Fees & approval", "Data quality", "Explore records"]
)

with tab_overview:
    left, right = st.columns([1.05, 1.25], gap="large")
    with left:
        st.markdown("## Opportunity mix")
        st.markdown('<div class="section-note">Rank categories by distinct opportunity count.</div>', unsafe_allow_html=True)
        cat_counts = count_by(filtered, "category_display").sort_values("Opportunities", ascending=True)
        fig = px.bar(
            cat_counts, x="Opportunities", y="category_display", orientation="h",
            text="Opportunities", color_discrete_sequence=[TEAL],
            labels={"Opportunities": "Opportunities", "category_display": ""},
        )
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
        fig.update_layout(showlegend=False, yaxis=dict(categoryorder="total ascending"), xaxis_title=None, yaxis_title=None)
        st.plotly_chart(style_fig(fig, max(320, 30 * len(cat_counts) + 90)), use_container_width=True)
    with right:
        st.markdown("## Creation activity")
        st.markdown('<div class="section-note">Monthly creation counts; gaps may reflect source coverage rather than demand.</div>', unsafe_allow_html=True)
        if "created_at" in filtered and filtered["created_at"].notna().any():
            monthly = filtered.dropna(subset=["created_at"]).assign(
                month=lambda x: x["created_at"].dt.to_period("M").dt.to_timestamp()
            ).groupby("month")["opportunity_id"].nunique().reset_index(name="Opportunities")
            fig = px.line(monthly, x="month", y="Opportunities", markers=True, color_discrete_sequence=[TEAL],
                          labels={"month": "Creation month", "Opportunities": "Opportunities"})
            fig.update_traces(line=dict(width=3), marker=dict(size=6))
            fig.update_layout(yaxis_title=None, xaxis_title=None, hovermode="x unified")
            st.plotly_chart(style_fig(fig, 380), use_container_width=True)
        else:
            st.info("No valid creation dates are available in this filtered view.")

    st.markdown("## Where opportunities are listed")
    st.markdown('<div class="section-note">Top recorded location/work-mode values. Unknowns are kept visible rather than silently removed.</div>', unsafe_allow_html=True)
    loc_counts = count_by(filtered, "location_display").sort_values("Opportunities", ascending=False).head(12)
    loc_counts = loc_counts.sort_values("Opportunities", ascending=True)
    fig = px.bar(loc_counts, x="Opportunities", y="location_display", orientation="h",
                 text="Opportunities", color_discrete_sequence=[BLUE],
                 labels={"location_display": "", "Opportunities": "Opportunities"})
    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
    fig.update_layout(showlegend=False, xaxis_title=None, yaxis_title=None)
    st.plotly_chart(style_fig(fig, max(300, 28 * len(loc_counts) + 90)), use_container_width=True)

    if "category_display" in filtered.columns:
        top_cat = filtered["category_display"].value_counts().head(1)
        if not top_cat.empty:
            st.markdown(
                f'<div class="insight"><b>Reading the mix:</b> {top_cat.index[0]} is the most common category in the current selection. Use the date and category filters together to compare like-for-like periods.</div>',
                unsafe_allow_html=True,
            )

with tab_fees:
    st.markdown("## Fee structure")
    st.markdown('<div class="section-note">Zero, positive and missing fees are distinct states. Monetary amounts are not pooled across currencies.</div>', unsafe_allow_html=True)
    a, b = st.columns([0.85, 1.4], gap="large")
    with a:
        fee_counts = filtered["fee_status"].value_counts().reindex(
            ["Free (zero fee)", "Paid (positive fee)", "Missing", "Review"], fill_value=0
        ).rename_axis("Fee status").reset_index(name="Opportunities")
        fig = px.bar(fee_counts, x="Opportunities", y="Fee status", orientation="h",
                     color="Fee status", color_discrete_map={
                         "Free (zero fee)": TEAL, "Paid (positive fee)": BLUE,
                         "Missing": GOLD, "Review": RED
                     }, text="Opportunities", labels={"Opportunities": "Opportunities"})
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
        fig.update_layout(showlegend=False, xaxis_title=None, yaxis_title=None, yaxis=dict(categoryorder="total ascending"))
        st.plotly_chart(style_fig(fig, 300), use_container_width=True)
    with b:
        mix = (filtered.groupby(["category_display", "fee_status"])["opportunity_id"].nunique()
               .reset_index(name="Opportunities"))
        fig = px.bar(mix, x="category_display", y="Opportunities", color="fee_status",
                     barmode="relative", color_discrete_map={
                         "Free (zero fee)": TEAL, "Paid (positive fee)": BLUE,
                         "Missing": GOLD, "Review": RED
                     }, labels={"category_display": "Category", "Opportunities": "Opportunities", "fee_status": "Fee status"})
        fig.update_layout(xaxis_title=None, yaxis_title=None, legend_title_text="Fee status")
        st.plotly_chart(style_fig(fig, 340), use_container_width=True)

    st.markdown("## Approval settings by category")
    if "is_auto_approve" in filtered:
        approval_df = filtered[filtered["is_auto_approve"].notna()].copy()
        if not approval_df.empty:
            approval_df["auto_num"] = approval_df["is_auto_approve"].astype(int)
            approval = approval_df.groupby("category_display").agg(
                Known_Settings=("opportunity_id", "nunique"),
                Auto_Approved=("auto_num", "sum"),
            ).reset_index()
            approval["Auto-approval rate"] = approval["Auto_Approved"] / approval["Known_Settings"]
            approval = approval.sort_values("Auto-approval rate", ascending=True)
            fig = px.bar(approval, x="Auto-approval rate", y="category_display", orientation="h",
                         color_discrete_sequence=[TEAL], text="Auto-approval rate",
                         hover_data=["Known_Settings", "Auto_Approved"],
                         labels={"category_display": "Category", "Auto-approval rate": "Auto-approval rate"})
            fig.update_traces(texttemplate="%{x:.1%}", textposition="outside", cliponaxis=False)
            fig.update_layout(xaxis_tickformat=".0%", xaxis_title=None, yaxis_title=None, showlegend=False)
            st.plotly_chart(style_fig(fig, max(320, 30 * len(approval) + 90)), use_container_width=True)
            st.caption("Rate denominator includes only records with a known approval setting. Hover for the known-setting count.")
        else:
            st.info("No known approval settings in this selection.")
    else:
        st.info("This workbook does not include an approval-setting column.")

    st.markdown('<div class="warning"><b>Interpretation guardrail:</b> fee amount averages and totals are intentionally omitted. The dataset contains USD, INR and EUR, and EDA found an extreme fee value that should be validated before monetary comparisons.</div>', unsafe_allow_html=True)

with tab_quality:
    st.markdown("## Data coverage")
    st.markdown('<div class="section-note">Missingness is measured on the filtered rows. Low-coverage fields should not be treated as representative of the full portfolio.</div>', unsafe_allow_html=True)
    quality_cols = [c for c in ["category", "currency_type", "fee", "is_auto_approve", "is_archived", "role", "role_responsibility", "location", "created_at", "last_date_to_apply", "duration_hours"] if c in filtered.columns]
    quality = pd.DataFrame({
        "Field": quality_cols,
        "Missing records": [int(filtered[c].isna().sum()) for c in quality_cols],
        "Missing share": [float(filtered[c].isna().mean()) if len(filtered) else 0 for c in quality_cols],
    }).sort_values("Missing share", ascending=False)
    qleft, qright = st.columns([1.1, 1], gap="large")
    with qleft:
        fig = px.bar(quality.sort_values("Missing share", ascending=True), x="Missing share", y="Field",
                     orientation="h", color_discrete_sequence=[GOLD], text="Missing share",
                     labels={"Missing share": "Missing share", "Field": ""})
        fig.update_traces(texttemplate="%{x:.1%}", textposition="outside", cliponaxis=False)
        fig.update_layout(xaxis_tickformat=".0%", xaxis_title=None, yaxis_title=None, showlegend=False)
        st.plotly_chart(style_fig(fig, max(330, 30 * len(quality) + 70)), use_container_width=True)
    with qright:
        st.markdown("### Coverage table")
        display_quality = quality.copy()
        display_quality["Missing share"] = display_quality["Missing share"].map(lambda x: f"{x:.1%}")
        st.dataframe(display_quality, hide_index=True, use_container_width=True, height=420)

    st.markdown("## Validation checks")
    checks = []
    if "opportunity_id" in filtered:
        dup_ids = int(filtered["opportunity_id"].duplicated().sum())
        checks.append({"Check": "Repeated opportunity IDs", "Result": fmt_int(dup_ids), "Interpretation": "Review IDs before using row counts as opportunity counts."})
    if "fee" in filtered:
        fee_max = filtered["fee"].max()
        checks.append({"Check": "Largest recorded fee", "Result": "—" if pd.isna(fee_max) else f"{fee_max:,.0f}", "Interpretation": "Validate extreme values; do not sum fees across currencies."})
        neg_fees = int(filtered["fee"].lt(0).sum())
        checks.append({"Check": "Negative fees", "Result": fmt_int(neg_fees), "Interpretation": "Unexpected negative values should be investigated."})
    if "last_date_to_apply" in filtered and "created_at" in filtered:
        bad_deadline = int((filtered["last_date_to_apply"] < filtered["created_at"]).sum())
        checks.append({"Check": "Deadline before creation", "Result": fmt_int(bad_deadline), "Interpretation": "Validate date anomalies before deadline reporting."})
    if "duration_hours" in filtered:
        dur_max = filtered["duration_hours"].max()
        checks.append({"Check": "Largest duration (hours)", "Result": "—" if pd.isna(dur_max) else f"{dur_max:,.0f}", "Interpretation": "Confirm units and outliers before using duration as a performance metric."})
    st.dataframe(pd.DataFrame(checks), hide_index=True, use_container_width=True)

with tab_data:
    st.markdown("## Explore the filtered records")
    st.markdown('<div class="section-note">Use this table for spot checks and export the currently filtered view for further analysis.</div>', unsafe_allow_html=True)
    search_text = st.text_input("Search opportunity name or ID", placeholder="Type a keyword or opportunity ID")
    view = filtered.copy()
    if search_text.strip():
        search = search_text.strip().casefold()
        name_series = view["name"].astype("string").fillna("") if "name" in view else pd.Series("", index=view.index)
        id_series = view["opportunity_id"].astype("string").fillna("")
        view = view[name_series.str.casefold().str.contains(search, regex=False) | id_series.str.casefold().str.contains(search, regex=False)]
    preferred = [c for c in [
        "opportunity_id", "name", "category_display", "currency_type", "fee", "fee_status",
        "is_auto_approve", "location_display", "created_at", "last_date_to_apply", "duration_hours"
    ] if c in view.columns]
    st.dataframe(view[preferred], hide_index=True, use_container_width=True, height=470)
    csv_data = view.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Download filtered records (CSV)", data=csv_data,
                       file_name="excelerate_filtered_opportunities.csv", mime="text/csv")
    st.caption(f"Showing {len(view):,} records after filters and search.")

st.markdown("---")
st.markdown(
    f'<div class="small-muted">EXCELERATE OPPORTUNITY INTELLIGENCE &nbsp;·&nbsp; Descriptive analytics only &nbsp;·&nbsp; Source workbook: {source_label}. '
    'Counts use distinct opportunity IDs where available. Unknown and missing values remain visible; monetary totals are intentionally excluded.</div>',
    unsafe_allow_html=True,
)
