# app.py

import os
from io import BytesIO

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots


# ============================================================
# STREAMLIT PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Supply Chain Simulation Performance Review",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# CONFIGURATION
# Place your Excel file in the SAME GitHub root folder as app.py.
# Streamlit Cloud will read it directly from the repository.
# ============================================================
PREFERRED_EXCEL_FILE = "FinanceReport (5)(1).xlsx"
REQUIRED_ROUNDS = [f"Round {i}" for i in range(7)]
REQUIRED_METRICS = [
    "ROI",
    "Realized Revenue",
    "Gross Margin",
    "Operating Profit",
    "Total Investment",
    "Total Penalties and Bonuses",
    "Overhead Costs",
    "Stock Costs",
    "Handling Costs",
    "Administration Costs",
    "Distribution Costs",
]
INDIRECT_COST_METRICS = [
    "Overhead Costs",
    "Stock Costs",
    "Handling Costs",
    "Administration Costs",
    "Distribution Costs",
]


# ============================================================
# EXECUTIVE CSS STYLING
# ============================================================
st.markdown(
    """
    <style>
        .main {
            background-color: #f7f9fc;
        }

        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

        .dashboard-title {
            font-size: 2.4rem;
            font-weight: 800;
            color: #102033;
            letter-spacing: -0.03em;
            margin-bottom: 0.2rem;
        }

        .dashboard-subtitle {
            font-size: 1rem;
            color: #5d6d7e;
            margin-bottom: 1.4rem;
        }

        .kpi-card {
            background: linear-gradient(135deg, #ffffff 0%, #f3f7fb 100%);
            border: 1px solid #e5edf5;
            border-radius: 18px;
            padding: 22px 20px;
            box-shadow: 0 8px 24px rgba(16, 32, 51, 0.06);
            min-height: 155px;
        }

        .kpi-title {
            color: #6c7a89;
            font-size: 0.85rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
        }

        .kpi-value {
            color: #102033;
            font-size: 2rem;
            font-weight: 850;
            margin-top: 0.45rem;
        }

        .kpi-delta-positive {
            color: #087f5b;
            font-size: 0.95rem;
            font-weight: 700;
            margin-top: 0.35rem;
        }

        .kpi-delta-neutral {
            color: #42647f;
            font-size: 0.95rem;
            font-weight: 700;
            margin-top: 0.35rem;
        }

        .section-header {
            color: #102033;
            font-size: 1.45rem;
            font-weight: 800;
            margin-top: 1rem;
            margin-bottom: 0.25rem;
        }

        .section-caption {
            color: #637488;
            font-size: 0.95rem;
            margin-bottom: 0.8rem;
        }

        .insight-box {
            background: #eef6ff;
            border-left: 6px solid #1f77b4;
            border-radius: 14px;
            padding: 18px 20px;
            margin-top: 0.3rem;
            margin-bottom: 1.7rem;
            box-shadow: 0 5px 16px rgba(31, 119, 180, 0.08);
        }

        .insight-title {
            font-weight: 850;
            color: #12395b;
            font-size: 1.02rem;
            margin-bottom: 0.35rem;
        }

        .insight-text {
            color: #27445e;
            font-size: 0.96rem;
            line-height: 1.55;
        }

        .small-note {
            color: #68798a;
            font-size: 0.88rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DATA LOADING HELPERS
# ============================================================
def find_excel_file() -> str:
    """Find the Excel file in the same folder as app.py."""
    if os.path.exists(PREFERRED_EXCEL_FILE):
        return PREFERRED_EXCEL_FILE

    excel_files = [
        file_name
        for file_name in os.listdir(".")
        if file_name.lower().endswith((".xlsx", ".xls")) and not file_name.startswith("~$")
    ]

    if excel_files:
        return excel_files[0]

    return ""


@st.cache_data(show_spinner=False)
def load_financial_data(excel_path: str) -> pd.DataFrame:
    """
    Load the financial simulation data from Excel.
    The app searches all sheets and keeps only Metric + Round 0 to Round 6.
    """
    if not excel_path:
        raise FileNotFoundError(
            "No Excel file found. Add your .xlsx file to the same GitHub root folder as app.py."
        )

    all_sheets = pd.read_excel(excel_path, sheet_name=None)
    selected_df = None
    selected_sheet = None

    for sheet_name, sheet_df in all_sheets.items():
        sheet_df = sheet_df.copy()
        sheet_df.columns = [str(col).strip() for col in sheet_df.columns]

        if "Metric" in sheet_df.columns and all(col in sheet_df.columns for col in REQUIRED_ROUNDS):
            selected_df = sheet_df[["Metric"] + REQUIRED_ROUNDS].copy()
            selected_sheet = sheet_name
            break

    if selected_df is None:
        raise ValueError(
            "Could not find a sheet with columns: Metric, Round 0, Round 1, ..., Round 6. "
            "Please keep your Excel data in that structure."
        )

    selected_df["Metric"] = selected_df["Metric"].astype(str).str.strip()
    selected_df = selected_df[selected_df["Metric"].isin(REQUIRED_METRICS)].copy()

    for col in REQUIRED_ROUNDS:
        selected_df[col] = pd.to_numeric(selected_df[col], errors="coerce")

    missing_metrics = [metric for metric in REQUIRED_METRICS if metric not in selected_df["Metric"].tolist()]
    if missing_metrics:
        raise ValueError(
            "The Excel file is missing required metrics: " + ", ".join(missing_metrics)
        )

    selected_df["Metric"] = pd.Categorical(
        selected_df["Metric"],
        categories=REQUIRED_METRICS,
        ordered=True,
    )
    selected_df = selected_df.sort_values("Metric").reset_index(drop=True)
    selected_df.attrs["source_file"] = excel_path
    selected_df.attrs["source_sheet"] = selected_sheet

    return selected_df


# ============================================================
# FORMAT HELPERS
# ============================================================
def currency_short(value: float) -> str:
    sign = "-" if value < 0 else ""
    value_abs = abs(value)

    if value_abs >= 1_000_000:
        return f"{sign}${value_abs / 1_000_000:.2f}M"
    if value_abs >= 1_000:
        return f"{sign}${value_abs / 1_000:.2f}K"
    return f"{sign}${value_abs:,.2f}"


def percent_fmt(value: float) -> str:
    return f"{value * 100:.2f}%"


def insight_box(text: str) -> None:
    st.markdown(
        f"""
        <div class="insight-box">
            <div class="insight-title">Supply Chain Head Actionable Insight</div>
            <div class="insight-text">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def chart_layout(fig: go.Figure, title=None, height=520) -> go.Figure:
    fig.update_layout(
        title=title,
        height=height,
        template="plotly_white",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        margin=dict(l=30, r=30, t=70, b=40),
        font=dict(family="Arial", size=13, color="#102033"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hoverlabel=dict(bgcolor="white", font_size=13, font_family="Arial"),
    )
    return fig


def dataframe_to_excel_bytes(display_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_df.to_excel(writer, index=False, sheet_name="Round 0 to Round 6")
    return output.getvalue()


# ============================================================
# LOAD DATA FROM EXCEL FILE IN GITHUB ROOT
# ============================================================
try:
    excel_file = find_excel_file()
    df = load_financial_data(excel_file)
except Exception as exc:
    st.error("Unable to load the Excel dataset.")
    st.info(
        "Upload your Excel workbook to the same GitHub root folder as `app.py`. "
        "The workbook must contain columns named `Metric`, `Round 0`, `Round 1`, ..., `Round 6`."
    )
    st.exception(exc)
    st.stop()

round_cols = REQUIRED_ROUNDS
wide = df.set_index("Metric")[round_cols].T.reset_index()
wide = wide.rename(columns={"index": "Round"})
wide["Round Number"] = wide["Round"].str.extract(r"(\d+)").astype(int)
wide["ROI (%)"] = wide["ROI"] * 100
wide["ROI Bubble Size"] = wide["ROI (%)"].abs() + 1.5


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 📦 Dashboard Controls")
    st.markdown(
        """
        This executive dashboard reviews the financial and operational progression
        of the supply chain simulation from **Round 0 to Round 6**.
        """
    )

    selected_view = st.radio(
        "Navigate",
        [
            "Executive Overview",
            "Turnaround & ROI",
            "SLA Reliability",
            "Cost Efficiency",
            "Margin & Capital Efficiency",
            "Data Table",
        ],
    )

    st.divider()

    st.markdown("### Excel Source")
    st.caption(f"File: `{df.attrs.get('source_file', excel_file)}`")
    st.caption(f"Sheet: `{df.attrs.get('source_sheet', 'Detected automatically')}`")

    st.divider()

    st.markdown("### Final Round Snapshot")
    st.metric("Final ROI", percent_fmt(wide.loc[6, "ROI"]))
    st.metric("Operating Profit", currency_short(wide.loc[6, "Operating Profit"]))
    st.metric("Realized Revenue", currency_short(wide.loc[6, "Realized Revenue"]))
    st.metric("Capital Investment", currency_short(wide.loc[6, "Total Investment"]))

    st.divider()

    st.markdown(
        """
        <span class="small-note">
        The app reads your Excel workbook directly from the GitHub repository root.
        No CSV conversion or manual upload inside the dashboard is required.
        </span>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# HEADER
# ============================================================
st.markdown(
    """
    <div class="dashboard-title">Supply Chain Simulation Performance Review</div>
    <div class="dashboard-subtitle">
        Executive-ready Streamlit Cloud dashboard tracking profitability, SLA recovery, cost discipline,
        margin conversion, and capital efficiency from Round 0 to Round 6.
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# KPI CARDS
# ============================================================
r0 = wide.loc[0]
r6 = wide.loc[6]

roi_delta = r6["ROI"] - r0["ROI"]
revenue_delta = r6["Realized Revenue"] - r0["Realized Revenue"]
profit_delta = r6["Operating Profit"] - r0["Operating Profit"]
investment_delta = r6["Total Investment"] - r0["Total Investment"]

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Final ROI</div>
            <div class="kpi-value">{percent_fmt(r6["ROI"])}</div>
            <div class="kpi-delta-positive">
                ▲ Up from {percent_fmt(r0["ROI"])} | +{roi_delta * 100:.2f} pts
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with kpi2:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Realized Revenue</div>
            <div class="kpi-value">{currency_short(r6["Realized Revenue"])}</div>
            <div class="kpi-delta-positive">
                ▲ {currency_short(revenue_delta)} vs Round 0
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with kpi3:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Operating Profit</div>
            <div class="kpi-value">{currency_short(r6["Operating Profit"])}</div>
            <div class="kpi-delta-positive">
                ▲ Up from loss of {currency_short(r0["Operating Profit"])} | +{currency_short(profit_delta)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with kpi4:
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Capital Investment</div>
            <div class="kpi-value">{currency_short(r6["Total Investment"])}</div>
            <div class="kpi-delta-neutral">
                ▼ {currency_short(abs(investment_delta))} less capital tied up
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)


# ============================================================
# CHART 1: MULTI-AXIS TURNAROUND JOURNEY
# ============================================================
def render_chart_1() -> None:
    st.markdown('<div class="section-header">1. The Multi-Axis Turnaround Journey</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">Operating profit recovery paired with ROI stabilization across simulation rounds.</div>',
        unsafe_allow_html=True,
    )

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=wide["Round"],
            y=wide["Operating Profit"],
            name="Operating Profit",
            marker=dict(
                color=wide["Operating Profit"],
                colorscale="RdYlGn",
                line=dict(color="rgba(16,32,51,0.2)", width=1),
            ),
            hovertemplate="<b>%{x}</b><br>Operating Profit: $%{y:,.0f}<extra></extra>",
        ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(
            x=wide["Round"],
            y=wide["ROI (%)"],
            name="ROI (%)",
            mode="lines+markers",
            line=dict(width=4, color="#1f77b4"),
            marker=dict(size=11, color="#1f77b4", line=dict(width=2, color="white")),
            hovertemplate="<b>%{x}</b><br>ROI: %{y:.2f}%<extra></extra>",
        ),
        secondary_y=True,
    )

    fig.add_hline(
        y=0,
        line_dash="dash",
        line_color="#6c757d",
        annotation_text="Break-even",
        annotation_position="bottom right",
    )

    fig.update_yaxes(title_text="Operating Profit ($)", secondary_y=False, tickprefix="$", separatethousands=True)
    fig.update_yaxes(title_text="ROI (%)", secondary_y=True, ticksuffix="%")
    fig.update_xaxes(title_text="Simulation Round")

    fig = chart_layout(fig, height=540)
    st.plotly_chart(fig, use_container_width=True)

    insight_box(
        f"""
        The simulation moved from a severe Round 0 operating loss of <b>{currency_short(r0['Operating Profit'])}</b>
        to a Round 6 profit of <b>{currency_short(r6['Operating Profit'])}</b>, while ROI improved from
        <b>{percent_fmt(r0['ROI'])}</b> to <b>{percent_fmt(r6['ROI'])}</b>. This indicates that supply chain
        interventions were not merely cosmetic cost cuts; they translated into sustained operating recovery.
        The leadership priority should now shift from emergency loss correction to repeatable profitability governance,
        using Round 4 to Round 6 as the emerging stabilization benchmark.
        """
    )


# ============================================================
# CHART 2: SLA & DELIVERY RELIABILITY PERFORMANCE
# ============================================================
def render_chart_2() -> None:
    st.markdown('<div class="section-header">2. SLA & Delivery Reliability Performance</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">Penalty-to-bonus conversion as a proxy for SLA reliability and OTIF discipline.</div>',
        unsafe_allow_html=True,
    )

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=wide["Round"],
            y=wide["Total Penalties and Bonuses"],
            mode="lines+markers",
            name="Penalties / Bonuses",
            fill="tozeroy",
            line=dict(width=4, color="#7b2cbf"),
            marker=dict(size=11, color="#7b2cbf", line=dict(width=2, color="white")),
            hovertemplate="<b>%{x}</b><br>Penalties / Bonuses: $%{y:,.0f}<extra></extra>",
        )
    )

    fig.add_hline(
        y=0,
        line_dash="dash",
        line_color="#6c757d",
        annotation_text="Penalty / Bonus Threshold",
        annotation_position="bottom right",
    )

    fig.update_yaxes(title_text="Total Penalties and Bonuses ($)", tickprefix="$", separatethousands=True)
    fig.update_xaxes(title_text="Simulation Round")

    fig = chart_layout(fig, height=520)
    st.plotly_chart(fig, use_container_width=True)

    insight_box(
        f"""
        Round 0 shows a major service failure signal with <b>{currency_short(r0['Total Penalties and Bonuses'])}</b>
        in penalties. From Round 1 onward, the business converts this into positive bonus territory, suggesting
        improved delivery cadence, stronger SLA adherence, and better OTIF execution. The action for the Supply Chain
        Head is to institutionalize the scheduling, allocation, and fulfillment rules that eliminated the penalty drag,
        while investigating why bonuses tapered from Round 2 to Round 6 despite profitability improving.
        """
    )


# ============================================================
# CHART 3: INDIRECT SUPPLY CHAIN EFFICIENCY MATRIX
# ============================================================
def render_chart_3() -> None:
    st.markdown('<div class="section-header">3. Indirect Supply Chain Efficiency Matrix</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">Interactive heatmap of indirect operating cost intensity by round.</div>',
        unsafe_allow_html=True,
    )

    heatmap_df = df[df["Metric"].isin(INDIRECT_COST_METRICS)].set_index("Metric")[round_cols]

    fig = go.Figure(
        data=go.Heatmap(
            z=heatmap_df.values,
            x=heatmap_df.columns,
            y=heatmap_df.index,
            colorscale="YlOrRd",
            colorbar=dict(title="Expense Intensity"),
            hovertemplate="<b>%{y}</b><br>%{x}: $%{z:,.0f}<extra></extra>",
        )
    )

    fig.update_xaxes(title_text="Simulation Round")
    fig.update_yaxes(title_text="Indirect Cost Category")

    fig = chart_layout(fig, height=520)
    st.plotly_chart(fig, use_container_width=True)

    stock_peak_round = wide.loc[wide["Stock Costs"].idxmax(), "Round"]
    stock_peak_value = wide["Stock Costs"].max()

    insight_box(
        f"""
        Most indirect cost categories remain relatively controlled, but <b>Stock Costs</b> show meaningful movement,
        rising from <b>{currency_short(r0['Stock Costs'])}</b> in Round 0 to a peak of
        <b>{currency_short(stock_peak_value)}</b> in <b>{stock_peak_round}</b> before moderating. This suggests that
        profitability improvement may have required a deliberate safety-stock or capacity-buffer strategy during the
        recovery phase. The next operational move is to test whether current inventory buffers are still needed or whether
        warehouse capacity and reorder policies can be tightened without reintroducing SLA penalties.
        """
    )


# ============================================================
# CHART 4: MARGIN OPTIMIZATION ANALYSIS
# ============================================================
def render_chart_4() -> None:
    st.markdown('<div class="section-header">4. Margin Optimization Analysis</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">Gross margin versus operating profit to evaluate bottom-line conversion efficiency.</div>',
        unsafe_allow_html=True,
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=wide["Round"],
            y=wide["Gross Margin"],
            name="Gross Margin",
            marker_color="#00a6a6",
            hovertemplate="<b>%{x}</b><br>Gross Margin: $%{y:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Bar(
            x=wide["Round"],
            y=wide["Operating Profit"],
            name="Operating Profit",
            marker_color="#f77f00",
            hovertemplate="<b>%{x}</b><br>Operating Profit: $%{y:,.0f}<extra></extra>",
        )
    )

    fig.add_hline(y=0, line_dash="dash", line_color="#6c757d")
    fig.update_layout(barmode="group")
    fig.update_yaxes(title_text="Financial Value ($)", tickprefix="$", separatethousands=True)
    fig.update_xaxes(title_text="Simulation Round")

    fig = chart_layout(fig, height=540)
    st.plotly_chart(fig, use_container_width=True)

    insight_box(
        f"""
        Gross margin improved sharply from <b>{currency_short(r0['Gross Margin'])}</b> in Round 0 to
        <b>{currency_short(r6['Gross Margin'])}</b> in Round 6, while operating profit moved from a loss into
        sustained positive territory. This demonstrates that production and commercial margin gains were increasingly
        protected through the operating cost structure instead of being consumed by penalties, stock inefficiencies,
        or administrative drag. The leadership focus should be to preserve this margin-to-profit conversion discipline
        while validating whether Round 6 revenue growth can be scaled without expanding indirect costs.
        """
    )


# ============================================================
# CHART 5: INVESTMENT RETURN EFFICIENCY
# ============================================================
def render_chart_5() -> None:
    st.markdown('<div class="section-header">5. Investment Return Efficiency</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">Capital tied up versus operating profit, with bubble size scaled by ROI magnitude.</div>',
        unsafe_allow_html=True,
    )

    fig = px.scatter(
        wide,
        x="Total Investment",
        y="Operating Profit",
        size="ROI Bubble Size",
        color="ROI (%)",
        text="Round",
        color_continuous_scale="Viridis",
        size_max=42,
        hover_data={
            "Total Investment": ":,.0f",
            "Operating Profit": ":,.0f",
            "ROI (%)": ":.2f",
            "ROI Bubble Size": False,
            "Round Number": False,
        },
    )

    fig.update_traces(textposition="top center", marker=dict(line=dict(width=1.5, color="white")))
    fig.update_xaxes(title_text="Total Investment ($)", tickprefix="$", separatethousands=True)
    fig.update_yaxes(title_text="Operating Profit ($)", tickprefix="$", separatethousands=True)

    fig.add_hline(
        y=0,
        line_dash="dash",
        line_color="#6c757d",
        annotation_text="Profit Break-even",
        annotation_position="bottom right",
    )

    fig = chart_layout(fig, height=540)
    st.plotly_chart(fig, use_container_width=True)

    insight_box(
        f"""
        The business reduced total investment from <b>{currency_short(r0['Total Investment'])}</b> in Round 0 to
        <b>{currency_short(r6['Total Investment'])}</b> in Round 6 while simultaneously improving operating profit by
        more than <b>{currency_short(profit_delta)}</b>. This is a strong capital-efficiency signal: less capital is tied
        up in the network, yet the operating system is generating materially better returns. The next executive question
        is whether the Round 4 to Round 6 investment band represents the optimal working-capital range, or whether further
        capital release would risk inventory availability and SLA performance.
        """
    )


# ============================================================
# VIEW ROUTING
# ============================================================
if selected_view == "Executive Overview":
    overview_col1, overview_col2 = st.columns(2)

    with overview_col1:
        render_chart_1()

    with overview_col2:
        render_chart_2()

    render_chart_3()
    render_chart_4()
    render_chart_5()

elif selected_view == "Turnaround & ROI":
    render_chart_1()

elif selected_view == "SLA Reliability":
    render_chart_2()

elif selected_view == "Cost Efficiency":
    render_chart_3()

elif selected_view == "Margin & Capital Efficiency":
    render_chart_4()
    render_chart_5()

elif selected_view == "Data Table":
    st.markdown('<div class="section-header">Underlying Excel Dataset</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-caption">Round 0 through Round 6 only, read directly from the Excel workbook in your GitHub repository.</div>',
        unsafe_allow_html=True,
    )

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        label="Download displayed data as Excel",
        data=dataframe_to_excel_bytes(df),
        file_name="supply_chain_simulation_round_0_to_6.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.markdown(
    """
    <span class="small-note">
    Built for Streamlit Cloud deployment. Keep <b>app.py</b>, <b>requirements.txt</b>, and your Excel workbook in the root of your GitHub repository.
    </span>
    """,
    unsafe_allow_html=True,
)
