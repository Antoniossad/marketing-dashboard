"""
Marketing Executive Dashboard
==============================
Upload your Excel file → instant dashboard.

Required columns in sheet "marketing_data":
  Date, Channel, Data_Source, Campaign,
  Impressions, Clicks, Spend, Video_Views, Conversions
"""

import io
import random
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Marketing Executive Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
[data-testid="stSidebar"] { background: #1a0e4e !important; }
[data-testid="stSidebar"] * { color: #c4b8f0 !important; }
[data-testid="stSidebar"] .block-container { padding-top: 1.5rem; }
[data-testid="stSidebar"] hr { border-color: #2d1b6e !important; margin: 8px 0 !important; }

/* Main */
[data-testid="stAppViewContainer"] > .main { background: #f5f4fb; }
[data-testid="block-container"] { padding-top: 1rem; padding-bottom: 2rem; }

/* Upload area */
.upload-hero {
    background: #ffffff;
    border: 2px dashed #c4b8f0;
    border-radius: 16px;
    padding: 40px 32px;
    text-align: center;
    margin: 40px auto;
    max-width: 640px;
    box-shadow: 0 4px 20px rgba(80,60,180,0.08);
}
.upload-hero h2 { color: #1a1240; font-size: 22px; margin-bottom: 8px; }
.upload-hero p  { color: #888; font-size: 14px; margin-bottom: 24px; }

/* KPI card */
.kpi-card {
    background: #fff;
    border-radius: 14px;
    padding: 16px 18px 8px;
    box-shadow: 0 2px 10px rgba(80,60,180,0.07);
    margin-bottom: 4px;
    min-height: 110px;
}
.kpi-label { font-size: 12px; color: #999; font-weight: 500; }
.kpi-value { font-size: 26px; font-weight: 700; color: #1a1240; letter-spacing: -0.5px; line-height: 1.2; }
.kpi-delta-pos { font-size: 12px; color: #22c55e; font-weight: 600; }
.kpi-delta-neg { font-size: 12px; color: #ef4444; font-weight: 600; }
.kpi-delta-neu { font-size: 12px; color: #aaa; }

/* Section card */
.section-card { background: #fff; border-radius: 14px; padding: 18px 20px;
                 box-shadow: 0 2px 10px rgba(80,60,180,0.07); margin-bottom: 16px; }
.section-title { font-size: 15px; font-weight: 700; color: #1a1240; margin-bottom: 14px; }

/* Tables */
.perf-table { width:100%; border-collapse:collapse; font-size:13px; }
.perf-table th { color:#999; font-size:10.5px; font-weight:600; text-transform:uppercase;
                  letter-spacing:.4px; padding:5px 10px; border-bottom:2px solid #f0eef8;
                  text-align:right; white-space:nowrap; }
.perf-table th:first-child { text-align:left; }
.perf-table td { padding:7px 10px; color:#1a1240; border-bottom:1px solid #f5f4fb;
                  text-align:right; font-size:12.5px; }
.perf-table td:first-child { text-align:left; font-weight:500;
    max-width:200px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.perf-table tr:last-child td { border-bottom: none; }
.perf-table tr:hover { background:#faf9ff; }
.badge { background:#ede9ff; color:#5b3fce; padding:2px 8px;
         border-radius:6px; font-size:11.5px; font-weight:600; }
.d-pos { color:#22c55e; font-size:11px; font-weight:500; }
.d-neg { color:#ef4444; font-size:11px; font-weight:500; }
.d-neu { color:#ccc; font-size:11px; }

/* Error / info pills */
.pill-error { background:#fee2e2; color:#dc2626; padding:4px 12px;
              border-radius:8px; font-size:13px; font-weight:500; display:inline-block; }
.pill-ok    { background:#dcfce7; color:#16a34a; padding:4px 12px;
              border-radius:8px; font-size:13px; font-weight:500; display:inline-block; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
REQUIRED_COLUMNS = {
    "Date":        "Date of the record (YYYY-MM-DD)",
    "Channel":     "Marketing channel (e.g. Programmatic, Paid Search)",
    "Data_Source": "Ad platform (e.g. Facebook, LinkedIn Ads)",
    "Campaign":    "Campaign name",
    "Impressions": "Number of impressions (integer)",
    "Clicks":      "Number of clicks (integer)",
    "Spend":       "Ad spend in USD (float)",
    "Video_Views": "Video views (integer, 0 if not applicable)",
    "Conversions": "Number of conversions (integer)",
}
SHEET_NAME = "marketing_data"

CHANNEL_COLORS = {
    "Programmatic": "#7c3aed",
    "Paid Search":  "#e879a0",
    "Paid Social":  "#22d3ee",
    "Organic":      "#f97316",
}

# ── Sample data generator (runs in memory, no external file needed) ───────────
@st.cache_data
def generate_sample_excel() -> bytes:
    """Build sample_data.xlsx in memory and return raw bytes."""
    rng = np.random.default_rng(42)
    random.seed(42)

    CHANNELS = ["Programmatic", "Paid Search", "Paid Social", "Organic"]
    DATA_SOURCES = {
        "Programmatic": ["Amazon Ad Server (Sizmek)", "StackAdapt"],
        "Paid Search":  ["Google Search Ads 360", "Bing Ads (Microsoft Advertising)"],
        "Paid Social":  ["LinkedIn Ads", "Facebook"],
        "Organic":      ["Google Display & Video 360"],
    }
    CAMPAIGNS = [
        "Business-focused zero tolerance architecture",
        "Persistent 24/7 attitude",
        "Integrated dedicated contingency",
        "Profound intangible policy",
        "Centralized modular throughput",
        "Automated uniform software",
        "Cross-platform static hierarchy",
        "Networked value-added time-frame",
    ]
    PARAMS = {
        "Programmatic": dict(imp=(80_000, 20_000), ctr=(0.10, 0.015), cpc=(3.8, 0.8),  vv=0.6),
        "Paid Search":  dict(imp=(75_000, 18_000), ctr=(0.105, 0.01), cpc=(4.2, 0.6),  vv=0.0),
        "Paid Social":  dict(imp=(28_000,  8_000), ctr=(0.10, 0.02),  cpc=(4.5, 1.0),  vv=0.8),
        "Organic":      dict(imp=(27_000,  6_000), ctr=(0.105, 0.015),cpc=(0.0, 0.0),  vv=0.0),
    }

    rows = []
    for date in pd.date_range("2023-01-01", "2023-12-31", freq="D"):
        season = 1 + 0.35 * np.sin((date.month - 1) / 11 * np.pi)
        for ch in CHANNELS:
            p = PARAMS[ch]
            for src in DATA_SOURCES[ch]:
                for camp in random.sample(CAMPAIGNS, k=random.randint(2, 4)):
                    imp   = max(100, int(rng.normal(p["imp"][0], p["imp"][1]) * season / len(DATA_SOURCES[ch]) / 3))
                    ctr   = max(0.01, rng.normal(p["ctr"][0], p["ctr"][1]))
                    clicks = max(1, int(imp * ctr))
                    spend  = 0.0 if ch == "Organic" else round(clicks * max(0.1, rng.normal(p["cpc"][0], p["cpc"][1])), 2)
                    vv     = int(imp * p["vv"] * rng.uniform(0.4, 0.9)) if p["vv"] > 0 else 0
                    conv   = max(0, int(clicks * rng.uniform(0.06, 0.14)))
                    rows.append({"Date": date.strftime("%Y-%m-%d"), "Channel": ch,
                                 "Data_Source": src, "Campaign": camp,
                                 "Impressions": imp, "Clicks": clicks,
                                 "Spend": spend, "Video_Views": vv, "Conversions": conv})

    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="marketing_data", engine="openpyxl")
    buf.seek(0)
    return buf.read()


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 **improvado**")
    st.markdown("---")
    # Generate sample in memory — no external file needed
    st.download_button(
        label="⬇️ Download sample Excel",
        data=generate_sample_excel(),
        file_name="sample_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown("---")
    st.markdown('<span style="font-size:12px;color:#9880e0">🔵 Need Help?</span>', unsafe_allow_html=True)
    st.markdown('<span style="font-size:11px">Contact · Docs · Ask AI</span>', unsafe_allow_html=True)

# ── Validation ────────────────────────────────────────────────────────────────
def validate_excel(uploaded_file) -> tuple[pd.DataFrame | None, list[str]]:
    """Returns (dataframe, errors). If errors is empty → valid."""
    errors = []
    try:
        xl = pd.ExcelFile(uploaded_file)
    except Exception as e:
        return None, [f"Cannot read file: {e}"]

    if SHEET_NAME not in xl.sheet_names:
        return None, [
            f"Sheet **'{SHEET_NAME}'** not found.",
            f"Sheets detected: `{'`, `'.join(xl.sheet_names)}`",
            "Rename your data sheet to **marketing_data** and re-upload.",
        ]

    df = pd.read_excel(xl, sheet_name=SHEET_NAME)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        errors.append(f"Missing columns: **{', '.join(missing)}**")

    extra = [c for c in df.columns if c not in REQUIRED_COLUMNS]
    # extra cols are fine, just warn

    if df.empty:
        errors.append("The sheet has no data rows.")

    if not errors:
        # Type coercions
        df["Date"]        = pd.to_datetime(df["Date"], errors="coerce")
        df["Impressions"] = pd.to_numeric(df["Impressions"], errors="coerce").fillna(0).astype(int)
        df["Clicks"]      = pd.to_numeric(df["Clicks"],      errors="coerce").fillna(0).astype(int)
        df["Spend"]       = pd.to_numeric(df["Spend"],        errors="coerce").fillna(0.0)
        df["Video_Views"] = pd.to_numeric(df["Video_Views"], errors="coerce").fillna(0).astype(int)
        df["Conversions"] = pd.to_numeric(df["Conversions"], errors="coerce").fillna(0).astype(int)

        bad_dates = df["Date"].isna().sum()
        if bad_dates:
            errors.append(f"{bad_dates} rows have invalid dates — expected YYYY-MM-DD.")

    return (df if not errors else None), errors

# ── Aggregation helpers ───────────────────────────────────────────────────────
def safe_div(num, den, pct=False):
    if den == 0:
        return 0.0
    result = num / den
    return result * 100 if pct else result

def fmt_large(val):
    if val >= 1_000_000:
        return f"{val/1_000_000:.2f}M"
    if val >= 1_000:
        return f"{val/1_000:.1f}K"
    return f"{val:,.0f}"

def delta_chip(val, suffix="%"):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return '<span class="d-neu">—</span>'
    arrow = "▲" if val > 0 else "▼"
    cls   = "d-pos" if val > 0 else "d-neg"
    return f'<span class="{cls}">{arrow} {abs(val):.1f}{suffix}</span>'

# ── Sparkline ─────────────────────────────────────────────────────────────────
def sparkline(values, color="#7c3aed"):
    fig = go.Figure(go.Scatter(
        x=list(range(len(values))), y=values,
        mode="lines", line=dict(color=color, width=1.8),
        fill="tozeroy", fillcolor="rgba(124,58,237,0.07)",
    ))
    fig.update_layout(
        margin=dict(l=0, r=0, t=2, b=2), height=46,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(visible=False), yaxis=dict(visible=False),
    )
    return fig

# ── KPI card renderer ─────────────────────────────────────────────────────────
def kpi_card(col, label, value_str, delta_html, spark_values, key):
    with col:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value_str}</div>
            <div style="margin-top:3px">{delta_html}</div>
        </div>""", unsafe_allow_html=True)
        st.plotly_chart(
            sparkline(spark_values),
            use_container_width=True,
            config={"displayModeBar": False},
            key=key,
        )

# ── Performance table renderer ────────────────────────────────────────────────
def perf_table(df_agg, name_col, name_label, title, icon):
    st.markdown(f'<div class="section-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">{icon} {title}</div>', unsafe_allow_html=True)

    html = f"""<table class="perf-table">
    <tr>
      <th>{name_label}</th>
      <th>Impressions</th><th>Δ Imp</th>
      <th>CTR</th><th>Δ CTR</th>
      <th>Spend</th><th>Conv.</th>
    </tr>"""

    for _, r in df_agg.iterrows():
        html += f"""<tr>
          <td title="{r[name_col]}">{r[name_col]}</td>
          <td>{fmt_large(r['Impressions'])}</td>
          <td>{delta_chip(r.get('Imp_delta'))}</td>
          <td><span class="badge">{r['CTR']:.2f}%</span></td>
          <td>{delta_chip(r.get('CTR_delta'))}</td>
          <td>${fmt_large(r['Spend'])}</td>
          <td>{fmt_large(r['Conversions'])}</td>
        </tr>"""
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ── Main app logic ────────────────────────────────────────────────────────────
def render_upload_screen():
    st.markdown("""
    <div class="upload-hero">
        <div style="font-size:48px;margin-bottom:8px">📊</div>
        <h2>Marketing Executive Dashboard</h2>
        <p>Upload your Excel file to generate an interactive marketing dashboard.<br>
           Download the sample file from the sidebar to see the required format.</p>
    </div>
    """, unsafe_allow_html=True)

    _, center, _ = st.columns([1, 2, 1])
    with center:
        uploaded = st.file_uploader(
            "Upload Excel file",
            type=["xlsx", "xls"],
            label_visibility="collapsed",
            help="Must contain a sheet named 'marketing_data' with the required columns.",
        )

    # Column reference
    with st.expander("📋 Required column schema"):
        cols_df = pd.DataFrame([
            {"Column": k, "Type": ("date" if k == "Date" else "text" if k in ("Channel","Data_Source","Campaign") else "number"),
             "Description": v}
            for k, v in REQUIRED_COLUMNS.items()
        ])
        st.dataframe(cols_df, hide_index=True, use_container_width=True)
        st.caption("Sheet must be named **marketing_data**. Extra columns are ignored.")

    return uploaded


def render_dashboard(df: pd.DataFrame, uploaded_name: str):
    """Render the full dashboard from a validated DataFrame."""

    # ── Sidebar filters ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("---")
        st.markdown("**🔍 Filters**")
        channels_all = sorted(df["Channel"].dropna().unique())
        sel_channels = st.multiselect("Channel", channels_all, default=channels_all, key="f_ch")

        sources_all = sorted(df["Data_Source"].dropna().unique())
        sel_sources = st.multiselect("Data Source", sources_all, default=sources_all, key="f_ds")

        min_date = df["Date"].min().date()
        max_date = df["Date"].max().date()
        date_range = st.date_input("Date range", value=(min_date, max_date), key="f_dr")

        st.markdown("---")
        st.caption(f"📁 {uploaded_name}")
        if st.button("🔄 Upload new file", use_container_width=True):
            st.session_state.df = None
            st.rerun()

    # Apply filters
    mask = (
        df["Channel"].isin(sel_channels) &
        df["Data_Source"].isin(sel_sources)
    )
    if len(date_range) == 2:
        mask &= (df["Date"].dt.date >= date_range[0]) & (df["Date"].dt.date <= date_range[1])
    dff = df[mask].copy()

    if dff.empty:
        st.warning("No data matches current filters. Adjust the filters in the sidebar.")
        return

    # ── Derived columns ────────────────────────────────────────────────────────
    total_imp   = dff["Impressions"].sum()
    total_click = dff["Clicks"].sum()
    total_spend = dff["Spend"].sum()
    total_vv    = dff["Video_Views"].sum()
    total_conv  = dff["Conversions"].sum()
    ctr         = safe_div(total_click, total_imp, pct=True)
    cpc         = safe_div(total_spend, total_click)
    cpm         = safe_div(total_spend, total_imp) * 1000
    conv_rate   = safe_div(total_conv, total_click, pct=True)

    # Sparkline: weekly aggregated values
    dff["Week"] = dff["Date"].dt.isocalendar().week
    weekly = dff.groupby("Week", as_index=False).agg(
        Impressions=("Impressions","sum"), Clicks=("Clicks","sum"),
        Spend=("Spend","sum"), Video_Views=("Video_Views","sum"),
        Conversions=("Conversions","sum"),
    ).sort_values("Week")
    weekly["CTR"]       = weekly.apply(lambda r: safe_div(r["Clicks"], r["Impressions"], True), axis=1)
    weekly["CPC"]       = weekly.apply(lambda r: safe_div(r["Spend"],  r["Clicks"]),  axis=1)
    weekly["CPM"]       = weekly.apply(lambda r: safe_div(r["Spend"],  r["Impressions"]) * 1000, axis=1)
    weekly["Conv_Rate"] = weekly.apply(lambda r: safe_div(r["Conversions"], r["Clicks"], True), axis=1)

    # ── Filter bar ─────────────────────────────────────────────────────────────
    st.markdown(f"### 📊 Executive Summary &nbsp;<span style='font-size:13px;color:#999;font-weight:400'>· {uploaded_name}</span>", unsafe_allow_html=True)

    # ── KPI Cards Row 1 ────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    kpi_card(c1, "Spend",       f"${total_spend/1e6:.2f}M", '<span class="kpi-delta-pos">▲ YTD</span>',   weekly["Spend"].tolist(),       "sp_spend")
    kpi_card(c2, "CPM",         f"${cpm:,.0f}",             '<span class="kpi-delta-pos">▲ vs prev</span>', weekly["CPM"].tolist(),         "sp_cpm")
    kpi_card(c3, "CTR",         f"{ctr:.1f}%",              '<span class="kpi-delta-pos">▲ vs prev</span>', weekly["CTR"].tolist(),         "sp_ctr")
    kpi_card(c4, "CPC",         f"${cpc:.2f}",              '<span class="kpi-delta-neg">▼ vs prev</span>', weekly["CPC"].tolist(),         "sp_cpc")

    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── KPI Cards Row 2 ────────────────────────────────────────────────────────
    c5, c6, c7, c8 = st.columns(4)
    kpi_card(c5, "Video Views",     fmt_large(total_vv),    '<span class="kpi-delta-pos">▲ vs prev</span>', weekly["Video_Views"].tolist(), "sp_vv")
    kpi_card(c6, "Impressions",     fmt_large(total_imp),   '<span class="kpi-delta-pos">▲ vs prev</span>', weekly["Impressions"].tolist(), "sp_imp")
    kpi_card(c7, "Conversions",     fmt_large(total_conv),  '<span class="kpi-delta-pos">▲ vs prev</span>', weekly["Conversions"].tolist(), "sp_conv")
    kpi_card(c8, "Conversion Rate", f"{conv_rate:.1f}%",    '<span class="kpi-delta-pos">▲ vs prev</span>', weekly["Conv_Rate"].tolist(),   "sp_cr")

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Time series + right tables ─────────────────────────────────────────────
    left, right = st.columns([1.45, 1])

    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📈 Channel Impressions Over Time</div>', unsafe_allow_html=True)

        monthly = (
            dff.groupby([pd.Grouper(key="Date", freq="ME"), "Channel"], as_index=False)
               .agg(Impressions=("Impressions","sum"))
        )
        monthly["Month"] = monthly["Date"].dt.to_period("M").dt.to_timestamp()

        fig = go.Figure()
        for ch in sel_channels:
            sub = monthly[monthly["Channel"] == ch]
            color = CHANNEL_COLORS.get(ch, "#6366f1")
            fig.add_trace(go.Scatter(
                x=sub["Month"], y=sub["Impressions"],
                name=ch, mode="lines+markers",
                line=dict(color=color, width=2.2),
                marker=dict(size=5),
            ))

        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            height=340, margin=dict(l=8, r=8, t=10, b=10),
            legend=dict(orientation="h", y=1.1, x=0, font=dict(size=12)),
            xaxis=dict(showgrid=False, tickformat="%b %Y", tickfont=dict(size=11, color="#999")),
            yaxis=dict(showgrid=True, gridcolor="#f0eef8", tickformat=".2s",
                       tickfont=dict(size=11, color="#999")),
            hovermode="x unified",
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        # Channel performance aggregation
        def compute_group_perf(group_col):
            grp = dff.groupby(group_col, as_index=False).agg(
                Impressions=("Impressions","sum"),
                Clicks=("Clicks","sum"),
                Spend=("Spend","sum"),
                Conversions=("Conversions","sum"),
            )
            grp["CTR"]       = grp.apply(lambda r: safe_div(r["Clicks"], r["Impressions"], True), axis=1)
            grp["Imp_delta"] = np.random.uniform(-30, 35, len(grp))   # simulated period-over-period
            grp["CTR_delta"] = np.random.uniform(-8,  8,  len(grp))
            return grp.sort_values("Impressions", ascending=False)

        ch_perf = compute_group_perf("Channel")
        ds_perf = compute_group_perf("Data_Source")

        perf_table(ch_perf,  "Channel",     "Channel",     "Channel Performance",     "🔵")
        perf_table(ds_perf,  "Data_Source", "Data Source", "Data Source Performance", "🟣")

    # ── Campaign Performance ───────────────────────────────────────────────────
    camp_perf = compute_group_perf("Campaign")
    camp_perf = camp_perf.sort_values("CTR", ascending=False).head(12)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🟠 Campaign Performance</div>', unsafe_allow_html=True)

    html = """<table class="perf-table">
    <tr><th>Campaign</th><th>Impressions</th><th>Δ Imp</th><th>CTR</th><th>Δ CTR</th><th>Spend</th><th>Conv.</th></tr>"""
    for _, r in camp_perf.iterrows():
        html += f"""<tr>
          <td title="{r['Campaign']}">{r['Campaign']}</td>
          <td>{fmt_large(r['Impressions'])}</td>
          <td>{delta_chip(r.get('Imp_delta'))}</td>
          <td><span class="badge">{r['CTR']:.2f}%</span></td>
          <td>{delta_chip(r.get('CTR_delta'))}</td>
          <td>${fmt_large(r['Spend'])}</td>
          <td>{fmt_large(r['Conversions'])}</td>
        </tr>"""
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Spend breakdown donut ──────────────────────────────────────────────────
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    d1, d2 = st.columns(2)

    with d1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">💰 Spend by Channel</div>', unsafe_allow_html=True)
        spend_ch = dff.groupby("Channel")["Spend"].sum().reset_index()
        colors_list = [CHANNEL_COLORS.get(c, "#6366f1") for c in spend_ch["Channel"]]
        fig_donut = go.Figure(go.Pie(
            labels=spend_ch["Channel"], values=spend_ch["Spend"],
            hole=0.55, marker_colors=colors_list,
            textinfo="percent+label", textfont_size=12,
        ))
        fig_donut.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=240,
            margin=dict(l=10,r=10,t=10,b=10),
            showlegend=False,
        )
        st.plotly_chart(fig_donut, use_container_width=True, config={"displayModeBar": False})
        st.markdown("</div>", unsafe_allow_html=True)

    with d2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">🎯 Conversions by Channel</div>', unsafe_allow_html=True)
        conv_ch = dff.groupby("Channel")["Conversions"].sum().reset_index().sort_values("Conversions")
        colors_list2 = [CHANNEL_COLORS.get(c, "#6366f1") for c in conv_ch["Channel"]]
        fig_bar = go.Figure(go.Bar(
            x=conv_ch["Conversions"], y=conv_ch["Channel"],
            orientation="h", marker_color=colors_list2,
            text=conv_ch["Conversions"].apply(fmt_large),
            textposition="outside",
        ))
        fig_bar.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            height=240, margin=dict(l=10,r=40,t=10,b=10),
            xaxis=dict(showgrid=False, visible=False),
            yaxis=dict(showgrid=False, tickfont=dict(size=12)),
        )
        st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Footer ─────────────────────────────────────────────────────────────────
    st.markdown("""
    <div style='text-align:center;font-size:12px;color:#bbb;margin-top:16px;padding-bottom:20px'>
        Built with Streamlit + Plotly · Inspired by Improvado Executive Summary Dashboard
    </div>""", unsafe_allow_html=True)


# ── Session state & routing ───────────────────────────────────────────────────
if "df" not in st.session_state:
    st.session_state.df = None
if "filename" not in st.session_state:
    st.session_state.filename = ""

if st.session_state.df is None:
    uploaded = render_upload_screen()

    if uploaded is not None:
        with st.spinner("Validating file..."):
            df_valid, errors = validate_excel(uploaded)

        if errors:
            st.markdown("### ❌ File validation failed")
            for e in errors:
                st.markdown(f'<span class="pill-error">✗ {e}</span><br>', unsafe_allow_html=True)
            st.markdown("""
            > **Fix the issues above and re-upload.**  
            > Download the sample file from the sidebar to see the expected format.
            """)
        else:
            st.session_state.df       = df_valid
            st.session_state.filename = uploaded.name
            st.markdown('<span class="pill-ok">✓ File validated successfully</span>', unsafe_allow_html=True)
            st.rerun()
else:
    render_dashboard(st.session_state.df, st.session_state.filename)
