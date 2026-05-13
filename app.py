"""
BTS Alarm Log Analysis Engine
Cleans, maps, and analyzes legacy network alarm logs cross-referenced with PKEY Master Mapping
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import re
import subprocess
import tempfile
import shutil
from io import BytesIO
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BTS Alarm Analysis Engine",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        color: white;
        padding: 20px 30px;
        border-radius: 12px;
        margin-bottom: 20px;
        border-left: 5px solid #e94560;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; }
    .main-header p  { margin: 4px 0 0; opacity: 0.75; font-size: 0.9rem; }

    .kpi-card {
        background: #1e293b;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 16px 20px;
        text-align: center;
    }
    .kpi-card .kpi-val { font-size: 2rem; font-weight: 700; color: #60a5fa; margin: 4px 0; }
    .kpi-card .kpi-label { font-size: 0.8rem; color: #94a3b8; text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-card .kpi-sub { font-size: 0.75rem; color: #64748b; margin-top: 2px; }

    .section-header {
        font-size: 1.1rem; font-weight: 600;
        color: #e2e8f0; margin: 24px 0 12px;
        border-bottom: 2px solid #e94560;
        padding-bottom: 6px;
    }
    .alert-box {
        background: #1e293b; border: 1px solid #475569;
        border-radius: 8px; padding: 12px 16px;
        margin: 8px 0;
    }
    .tag-2g  { background:#3b82f6; color:white; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600; }
    .tag-3g  { background:#10b981; color:white; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600; }
    .tag-4g  { background:#f59e0b; color:white; padding:2px 8px; border-radius:4px; font-size:0.75rem; font-weight:600; }
    .tag-crit { background:#ef4444; color:white; padding:2px 8px; border-radius:4px; font-size:0.75rem; }
    .tag-warn { background:#f97316; color:white; padding:2px 8px; border-radius:4px; font-size:0.75rem; }

    div[data-testid="stTabs"] button { font-weight: 600; }
    .stDataFrame { border-radius: 8px; }
    div[data-testid="metric-container"] { background:#1e293b; border-radius:8px; padding:12px; border:1px solid #334155; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def convert_xls_to_df(file_bytes: bytes, filename: str) -> pd.DataFrame | None:
    """Convert legacy .xls to DataFrame via LibreOffice."""
    tmpdir = tempfile.mkdtemp()
    try:
        xls_path = os.path.join(tmpdir, filename)
        with open(xls_path, "wb") as f:
            f.write(file_bytes)
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "xlsx", xls_path, "--outdir", tmpdir],
            capture_output=True, timeout=60
        )
        xlsx_path = xls_path.replace(".xls", ".xlsx")
        if not os.path.exists(xlsx_path):
            return None
        df = pd.read_excel(xlsx_path)
        return df
    except Exception as e:
        st.warning(f"Could not read {filename}: {e}")
        return None
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


@st.cache_data(show_spinner=False)
def load_pkey(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Load PKEY master mapping file."""
    buf = BytesIO(file_bytes)
    if filename.lower().endswith(".csv"):
        df = pd.read_csv(buf)
    else:
        df = pd.read_excel(buf)
    df.columns = [c.strip() for c in df.columns]
    return df


def parse_down_period(s: str) -> float:
    """Parse '0 days 5 hours 48 minutes' → total minutes (float)."""
    if pd.isna(s) or not str(s).strip():
        return np.nan
    s = str(s).lower().strip()
    days = re.search(r"(\d+)\s*day", s)
    hours = re.search(r"(\d+)\s*hour", s)
    mins = re.search(r"(\d+)\s*min", s)
    total = 0.0
    if days:  total += int(days.group(1)) * 1440
    if hours: total += int(hours.group(1)) * 60
    if mins:  total += int(mins.group(1))
    return total


def classify_duration(mins: float) -> str:
    if pd.isna(mins): return "Unknown"
    if mins < 60:     return "< 1 hr"
    if mins < 240:    return "1–4 hrs"
    if mins < 480:    return "4–8 hrs"
    if mins < 1440:   return "8–24 hrs"
    return "> 24 hrs"


def severity_color(mins: float) -> str:
    if pd.isna(mins): return "#64748b"
    if mins < 60:   return "#22c55e"
    if mins < 240:  return "#84cc16"
    if mins < 480:  return "#f59e0b"
    if mins < 1440: return "#ef4444"
    return "#7c3aed"


def process_logs(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """Concatenate and clean multiple alarm log DataFrames."""
    combined = pd.concat(dfs, ignore_index=True)
    combined.columns = [c.strip() for c in combined.columns]

    # Datetime parsing
    for col in ["bts_down_dt", "bts_up_dt"]:
        if col in combined.columns:
            combined[col] = pd.to_datetime(combined[col], errors="coerce")

    # Parse downPeriod string → numeric minutes
    if "downPeriod" in combined.columns:
        combined["down_minutes"] = combined["downPeriod"].apply(parse_down_period)
        combined["down_hours"] = (combined["down_minutes"] / 60).round(2)
        combined["duration_band"] = combined["down_minutes"].apply(classify_duration)
    
    # Derive log_date from file
    if "bts_down_dt" in combined.columns:
        combined["log_date"] = combined["bts_down_dt"].dt.date

    # Normalise text columns
    for col in ["bts_type", "vendor", "fault_type", "sdca_name", "ssa_name"]:
        if col in combined.columns:
            combined[col] = combined[col].astype(str).str.strip().str.title()

    # De-duplicate (same bts_id + down_dt from overlapping daily files)
    key_cols = [c for c in ["bts_id", "bts_down_dt", "bts_up_dt"] if c in combined.columns]
    if key_cols:
        combined = combined.drop_duplicates(subset=key_cols)

    return combined.reset_index(drop=True)


def merge_pkey(df_alarms: pd.DataFrame, df_pkey: pd.DataFrame) -> pd.DataFrame:
    """Cross-reference alarms with PKEY master to get incharge details."""
    pkey_cols = ["BTSIPID", "incharge", "JTO INCHARGE", "SITENAME", "LOCATION", "SDCA", "PKEY"]
    pkey_sub = df_pkey[[c for c in pkey_cols if c in df_pkey.columns]].copy()
    pkey_sub = pkey_sub.drop_duplicates(subset=["BTSIPID"])
    pkey_sub["BTSIPID"] = pkey_sub["BTSIPID"].astype(str).str.strip()
    df_alarms["bts_ip_id_clean"] = df_alarms["bts_ip_id"].astype(str).str.strip()
    merged = df_alarms.merge(pkey_sub, left_on="bts_ip_id_clean", right_on="BTSIPID", how="left")
    merged.drop(columns=["bts_ip_id_clean", "BTSIPID"], errors="ignore", inplace=True)
    return merged


PALETTE = px.colors.qualitative.Bold
DARK_BG  = "#0f172a"
CARD_BG  = "#1e293b"
GRID_COL = "#334155"

def dark_layout(fig, title="", height=400):
    fig.update_layout(
        title=dict(text=title, font=dict(color="#e2e8f0", size=14)),
        paper_bgcolor=DARK_BG, plot_bgcolor=CARD_BG,
        font=dict(color="#94a3b8"),
        height=height,
        legend=dict(bgcolor=CARD_BG, bordercolor=GRID_COL),
        margin=dict(l=40, r=20, t=50, b=40),
        xaxis=dict(gridcolor=GRID_COL, zerolinecolor=GRID_COL),
        yaxis=dict(gridcolor=GRID_COL, zerolinecolor=GRID_COL),
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📡 BTS Alarm Engine")
    st.markdown("---")
    st.markdown("### 1. PKEY Master Mapping")
    pkey_file = st.file_uploader("Upload PKEY (.xlsx / .csv)", type=["xlsx", "xls", "csv"], key="pkey")

    st.markdown("### 2. Alarm Log Files")
    st.caption("Upload 1–31 daily .xls files")
    alarm_files = st.file_uploader(
        "Upload Alarm Logs (.xls)", type=["xls", "xlsx"],
        accept_multiple_files=True, key="alarms"
    )

    st.markdown("---")
    st.markdown("### 3. Filters")
    bts_types_sel  = st.multiselect("BTS Type", ["2G", "3G", "4G"], default=["2G", "3G", "4G"])
    duration_bands = ["< 1 hr", "1–4 hrs", "4–8 hrs", "8–24 hrs", "> 24 hrs"]
    dur_sel        = st.multiselect("Duration Band", duration_bands, default=duration_bands)

    st.markdown("---")
    st.caption("BSNL TN Circle • Network Ops • v2.0")


# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>📡 BTS Alarm Log Analysis Engine</h1>
  <p>Legacy XLS processing • PKEY cross-referencing • Fault &amp; Duration intelligence • Incharge analytics</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# LANDING STATE
# ─────────────────────────────────────────────────────────────────────────────
if not pkey_file or not alarm_files:
    col1, col2 = st.columns(2)
    with col1:
        st.info("👈 **Step 1:** Upload the PKEY master mapping file (Excel or CSV).")
        st.markdown("""
        **PKEY file columns expected:**
        - `PKEY`, `BTSIPID`, `SSACODE`, `SSANAME`
        - `SDCANAME`, `SITENAME`, `LOCATION`, `SDCA`
        - `incharge`, `JTO INCHARGE`
        """)
    with col2:
        st.info("👈 **Step 2:** Upload one or more daily alarm log `.xls` files.")
        st.markdown("""
        **Alarm log columns expected:**
        - `bts_id`, `bts_name`, `bts_ip_id`, `bts_type`
        - `vendor`, `bts_down_dt`, `bts_up_dt`
        - `downPeriod`, `fault_type`, `sdca_name`
        """)
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
with st.spinner("🔄 Loading & processing files…"):

    # PKEY
    df_pkey = load_pkey(pkey_file.read(), pkey_file.name)

    # Alarm logs
    alarm_dfs = []
    failed_files = []
    progress = st.progress(0, text="Converting alarm files…")
    for i, f in enumerate(alarm_files):
        file_bytes = f.read()
        name = f.name
        if name.lower().endswith(".xls"):
            df_raw = convert_xls_to_df(file_bytes, name)
        else:
            df_raw = pd.read_excel(BytesIO(file_bytes))
        if df_raw is not None and len(df_raw) > 0:
            df_raw["_source_file"] = name
            alarm_dfs.append(df_raw)
        else:
            failed_files.append(name)
        progress.progress((i + 1) / len(alarm_files), text=f"Processing {name}…")
    progress.empty()

    if not alarm_dfs:
        st.error("❌ No valid alarm files could be loaded. Please check your files.")
        st.stop()

    # Process & merge
    df_all = process_logs(alarm_dfs)
    df_all = merge_pkey(df_all, df_pkey)

    # Apply sidebar filters
    if "bts_type" in df_all.columns and bts_types_sel:
        df = df_all[df_all["bts_type"].str.upper().isin([x.upper() for x in bts_types_sel])].copy()
    else:
        df = df_all.copy()
    if "duration_band" in df.columns and dur_sel:
        df = df[df["duration_band"].isin(dur_sel)].copy()

if failed_files:
    with st.expander(f"⚠️ {len(failed_files)} file(s) failed to load"):
        for f in failed_files:
            st.write(f"• {f}")

# ─────────────────────────────────────────────────────────────────────────────
# KPI ROW
# ─────────────────────────────────────────────────────────────────────────────
total_events    = len(df)
unique_bts      = df["bts_id"].nunique() if "bts_id" in df.columns else 0
total_down_hrs  = df["down_hours"].sum() if "down_hours" in df.columns else 0
avg_down_hrs    = df["down_hours"].mean() if "down_hours" in df.columns else 0
files_loaded    = len(alarm_dfs)
unmatched_pkey  = df["incharge"].isna().sum() if "incharge" in df.columns else 0

c1, c2, c3, c4, c5, c6 = st.columns(6)
def kpi(col, val, label, sub=""):
    col.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-val">{val}</div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

kpi(c1, f"{total_events:,}",  "Alarm Events",      f"{files_loaded} file(s) loaded")
kpi(c2, f"{unique_bts:,}",    "Unique BTS",         "Affected sites")
kpi(c3, f"{total_down_hrs:,.0f} h", "Total Down Hours",  "Cumulative outage")
kpi(c4, f"{avg_down_hrs:.1f} h",    "Avg Down / Event",  "Mean duration")
kpi(c5, f"{df['fault_type'].nunique() if 'fault_type' in df.columns else 0}", "Fault Types", "Distinct categories")
kpi(c6, f"{unmatched_pkey}",  "PKEY Mismatches",    "No incharge mapped")

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📊 Overview",
    "📡 BTS Type & Vendor",
    "⚠️ Fault Analysis",
    "⏱️ Duration Deep-Dive",
    "👤 Incharge Analytics",
    "📅 Daily Trend",
    "📋 Summary Report",
    "🗃️ Raw Data"
])

# ══════════════════════════════════════════════════════════════════
# TAB 0: OVERVIEW
# ══════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="section-header">Fleet Overview</div>', unsafe_allow_html=True)
    
    col_l, col_r = st.columns([3, 2])

    with col_l:
        # BTS type distribution
        if "bts_type" in df.columns:
            bts_dist = df.groupby("bts_type").agg(
                Events=("bts_id", "count"),
                Unique_BTS=("bts_id", "nunique"),
                Total_Down_Hrs=("down_hours", "sum"),
            ).reset_index()
            fig = px.bar(bts_dist, x="bts_type", y="Events",
                         color="bts_type", text="Events",
                         color_discrete_sequence=PALETTE,
                         title="Alarm Events by BTS Type")
            fig.update_traces(textposition="outside")
            dark_layout(fig, height=320)
            st.plotly_chart(fig, use_container_width=True)

    with col_r:
        if "vendor" in df.columns:
            vend_cnt = df["vendor"].value_counts().reset_index()
            vend_cnt.columns = ["vendor", "count"]
            fig2 = px.pie(vend_cnt, values="count", names="vendor",
                          title="Events by Vendor",
                          color_discrete_sequence=PALETTE, hole=0.4)
            dark_layout(fig2, height=320)
            fig2.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig2, use_container_width=True)

    # SDCA map
    if "sdca_name" in df.columns:
        st.markdown('<div class="section-header">SDCA-wise Alarm Distribution</div>', unsafe_allow_html=True)
        sdca_df = df.groupby("sdca_name").agg(
            Events=("bts_id","count"),
            Unique_BTS=("bts_id","nunique"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean")
        ).reset_index().sort_values("Total_Down_Hrs", ascending=False)
        sdca_df["Total_Down_Hrs"] = sdca_df["Total_Down_Hrs"].round(1)
        sdca_df["Avg_Down_Hrs"] = sdca_df["Avg_Down_Hrs"].round(1)
        fig3 = px.bar(sdca_df, x="sdca_name", y="Total_Down_Hrs",
                      color="Events", text="Events",
                      color_continuous_scale="Blues",
                      title="Total Down Hours per SDCA")
        fig3.update_traces(textposition="outside")
        dark_layout(fig3, height=350)
        st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# TAB 1: BTS TYPE & VENDOR
# ══════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="section-header">BTS Type × Vendor Analysis</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        if "bts_type" in df.columns and "vendor" in df.columns:
            cross = df.groupby(["bts_type", "vendor"]).agg(
                Events=("bts_id","count"),
                Avg_Down_Hrs=("down_hours","mean"),
                Total_Down_Hrs=("down_hours","sum")
            ).reset_index()
            fig = px.bar(cross, x="bts_type", y="Events", color="vendor",
                         barmode="group", title="Events by BTS Type & Vendor",
                         color_discrete_sequence=PALETTE)
            dark_layout(fig, height=360)
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        if "bts_type" in df.columns and "vendor" in df.columns:
            fig2 = px.bar(cross, x="bts_type", y="Total_Down_Hrs", color="vendor",
                          barmode="stack", title="Total Down Hours by BTS Type & Vendor",
                          color_discrete_sequence=PALETTE)
            dark_layout(fig2, height=360)
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown('<div class="section-header">Vendor-wise Mean Downtime</div>', unsafe_allow_html=True)
    c3, c4 = st.columns(2)

    with c3:
        if "vendor" in df.columns:
            vd = df.groupby("vendor").agg(
                Events=("bts_id","count"),
                Unique_BTS=("bts_id","nunique"),
                Total_Down_Hrs=("down_hours","sum"),
                Avg_Down_Hrs=("down_hours","mean"),
                Max_Down_Hrs=("down_hours","max"),
            ).reset_index().round(2)
            fig3 = px.scatter(vd, x="Events", y="Avg_Down_Hrs",
                              size="Total_Down_Hrs", color="vendor",
                              text="vendor", title="Vendor: Events vs Avg Duration",
                              color_discrete_sequence=PALETTE)
            fig3.update_traces(textposition="top center")
            dark_layout(fig3, height=360)
            st.plotly_chart(fig3, use_container_width=True)

    with c4:
        if "bts_type" in df.columns:
            bt = df.groupby(["bts_type"]).agg(
                Events=("bts_id","count"),
                Unique_BTS=("bts_id","nunique"),
                Total_Down_Hrs=("down_hours","sum"),
                Avg_Down_Hrs=("down_hours","mean"),
            ).reset_index().round(2)
            fig4 = px.funnel(bt, x="Total_Down_Hrs", y="bts_type",
                             title="Total Down Hours Funnel by BTS Type",
                             color="bts_type", color_discrete_sequence=PALETTE)
            dark_layout(fig4, height=360)
            st.plotly_chart(fig4, use_container_width=True)

    # Summary table
    st.markdown('<div class="section-header">Detailed BTS Type × Vendor Table</div>', unsafe_allow_html=True)
    if "bts_type" in df.columns and "vendor" in df.columns:
        tbl = df.groupby(["bts_type", "vendor", "site_type"] if "site_type" in df.columns else ["bts_type","vendor"]).agg(
            Events=("bts_id","count"),
            Unique_BTS=("bts_id","nunique"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
            Max_Down_Hrs=("down_hours","max"),
        ).reset_index().round(2).sort_values("Total_Down_Hrs", ascending=False)
        st.dataframe(tbl, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════
# TAB 2: FAULT ANALYSIS
# ══════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="section-header">Fault Type Distribution</div>', unsafe_allow_html=True)

    if "fault_type" not in df.columns:
        st.warning("fault_type column not found in data.")
    else:
        fault_agg = df.groupby("fault_type").agg(
            Events=("bts_id","count"),
            Unique_BTS=("bts_id","nunique"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
        ).reset_index().round(2).sort_values("Events", ascending=False)

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(fault_agg, x="Events", y="fault_type",
                         orientation="h", color="Events",
                         color_continuous_scale="Reds",
                         title="Fault Events (Pareto)")
            fig.update_layout(yaxis=dict(autorange="reversed"))
            dark_layout(fig, height=400)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.bar(fault_agg, x="Total_Down_Hrs", y="fault_type",
                          orientation="h", color="Avg_Down_Hrs",
                          color_continuous_scale="Oranges",
                          title="Total Down Hours by Fault")
            fig2.update_layout(yaxis=dict(autorange="reversed"))
            dark_layout(fig2, height=400)
            st.plotly_chart(fig2, use_container_width=True)

        # Fault × BTS type heatmap
        st.markdown('<div class="section-header">Fault Type × BTS Type Heatmap</div>', unsafe_allow_html=True)
        if "bts_type" in df.columns:
            heat = df.groupby(["fault_type","bts_type"])["bts_id"].count().unstack(fill_value=0)
            fig3 = px.imshow(heat, text_auto=True, color_continuous_scale="Blues",
                             title="Event Count: Fault Type × BTS Type")
            dark_layout(fig3, height=420)
            st.plotly_chart(fig3, use_container_width=True)

        # Fault × Vendor
        st.markdown('<div class="section-header">Fault Type × Vendor</div>', unsafe_allow_html=True)
        if "vendor" in df.columns:
            fv = df.groupby(["fault_type","vendor"]).agg(
                Events=("bts_id","count"),
                Total_Down_Hrs=("down_hours","sum"),
            ).reset_index().round(2)
            fig4 = px.bar(fv, x="fault_type", y="Events", color="vendor",
                          barmode="stack", title="Fault Events by Vendor",
                          color_discrete_sequence=PALETTE)
            fig4.update_xaxes(tickangle=30)
            dark_layout(fig4, height=380)
            st.plotly_chart(fig4, use_container_width=True)

        st.dataframe(fault_agg.rename(columns={
            "fault_type":"Fault Type","Events":"# Events",
            "Unique_BTS":"Unique BTS","Total_Down_Hrs":"Total Down Hrs",
            "Avg_Down_Hrs":"Avg Down Hrs"
        }), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════
# TAB 3: DURATION DEEP-DIVE
# ══════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-header">Downtime Duration Analysis</div>', unsafe_allow_html=True)

    if "down_hours" not in df.columns:
        st.warning("downPeriod column missing or unparseable.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.histogram(df, x="down_hours", nbins=40,
                               color_discrete_sequence=["#60a5fa"],
                               title="Distribution of Down Hours per Event",
                               labels={"down_hours":"Down Hours"})
            dark_layout(fig, height=330)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            band_cnt = df["duration_band"].value_counts().reindex(
                ["< 1 hr","1–4 hrs","4–8 hrs","8–24 hrs","> 24 hrs"]
            ).fillna(0).reset_index()
            band_cnt.columns = ["band","count"]
            colors = ["#22c55e","#84cc16","#f59e0b","#ef4444","#7c3aed"]
            fig2 = px.pie(band_cnt, values="count", names="band", hole=0.45,
                          title="Events by Duration Band",
                          color_discrete_sequence=colors)
            fig2.update_traces(textinfo="percent+label")
            dark_layout(fig2, height=330)
            st.plotly_chart(fig2, use_container_width=True)

        # Box plot by BTS type
        st.markdown('<div class="section-header">Duration Distribution by BTS Type</div>', unsafe_allow_html=True)
        c3, c4 = st.columns(2)
        with c3:
            if "bts_type" in df.columns:
                fig3 = px.box(df, x="bts_type", y="down_hours", color="bts_type",
                              title="Down Hours Box Plot by BTS Type",
                              color_discrete_sequence=PALETTE)
                dark_layout(fig3, height=350)
                st.plotly_chart(fig3, use_container_width=True)
        with c4:
            if "vendor" in df.columns:
                fig4 = px.violin(df, x="vendor", y="down_hours", color="vendor",
                                 box=True, title="Duration Violin Plot by Vendor",
                                 color_discrete_sequence=PALETTE)
                dark_layout(fig4, height=350)
                st.plotly_chart(fig4, use_container_width=True)

        # Worst BTS by total down hours
        st.markdown('<div class="section-header">🔴 Top 20 BTS by Total Downtime</div>', unsafe_allow_html=True)
        worst = df.groupby(["bts_name" if "bts_name" in df.columns else "bts_id", "bts_type", "vendor"]).agg(
            Events=("bts_id","count"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
        ).reset_index().sort_values("Total_Down_Hrs", ascending=False).head(20)
        worst = worst.round(2)
        name_col = "bts_name" if "bts_name" in worst.columns else "bts_id"
        fig5 = px.bar(worst, x=name_col, y="Total_Down_Hrs",
                      color="bts_type", text="Total_Down_Hrs",
                      title="Top 20 Worst BTS – Total Down Hours",
                      color_discrete_sequence=PALETTE)
        fig5.update_xaxes(tickangle=40)
        fig5.update_traces(texttemplate="%{text:.1f}h", textposition="outside")
        dark_layout(fig5, height=400)
        st.plotly_chart(fig5, use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# TAB 4: INCHARGE ANALYTICS
# ══════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="section-header">Incharge / JTO Responsibility Analysis</div>', unsafe_allow_html=True)

    has_incharge = "incharge" in df.columns
    has_jto      = "JTO INCHARGE" in df.columns

    if not has_incharge:
        st.warning("No `incharge` column found after PKEY merge. Check PKEY file BTSIPID mapping.")
    else:
        unmapped = df["incharge"].isna().sum()
        mapped   = df["incharge"].notna().sum()
        st.info(f"✅ Mapped: **{mapped}** events  |  ❌ Unmapped: **{unmapped}** events (no PKEY match)")

        ic_agg = df[df["incharge"].notna()].groupby("incharge").agg(
            Events=("bts_id","count"),
            Unique_BTS=("bts_id","nunique"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
            Max_Down_Hrs=("down_hours","max"),
        ).reset_index().round(2).sort_values("Total_Down_Hrs", ascending=False)

        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(ic_agg, x="incharge", y="Events",
                         color="Total_Down_Hrs", text="Events",
                         color_continuous_scale="Reds",
                         title="Alarm Events per Incharge")
            fig.update_traces(textposition="outside")
            dark_layout(fig, height=370)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.bar(ic_agg, x="incharge", y="Total_Down_Hrs",
                          color="Avg_Down_Hrs", text="Total_Down_Hrs",
                          color_continuous_scale="Oranges",
                          title="Total Down Hours per Incharge")
            fig2.update_traces(texttemplate="%{text:.1f}h", textposition="outside")
            dark_layout(fig2, height=370)
            st.plotly_chart(fig2, use_container_width=True)

        # Incharge × fault type
        st.markdown('<div class="section-header">Incharge × Fault Type Matrix</div>', unsafe_allow_html=True)
        if "fault_type" in df.columns:
            ic_fault = df[df["incharge"].notna()].groupby(
                ["incharge","fault_type"])["bts_id"].count().unstack(fill_value=0)
            fig3 = px.imshow(ic_fault, text_auto=True, color_continuous_scale="YlOrRd",
                             title="Event Count: Incharge × Fault Type",
                             aspect="auto")
            dark_layout(fig3, height=420)
            st.plotly_chart(fig3, use_container_width=True)

        # JTO analysis
        if has_jto:
            st.markdown('<div class="section-header">JTO Incharge Analysis</div>', unsafe_allow_html=True)
            jto_agg = df[df["JTO INCHARGE"].notna()].groupby("JTO INCHARGE").agg(
                Events=("bts_id","count"),
                Unique_BTS=("bts_id","nunique"),
                Total_Down_Hrs=("down_hours","sum"),
                Avg_Down_Hrs=("down_hours","mean"),
            ).reset_index().round(2).sort_values("Total_Down_Hrs", ascending=False)
            fig4 = px.bar(jto_agg, x="JTO INCHARGE", y="Total_Down_Hrs",
                          color="Events", text="Total_Down_Hrs",
                          color_continuous_scale="Blues",
                          title="Total Down Hours per JTO Incharge")
            fig4.update_traces(texttemplate="%{text:.1f}h", textposition="outside")
            dark_layout(fig4, height=360)
            st.plotly_chart(fig4, use_container_width=True)

        st.markdown('<div class="section-header">Incharge Summary Table</div>', unsafe_allow_html=True)
        st.dataframe(ic_agg, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════
# TAB 5: DAILY TREND
# ══════════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<div class="section-header">Daily Alarm Trend (Monthly View)</div>', unsafe_allow_html=True)

    if "log_date" not in df.columns or df["log_date"].isna().all():
        st.warning("No date information could be parsed. Ensure bts_down_dt column is present.")
    else:
        daily = df.groupby("log_date").agg(
            Events=("bts_id","count"),
            Unique_BTS=("bts_id","nunique"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
        ).reset_index().sort_values("log_date")
        daily["log_date"] = pd.to_datetime(daily["log_date"])

        fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                            subplot_titles=("Daily Event Count", "Daily Total Down Hours"),
                            vertical_spacing=0.1)
        fig.add_trace(go.Bar(x=daily["log_date"], y=daily["Events"],
                             name="Events", marker_color="#60a5fa"), row=1, col=1)
        fig.add_trace(go.Scatter(x=daily["log_date"], y=daily["Total_Down_Hrs"],
                                 mode="lines+markers", name="Down Hrs",
                                 line=dict(color="#f97316", width=2)), row=2, col=1)
        fig.update_layout(
            paper_bgcolor=DARK_BG, plot_bgcolor=CARD_BG,
            font=dict(color="#94a3b8"), height=500,
            legend=dict(bgcolor=CARD_BG)
        )
        st.plotly_chart(fig, use_container_width=True)

        # BTS type trend
        if "bts_type" in df.columns:
            st.markdown('<div class="section-header">Daily Events by BTS Type</div>', unsafe_allow_html=True)
            daily_bt = df.groupby(["log_date","bts_type"])["bts_id"].count().reset_index()
            daily_bt["log_date"] = pd.to_datetime(daily_bt["log_date"])
            daily_bt.columns = ["date","bts_type","events"]
            fig2 = px.area(daily_bt, x="date", y="events", color="bts_type",
                           title="Daily Events by BTS Type",
                           color_discrete_sequence=PALETTE)
            dark_layout(fig2, height=360)
            st.plotly_chart(fig2, use_container_width=True)

        # Fault type trend
        if "fault_type" in df.columns:
            st.markdown('<div class="section-header">Top Fault Types Over Time</div>', unsafe_allow_html=True)
            top_faults = df["fault_type"].value_counts().head(5).index.tolist()
            ft_trend = df[df["fault_type"].isin(top_faults)].groupby(
                ["log_date","fault_type"])["bts_id"].count().reset_index()
            ft_trend["log_date"] = pd.to_datetime(ft_trend["log_date"])
            ft_trend.columns = ["date","fault_type","count"]
            fig3 = px.line(ft_trend, x="date", y="count", color="fault_type",
                           markers=True, title="Top 5 Fault Types – Daily Count",
                           color_discrete_sequence=PALETTE)
            dark_layout(fig3, height=360)
            st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# TAB 6: SUMMARY REPORT
# ══════════════════════════════════════════════════════════════════
with tabs[6]:
    st.markdown('<div class="section-header">📋 Executive Summary Report</div>', unsafe_allow_html=True)

    # Date range
    if "log_date" in df.columns and df["log_date"].notna().any():
        d_min = df["log_date"].min()
        d_max = df["log_date"].max()
        period_str = f"{d_min} → {d_max}"
    else:
        period_str = "N/A"

    # Top fault
    top_fault = df["fault_type"].value_counts().idxmax() if "fault_type" in df.columns else "N/A"
    top_vendor_hrs = df.groupby("vendor")["down_hours"].sum().idxmax() if "vendor" in df.columns else "N/A"
    top_bts_type = df.groupby("bts_type")["down_hours"].sum().idxmax() if "bts_type" in df.columns else "N/A"
    worst_bts_name = df.groupby("bts_name" if "bts_name" in df.columns else "bts_id")["down_hours"].sum().idxmax() if "down_hours" in df.columns else "N/A"
    
    if "incharge" in df.columns:
        worst_ic = df.groupby("incharge")["down_hours"].sum().idxmax() if df["incharge"].notna().any() else "N/A"
    else:
        worst_ic = "N/A"

    # Critical BTS (> 24h total down)
    crit_bts = 0
    if "bts_name" in df.columns and "down_hours" in df.columns:
        crit_bts = int((df.groupby("bts_name")["down_hours"].sum() > 24).sum())

    st.markdown(f"""
    <div class="alert-box">
    <b>📅 Period:</b> {period_str} &nbsp;|&nbsp; <b>Files Loaded:</b> {files_loaded} daily files
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 🔑 Key Metrics")
        metrics = {
            "Total Alarm Events": f"{total_events:,}",
            "Unique BTS Affected": f"{unique_bts:,}",
            "Total Cumulative Down Hours": f"{total_down_hrs:,.1f} hrs",
            "Average Down per Event": f"{avg_down_hrs:.2f} hrs",
            "Critical BTS (>24h total)": f"{crit_bts}",
            "Distinct Fault Types": f"{df['fault_type'].nunique() if 'fault_type' in df.columns else 'N/A'}",
            "PKEY Unmatched Events": f"{unmatched_pkey}",
        }
        for k, v in metrics.items():
            st.markdown(f"- **{k}:** {v}")

    with col2:
        st.markdown("#### 🏆 Worst Performers")
        worst_items = {
            "Top Fault Category": top_fault,
            "Vendor with Most Downtime": top_vendor_hrs,
            "BTS Type with Most Downtime": top_bts_type,
            "Worst Single BTS": worst_bts_name,
            "Incharge with Highest Downtime": worst_ic,
        }
        for k, v in worst_items.items():
            st.markdown(f"- **{k}:** {v}")

    # Duration band summary
    st.markdown("#### ⏱️ Duration Band Summary")
    if "duration_band" in df.columns:
        band_order = ["< 1 hr","1–4 hrs","4–8 hrs","8–24 hrs","> 24 hrs"]
        band_s = df.groupby("duration_band").agg(
            Events=("bts_id","count"),
            Total_Down_Hrs=("down_hours","sum"),
        ).reindex(band_order).fillna(0).astype({"Events":int}).reset_index()
        band_s["Total_Down_Hrs"] = band_s["Total_Down_Hrs"].round(1)
        band_s["% of Events"] = (band_s["Events"] / band_s["Events"].sum() * 100).round(1)
        st.dataframe(band_s, use_container_width=True, hide_index=True)

    # Fault type summary
    st.markdown("#### ⚠️ Fault Type Summary")
    if "fault_type" in df.columns:
        ft_s = df.groupby("fault_type").agg(
            Events=("bts_id","count"),
            Unique_BTS=("bts_id","nunique"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
        ).reset_index().sort_values("Events", ascending=False).round(2)
        ft_s["% Share"] = (ft_s["Events"] / ft_s["Events"].sum() * 100).round(1)
        st.dataframe(ft_s, use_container_width=True, hide_index=True)

    # Troubleshooting category analysis
    st.markdown("#### 🔧 Troubleshooting Category Analysis")
    # ── Keyword order matters: first match wins, so put longer/specific phrases
    #    before short ones to avoid partial false-matches.
    #    Categories are evaluated top-to-bottom; Power/Battery before Hardware
    #    so "Pp Control Panel Fault" (pp control) hits Power/Battery, not Hardware.
    cat_map = {
        "Power/Battery":   [
            "battery", "mains", "power plant", "pp control",
            "dg", "generator", "eb pole", "eb supply", "eb ",
        ],
        "Transmission":    [
            "ofc",          # optical fibre cable breaks (SSA OFC, CNTX-Zone OFC)
            "e1",           # E1 link failures
            "aggregation",  # aggregation site issues
            "hub site",     # due-to-hub-site outages
            "hub",
            "ssa media",    # SSA media issues
            "media issue",  # generic media issues
            "cntx",         # concentration-zone breaks
            "rrh link",     # RRH link-down
            "link down",
            "transmission",
            "fiber", "fibre",
            "microwave",
        ],
        "Hardware":        [
            "cpan",         # control panel faults
            "tcs",          # TCS/Tejas vendor visits → hardware rectification
            "tejas",
            "hardware",
            "board", "card", "equipment",
        ],
        "Environment":     [
            "tempature",    # common field typo for "temperature"
            "temperature",
            "ac ", "cooling", "fire", "flood",
        ],
        "Software/Config": [
            "resetting",    # site resetting issue
            "software", "config", "reset", "upgrade", "parameter",
        ],
        "Unknown/Other":   []
    }
    if "fault_type" in df.columns:
        def categorize(ft):
            ft_lower = str(ft).lower()
            for cat, kws in cat_map.items():
                if any(kw in ft_lower for kw in kws):
                    return cat
            return "Unknown/Other"
        df["trouble_category"] = df["fault_type"].apply(categorize)
        tc = df.groupby("trouble_category").agg(
            Events=("bts_id","count"),
            Total_Down_Hrs=("down_hours","sum"),
            Avg_Down_Hrs=("down_hours","mean"),
        ).reset_index().sort_values("Events", ascending=False).round(2)
        tc["% Share"] = (tc["Events"] / tc["Events"].sum() * 100).round(1)

        c1, c2 = st.columns(2)
        with c1:
            fig = px.pie(tc, values="Events", names="trouble_category", hole=0.4,
                         title="Troubleshooting Category – Event Share",
                         color_discrete_sequence=PALETTE)
            fig.update_traces(textinfo="percent+label")
            dark_layout(fig, height=350)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig2 = px.bar(tc, x="trouble_category", y="Total_Down_Hrs",
                          color="trouble_category", text="Total_Down_Hrs",
                          title="Down Hours by Troubleshooting Category",
                          color_discrete_sequence=PALETTE)
            fig2.update_traces(texttemplate="%{text:.0f}h", textposition="outside")
            dark_layout(fig2, height=350)
            st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(tc, use_container_width=True, hide_index=True)

    # Export button
    st.markdown("---")
    st.markdown("#### 📥 Export Processed Data")

    @st.cache_data
    def to_excel_export(dataframe):
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            dataframe.to_excel(writer, sheet_name="Alarm_Data", index=False)
            if "fault_type" in dataframe.columns:
                dataframe.groupby("fault_type").agg(
                    Events=("bts_id","count"),
                    Total_Down_Hrs=("down_hours","sum"),
                    Avg_Down_Hrs=("down_hours","mean"),
                ).reset_index().to_excel(writer, sheet_name="Fault_Summary", index=False)
            if "incharge" in dataframe.columns and dataframe["incharge"].notna().any():
                dataframe[dataframe["incharge"].notna()].groupby("incharge").agg(
                    Events=("bts_id","count"),
                    Total_Down_Hrs=("down_hours","sum"),
                ).reset_index().to_excel(writer, sheet_name="Incharge_Summary", index=False)
        return buf.getvalue()

    export_bytes = to_excel_export(df)
    st.download_button(
        label="📥 Download Full Analysis (Excel)",
        data=export_bytes,
        file_name=f"BTS_Alarm_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ══════════════════════════════════════════════════════════════════
# TAB 7: RAW DATA
# ══════════════════════════════════════════════════════════════════
with tabs[7]:
    st.markdown('<div class="section-header">🗃️ Processed & Merged Dataset</div>', unsafe_allow_html=True)
    st.caption(f"Showing {len(df):,} rows after filters. Total loaded (before filter): {len(df_all):,}")

    search = st.text_input("🔍 Filter rows (searches bts_name, fault_type, incharge)", "")
    disp = df.copy()
    if search:
        mask = pd.Series([False] * len(disp))
        for col in ["bts_name","fault_type","incharge","vendor","bts_type","sdca_name"]:
            if col in disp.columns:
                mask |= disp[col].astype(str).str.contains(search, case=False, na=False)
        disp = disp[mask]

    st.dataframe(disp, use_container_width=True, height=500)
    st.caption(f"Displaying {len(disp):,} rows")
