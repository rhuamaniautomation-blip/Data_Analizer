"""
╔══════════════════════════════════════════════════════════════════════════════╗
║          PROCESS DATA ANALYZER — Industrial Engineering Platform             ║
║          Análisis avanzado de variables de proceso: T, P, F, Q, etc.        ║
╚══════════════════════════════════════════════════════════════════════════════╝
Autor  : Claude / Anthropic
Versión: 2.0.0
Python : 3.9+
Deps   : streamlit, pandas, numpy, plotly, scipy, openpyxl, xlrd, kaleido
"""

# ─────────────────────────────────────────────────────────────────────────────
# 0. IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import io
import warnings
import base64
from datetime import datetime, timedelta
from typing import Optional, Tuple, List, Dict, Any

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import scipy.stats as stats
from scipy.signal import savgol_filter
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# 1. PAGE CONFIG & THEME
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Process Data Analyzer",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "Get Help": None,
        "Report a bug": None,
        "About": "# Process Data Analyzer v2.0\nAnálisis avanzado de datos de proceso industrial.",
    },
)

# ─────────────────────────────────────────────────────────────────────────────
# 2. CUSTOM CSS — DARK INDUSTRIAL THEME
# ─────────────────────────────────────────────────────────────────────────────
CUSTOM_CSS = """
<style>
/* ── Google Fonts ─────────────────────────────────────────────────────────── */
@import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Share+Tech+Mono&family=Nunito:wght@300;400;600;700&display=swap');

/* ── Root Variables ────────────────────────────────────────────────────────── */
:root {
    --bg-base:        #0a0e17;
    --bg-panel:       #0f1520;
    --bg-card:        #141c2e;
    --bg-card-hover:  #1a2438;
    --border:         #1e2d4a;
    --border-accent:  #2a4080;
    --primary:        #00d4ff;
    --primary-glow:   rgba(0,212,255,0.18);
    --secondary:      #ff6b35;
    --success:        #00e676;
    --warning:        #ffd740;
    --danger:         #ff4444;
    --text-main:      #e8f0fe;
    --text-muted:     #7a8fb5;
    --text-label:     #a0b4d6;
    --font-display:   'Rajdhani', sans-serif;
    --font-mono:      'Share Tech Mono', monospace;
    --font-body:      'Nunito', sans-serif;
    --radius:         8px;
    --radius-lg:      14px;
    --shadow:         0 4px 24px rgba(0,0,0,0.6);
    --shadow-glow:    0 0 20px rgba(0,212,255,0.15);
}

/* ── Global Reset ──────────────────────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: var(--font-body) !important;
    background-color: var(--bg-base) !important;
    color: var(--text-main) !important;
}

/* ── Main container ────────────────────────────────────────────────────────── */
.main .block-container {
    padding: 1.5rem 2rem 3rem 2rem !important;
    max-width: 1700px !important;
    background: var(--bg-base) !important;
}

/* ── Sidebar ───────────────────────────────────────────────────────────────── */
section[data-testid="stSidebar"] {
    background: var(--bg-panel) !important;
    border-right: 1px solid var(--border) !important;
    padding-top: 1rem !important;
}
section[data-testid="stSidebar"] > div {
    padding: 1.2rem 1rem !important;
}

/* ── Header banner ─────────────────────────────────────────────────────────── */
.header-banner {
    background: linear-gradient(135deg, #0a1628 0%, #0e1f3e 50%, #0a1628 100%);
    border: 1px solid var(--border-accent);
    border-radius: var(--radius-lg);
    padding: 1.6rem 2.2rem;
    margin-bottom: 1.8rem;
    position: relative;
    overflow: hidden;
    box-shadow: var(--shadow), var(--shadow-glow);
}
.header-banner::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(90deg, transparent, var(--primary), transparent);
}
.header-banner::after {
    content: '';
    position: absolute; bottom: 0; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, transparent, var(--border-accent), transparent);
}
.header-title {
    font-family: var(--font-display) !important;
    font-size: 2.2rem !important;
    font-weight: 700 !important;
    letter-spacing: 3px !important;
    color: var(--primary) !important;
    text-shadow: 0 0 30px rgba(0,212,255,0.4);
    margin: 0 !important;
    line-height: 1.1 !important;
}
.header-sub {
    font-family: var(--font-mono) !important;
    font-size: 0.75rem !important;
    color: var(--text-muted) !important;
    letter-spacing: 2px !important;
    margin-top: 0.4rem !important;
}
.header-tag {
    display: inline-block;
    background: rgba(0,212,255,0.1);
    border: 1px solid rgba(0,212,255,0.3);
    color: var(--primary);
    font-family: var(--font-mono);
    font-size: 0.65rem;
    padding: 2px 10px;
    border-radius: 20px;
    letter-spacing: 1.5px;
    margin-right: 6px;
    margin-top: 8px;
}

/* ── KPI Cards ─────────────────────────────────────────────────────────────── */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 12px;
    margin-bottom: 1.4rem;
}
.kpi-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1rem 1.2rem;
    transition: all 0.25s ease;
    position: relative;
    overflow: hidden;
}
.kpi-card:hover {
    background: var(--bg-card-hover);
    border-color: var(--primary);
    box-shadow: 0 0 16px rgba(0,212,255,0.12);
    transform: translateY(-2px);
}
.kpi-card::before {
    content: '';
    position: absolute; top: 0; left: 0; width: 3px; height: 100%;
    background: var(--primary);
    opacity: 0.6;
}
.kpi-label {
    font-family: var(--font-mono);
    font-size: 0.62rem;
    letter-spacing: 1.5px;
    color: var(--text-muted);
    text-transform: uppercase;
    margin-bottom: 6px;
}
.kpi-value {
    font-family: var(--font-display);
    font-size: 1.55rem;
    font-weight: 700;
    color: var(--primary);
    line-height: 1;
}
.kpi-unit {
    font-family: var(--font-mono);
    font-size: 0.7rem;
    color: var(--text-muted);
    margin-top: 4px;
}
.kpi-card.warning .kpi-value { color: var(--warning); }
.kpi-card.warning::before { background: var(--warning); }
.kpi-card.danger  .kpi-value { color: var(--danger);  }
.kpi-card.danger::before  { background: var(--danger);  }
.kpi-card.success .kpi-value { color: var(--success); }
.kpi-card.success::before { background: var(--success); }
.kpi-card.secondary .kpi-value { color: var(--secondary); }
.kpi-card.secondary::before { background: var(--secondary); }

/* ── Section Headers ───────────────────────────────────────────────────────── */
.section-header {
    display: flex; align-items: center; gap: 10px;
    margin: 1.6rem 0 0.8rem 0;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--border);
}
.section-icon {
    font-size: 1.1rem;
    filter: drop-shadow(0 0 6px var(--primary));
}
.section-title {
    font-family: var(--font-display) !important;
    font-size: 1.15rem !important;
    font-weight: 600 !important;
    letter-spacing: 2px !important;
    color: var(--text-main) !important;
    text-transform: uppercase !important;
    margin: 0 !important;
}
.section-pill {
    margin-left: auto;
    font-family: var(--font-mono);
    font-size: 0.6rem;
    color: var(--text-muted);
    border: 1px solid var(--border);
    padding: 2px 8px;
    border-radius: 20px;
}

/* ── Stats Table ───────────────────────────────────────────────────────────── */
.stats-table {
    width: 100%;
    border-collapse: collapse;
    font-family: var(--font-body);
    font-size: 0.85rem;
}
.stats-table th {
    font-family: var(--font-mono);
    font-size: 0.65rem;
    letter-spacing: 1.5px;
    color: var(--text-muted);
    text-transform: uppercase;
    padding: 8px 14px;
    border-bottom: 1px solid var(--border-accent);
    text-align: left;
}
.stats-table td {
    padding: 9px 14px;
    border-bottom: 1px solid var(--border);
    color: var(--text-main);
}
.stats-table tr:hover td { background: var(--bg-card-hover); }
.stats-table td.val {
    font-family: var(--font-mono);
    color: var(--primary);
    font-weight: 600;
}
.stats-table td.bad { color: var(--danger); }
.stats-table td.ok  { color: var(--success); }
.stats-table td.warn { color: var(--warning); }

/* ── Info boxes ────────────────────────────────────────────────────────────── */
.info-box {
    background: rgba(0,212,255,0.05);
    border: 1px solid rgba(0,212,255,0.2);
    border-radius: var(--radius);
    padding: 0.8rem 1.1rem;
    font-size: 0.82rem;
    color: var(--text-label);
    margin: 0.6rem 0;
}
.warn-box {
    background: rgba(255,215,64,0.06);
    border: 1px solid rgba(255,215,64,0.25);
    border-radius: var(--radius);
    padding: 0.8rem 1.1rem;
    font-size: 0.82rem;
    color: var(--warning);
    margin: 0.6rem 0;
}
.err-box {
    background: rgba(255,68,68,0.06);
    border: 1px solid rgba(255,68,68,0.25);
    border-radius: var(--radius);
    padding: 0.8rem 1.1rem;
    font-size: 0.82rem;
    color: var(--danger);
    margin: 0.6rem 0;
}

/* ── Streamlit overrides ───────────────────────────────────────────────────── */
div[data-testid="stMetric"] {
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
    padding: 0.8rem 1rem !important;
}
.stSelectbox > div > div, .stMultiSelect > div > div {
    background: var(--bg-card) !important;
    border-color: var(--border) !important;
    color: var(--text-main) !important;
}
.stSlider > div > div { color: var(--primary) !important; }
div.stButton > button {
    background: linear-gradient(135deg, rgba(0,212,255,0.15), rgba(0,212,255,0.08)) !important;
    border: 1px solid rgba(0,212,255,0.4) !important;
    color: var(--primary) !important;
    font-family: var(--font-display) !important;
    font-weight: 600 !important;
    letter-spacing: 1.5px !important;
    border-radius: var(--radius) !important;
    transition: all 0.2s ease !important;
}
div.stButton > button:hover {
    background: linear-gradient(135deg, rgba(0,212,255,0.3), rgba(0,212,255,0.15)) !important;
    box-shadow: 0 0 18px rgba(0,212,255,0.25) !important;
    transform: translateY(-1px) !important;
}
div.stDownloadButton > button {
    background: linear-gradient(135deg, rgba(0,230,118,0.15), rgba(0,230,118,0.06)) !important;
    border: 1px solid rgba(0,230,118,0.4) !important;
    color: var(--success) !important;
    font-family: var(--font-display) !important;
    font-weight: 600 !important;
    letter-spacing: 1.5px !important;
    border-radius: var(--radius) !important;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: var(--bg-panel) !important;
    border-radius: var(--radius) !important;
    padding: 4px !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border: none !important;
    color: var(--text-muted) !important;
    font-family: var(--font-display) !important;
    font-weight: 600 !important;
    letter-spacing: 1.5px !important;
    font-size: 0.85rem !important;
    padding: 6px 18px !important;
    border-radius: 6px !important;
}
.stTabs [aria-selected="true"] {
    background: rgba(0,212,255,0.12) !important;
    color: var(--primary) !important;
}
.stExpander {
    background: var(--bg-card) !important;
    border: 1px solid var(--border) !important;
    border-radius: var(--radius) !important;
}
.stDataFrame { background: var(--bg-card) !important; }
[data-testid="stFileUploader"] {
    background: var(--bg-card) !important;
    border: 2px dashed var(--border-accent) !important;
    border-radius: var(--radius-lg) !important;
}
.stCheckbox > label { color: var(--text-label) !important; }
.stRadio > label { color: var(--text-label) !important; }
h1, h2, h3, h4 {
    font-family: var(--font-display) !important;
    color: var(--text-main) !important;
}
/* Sidebar labels */
.css-1544g2n p, .stSidebar p, .stSidebar label {
    color: var(--text-label) !important;
    font-size: 0.82rem !important;
}
/* Plotly modebar */
.modebar { background: transparent !important; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# 3. PLOTLY LAYOUT TEMPLATE (dark industrial)
# ─────────────────────────────────────────────────────────────────────────────
PLOTLY_TEMPLATE = dict(
    layout=go.Layout(
        paper_bgcolor="#0a0e17",
        plot_bgcolor="#0f1520",
        font=dict(family="Rajdhani, Share Tech Mono, monospace", color="#e8f0fe", size=12),
        xaxis=dict(
            gridcolor="#1e2d4a", gridwidth=1, zerolinecolor="#1e2d4a",
            linecolor="#2a4080", tickcolor="#7a8fb5",
            showspikes=True, spikethickness=1, spikecolor="#00d4ff",
            spikedash="dot", spikemode="across",
        ),
        yaxis=dict(
            gridcolor="#1e2d4a", gridwidth=1, zerolinecolor="#1e2d4a",
            linecolor="#2a4080", tickcolor="#7a8fb5",
            showspikes=True, spikethickness=1, spikecolor="#00d4ff",
            spikedash="dot", spikemode="across",
        ),
        hovermode="x unified",
        hoverlabel=dict(
            bgcolor="#141c2e", bordercolor="#2a4080",
            font=dict(family="Share Tech Mono, monospace", color="#e8f0fe", size=11),
        ),
        legend=dict(
            bgcolor="rgba(10,14,23,0.85)", bordercolor="#2a4080", borderwidth=1,
            font=dict(family="Share Tech Mono", color="#a0b4d6", size=10),
        ),
        margin=dict(l=60, r=30, t=50, b=60),
        title=dict(font=dict(family="Rajdhani", size=16, color="#e8f0fe")),
    )
)

COLORS_PRIMARY = [
    "#00d4ff", "#ff6b35", "#00e676", "#ffd740", "#ea80fc",
    "#40c4ff", "#ff4081", "#69f0ae", "#ffab40", "#e040fb",
]

# ─────────────────────────────────────────────────────────────────────────────
# 4. UTILITY FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def detect_datetime_columns(df: pd.DataFrame) -> List[str]:
    """Heuristic detection of datetime-like columns."""
    candidates = []
    for col in df.columns:
        if df[col].dtype == "object":
            sample = df[col].dropna().head(20).astype(str)
            hits = 0
            for v in sample:
                try:
                    pd.to_datetime(v, infer_datetime_format=True)
                    hits += 1
                except Exception:
                    pass
            if hits >= len(sample) * 0.75:
                candidates.append(col)
        elif np.issubdtype(df[col].dtype, np.datetime64):
            candidates.append(col)
    return candidates


def detect_numeric_columns(df: pd.DataFrame) -> List[str]:
    return [c for c in df.columns if np.issubdtype(df[c].dtype, np.number)]


def parse_datetime_column(series: pd.Series) -> pd.Series:
    """Try multiple datetime parsing strategies."""
    try:
        return pd.to_datetime(series, infer_datetime_format=True)
    except Exception:
        pass
    formats = [
        "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M", "%d-%m-%Y %H:%M:%S",
        "%Y/%m/%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d/%m/%Y",
        "%m/%d/%Y", "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return pd.to_datetime(series, format=fmt)
        except Exception:
            continue
    return pd.to_datetime(series, errors="coerce")


@st.cache_data(show_spinner=False)
def load_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Load CSV or Excel into DataFrame."""
    ext = filename.rsplit(".", 1)[-1].lower()
    if ext == "csv":
        for enc in ("utf-8", "latin-1", "cp1252", "iso-8859-1"):
            for sep in (",", ";", "\t", "|"):
                try:
                    df = pd.read_csv(io.BytesIO(file_bytes), sep=sep, encoding=enc,
                                     low_memory=False)
                    if df.shape[1] > 1:
                        return df
                except Exception:
                    pass
        return pd.read_csv(io.BytesIO(file_bytes), sep=None, engine="python",
                           encoding="utf-8", errors="replace")
    else:  # xlsx / xls
        try:
            xf = pd.ExcelFile(io.BytesIO(file_bytes))
            if len(xf.sheet_names) > 1:
                sheet = st.sidebar.selectbox("📄 Hoja Excel", xf.sheet_names, key="excel_sheet")
            else:
                sheet = xf.sheet_names[0]
            return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
        except Exception as e:
            st.error(f"Error leyendo Excel: {e}")
            return pd.DataFrame()


def compute_statistics(series: pd.Series, lsl: float = None,
                        usl: float = None) -> Dict[str, Any]:
    """Compute comprehensive process statistics."""
    s = series.dropna()
    n = len(s)
    if n == 0:
        return {}

    mean     = float(s.mean())
    median   = float(s.median())
    std      = float(s.std(ddof=1)) if n > 1 else 0.0
    variance = float(s.var(ddof=1)) if n > 1 else 0.0
    minimum  = float(s.min())
    maximum  = float(s.max())
    rng      = maximum - minimum
    q1       = float(s.quantile(0.25))
    q3       = float(s.quantile(0.75))
    iqr      = q3 - q1
    skew     = float(s.skew())
    kurt     = float(s.kurtosis())
    cv       = (std / mean * 100) if mean != 0 else float("nan")

    # Percentiles
    p5  = float(s.quantile(0.05))
    p95 = float(s.quantile(0.95))
    p10 = float(s.quantile(0.10))
    p90 = float(s.quantile(0.90))

    # Process Capability (if limits provided)
    cp = cpk = cpl = cpu = pp = ppk = float("nan")
    sigma_level = float("nan")
    ppm_est     = float("nan")
    yield_pct   = float("nan")

    if lsl is not None and usl is not None and std > 0:
        spec_range = usl - lsl
        cp  = spec_range / (6 * std)
        cpl = (mean - lsl) / (3 * std)
        cpu = (usl - mean) / (3 * std)
        cpk = min(cpl, cpu)
        pp  = spec_range / (6 * std)   # same as Cp when using sample std
        ppk = min((mean - lsl) / (3 * std), (usl - mean) / (3 * std))
        sigma_level = cpk * 3
        # PPM estimate
        ppm_est = (stats.norm.cdf((lsl - mean) / std) +
                   (1 - stats.norm.cdf((usl - mean) / std))) * 1_000_000
        yield_pct = 100 - ppm_est / 10_000
    elif lsl is not None and std > 0:
        cpl = (mean - lsl) / (3 * std)
        cpk = cpl
    elif usl is not None and std > 0:
        cpu = (usl - mean) / (3 * std)
        cpk = cpu

    # Normality test (Shapiro–Wilk, max 5000 samples)
    sw_stat = sw_p = float("nan")
    if 3 <= n <= 5000:
        try:
            sw_stat, sw_p = stats.shapiro(s.sample(min(n, 5000), random_state=42))
        except Exception:
            pass

    # Anderson–Darling
    ad_stat = float("nan")
    try:
        ad_res  = stats.anderson(s, dist="norm")
        ad_stat = float(ad_res.statistic)
    except Exception:
        pass

    # Missing / outliers (IQR method)
    n_missing = int(series.isna().sum())
    fence_lo  = q1 - 1.5 * iqr
    fence_hi  = q3 + 1.5 * iqr
    n_outliers = int(((s < fence_lo) | (s > fence_hi)).sum())

    # Moving statistics (last 20% of data)
    tail_n = max(1, n // 5)
    recent_mean = float(s.iloc[-tail_n:].mean())
    recent_std  = float(s.iloc[-tail_n:].std(ddof=1)) if tail_n > 1 else 0.0

    return dict(
        n=n, mean=mean, median=median, std=std, variance=variance,
        minimum=minimum, maximum=maximum, range=rng,
        q1=q1, q3=q3, iqr=iqr, p5=p5, p10=p10, p90=p90, p95=p95,
        skewness=skew, kurtosis=kurt, cv=cv,
        cp=cp, cpk=cpk, cpl=cpl, cpu=cpu, pp=pp, ppk=ppk,
        sigma_level=sigma_level, ppm_est=ppm_est, yield_pct=yield_pct,
        sw_stat=sw_stat, sw_p=sw_p, ad_stat=ad_stat,
        n_missing=n_missing, n_outliers=n_outliers,
        recent_mean=recent_mean, recent_std=recent_std,
        lsl=lsl, usl=usl,
    )


def fmt(val: Any, decimals: int = 4) -> str:
    """Format numeric value for display."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "—"
    if isinstance(val, (int, np.integer)):
        return f"{val:,}"
    return f"{val:,.{decimals}f}"


def smooth_series(s: pd.Series, window: int = 5) -> pd.Series:
    """Savitzky–Golay smoothing."""
    if len(s) < window + 2:
        return s.rolling(window=max(3, window), min_periods=1).mean()
    poly = min(3, window - 1)
    try:
        return pd.Series(savgol_filter(s.values, window_length=window,
                                       polyorder=poly), index=s.index)
    except Exception:
        return s.rolling(window=window, min_periods=1).mean()


def fig_to_png(fig: go.Figure) -> bytes:
    """Export Plotly figure to PNG bytes using kaleido."""
    try:
        return fig.to_image(format="png", width=1400, height=700, scale=2)
    except Exception:
        return b""


def fig_to_html(fig: go.Figure) -> str:
    """Export Plotly figure as interactive HTML."""
    return fig.to_html(full_html=True, include_plotlyjs="cdn")


def apply_template(fig: go.Figure) -> go.Figure:
    """Apply the global Plotly template to a figure."""
    fig.update_layout(**PLOTLY_TEMPLATE["layout"].to_plotly_json())
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# 5. CHART BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def build_trend_chart(
    df: pd.DataFrame,
    time_col: str,
    var_cols: List[str],
    show_smooth: bool = False,
    smooth_window: int = 11,
    show_markers: bool = False,
    lsl: Dict[str, float] = None,
    usl: Dict[str, float] = None,
    target: Dict[str, float] = None,
    show_mean_line: bool = True,
    y_range: Tuple[float, float] = None,
) -> go.Figure:
    """Main time-series trend chart with interactive hover (MATLAB-like crosshair)."""
    n_vars = len(var_cols)
    if n_vars == 0:
        return go.Figure()

    # Shared x-axis (first variable) or separate rows
    rows = n_vars
    fig = make_subplots(
        rows=rows, cols=1,
        shared_xaxes=True,
        vertical_spacing=0.04,
        subplot_titles=[f"▸  {v}" for v in var_cols],
    )

    for i, col in enumerate(var_cols, start=1):
        color = COLORS_PRIMARY[(i - 1) % len(COLORS_PRIMARY)]
        s = df[col].copy()

        # Raw trace
        fig.add_trace(
            go.Scatter(
                x=df[time_col], y=s,
                mode="lines+markers" if show_markers else "lines",
                name=col,
                line=dict(color=color, width=1.8),
                marker=dict(size=4, color=color) if show_markers else dict(size=0),
                hovertemplate=(
                    f"<b style='color:{color}'>{col}</b><br>"
                    "📅 %{x|%Y-%m-%d %H:%M:%S}<br>"
                    f"📊 %{{y:,.4f}}<extra></extra>"
                ),
            ),
            row=i, col=1,
        )

        # Smoothed overlay
        if show_smooth and len(s.dropna()) > smooth_window:
            s_sm = smooth_series(s.dropna(), window=smooth_window)
            fig.add_trace(
                go.Scatter(
                    x=df.loc[s.dropna().index, time_col], y=s_sm,
                    mode="lines",
                    name=f"{col} (smooth)",
                    line=dict(color=color, width=2.5, dash="dot"),
                    opacity=0.7,
                    hoverinfo="skip",
                    showlegend=True,
                ),
                row=i, col=1,
            )

        # Mean line
        if show_mean_line:
            mean_val = float(s.mean())
            fig.add_hline(
                y=mean_val, row=i, col=1,
                line=dict(color=color, width=1, dash="dash"),
                annotation_text=f"μ={mean_val:,.3f}",
                annotation_font_color=color,
                annotation_font_size=9,
                annotation_position="right",
                opacity=0.5,
            )

        # Spec limits
        if lsl and col in lsl:
            fig.add_hline(
                y=lsl[col], row=i, col=1,
                line=dict(color="#ff4444", width=1.5, dash="longdash"),
                annotation_text=f"LSL={lsl[col]:,.3f}",
                annotation_font_color="#ff4444",
                annotation_font_size=9,
                annotation_position="right",
                opacity=0.8,
            )
        if usl and col in usl:
            fig.add_hline(
                y=usl[col], row=i, col=1,
                line=dict(color="#ff4444", width=1.5, dash="longdash"),
                annotation_text=f"USL={usl[col]:,.3f}",
                annotation_font_color="#ff4444",
                annotation_font_size=9,
                annotation_position="right",
                opacity=0.8,
            )
        if target and col in target:
            fig.add_hline(
                y=target[col], row=i, col=1,
                line=dict(color="#ffd740", width=1.2, dash="dot"),
                annotation_text=f"TGT={target[col]:,.3f}",
                annotation_font_color="#ffd740",
                annotation_font_size=9,
                annotation_position="right",
                opacity=0.8,
            )

        # Y-axis range
        if y_range:
            fig.update_yaxes(range=[y_range[0], y_range[1]], row=i, col=1)

    apply_template(fig)
    fig.update_layout(
        height=300 * rows + 60,
        title=dict(
            text="📈  TENDENCIA DE VARIABLES DE PROCESO — Serie Temporal",
            font=dict(family="Rajdhani", size=15, color="#e8f0fe"),
            x=0.01,
        ),
        hovermode="x unified",
        showlegend=True,
        dragmode="zoom",
    )
    for i in range(1, rows + 1):
        fig.update_xaxes(
            showspikes=True, spikethickness=1, spikecolor="#00d4ff",
            spikedash="dot", spikemode="across",
            row=i, col=1,
        )
        fig.update_yaxes(
            showspikes=True, spikethickness=1, spikecolor="#00d4ff",
            spikedash="dot", spikemode="across",
            row=i, col=1,
        )
    return fig


def build_histogram(
    df: pd.DataFrame,
    col: str,
    lsl: float = None,
    usl: float = None,
    target: float = None,
    nbins: int = 50,
) -> go.Figure:
    """Histogram with KDE overlay and spec limits."""
    s = df[col].dropna()
    color = COLORS_PRIMARY[0]

    fig = go.Figure()

    # Histogram
    fig.add_trace(go.Histogram(
        x=s, nbinsx=nbins, name="Frecuencia",
        marker=dict(color=color, opacity=0.7,
                    line=dict(color="rgba(0,212,255,0.3)", width=0.5)),
        hovertemplate="Valor: %{x}<br>Conteo: %{y}<extra></extra>",
        histnorm="probability density",
    ))

    # KDE overlay
    if len(s) > 5:
        try:
            kde = stats.gaussian_kde(s)
            x_range = np.linspace(s.min(), s.max(), 400)
            fig.add_trace(go.Scatter(
                x=x_range, y=kde(x_range),
                mode="lines", name="KDE",
                line=dict(color="#ffd740", width=2),
                hovertemplate="x: %{x:,.4f}<br>Densidad: %{y:,.6f}<extra></extra>",
            ))
            # Normal fit
            mu, sigma = s.mean(), s.std()
            norm_y = stats.norm.pdf(x_range, mu, sigma)
            fig.add_trace(go.Scatter(
                x=x_range, y=norm_y,
                mode="lines", name="Normal ajustada",
                line=dict(color="#ea80fc", width=2, dash="dash"),
                hovertemplate="x: %{x:,.4f}<br>Normal: %{y:,.6f}<extra></extra>",
            ))
        except Exception:
            pass

    # Spec lines
    for val, label, clr in [
        (lsl, "LSL", "#ff4444"),
        (usl, "USL", "#ff4444"),
        (target, "TGT", "#ffd740"),
    ]:
        if val is not None:
            fig.add_vline(x=val, line=dict(color=clr, width=2, dash="longdash"),
                          annotation_text=label,
                          annotation_font_color=clr,
                          annotation_font_size=10)

    # Mean
    fig.add_vline(x=s.mean(), line=dict(color=color, width=1.5, dash="dot"),
                  annotation_text=f"μ={s.mean():,.4f}",
                  annotation_font_color=color,
                  annotation_font_size=9)

    apply_template(fig)
    fig.update_layout(
        title=f"📊  Distribución — {col}",
        xaxis_title=col,
        yaxis_title="Densidad de Probabilidad",
        height=420,
        bargap=0.02,
    )
    return fig


def build_box_violin(df: pd.DataFrame, var_cols: List[str]) -> go.Figure:
    """Combined box + violin plot for multiple variables."""
    fig = go.Figure()
    for i, col in enumerate(var_cols):
        color = COLORS_PRIMARY[i % len(COLORS_PRIMARY)]
        s = df[col].dropna()
        fig.add_trace(go.Violin(
            y=s, x=[col] * len(s),
            name=col,
            line_color=color,
            fillcolor=f"rgba({int(color[1:3],16)},{int(color[3:5],16)},{int(color[5:7],16)},0.15)",
            meanline_visible=True,
            box_visible=True,
            points="outliers",
            pointpos=0,
            marker=dict(color=color, size=3, opacity=0.6),
            hovertemplate=f"<b>{col}</b><br>Valor: %{{y:,.4f}}<extra></extra>",
        ))
    apply_template(fig)
    fig.update_layout(
        title="🎻  Box-Violin Plot — Distribución por Variable",
        yaxis_title="Valor",
        height=480,
        violingap=0.15, violingroupgap=0.05,
    )
    return fig


def build_correlation_heatmap(df: pd.DataFrame, var_cols: List[str]) -> go.Figure:
    """Correlation matrix heatmap."""
    if len(var_cols) < 2:
        return go.Figure()
    corr = df[var_cols].corr()
    fig = go.Figure(go.Heatmap(
        z=corr.values,
        x=corr.columns.tolist(),
        y=corr.index.tolist(),
        colorscale=[[0, "#ff4444"], [0.5, "#0f1520"], [1, "#00d4ff"]],
        zmid=0,
        text=[[f"{v:.3f}" for v in row] for row in corr.values],
        texttemplate="%{text}",
        textfont_size=10,
        hovertemplate="<b>%{x}</b> vs <b>%{y}</b><br>r = %{z:.4f}<extra></extra>",
        colorbar=dict(
            tickfont=dict(color="#a0b4d6"),
            title=dict(text="Pearson r", font=dict(color="#a0b4d6")),
        ),
    ))
    apply_template(fig)
    fig.update_layout(
        title="🔗  Matriz de Correlación",
        height=max(400, 60 * len(var_cols) + 80),
        xaxis=dict(tickangle=-35),
    )
    return fig


def build_scatter_matrix(df: pd.DataFrame, var_cols: List[str],
                          color_col: str = None) -> go.Figure:
    """Scatter matrix / pairs plot."""
    cols = var_cols[:6]  # limit for performance
    kw = dict(
        dimensions=[dict(label=c, values=df[c]) for c in cols],
        showupperhalf=False,
        diagonal_visible=True,
    )
    if color_col and color_col in df.columns:
        kw["color"] = df[color_col]
    fig = go.Figure(go.Splom(
        **kw,
        marker=dict(size=4, opacity=0.5, colorscale="Plasma",
                    line=dict(width=0)),
        hovertemplate=None,
    ))
    apply_template(fig)
    fig.update_layout(
        title="🔵  Scatter Matrix",
        height=600,
        dragmode="select",
    )
    return fig


def build_control_chart(df: pd.DataFrame, time_col: str, col: str,
                         sigma_level: int = 3) -> go.Figure:
    """X-bar / Individuals Control Chart (Shewhart)."""
    s = df[[time_col, col]].dropna(subset=[col]).copy()
    s = s.sort_values(time_col)

    x_vals = s[time_col].values
    y_vals = s[col].values.astype(float)
    mean   = y_vals.mean()
    std    = y_vals.std(ddof=1)
    ucl    = mean + sigma_level * std
    lcl    = mean - sigma_level * std
    u2     = mean + 2 * std
    l2     = mean - 2 * std

    # Identify out-of-control points
    ooc = (y_vals > ucl) | (y_vals < lcl)

    fig = go.Figure()

    # Fill UCL-LCL band
    fig.add_trace(go.Scatter(
        x=np.concatenate([x_vals, x_vals[::-1]]),
        y=np.concatenate([np.full_like(y_vals, ucl), np.full_like(y_vals, lcl)[::-1]]),
        fill="toself",
        fillcolor="rgba(0,212,255,0.04)",
        line=dict(width=0),
        name="Zona control",
        hoverinfo="skip",
    ))

    # Main series
    fig.add_trace(go.Scatter(
        x=x_vals, y=y_vals, mode="lines+markers",
        name=col,
        line=dict(color="#00d4ff", width=1.5),
        marker=dict(color=np.where(ooc, "#ff4444", "#00d4ff"),
                    size=np.where(ooc, 8, 5).tolist(),
                    symbol=np.where(ooc, "x", "circle").tolist()),
        hovertemplate="%{x|%Y-%m-%d %H:%M:%S}<br>Valor: %{y:,.4f}<extra></extra>",
    ))

    # OOC markers
    if ooc.any():
        fig.add_trace(go.Scatter(
            x=x_vals[ooc], y=y_vals[ooc], mode="markers",
            name="Fuera de control ⚠",
            marker=dict(color="#ff4444", size=10, symbol="x-open",
                        line=dict(color="#ff4444", width=2)),
            hovertemplate="⚠ OOC: %{y:,.4f}<extra></extra>",
        ))

    for val, label, clr, dash in [
        (ucl, f"UCL={ucl:,.3f}", "#ff4444", "longdash"),
        (lcl, f"LCL={lcl:,.3f}", "#ff4444", "longdash"),
        (mean, f"CL={mean:,.3f}", "#ffd740", "dot"),
        (u2, "±2σ", "#ff6b35", "dash"),
        (l2, "±2σ", "#ff6b35", "dash"),
    ]:
        fig.add_hline(y=val, line=dict(color=clr, width=1.4, dash=dash),
                      annotation_text=label,
                      annotation_font_color=clr,
                      annotation_font_size=9,
                      annotation_position="right",
                      opacity=0.8)

    apply_template(fig)
    fig.update_layout(
        title=f"📉  Carta de Control Individuales — {col}  (±{sigma_level}σ)",
        yaxis_title=col,
        height=450,
    )
    return fig


def build_moving_range_chart(df: pd.DataFrame, time_col: str,
                              col: str) -> go.Figure:
    """Moving Range (MR) chart."""
    s = df[[time_col, col]].dropna(subset=[col]).sort_values(time_col)
    y = s[col].values.astype(float)
    mr = np.abs(np.diff(y))
    x  = s[time_col].values[1:]

    mr_bar = mr.mean()
    d2     = 1.128
    sigma  = mr_bar / d2
    ucl_mr = 3.267 * mr_bar

    ooc = mr > ucl_mr

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=mr, mode="lines+markers",
        name="Moving Range",
        line=dict(color="#ff6b35", width=1.5),
        marker=dict(color=np.where(ooc, "#ff4444", "#ff6b35"),
                    size=np.where(ooc, 8, 4).tolist()),
        hovertemplate="%{x|%Y-%m-%d %H:%M:%S}<br>MR: %{y:,.4f}<extra></extra>",
    ))
    for val, label, clr, dash in [
        (ucl_mr, f"UCL={ucl_mr:,.3f}", "#ff4444", "longdash"),
        (mr_bar, f"MR̄={mr_bar:,.3f}", "#ffd740", "dot"),
        (0, "LCL=0", "#ff4444", "longdash"),
    ]:
        fig.add_hline(y=val, line=dict(color=clr, width=1.2, dash=dash),
                      annotation_text=label, annotation_font_color=clr,
                      annotation_font_size=9, annotation_position="right",
                      opacity=0.8)

    apply_template(fig)
    fig.update_layout(
        title=f"📐  Carta de Rango Móvil — {col}",
        yaxis_title="Rango Móvil", height=340,
    )
    return fig


def build_autocorrelation(df: pd.DataFrame, col: str,
                           max_lags: int = 50) -> go.Figure:
    """Autocorrelation (ACF) chart."""
    s = df[col].dropna().values.astype(float)
    n = len(s)
    lags = min(max_lags, n // 4)
    acf_vals = [1.0]
    mean_s = s.mean()
    var_s  = np.var(s)
    if var_s == 0:
        return go.Figure()
    for lag in range(1, lags + 1):
        cov = np.mean((s[:-lag] - mean_s) * (s[lag:] - mean_s))
        acf_vals.append(cov / var_s)

    conf = 1.96 / np.sqrt(n)
    x_lags = list(range(len(acf_vals)))

    fig = go.Figure()
    # Bars
    for xi, yi in zip(x_lags, acf_vals):
        fig.add_trace(go.Bar(
            x=[xi], y=[yi],
            marker_color="#00d4ff" if abs(yi) <= conf else "#ff4444",
            showlegend=False,
            hovertemplate=f"Lag: {xi}<br>ACF: {yi:.4f}<extra></extra>",
        ))
    # Confidence bounds
    for sign in [1, -1]:
        fig.add_hline(y=sign * conf, line=dict(color="#ffd740", width=1.5,
                                                dash="dash"),
                      annotation_text=f"±{conf:.4f}",
                      annotation_font_color="#ffd740",
                      annotation_font_size=8)

    apply_template(fig)
    fig.update_layout(
        title=f"📡  Autocorrelación (ACF) — {col}",
        xaxis_title="Lag", yaxis_title="ACF",
        barmode="overlay", height=370,
    )
    return fig


def build_cumulative_sum(df: pd.DataFrame, time_col: str,
                          col: str, target: float = None) -> go.Figure:
    """CUSUM chart."""
    s = df[[time_col, col]].dropna(subset=[col]).sort_values(time_col)
    y   = s[col].values.astype(float)
    ref = target if target is not None else y.mean()
    k   = y.std() / 2  # allowance
    cusum_pos = np.zeros(len(y))
    cusum_neg = np.zeros(len(y))
    for i in range(1, len(y)):
        cusum_pos[i] = max(0, cusum_pos[i-1] + (y[i] - ref) - k)
        cusum_neg[i] = min(0, cusum_neg[i-1] + (y[i] - ref) + k)

    h = 4 * y.std()  # decision interval

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=s[time_col], y=cusum_pos, mode="lines",
        name="CUSUM+", line=dict(color="#00e676", width=2),
        hovertemplate="%{x|%Y-%m-%d %H:%M:%S}<br>C+: %{y:,.4f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=s[time_col], y=cusum_neg, mode="lines",
        name="CUSUM−", line=dict(color="#ff6b35", width=2),
        hovertemplate="%{x|%Y-%m-%d %H:%M:%S}<br>C−: %{y:,.4f}<extra></extra>",
    ))
    for val, lbl, clr in [(h, f"H={h:.3f}", "#ff4444"),
                           (-h, f"-H={-h:.3f}", "#ff4444")]:
        fig.add_hline(y=val, line=dict(color=clr, width=1.5, dash="longdash"),
                      annotation_text=lbl, annotation_font_color=clr,
                      annotation_font_size=9, annotation_position="right")
    apply_template(fig)
    fig.update_layout(title=f"📈  CUSUM Chart — {col}", height=380)
    return fig


def build_ewma_chart(df: pd.DataFrame, time_col: str,
                     col: str, lam: float = 0.2) -> go.Figure:
    """EWMA (Exponentially Weighted Moving Average) chart."""
    s = df[[time_col, col]].dropna(subset=[col]).sort_values(time_col)
    y    = s[col].values.astype(float)
    mu   = y.mean()
    sig  = y.std(ddof=1)
    ewma = np.zeros(len(y))
    ewma[0] = y[0]
    for i in range(1, len(y)):
        ewma[i] = lam * y[i] + (1 - lam) * ewma[i - 1]
    L = 3.0
    ucl = mu + L * sig * np.sqrt(lam / (2 - lam))
    lcl = mu - L * sig * np.sqrt(lam / (2 - lam))

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=s[time_col], y=y, mode="lines",
        name=col, line=dict(color="rgba(0,212,255,0.35)", width=1),
        hovertemplate="%{x|%Y-%m-%d %H:%M:%S}<br>Raw: %{y:,.4f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=s[time_col], y=ewma, mode="lines",
        name=f"EWMA (λ={lam})", line=dict(color="#00d4ff", width=2.5),
        hovertemplate="%{x|%Y-%m-%d %H:%M:%S}<br>EWMA: %{y:,.4f}<extra></extra>",
    ))
    for val, lbl, clr, dash in [
        (ucl, f"UCL={ucl:.3f}", "#ff4444", "longdash"),
        (lcl, f"LCL={lcl:.3f}", "#ff4444", "longdash"),
        (mu,  f"CL={mu:.3f}",   "#ffd740", "dot"),
    ]:
        fig.add_hline(y=val, line=dict(color=clr, width=1.3, dash=dash),
                      annotation_text=lbl, annotation_font_color=clr,
                      annotation_font_size=9, annotation_position="right",
                      opacity=0.85)
    apply_template(fig)
    fig.update_layout(
        title=f"〰️  EWMA Chart — {col}  (λ={lam})",
        height=400,
    )
    return fig


def build_qq_plot(df: pd.DataFrame, col: str) -> go.Figure:
    """Q-Q Normal plot."""
    s = df[col].dropna().values.astype(float)
    (osm, osr), (slope, intercept, r) = stats.probplot(s, dist="norm")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=osm, y=osr, mode="markers",
        name="Datos",
        marker=dict(color="#00d4ff", size=5, opacity=0.7),
        hovertemplate="Teórico: %{x:.3f}<br>Muestra: %{y:.4f}<extra></extra>",
    ))
    x_line = np.array([osm.min(), osm.max()])
    fig.add_trace(go.Scatter(
        x=x_line, y=slope * x_line + intercept,
        mode="lines", name=f"Normal (r={r:.4f})",
        line=dict(color="#ffd740", width=2, dash="dash"),
    ))
    apply_template(fig)
    fig.update_layout(
        title=f"📐  Q-Q Normal — {col}",
        xaxis_title="Cuantiles Teóricos",
        yaxis_title="Cuantiles Observados",
        height=420,
    )
    return fig


def build_capability_chart(stats_dict: Dict, col: str) -> go.Figure:
    """Process Capability graphical overview (Normal curve + spec limits)."""
    if not stats_dict or np.isnan(stats_dict.get("std", np.nan)):
        return go.Figure()
    mean = stats_dict["mean"]
    std  = stats_dict["std"]
    lsl  = stats_dict.get("lsl")
    usl  = stats_dict.get("usl")

    x = np.linspace(mean - 5 * std, mean + 5 * std, 500)
    y = stats.norm.pdf(x, mean, std)

    fig = go.Figure()

    # Fill out-of-spec zones
    if lsl is not None:
        mask = x <= lsl
        fig.add_trace(go.Scatter(
            x=np.concatenate([[lsl], x[mask], [lsl]]),
            y=np.concatenate([[0], y[mask], [0]]),
            fill="toself", fillcolor="rgba(255,68,68,0.25)",
            line=dict(width=0), name="Fuera spec (LSL)", hoverinfo="skip",
        ))
    if usl is not None:
        mask = x >= usl
        fig.add_trace(go.Scatter(
            x=np.concatenate([[usl], x[mask], [usl]]),
            y=np.concatenate([[0], y[mask], [0]]),
            fill="toself", fillcolor="rgba(255,68,68,0.25)",
            line=dict(width=0), name="Fuera spec (USL)", hoverinfo="skip",
        ))

    # Normal curve
    fig.add_trace(go.Scatter(
        x=x, y=y, mode="lines", name="Distribución Normal",
        line=dict(color="#00d4ff", width=2.5),
        hovertemplate="x: %{x:.4f}<br>f(x): %{y:.6f}<extra></extra>",
    ))

    for val, lbl, clr, dash in [
        (mean,  f"μ={mean:.4f}", "#ffd740", "dot"),
        (lsl,   "LSL",          "#ff4444", "longdash"),
        (usl,   "USL",          "#ff4444", "longdash"),
    ]:
        if val is not None:
            fig.add_vline(x=val, line=dict(color=clr, width=1.8, dash=dash),
                          annotation_text=lbl,
                          annotation_font_color=clr, annotation_font_size=10)

    # Sigma lines
    for n_sig in [1, 2, 3]:
        for sign in [1, -1]:
            fig.add_vline(
                x=mean + sign * n_sig * std,
                line=dict(color="rgba(160,180,214,0.25)", width=1, dash="dot"),
                annotation_text=f"{'+' if sign > 0 else '-'}{n_sig}σ",
                annotation_font_color="#7a8fb5", annotation_font_size=8,
                annotation_position="top" if sign > 0 else "top left",
            )

    apply_template(fig)
    fig.update_layout(
        title=f"⚙️  Capacidad de Proceso — {col}",
        xaxis_title=col, yaxis_title="Densidad",
        height=400,
    )
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# 6. STATISTICS TABLE RENDERER
# ─────────────────────────────────────────────────────────────────────────────

def render_stats_table(st_dict: Dict, col: str) -> None:
    """Render the stats HTML table inside Streamlit."""
    if not st_dict:
        st.warning("Sin datos suficientes para calcular estadísticas.")
        return

    cpk = st_dict.get("cpk", float("nan"))
    cp  = st_dict.get("cp",  float("nan"))

    def cpk_class(v):
        if np.isnan(v): return ""
        return "ok" if v >= 1.67 else ("warn" if v >= 1.33 else "bad")

    rows_basic = [
        ("N (muestras)",      fmt(st_dict.get("n")),          "—",    ""),
        ("Mínimo",            fmt(st_dict.get("minimum")),     "—",    ""),
        ("Máximo",            fmt(st_dict.get("maximum")),     "—",    ""),
        ("Rango",             fmt(st_dict.get("range")),       "—",    ""),
        ("Media (μ)",         fmt(st_dict.get("mean")),        "—",    ""),
        ("Mediana",           fmt(st_dict.get("median")),      "—",    ""),
        ("Desv. Estándar (σ)",fmt(st_dict.get("std")),         "—",    ""),
        ("Varianza (σ²)",     fmt(st_dict.get("variance")),    "—",    ""),
        ("CV (%)",            fmt(st_dict.get("cv"), 2),       "< 10%","warn" if (st_dict.get("cv") or 0) > 10 else "ok"),
        ("Q1 (25%)",          fmt(st_dict.get("q1")),          "—",    ""),
        ("Q3 (75%)",          fmt(st_dict.get("q3")),          "—",    ""),
        ("IQR",               fmt(st_dict.get("iqr")),         "—",    ""),
        ("P5",                fmt(st_dict.get("p5")),          "—",    ""),
        ("P10",               fmt(st_dict.get("p10")),         "—",    ""),
        ("P90",               fmt(st_dict.get("p90")),         "—",    ""),
        ("P95",               fmt(st_dict.get("p95")),         "—",    ""),
        ("Asimetría (Skew)",  fmt(st_dict.get("skewness"), 4),
         "≈0 normal",   "warn" if abs(st_dict.get("skewness") or 0) > 1 else "ok"),
        ("Curtosis",          fmt(st_dict.get("kurtosis"), 4),
         "≈3 normal",   "warn" if abs(st_dict.get("kurtosis") or 0) > 2 else "ok"),
        ("Datos faltantes",   fmt(st_dict.get("n_missing")),   "= 0",
         "bad" if (st_dict.get("n_missing") or 0) > 0 else "ok"),
        ("Outliers (IQR)",    fmt(st_dict.get("n_outliers")),  "= 0",
         "warn" if (st_dict.get("n_outliers") or 0) > 0 else "ok"),
    ]

    rows_proc = [
        ("Cp",            fmt(st_dict.get("cp"), 4),     "≥ 1.67", cpk_class(cp)),
        ("Cpk",           fmt(st_dict.get("cpk"), 4),    "≥ 1.67", cpk_class(cpk)),
        ("Cpl",           fmt(st_dict.get("cpl"), 4),    "≥ 1.33", ""),
        ("Cpu",           fmt(st_dict.get("cpu"), 4),    "≥ 1.33", ""),
        ("Pp",            fmt(st_dict.get("pp"), 4),     "≥ 1.67", ""),
        ("Ppk",           fmt(st_dict.get("ppk"), 4),    "≥ 1.67", cpk_class(st_dict.get("ppk", float("nan")))),
        ("Nivel Sigma",   fmt(st_dict.get("sigma_level"), 2), "≥ 5.0", ""),
        ("PPM Estimado",  fmt(st_dict.get("ppm_est"), 1), "< 233",  ""),
        ("Yield (%)",     fmt(st_dict.get("yield_pct"), 4), "> 99.977%",""),
        ("LSL",           fmt(st_dict.get("lsl")),        "Especificado",""),
        ("USL",           fmt(st_dict.get("usl")),        "Especificado",""),
    ]

    rows_norm = [
        ("Shapiro-Wilk W",  fmt(st_dict.get("sw_stat"), 5), "—", ""),
        ("Shapiro-Wilk p",  fmt(st_dict.get("sw_p"), 5),
         "> 0.05",   "ok" if (st_dict.get("sw_p") or 0) > 0.05 else "warn"),
        ("Anderson-Darling",fmt(st_dict.get("ad_stat"), 5), "—", ""),
    ]

    def table_html(rows, title_label):
        rows_str = ""
        for name, val, ref, cls in rows:
            rows_str += (
                f"<tr><td>{name}</td>"
                f"<td class='val {cls}'>{val}</td>"
                f"<td style='color:#7a8fb5;font-size:0.75rem'>{ref}</td></tr>"
            )
        return f"""
        <div style="margin-bottom:1.2rem;">
          <div style="font-family:'Rajdhani',sans-serif;font-size:0.75rem;
                      letter-spacing:2px;color:#7a8fb5;text-transform:uppercase;
                      margin-bottom:6px;">{title_label}</div>
          <table class='stats-table'>
            <thead><tr>
              <th>Indicador</th><th>Valor</th><th>Referencia</th>
            </tr></thead>
            <tbody>{rows_str}</tbody>
          </table>
        </div>"""

    html = (
        table_html(rows_basic, "📊 Estadísticas Descriptivas") +
        table_html(rows_proc,  "⚙️ Capacidad de Proceso (SPC)") +
        table_html(rows_norm,  "🔬 Prueba de Normalidad")
    )
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# 7. MULTI-VARIABLE EXPORT HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def export_stats_excel(df: pd.DataFrame, var_cols: List[str],
                        lsl_d: Dict, usl_d: Dict) -> bytes:
    """Generate Excel with statistics for all variables."""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        all_stats = []
        for col in var_cols:
            sd = compute_statistics(df[col],
                                     lsl=lsl_d.get(col),
                                     usl=usl_d.get(col))
            sd["Variable"] = col
            all_stats.append(sd)
        stats_df = pd.DataFrame(all_stats).set_index("Variable")
        stats_df.to_excel(writer, sheet_name="Estadísticas")
        df[var_cols].describe().T.to_excel(writer, sheet_name="Describe")
        corr = df[var_cols].corr()
        corr.to_excel(writer, sheet_name="Correlación")
    return out.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# 8. HEADER COMPONENT
# ─────────────────────────────────────────────────────────────────────────────

def render_header() -> None:
    st.markdown("""
    <div class="header-banner">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;">
        <div>
          <div class="header-title">⚙ PROCESS DATA ANALYZER</div>
          <div class="header-sub">INDUSTRIAL ENGINEERING ANALYTICS PLATFORM  //  v2.0.0</div>
          <div style="margin-top:10px;">
            <span class="header-tag">SPC</span>
            <span class="header-tag">TENDENCIAS</span>
            <span class="header-tag">CAPACIDAD</span>
            <span class="header-tag">CORRELACIÓN</span>
            <span class="header-tag">CONTROL CHARTS</span>
            <span class="header-tag">CUSUM / EWMA</span>
          </div>
        </div>
        <div style="text-align:right;font-family:'Share Tech Mono',monospace;
                    font-size:0.7rem;color:#7a8fb5;line-height:2;">
          <div>MÓDULOS ACTIVOS: 12</div>
          <div>MOTOR: PLOTLY + SCIPY</div>
          <div>BUILD: STREAMLIT</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# 9. SECTION HELPER
# ─────────────────────────────────────────────────────────────────────────────

def section(icon: str, title: str, pill: str = "") -> None:
    st.markdown(f"""
    <div class="section-header">
      <span class="section-icon">{icon}</span>
      <span class="section-title">{title}</span>
      {"<span class='section-pill'>" + pill + "</span>" if pill else ""}
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# 10. MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    render_header()

    # ── SIDEBAR ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("""
        <div style="font-family:'Rajdhani',sans-serif;font-size:1.0rem;
                    font-weight:700;letter-spacing:2px;color:#00d4ff;
                    text-transform:uppercase;margin-bottom:1rem;
                    padding-bottom:8px;border-bottom:1px solid #1e2d4a;">
          ⚙ Panel de Control
        </div>""", unsafe_allow_html=True)

        # File upload
        uploaded = st.file_uploader(
            "📂 Cargar archivo de datos",
            type=["csv", "xlsx", "xls"],
            help="CSV, XLSX o XLS. Máximo 200 MB.",
        )

        if uploaded is None:
            st.markdown("""
            <div class="info-box">
              <b>Instrucciones:</b><br>
              1. Carga tu archivo CSV o Excel.<br>
              2. Selecciona la columna de fecha/hora.<br>
              3. Elige las variables a analizar.<br>
              4. Ajusta el intervalo temporal.<br>
              5. Explora los módulos en las pestañas.
            </div>""", unsafe_allow_html=True)
            st.stop()

        raw_bytes = uploaded.read()
        df_raw = load_file(raw_bytes, uploaded.name)

        if df_raw is None or df_raw.empty:
            st.error("❌ No se pudo leer el archivo. Verifica el formato.")
            st.stop()

        st.markdown(f"""
        <div class="info-box">
          📄 <b>{uploaded.name}</b><br>
          Filas: {df_raw.shape[0]:,} | Columnas: {df_raw.shape[1]}
        </div>""", unsafe_allow_html=True)

        # Datetime column
        dt_cols   = detect_datetime_columns(df_raw)
        num_cols  = detect_numeric_columns(df_raw)
        all_cols  = list(df_raw.columns)

        st.markdown("---")
        st.markdown("**📅 Columna de Fecha/Hora**")
        dt_col_choice = st.selectbox(
            "Seleccionar columna temporal",
            options=dt_cols + [c for c in all_cols if c not in dt_cols],
            index=0 if dt_cols else 0,
            label_visibility="collapsed",
        )

        # Parse datetime
        df_raw["__time__"] = parse_datetime_column(df_raw[dt_col_choice])
        df_valid = df_raw.dropna(subset=["__time__"]).copy()
        df_valid = df_valid.sort_values("__time__").reset_index(drop=True)
        TIME_COL = "__time__"

        if df_valid.empty:
            st.error("❌ La columna seleccionada no contiene fechas válidas.")
            st.stop()

        t_min = df_valid[TIME_COL].min()
        t_max = df_valid[TIME_COL].max()

        st.markdown("**📊 Variables de proceso**")
        numeric_cols_avail = [c for c in num_cols if c != dt_col_choice]
        if not numeric_cols_avail:
            st.error("❌ No se encontraron columnas numéricas.")
            st.stop()

        selected_vars = st.multiselect(
            "Seleccionar variables",
            options=numeric_cols_avail,
            default=numeric_cols_avail[:min(4, len(numeric_cols_avail))],
            label_visibility="collapsed",
        )
        if not selected_vars:
            st.warning("⚠ Selecciona al menos una variable.")
            st.stop()

        # ── Date/time range ───────────────────────────────────────────────
        st.markdown("---")
        st.markdown("**🗓 Intervalo de Fecha**")
        d_start = st.date_input("Desde", value=t_min.date(),
                                 min_value=t_min.date(), max_value=t_max.date(),
                                 key="date_start")
        d_end   = st.date_input("Hasta", value=t_max.date(),
                                 min_value=t_min.date(), max_value=t_max.date(),
                                 key="date_end")

        st.markdown("**⏰ Intervalo de Hora**")
        col1, col2 = st.columns(2)
        with col1:
            h_start = st.number_input("H inicio", 0, 23, 0, key="h_start")
            m_start = st.number_input("M inicio", 0, 59, 0, key="m_start")
        with col2:
            h_end   = st.number_input("H fin",    0, 23, 23, key="h_end")
            m_end   = st.number_input("M fin",    0, 59, 59, key="m_end")

        dt_start = datetime.combine(d_start, datetime.min.time()).replace(
            hour=h_start, minute=m_start)
        dt_end   = datetime.combine(d_end, datetime.min.time()).replace(
            hour=h_end, minute=m_end, second=59)

        mask = (df_valid[TIME_COL] >= dt_start) & (df_valid[TIME_COL] <= dt_end)
        df = df_valid[mask].copy()

        if df.empty:
            st.error("❌ Sin datos en el intervalo seleccionado.")
            st.stop()

        st.markdown(f"""
        <div class="info-box">
          ⏱ <b>Intervalo activo</b><br>
          Desde: {dt_start:%Y-%m-%d %H:%M}<br>
          Hasta: {dt_end:%Y-%m-%d %H:%M}<br>
          Registros: <b>{len(df):,}</b>
        </div>""", unsafe_allow_html=True)

        # ── Chart options ─────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("**🎛 Opciones de Gráfica**")
        show_smooth   = st.checkbox("Suavizado (Savitzky-Golay)", value=False)
        smooth_window = st.slider("Ventana suavizado", 5, 51, 11, 2,
                                  disabled=not show_smooth)
        show_markers  = st.checkbox("Mostrar marcadores", value=False)
        show_mean_ln  = st.checkbox("Línea de media", value=True)
        sigma_ctrl    = st.selectbox("Sigma cartas control", [2, 3], index=1)

        st.markdown("**📐 Límites de Especificación (por variable)**")
        lsl_dict: Dict[str, float] = {}
        usl_dict: Dict[str, float] = {}
        tgt_dict: Dict[str, float] = {}
        with st.expander("Definir LSL / USL / Target", expanded=False):
            for var in selected_vars:
                st.markdown(f"**{var}**")
                c1, c2, c3 = st.columns(3)
                lv = c1.text_input("LSL", value="", key=f"lsl_{var}")
                uv = c2.text_input("USL", value="", key=f"usl_{var}")
                tv = c3.text_input("TGT", value="", key=f"tgt_{var}")
                try: lsl_dict[var] = float(lv)
                except ValueError: pass
                try: usl_dict[var] = float(uv)
                except ValueError: pass
                try: tgt_dict[var] = float(tv)
                except ValueError: pass

        st.markdown("---")
        st.markdown("**📤 Exportar Estadísticas**")
        if st.button("📊 Generar Excel de Estadísticas"):
            xlsx_bytes = export_stats_excel(df, selected_vars, lsl_dict, usl_dict)
            st.download_button(
                "⬇ Descargar Excel",
                data=xlsx_bytes,
                file_name=f"estadisticas_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ── END SIDEBAR ──────────────────────────────────────────────────────────

    # ── KPI ROW ───────────────────────────────────────────────────────────────
    n_vars = len(selected_vars)
    dur_h  = (dt_end - dt_start).total_seconds() / 3600
    n_recs = len(df)
    freq_s = (df[TIME_COL].diff().dt.total_seconds().median()
              if n_recs > 1 else float("nan"))

    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">Registros</div>
        <div class="kpi-value">{n_recs:,}</div>
        <div class="kpi-unit">en intervalo</div>
      </div>
      <div class="kpi-card secondary">
        <div class="kpi-label">Variables</div>
        <div class="kpi-value">{n_vars}</div>
        <div class="kpi-unit">seleccionadas</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Duración</div>
        <div class="kpi-value">{dur_h:,.1f}</div>
        <div class="kpi-unit">horas</div>
      </div>
      <div class="kpi-card success">
        <div class="kpi-label">Frec. Muestreo</div>
        <div class="kpi-value">{"—" if np.isnan(freq_s) else f"{freq_s:.0f}"}</div>
        <div class="kpi-unit">segundos</div>
      </div>
      <div class="kpi-card {'danger' if df[selected_vars].isna().any().any() else 'success'}">
        <div class="kpi-label">Datos Faltantes</div>
        <div class="kpi-value">{int(df[selected_vars].isna().sum().sum())}</div>
        <div class="kpi-unit">valores nulos</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Inicio</div>
        <div class="kpi-value" style="font-size:0.85rem">{df[TIME_COL].min():%d/%m/%y %H:%M}</div>
        <div class="kpi-unit">timestamp</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Fin</div>
        <div class="kpi-value" style="font-size:0.85rem">{df[TIME_COL].max():%d/%m/%y %H:%M}</div>
        <div class="kpi-unit">timestamp</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── TABS ──────────────────────────────────────────────────────────────────
    tabs = st.tabs([
        "📈 Tendencia",
        "📊 Distribución",
        "📉 Control SPC",
        "🔗 Correlación",
        "⚙️ Capacidad",
        "📐 Avanzado",
        "🗂 Datos Crudos",
    ])

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 1 — TREND
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[0]:
        section("📈", "TENDENCIA TEMPORAL", f"{n_recs:,} pts")

        fig_trend = build_trend_chart(
            df=df, time_col=TIME_COL, var_cols=selected_vars,
            show_smooth=show_smooth, smooth_window=smooth_window,
            show_markers=show_markers,
            lsl=lsl_dict, usl=usl_dict, target=tgt_dict,
            show_mean_line=show_mean_ln,
        )
        st.plotly_chart(fig_trend, use_container_width=True,
                        config={"scrollZoom": True,
                                "displayModeBar": True,
                                "modeBarButtonsToAdd": ["drawline",
                                                        "drawopenpath",
                                                        "eraseshape"],
                                "toImageButtonOptions": {
                                    "format": "png", "width": 1400,
                                    "height": 700, "scale": 2,
                                    "filename": "tendencia_proceso",
                                }})

        # Per-variable mini stats under trend
        st.markdown("<br>", unsafe_allow_html=True)
        section("🔢", "RESUMEN RÁPIDO POR VARIABLE")

        cols_st = st.columns(min(n_vars, 4))
        for idx, var in enumerate(selected_vars):
            s_d = compute_statistics(df[var],
                                      lsl=lsl_dict.get(var),
                                      usl=usl_dict.get(var))
            clr = COLORS_PRIMARY[idx % len(COLORS_PRIMARY)]
            cpk_v = s_d.get("cpk", float("nan"))
            with cols_st[idx % len(cols_st)]:
                cpk_color = (
                    "#00e676" if (not np.isnan(cpk_v) and cpk_v >= 1.67)
                    else "#ffd740" if (not np.isnan(cpk_v) and cpk_v >= 1.33)
                    else "#ff4444" if not np.isnan(cpk_v)
                    else "#7a8fb5"
                )
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color:{clr};">
                  <div class="kpi-label" style="color:{clr}">{var}</div>
                  <div style="font-family:'Share Tech Mono';font-size:0.78rem;
                              color:#a0b4d6;line-height:1.9;margin-top:4px;">
                    MIN: <span style="color:{clr}">{fmt(s_d.get('minimum'))}</span><br>
                    MAX: <span style="color:{clr}">{fmt(s_d.get('maximum'))}</span><br>
                    μ:   <span style="color:{clr}">{fmt(s_d.get('mean'))}</span><br>
                    σ:   <span style="color:{clr}">{fmt(s_d.get('std'))}</span><br>
                    Cpk: <span style="color:{cpk_color};font-weight:700">
                           {fmt(cpk_v)}</span>
                  </div>
                </div>
                """, unsafe_allow_html=True)

        # Export trend
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns([1, 6])
        with c1:
            html_str = fig_to_html(fig_trend)
            st.download_button(
                "⬇ HTML Interactivo",
                data=html_str.encode(),
                file_name=f"tendencia_{datetime.now():%Y%m%d_%H%M%S}.html",
                mime="text/html",
            )
        with c2:
            png_b = fig_to_png(fig_trend)
            if png_b:
                st.download_button(
                    "⬇ PNG (alta resolución)",
                    data=png_b,
                    file_name=f"tendencia_{datetime.now():%Y%m%d_%H%M%S}.png",
                    mime="image/png",
                )

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 2 — DISTRIBUTION
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[1]:
        section("📊", "DISTRIBUCIÓN Y ESTADÍSTICAS")

        var_dist = st.selectbox(
            "Variable para análisis de distribución",
            options=selected_vars, key="dist_var",
        )
        nbins = st.slider("Número de bins", 10, 200, 50, key="hist_bins")

        col_h, col_s = st.columns([3, 2])
        with col_h:
            fig_h = build_histogram(
                df, var_dist,
                lsl=lsl_dict.get(var_dist),
                usl=usl_dict.get(var_dist),
                target=tgt_dict.get(var_dist),
                nbins=nbins,
            )
            st.plotly_chart(fig_h, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 1000, "height": 500,
                                "scale": 2, "filename": f"histograma_{var_dist}"}})
            c1, c2 = st.columns(2)
            with c1:
                html_str = fig_to_html(fig_h)
                st.download_button("⬇ HTML Histograma", data=html_str.encode(),
                                   file_name=f"hist_{var_dist}.html",
                                   mime="text/html")
            with c2:
                png_b = fig_to_png(fig_h)
                if png_b:
                    st.download_button("⬇ PNG Histograma", data=png_b,
                                       file_name=f"hist_{var_dist}.png",
                                       mime="image/png")

        with col_s:
            section("🔢", "INDICADORES ESTADÍSTICOS")
            s_dict = compute_statistics(
                df[var_dist],
                lsl=lsl_dict.get(var_dist),
                usl=usl_dict.get(var_dist),
            )
            render_stats_table(s_dict, var_dist)

        # Q-Q plot
        st.markdown("<br>", unsafe_allow_html=True)
        section("📐", "Q-Q PLOT — NORMALIDAD")
        fig_qq = build_qq_plot(df, var_dist)
        st.plotly_chart(fig_qq, use_container_width=True,
                        config={"toImageButtonOptions": {
                            "format": "png", "width": 900, "height": 450,
                            "scale": 2, "filename": f"qq_{var_dist}"}})
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇ HTML Q-Q", data=fig_to_html(fig_qq).encode(),
                               file_name=f"qq_{var_dist}.html", mime="text/html")
        with c2:
            png_b = fig_to_png(fig_qq)
            if png_b:
                st.download_button("⬇ PNG Q-Q", data=png_b,
                                   file_name=f"qq_{var_dist}.png", mime="image/png")

        # Box-Violin
        if len(selected_vars) > 1:
            st.markdown("<br>", unsafe_allow_html=True)
            section("🎻", "BOX-VIOLIN — TODAS LAS VARIABLES")
            fig_bv = build_box_violin(df, selected_vars)
            st.plotly_chart(fig_bv, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 1200, "height": 550,
                                "scale": 2, "filename": "boxviolin"}})
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ HTML Box-Violin",
                                   data=fig_to_html(fig_bv).encode(),
                                   file_name="boxviolin.html", mime="text/html")
            with c2:
                png_b = fig_to_png(fig_bv)
                if png_b:
                    st.download_button("⬇ PNG Box-Violin", data=png_b,
                                       file_name="boxviolin.png", mime="image/png")

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 3 — SPC CONTROL CHARTS
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[2]:
        section("📉", "CARTAS DE CONTROL SPC", f"±{sigma_ctrl}σ")

        var_ctrl = st.selectbox(
            "Variable para Carta de Control",
            options=selected_vars, key="ctrl_var",
        )

        # Individuals chart
        fig_ic = build_control_chart(df, TIME_COL, var_ctrl, sigma_ctrl)
        st.plotly_chart(fig_ic, use_container_width=True,
                        config={"scrollZoom": True,
                                "toImageButtonOptions": {
                                    "format": "png", "width": 1400, "height": 500,
                                    "scale": 2, "filename": f"control_{var_ctrl}"}})
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇ HTML Carta Control",
                               data=fig_to_html(fig_ic).encode(),
                               file_name=f"ctrl_{var_ctrl}.html", mime="text/html")
        with c2:
            png_b = fig_to_png(fig_ic)
            if png_b:
                st.download_button("⬇ PNG Carta Control", data=png_b,
                                   file_name=f"ctrl_{var_ctrl}.png", mime="image/png")

        # Moving Range
        st.markdown("<br>", unsafe_allow_html=True)
        section("📐", "CARTA DE RANGO MÓVIL (MR)")
        fig_mr = build_moving_range_chart(df, TIME_COL, var_ctrl)
        st.plotly_chart(fig_mr, use_container_width=True,
                        config={"toImageButtonOptions": {
                            "format": "png", "width": 1400, "height": 380,
                            "scale": 2, "filename": f"mr_{var_ctrl}"}})
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇ HTML MR Chart",
                               data=fig_to_html(fig_mr).encode(),
                               file_name=f"mr_{var_ctrl}.html", mime="text/html")
        with c2:
            png_b = fig_to_png(fig_mr)
            if png_b:
                st.download_button("⬇ PNG MR Chart", data=png_b,
                                   file_name=f"mr_{var_ctrl}.png", mime="image/png")

        # EWMA
        st.markdown("<br>", unsafe_allow_html=True)
        section("〰️", "CARTA EWMA")
        lam_val = st.slider("Lambda (λ) EWMA", 0.05, 0.50, 0.20, 0.05,
                             key="ewma_lam")
        fig_ew = build_ewma_chart(df, TIME_COL, var_ctrl, lam=lam_val)
        st.plotly_chart(fig_ew, use_container_width=True,
                        config={"toImageButtonOptions": {
                            "format": "png", "width": 1400, "height": 440,
                            "scale": 2, "filename": f"ewma_{var_ctrl}"}})
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇ HTML EWMA",
                               data=fig_to_html(fig_ew).encode(),
                               file_name=f"ewma_{var_ctrl}.html", mime="text/html")
        with c2:
            png_b = fig_to_png(fig_ew)
            if png_b:
                st.download_button("⬇ PNG EWMA", data=png_b,
                                   file_name=f"ewma_{var_ctrl}.png", mime="image/png")

        # CUSUM
        st.markdown("<br>", unsafe_allow_html=True)
        section("📈", "CARTA CUSUM")
        tgt_cusum = tgt_dict.get(var_ctrl)
        fig_cs = build_cumulative_sum(df, TIME_COL, var_ctrl, target=tgt_cusum)
        st.plotly_chart(fig_cs, use_container_width=True,
                        config={"toImageButtonOptions": {
                            "format": "png", "width": 1400, "height": 430,
                            "scale": 2, "filename": f"cusum_{var_ctrl}"}})
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("⬇ HTML CUSUM",
                               data=fig_to_html(fig_cs).encode(),
                               file_name=f"cusum_{var_ctrl}.html", mime="text/html")
        with c2:
            png_b = fig_to_png(fig_cs)
            if png_b:
                st.download_button("⬇ PNG CUSUM", data=png_b,
                                   file_name=f"cusum_{var_ctrl}.png", mime="image/png")

        # OOC summary
        st.markdown("<br>", unsafe_allow_html=True)
        section("⚠️", "RESUMEN PUNTOS FUERA DE CONTROL")
        s_ctrl = df[[TIME_COL, var_ctrl]].dropna(subset=[var_ctrl]).copy()
        y_c    = s_ctrl[var_ctrl].values.astype(float)
        mu_c   = y_c.mean()
        sig_c  = y_c.std(ddof=1)
        ucl_c  = mu_c + sigma_ctrl * sig_c
        lcl_c  = mu_c - sigma_ctrl * sig_c
        ooc_m  = (y_c > ucl_c) | (y_c < lcl_c)
        n_ooc  = ooc_m.sum()
        pct_ooc = n_ooc / len(y_c) * 100 if len(y_c) > 0 else 0

        st.markdown(f"""
        <div class="{'err-box' if n_ooc > 0 else 'info-box'}">
          <b>Variable: {var_ctrl}</b><br>
          Puntos totales: {len(y_c):,} |
          Fuera de control: <b>{n_ooc}</b> ({pct_ooc:.2f}%)<br>
          UCL = {ucl_c:.4f} | LCL = {lcl_c:.4f} | CL = {mu_c:.4f}
        </div>""", unsafe_allow_html=True)

        if n_ooc > 0:
            ooc_df = s_ctrl[ooc_m].copy()
            ooc_df.columns = ["Timestamp", var_ctrl]
            ooc_df["Estado"] = np.where(
                ooc_df[var_ctrl].values > ucl_c, "⬆ Por encima UCL",
                "⬇ Por debajo LCL"
            )
            st.dataframe(ooc_df.head(200), use_container_width=True,
                         hide_index=True)

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 4 — CORRELATION
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[3]:
        section("🔗", "ANÁLISIS DE CORRELACIÓN")

        if len(selected_vars) < 2:
            st.markdown('<div class="warn-box">Selecciona al menos 2 variables '
                        'para el análisis de correlación.</div>',
                        unsafe_allow_html=True)
        else:
            # Heatmap
            fig_corr = build_correlation_heatmap(df, selected_vars)
            st.plotly_chart(fig_corr, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 900, "height": 700,
                                "scale": 2, "filename": "correlacion"}})
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ HTML Correlación",
                                   data=fig_to_html(fig_corr).encode(),
                                   file_name="correlacion.html", mime="text/html")
            with c2:
                png_b = fig_to_png(fig_corr)
                if png_b:
                    st.download_button("⬇ PNG Correlación", data=png_b,
                                       file_name="correlacion.png", mime="image/png")

            # Scatter matrix
            st.markdown("<br>", unsafe_allow_html=True)
            section("🔵", "SCATTER MATRIX — PARES DE VARIABLES",
                    "máx. 6 vars")
            fig_sm = build_scatter_matrix(df, selected_vars)
            st.plotly_chart(fig_sm, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 1100, "height": 700,
                                "scale": 2, "filename": "scatter_matrix"}})
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ HTML Scatter Matrix",
                                   data=fig_to_html(fig_sm).encode(),
                                   file_name="scatter_matrix.html",
                                   mime="text/html")
            with c2:
                png_b = fig_to_png(fig_sm)
                if png_b:
                    st.download_button("⬇ PNG Scatter Matrix", data=png_b,
                                       file_name="scatter_matrix.png",
                                       mime="image/png")

            # Correlation table
            st.markdown("<br>", unsafe_allow_html=True)
            section("📋", "TABLA DE CORRELACIONES")
            corr_df = df[selected_vars].corr().round(4)
            st.dataframe(corr_df.style.background_gradient(
                cmap="RdYlGn_r", axis=None).format("{:.4f}"),
                use_container_width=True)

            # Top pairs
            st.markdown("<br>", unsafe_allow_html=True)
            section("⭐", "PARES MÁS CORRELACIONADOS")
            pairs = []
            for i, v1 in enumerate(selected_vars):
                for j, v2 in enumerate(selected_vars):
                    if j > i:
                        r = corr_df.loc[v1, v2]
                        pairs.append((v1, v2, r, abs(r)))
            pairs.sort(key=lambda x: x[3], reverse=True)
            pairs_df = pd.DataFrame(pairs,
                                    columns=["Variable 1", "Variable 2",
                                             "Pearson r", "|r|"])
            pairs_df["Interpretación"] = pairs_df["Pearson r"].apply(
                lambda r: (
                    "Muy alta +" if r >= 0.9 else
                    "Alta +" if r >= 0.7 else
                    "Moderada +" if r >= 0.4 else
                    "Débil +" if r >= 0.1 else
                    "Muy alta −" if r <= -0.9 else
                    "Alta −" if r <= -0.7 else
                    "Moderada −" if r <= -0.4 else
                    "Débil −" if r <= -0.1 else
                    "Sin correlación"
                )
            )
            st.dataframe(pairs_df.head(20), use_container_width=True,
                         hide_index=True)

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 5 — PROCESS CAPABILITY
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[4]:
        section("⚙️", "CAPACIDAD DE PROCESO — SPC")

        if not lsl_dict and not usl_dict:
            st.markdown("""
            <div class="warn-box">
              ⚠ Define al menos un límite de especificación (LSL / USL) en el
              Panel de Control del sidebar para calcular índices de capacidad.
            </div>""", unsafe_allow_html=True)

        var_cap = st.selectbox(
            "Variable para análisis de capacidad",
            options=selected_vars, key="cap_var",
        )
        s_dict = compute_statistics(
            df[var_cap],
            lsl=lsl_dict.get(var_cap),
            usl=usl_dict.get(var_cap),
        )

        col_a, col_b = st.columns([2, 1])
        with col_a:
            fig_cap = build_capability_chart(s_dict, var_cap)
            st.plotly_chart(fig_cap, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 1000, "height": 450,
                                "scale": 2, "filename": f"cap_{var_cap}"}})
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ HTML Capacidad",
                                   data=fig_to_html(fig_cap).encode(),
                                   file_name=f"cap_{var_cap}.html",
                                   mime="text/html")
            with c2:
                png_b = fig_to_png(fig_cap)
                if png_b:
                    st.download_button("⬇ PNG Capacidad", data=png_b,
                                       file_name=f"cap_{var_cap}.png",
                                       mime="image/png")

        with col_b:
            section("🔢", "ÍNDICES DE CAPACIDAD")
            render_stats_table(s_dict, var_cap)

        # Capability summary for all vars
        if len(selected_vars) > 1 and (lsl_dict or usl_dict):
            st.markdown("<br>", unsafe_allow_html=True)
            section("📊", "RESUMEN DE CAPACIDAD — TODAS LAS VARIABLES")
            cap_rows = []
            for v in selected_vars:
                sd = compute_statistics(df[v], lsl=lsl_dict.get(v),
                                         usl=usl_dict.get(v))
                cap_rows.append({
                    "Variable": v,
                    "N":      sd.get("n"),
                    "μ":      round(sd.get("mean", 0), 4),
                    "σ":      round(sd.get("std",  0), 4),
                    "Cp":     round(sd.get("cp",   float("nan")), 4),
                    "Cpk":    round(sd.get("cpk",  float("nan")), 4),
                    "Pp":     round(sd.get("pp",   float("nan")), 4),
                    "Ppk":    round(sd.get("ppk",  float("nan")), 4),
                    "Sigma":  round(sd.get("sigma_level", float("nan")), 2),
                    "PPM":    round(sd.get("ppm_est", float("nan")), 1),
                    "Yield%": round(sd.get("yield_pct", float("nan")), 4),
                })
            cap_df = pd.DataFrame(cap_rows)

            def color_cpk(val):
                if pd.isna(val): return ""
                if val >= 1.67: return "background-color: rgba(0,230,118,0.2)"
                if val >= 1.33: return "background-color: rgba(255,215,64,0.2)"
                return "background-color: rgba(255,68,68,0.2)"

            st.dataframe(
                cap_df.style.applymap(color_cpk, subset=["Cpk", "Ppk"]),
                use_container_width=True, hide_index=True,
            )

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 6 — ADVANCED
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[5]:
        section("📐", "ANÁLISIS AVANZADO")

        adv_tabs = st.tabs([
            "📡 Autocorrelación",
            "📈 CUSUM Multivariable",
            "🔄 Tendencias Superpuestas",
            "🕰 Análisis Periódico",
        ])

        # — Autocorrelation —
        with adv_tabs[0]:
            section("📡", "AUTOCORRELACIÓN (ACF)")
            var_acf = st.selectbox("Variable ACF", options=selected_vars,
                                    key="acf_var")
            max_lag = st.slider("Máx. Lags", 10, 100, 50, key="acf_lags")
            fig_acf = build_autocorrelation(df, var_acf, max_lags=max_lag)
            st.plotly_chart(fig_acf, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 1200, "height": 420,
                                "scale": 2, "filename": f"acf_{var_acf}"}})
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ HTML ACF",
                                   data=fig_to_html(fig_acf).encode(),
                                   file_name=f"acf_{var_acf}.html",
                                   mime="text/html")
            with c2:
                png_b = fig_to_png(fig_acf)
                if png_b:
                    st.download_button("⬇ PNG ACF", data=png_b,
                                       file_name=f"acf_{var_acf}.png",
                                       mime="image/png")

        # — Multi CUSUM —
        with adv_tabs[1]:
            section("📈", "CUSUM — COMPARATIVA MULTIVARIABLE")
            cusum_vars = st.multiselect(
                "Variables CUSUM", options=selected_vars,
                default=selected_vars[:min(3, len(selected_vars))],
                key="cusum_vars",
            )
            for v in cusum_vars:
                fig_cs2 = build_cumulative_sum(df, TIME_COL, v,
                                               target=tgt_dict.get(v))
                st.plotly_chart(fig_cs2, use_container_width=True,
                                config={"toImageButtonOptions": {
                                    "format": "png", "width": 1300,
                                    "height": 380, "scale": 2,
                                    "filename": f"cusum_{v}"}})
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(f"⬇ HTML CUSUM {v}",
                                       data=fig_to_html(fig_cs2).encode(),
                                       file_name=f"cusum_{v}.html",
                                       mime="text/html", key=f"dl_cusum_{v}")
                with c2:
                    png_b = fig_to_png(fig_cs2)
                    if png_b:
                        st.download_button(f"⬇ PNG CUSUM {v}", data=png_b,
                                           file_name=f"cusum_{v}.png",
                                           mime="image/png",
                                           key=f"dl_cusum_png_{v}")

        # — Overlapping trends —
        with adv_tabs[2]:
            section("🔄", "TENDENCIAS SUPERPUESTAS — EJE Y DOBLE")
            if len(selected_vars) >= 2:
                v_left  = st.selectbox("Variable eje izquierdo",
                                        options=selected_vars, index=0,
                                        key="ov_left")
                v_right = st.selectbox("Variable eje derecho",
                                        options=[v for v in selected_vars
                                                 if v != v_left],
                                        index=0, key="ov_right")

                fig_ov = make_subplots(specs=[[{"secondary_y": True}]])
                fig_ov.add_trace(
                    go.Scatter(
                        x=df[TIME_COL], y=df[v_left], mode="lines",
                        name=v_left,
                        line=dict(color=COLORS_PRIMARY[0], width=2),
                        hovertemplate=f"{v_left}: %{{y:,.4f}}<extra></extra>",
                    ), secondary_y=False,
                )
                fig_ov.add_trace(
                    go.Scatter(
                        x=df[TIME_COL], y=df[v_right], mode="lines",
                        name=v_right,
                        line=dict(color=COLORS_PRIMARY[1], width=2,
                                  dash="dash"),
                        hovertemplate=f"{v_right}: %{{y:,.4f}}<extra></extra>",
                    ), secondary_y=True,
                )
                apply_template(fig_ov)
                fig_ov.update_layout(
                    title=f"🔄  {v_left} vs {v_right} — Doble Eje Y",
                    height=460,
                    hovermode="x unified",
                )
                fig_ov.update_yaxes(title_text=v_left, secondary_y=False,
                                    title_font_color=COLORS_PRIMARY[0])
                fig_ov.update_yaxes(title_text=v_right, secondary_y=True,
                                    title_font_color=COLORS_PRIMARY[1])
                st.plotly_chart(fig_ov, use_container_width=True,
                                config={"scrollZoom": True,
                                        "toImageButtonOptions": {
                                            "format": "png", "width": 1400,
                                            "height": 500, "scale": 2,
                                            "filename": "doble_eje"}})
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("⬇ HTML Doble Eje",
                                       data=fig_to_html(fig_ov).encode(),
                                       file_name="doble_eje.html",
                                       mime="text/html")
                with c2:
                    png_b = fig_to_png(fig_ov)
                    if png_b:
                        st.download_button("⬇ PNG Doble Eje", data=png_b,
                                           file_name="doble_eje.png",
                                           mime="image/png")
            else:
                st.markdown('<div class="warn-box">Necesitas al menos 2 variables.</div>',
                            unsafe_allow_html=True)

        # — Periodic analysis —
        with adv_tabs[3]:
            section("🕰", "ANÁLISIS PERIÓDICO — PROMEDIO POR HORA / DÍA")
            var_per = st.selectbox("Variable", options=selected_vars, key="per_var")
            agg_by  = st.radio("Agregar por", ["Hora del día", "Día de semana",
                                                "Mes"], horizontal=True)

            df_t = df[[TIME_COL, var_per]].dropna().copy()
            if agg_by == "Hora del día":
                df_t["periodo"] = df_t[TIME_COL].dt.hour
                xlabel = "Hora del día (0-23)"
            elif agg_by == "Día de semana":
                df_t["periodo"] = df_t[TIME_COL].dt.day_name()
                xlabel = "Día de la semana"
            else:
                df_t["periodo"] = df_t[TIME_COL].dt.month_name()
                xlabel = "Mes"

            agg = df_t.groupby("periodo")[var_per].agg(
                ["mean", "std", "min", "max", "count"]
            ).reset_index()
            agg.columns = ["periodo", "media", "std", "min", "max", "n"]

            fig_per = go.Figure()
            fig_per.add_trace(go.Bar(
                x=agg["periodo"], y=agg["media"],
                name="Media",
                marker_color=COLORS_PRIMARY[0],
                error_y=dict(type="data", array=agg["std"].fillna(0).tolist(),
                             color="rgba(255,215,64,0.6)", thickness=1.5,
                             width=4),
                hovertemplate=(
                    "%{x}<br>μ: %{y:,.4f}<br>"
                    "σ: %{error_y.array:,.4f}<extra></extra>"
                ),
            ))
            apply_template(fig_per)
            fig_per.update_layout(
                title=f"📊  {var_per} — Media por {agg_by}",
                xaxis_title=xlabel, yaxis_title=var_per,
                height=420,
            )
            st.plotly_chart(fig_per, use_container_width=True,
                            config={"toImageButtonOptions": {
                                "format": "png", "width": 1200, "height": 480,
                                "scale": 2, "filename": f"period_{var_per}"}})
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇ HTML Periódico",
                                   data=fig_to_html(fig_per).encode(),
                                   file_name=f"period_{var_per}.html",
                                   mime="text/html")
            with c2:
                png_b = fig_to_png(fig_per)
                if png_b:
                    st.download_button("⬇ PNG Periódico", data=png_b,
                                       file_name=f"period_{var_per}.png",
                                       mime="image/png")

            st.dataframe(agg, use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════════════
    # TAB 7 — RAW DATA
    # ═══════════════════════════════════════════════════════════════════════
    with tabs[6]:
        section("🗂", "DATOS CRUDOS", f"{len(df):,} registros")

        display_cols = [TIME_COL] + selected_vars
        df_show = df[display_cols].copy()
        df_show.columns = [dt_col_choice] + selected_vars

        # Search / filter
        with st.expander("🔍 Filtros rápidos", expanded=False):
            fc1, fc2 = st.columns(2)
            with fc1:
                fvar = st.selectbox("Filtrar por variable", options=selected_vars,
                                     key="flt_var")
                fop  = st.selectbox("Operador",
                                     ["mayor que", "menor que",
                                      "igual a", "entre"],
                                     key="flt_op")
            with fc2:
                fval = st.number_input("Valor 1", key="flt_val1")
                fval2 = st.number_input("Valor 2 (solo 'entre')", key="flt_val2")

            if st.button("Aplicar filtro"):
                orig_df = df_show.copy()
                try:
                    if fop == "mayor que":
                        df_show = df_show[df_show[fvar] > fval]
                    elif fop == "menor que":
                        df_show = df_show[df_show[fvar] < fval]
                    elif fop == "igual a":
                        df_show = df_show[df_show[fvar] == fval]
                    elif fop == "entre":
                        df_show = df_show[df_show[fvar].between(fval, fval2)]
                    st.info(f"Filtrando {fvar} {fop} {fval}: "
                            f"{len(df_show):,} de {len(orig_df):,} registros")
                except Exception as e:
                    st.error(f"Error filtro: {e}")

        st.dataframe(df_show, use_container_width=True, height=420)

        # Download raw data
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            csv_buf = df_show.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇ Descargar CSV filtrado",
                data=csv_buf,
                file_name=f"datos_{datetime.now():%Y%m%d_%H%M%S}.csv",
                mime="text/csv",
            )
        with c2:
            xl_buf = io.BytesIO()
            df_show.to_excel(xl_buf, index=False)
            st.download_button(
                "⬇ Descargar Excel filtrado",
                data=xl_buf.getvalue(),
                file_name=f"datos_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # Describe
        st.markdown("<br>", unsafe_allow_html=True)
        section("📋", "RESUMEN ESTADÍSTICO RÁPIDO (describe)")
        st.dataframe(df_show[selected_vars].describe().T.round(4),
                     use_container_width=True)

    # ── FOOTER ───────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="margin-top:3rem;padding:1.2rem 2rem;
                border-top:1px solid #1e2d4a;
                font-family:'Share Tech Mono',monospace;
                font-size:0.65rem;color:#4a5a7a;
                display:flex;justify-content:space-between;
                align-items:center;">
      <div>⚙ PROCESS DATA ANALYZER v2.0 · INDUSTRIAL ENGINEERING PLATFORM</div>
      <div>MOTOR: PLOTLY · SCIPY · PANDAS · STREAMLIT</div>
      <div>© 2025 · ANÁLISIS DE PROCESOS INDUSTRIALES</div>
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# 11. ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
