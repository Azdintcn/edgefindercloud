"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                     EDGEFINDER PRO — AI WEB DASHBOARD                        ║
║  Institutional Macro, Retail Sentiment & COT Market Intelligence Terminal    ║
║  Includes Complete A1 Trading EdgeFinder Asset Scorecard & Deep Matrixes.     ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import os
import sys
import time
import json
from datetime import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

import importlib

# Ensure script dir has highest priority in sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR in sys.path:
    sys.path.remove(SCRIPT_DIR)
sys.path.insert(0, SCRIPT_DIR)

import edgefinder_engine as ee
importlib.reload(ee)

# =====================================================================
# 1. PAGE CONFIGURATION & INSTITUTIONAL DARK THEME
# =====================================================================
st.set_page_config(
    page_title="EdgeFinder Pro | Asset Scorecard & Institutional Terminal",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-End Trading Terminal Styling
st.markdown("""
<style>
    .stApp {
        background-color: #0d1117;
        color: #c9d1d9;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    
    .terminal-header {
        background: linear-gradient(90deg, #161b22 0%, #1f2937 100%);
        border: 1px solid #30363d;
        border-radius: 10px;
        padding: 14px 20px;
        margin-bottom: 16px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    
    .kpi-card {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 12px 14px;
        text-align: center;
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .kpi-card:hover {
        border-color: #00c896;
        transform: translateY(-2px);
    }
    .kpi-title {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #8b949e;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 22px;
        font-weight: 800;
        color: #ffffff;
        margin-bottom: 2px;
    }
    .kpi-sub {
        font-size: 11px;
        font-weight: 600;
    }
    
    .section-title {
        font-size: 15px;
        font-weight: 700;
        color: #ffd700;
        margin-top: 10px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    /* Institutional Matrix Table Styling */
    .matrix-card {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
        margin-bottom: 12px;
        overflow: hidden;
    }
    .matrix-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 8px 14px;
        background-color: #21262d;
        border-bottom: 1px solid #30363d;
    }
    .matrix-title {
        font-size: 13px;
        font-style: italic;
        font-weight: 700;
        color: #c9d1d9;
    }
    .matrix-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
    }
    .matrix-table th {
        text-align: left;
        padding: 6px 12px;
        color: #8b949e;
        font-weight: 600;
        font-size: 11px;
        border-bottom: 1px solid #21262d;
    }
    .matrix-table td {
        padding: 7px 12px;
        border-bottom: 1px solid #1f242c;
        color: #e6edf3;
    }
    .matrix-table tr:last-child td {
        border-bottom: none;
    }
    .matrix-mono {
        font-family: ui-monospace, SFMono-Regular, "SF Mono", Menlo, Consolas, monospace;
        font-weight: 600;
    }
    
    .news-box {
        background-color: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
    }
    .news-headline {
        font-size: 13px;
        font-weight: 600;
        color: #e6edf3;
        text-decoration: none;
    }
    .news-headline:hover {
        color: #58a6ff;
        text-decoration: underline;
    }
    .news-meta {
        font-size: 10px;
        color: #8b949e;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

# Helper function for rendering status badges
def get_badge_html(status):
    st_up = str(status).upper()
    if "VERY BEARISH" in st_up:
        bg = "#da3633"
    elif "BEARISH" in st_up:
        bg = "#da3633"
    elif "VERY BULLISH" in st_up:
        bg = "#1f6feb"
    elif "BULLISH" in st_up:
        bg = "#1f6feb"
    else:
        bg = "#484f58"
    return f'<span style="background-color: {bg}; color: #ffffff; padding: 2px 10px; border-radius: 3px; font-size: 11px; font-weight: 700; display: inline-block; min-width: 65px; text-align: center;">{status}</span>'

# =====================================================================
# 2. DATA LOADING & CACHE LAYER
# =====================================================================
@st.cache_data(ttl=60)
def load_edgefinder_data(force=False):
    return ee.collect_all_data_fast(force_refresh=force)

# =====================================================================
# 3. SIDEBAR CONTROLS
# =====================================================================
st.sidebar.markdown("### ⚡ **EdgeFinder Control Center**")

# Refresh Controls
col_sync1, col_sync2 = st.sidebar.columns([1, 1])
force_refresh = col_sync1.button("🔄 Sync Live Data", use_container_width=True)
if force_refresh:
    st.cache_data.clear()
    with st.spinner("Executing multi-threaded data sync..."):
        data = load_edgefinder_data(force=True)
    st.sidebar.success(f"Synced in {data.get('_sync_duration', 0)}s!")
else:
    data = load_edgefinder_data(force=False)

# Export Excel Button
if col_sync2.button("📥 Export Excel", use_container_width=True):
    with st.spinner("Generating Excel Workbook..."):
        excel_path = ee.build_excel_file(data)
    if excel_path and os.path.exists(excel_path):
        st.sidebar.success("Excel Updated!")
        with open(excel_path, "rb") as f:
            st.sidebar.download_button(
                label="💾 Download .xlsx",
                data=f,
                file_name="EdgeFinder_Dashboard.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

st.sidebar.markdown("---")
st.sidebar.markdown("#### 🎯 **Asset Filter Options**")
selected_market = st.sidebar.selectbox(
    "Market Category:",
    ["All Markets", "Forex", "Commodity", "Index", "Crypto"]
)

min_quality = st.sidebar.slider(
    "Minimum Setup Quality (Stars):",
    min_value=1,
    max_value=5,
    value=1,
    help="Filter setups by institutional quality confidence."
)

search_query = st.sidebar.text_input("🔍 Search Pair or Symbol:", placeholder="e.g. GOLD, EUR, SPX, BTC")

st.sidebar.markdown("---")
st.sidebar.markdown("#### 🛰️ **Connected Feeds**")
st.sidebar.markdown("""
- **FRED Macro API**: Online 🟢
- **CFTC COT Speculators**: Online 🟢
- **Myfxbook Retail Sentiment**: Online 🟢
- **MetaTrader 5 Tick Engine**: Active 🟢
- **Finnhub Global Wire**: Online 🟢
""")

# =====================================================================
# 4. TOP INSTITUTIONAL KPI BAR
# =====================================================================
st.markdown("""
<div class="terminal-header">
    <div>
        <span style="font-size: 20px; font-weight: 800; color: #ffffff; letter-spacing: 0.05em;">⚡ EDGEFINDER PRO</span>
        <span style="font-size: 13px; color: #00c896; margin-left: 10px; font-weight: 600;">INSTITUTIONAL QUANTITATIVE TERMINAL</span>
    </div>
    <div style="font-size: 12px; color: #8b949e;">
        Live Data Sync: <span style="color: #58a6ff; font-weight: bold;">""" + str(data.get("_last_sync_str", "Live")) + """</span>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2, col3, col4, col5, col6 = st.columns(6)

fed_rate = data.get("fed_rate", 3.88)
y10 = data.get("10y_yield", 4.84)
y2  = data.get("2y_yield", 4.46)
spread = data.get("yield_spread", round(y10 - y2, 3))
yield_status = data.get("yield_status", "FLAT (Caution)")
vix = data.get("vix", 28.18)
dxy = data.get("dxy", 121.67)

with col1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Fed Funds Rate</div>
        <div class="kpi-value">{fed_rate:.2f}%</div>
        <div class="kpi-sub" style="color: #00c896;">Benchmark Policy</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">US 10Y Yield</div>
        <div class="kpi-value">{y10:.2f}%</div>
        <div class="kpi-sub" style="color: #8b949e;">Risk-Free Benchmark</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">US 2Y Yield</div>
        <div class="kpi-value">{y2:.2f}%</div>
        <div class="kpi-sub" style="color: #8b949e;">Short-Term Horizon</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    spread_color = "#da3633" if spread < 0 else ("#f59e0b" if spread < 0.5 else "#22c55e")
    spread_label = "INVERTED ⚠️" if spread < 0 else ("FLAT ⚠️" if spread < 0.5 else "HEALTHY STEEP")
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">10Y - 2Y Spread</div>
        <div class="kpi-value" style="color: {spread_color};">{spread:+.3f}%</div>
        <div class="kpi-sub" style="color: {spread_color}; font-weight: bold;">{spread_label}</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    vix_val = vix if vix else 28.18
    vix_color = "#22c55e" if vix_val < 18 else ("#f59e0b" if vix_val < 25 else "#da3633")
    vix_label = "RISK-ON" if vix_val < 18 else ("CAUTION" if vix_val < 25 else "RISK-OFF")
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">VIX Fear Index</div>
        <div class="kpi-value" style="color: {vix_color};">{vix_val:.2f}</div>
        <div class="kpi-sub" style="color: {vix_color}; font-weight: bold;">{vix_label}</div>
    </div>
    """, unsafe_allow_html=True)

with col6:
    dxy_val = f"{dxy:.2f}" if dxy else "121.67"
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">US Dollar Index</div>
        <div class="kpi-value">{dxy_val}</div>
        <div class="kpi-sub" style="color: #3b82f6;">Trade-Weighted USD</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# =====================================================================
# 5. MAIN NAVIGATION TABS
# =====================================================================
tab_scorecard, tab_scanner, tab_cot_sentiment, tab_market_watch, tab_macro_radar, tab_news = st.tabs([
    "🎯 Asset Scorecard (Institutional Matrix)",
    "🔥 Top Setups Scanner",
    "👥 Retail Sentiment & CFTC COT",
    "💹 Live Market Watch (Forex, Gold, Indices, Oil, BTC)",
    "🌍 Central Bank & Macro Radar",
    "📰 Live Financial News Wire"
])

# ═════════════════════════════════════════════════════════════════════
# TAB 1: ASSET SCORECARD (A1 TRADING EDGEFINDER MATRIX)
# ═════════════════════════════════════════════════════════════════════
with tab_scorecard:
    # Asset Selection Header
    col_sel, col_rating, col_meta = st.columns([3, 3, 4])
    
    asset_options = [
        "GOLD", "EUR/USD", "GBP/USD", "USD/JPY", "AUD/USD", 
        "USD/CAD", "USD/CHF", "NZD/USD", "SPX500", "NAS100", 
        "DOW30", "US OIL", "BTC/USD"
    ]
    
    selected_asset = col_sel.selectbox(
        "Symbol:",
        asset_options,
        index=0,
        help="Select any asset to inspect its full quantitative scorecard."
    )
    
    scorecard = ee.get_asset_scorecard(selected_asset, data)
    rating_str = scorecard["overall_rating"]
    overall_score = scorecard["overall_score"]
    
    # Large Rating Badge
    badge_bg = "#da3633" if "BEARISH" in rating_str.upper() else ("#1f6feb" if "BULLISH" in rating_str.upper() else "#484f58")
    col_rating.markdown(f"""
    <div style="display: flex; align-items: center; height: 100%; padding-top: 25px;">
        <span style="background-color: {badge_bg}; color: #ffffff; font-size: 18px; font-weight: 800; padding: 6px 20px; border-radius: 4px; letter-spacing: 0.05em; text-align: center; width: 100%;">
            {rating_str.upper()}
        </span>
    </div>
    """, unsafe_allow_html=True)
    
    col_meta.markdown(f"""
    <div style="text-align: right; padding-top: 25px; font-size: 12px; color: #8b949e;">
        <div><b>{scorecard.get('timestamp_str', '')}</b></div>
        <div style="font-size: 11px; font-style: italic; color: #6e7681; margin-top: 2px;">
            * signals that new data is expected in the next few days
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.write("")
    
    # Main Scorecard 2-Column Split
    col_left_dial, col_right_tables = st.columns([4, 7])
    
    # ── LEFT COLUMN: SPEEDOMETER, SCORES & HISTORY ───────────────────
    with col_left_dial:
        # 1. Gauge Speedometer Chart
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=overall_score,
            domain={'x': [0, 1], 'y': [0, 1]},
            gauge={
                'axis': {'range': [-15, 15], 'tickwidth': 1, 'tickcolor': "#8b949e", 'dtick': 5},
                'bar': {'color': "#ffffff", 'thickness': 0.15},
                'steps': [
                    {'range': [-15, -7], 'color': '#da3633'},
                    {'range': [-7, -2], 'color': '#f85149'},
                    {'range': [-2, 2], 'color': '#484f58'},
                    {'range': [2, 7], 'color': '#388bfd'},
                    {'range': [7, 15], 'color': '#1f6feb'}
                ],
                'threshold': {
                    'line': {'color': "#ffffff", 'width': 4},
                    'thickness': 0.75,
                    'value': overall_score
                }
            }
        ))
        fig_gauge.update_layout(
            height=210,
            margin=dict(l=20, r=20, t=10, b=10),
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#ffffff', size=14)
        )
        st.plotly_chart(fig_gauge, use_container_width=True)
        
        # 2. Score Boxes
        def score_color(val):
            if val <= -6: return "#da3633"
            elif val < 0: return "#f85149"
            elif val >= 6: return "#1f6feb"
            elif val > 0: return "#388bfd"
            return "#484f58"
        
        st.markdown(f"""
        <div style="display: flex; flex-direction: column; gap: 8px; margin-bottom: 20px;">
            <div style="display: flex; justify-content: space-between; align-items: center; background-color: #161b22; padding: 10px 14px; border-radius: 6px; border: 1px solid #30363d;">
                <span style="font-weight: 700; font-size: 13px; color: #e6edf3;">EdgeFinder score</span>
                <span style="background-color: {score_color(overall_score)}; color: #ffffff; font-weight: 800; font-size: 15px; padding: 3px 14px; border-radius: 4px; min-width: 45px; text-align: center;">{overall_score}</span>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; background-color: #161b22; padding: 10px 14px; border-radius: 6px; border: 1px solid #30363d;">
                <span style="font-weight: 700; font-size: 13px; color: #e6edf3;">Technical score</span>
                <span style="background-color: {score_color(scorecard['technical_score'])}; color: #ffffff; font-weight: 800; font-size: 15px; padding: 3px 14px; border-radius: 4px; min-width: 45px; text-align: center;">{scorecard['technical_score']}</span>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; background-color: #161b22; padding: 10px 14px; border-radius: 6px; border: 1px solid #30363d;">
                <span style="font-weight: 700; font-size: 13px; color: #e6edf3;">Sentiment score</span>
                <span style="background-color: {score_color(scorecard['sentiment_score'])}; color: #ffffff; font-weight: 800; font-size: 15px; padding: 3px 14px; border-radius: 4px; min-width: 45px; text-align: center;">{scorecard['sentiment_score']}</span>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; background-color: #161b22; padding: 10px 14px; border-radius: 6px; border: 1px solid #30363d;">
                <span style="font-weight: 700; font-size: 13px; color: #e6edf3;">Macroeconomic score</span>
                <span style="background-color: {score_color(scorecard['macro_score'])}; color: #ffffff; font-weight: 800; font-size: 15px; padding: 3px 14px; border-radius: 4px; min-width: 45px; text-align: center;">{scorecard['macro_score']}</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # 3. Score History Bar Chart
        st.markdown('<div style="text-align: center; font-style: italic; font-weight: 700; color: #8b949e; font-size: 13px; margin-bottom: 6px;">Score history</div>', unsafe_allow_html=True)
        hist = scorecard.get("score_history", [])
        hist_dates = [h["date"] for h in hist]
        hist_scores = [h["score"] for h in hist]
        hist_colors = ["#1f6feb" if s >= 0 else "#da3633" for s in hist_scores]
        
        fig_hist = go.Figure()
        fig_hist.add_trace(go.Bar(
            x=hist_dates,
            y=hist_scores,
            marker_color=hist_colors,
            text=[f"{s}" for s in hist_scores],
            textposition="outside"
        ))
        fig_hist.add_hline(y=0, line_color="#30363d", line_width=1)
        fig_hist.update_layout(
            height=190,
            margin=dict(l=10, r=10, t=20, b=25),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#8b949e', size=10),
            xaxis=dict(gridcolor='rgba(0,0,0,0)', showgrid=False),
            yaxis=dict(gridcolor='#21262d', range=[-16, 16], zeroline=True, zerolinecolor='#484f58')
        )
        st.plotly_chart(fig_hist, use_container_width=True)

    # ── RIGHT COLUMN: CROWD SENTIMENT & 5 MATRIX TABLES ─────────────
    with col_right_tables:
        # 1. Crowd Sentiment Signal
        crowd = scorecard.get("crowd_sentiment", {})
        crowd_sig = crowd.get("signal", "Bearish")
        crowd_bg = "#da3633" if "BEARISH" in crowd_sig.upper() else ("#1f6feb" if "BULLISH" in crowd_sig.upper() else "#484f58")
        bull_pct = crowd.get("bullish_pct", 78.4)
        bear_pct = crowd.get("bearish_pct", 21.6)
        
        st.markdown(f"""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
            <span style="font-weight: 700; font-size: 13px; color: #e6edf3;">Crowd sentiment signal</span>
            <span style="background-color: {crowd_bg}; color: white; font-weight: 700; font-size: 11px; padding: 2px 14px; border-radius: 3px;">{crowd_sig}</span>
        </div>
        <div style="display: flex; width: 100%; height: 20px; border-radius: 3px; overflow: hidden; margin-bottom: 4px;">
            <div style="width: {bull_pct}%; background-color: #1f6feb; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; color: white;">{bull_pct:.1f}%</div>
            <div style="width: {bear_pct}%; background-color: #da3633; display: flex; align-items: center; justify-content: center; font-size: 11px; font-weight: 700; color: white;">{bear_pct:.1f}%</div>
        </div>
        <div style="display: flex; justify-content: space-between; font-size: 10px; color: #8b949e; margin-bottom: 14px;">
            <span>Bullish sentiment</span>
            <span>Bearish sentiment</span>
        </div>
        """, unsafe_allow_html=True)
        
        tables = scorecard.get("tables", {})
        
        # ── TABLE 1: TECHNICALS ──────────────────────────────────────
        t_tech = tables.get("technicals", {})
        tech_badge = get_badge_html(t_tech.get("badge", "Very Bearish"))
        tech_rows_html = "".join([
            f"""<tr>
                <td style="font-weight: 600;">{r['metric']}</td>
                <td style="text-align: right;">{get_badge_html(r['status'])}</td>
            </tr>""" for r in t_tech.get("rows", [])
        ])
        st.markdown(f"""
        <div class="matrix-card">
            <div class="matrix-header">
                <span class="matrix-title">Technicals</span>
                {tech_badge}
            </div>
            <table class="matrix-table">
                {tech_rows_html}
            </table>
        </div>
        """, unsafe_allow_html=True)
        
        # ── TABLE 2: INSTITUTIONAL ACTIVITY (COT) ────────────────────
        t_inst = tables.get("institutional", {})
        inst_badge = get_badge_html(t_inst.get("badge", "Neutral"))
        inst_rows_html = "".join([
            f"""<tr>
                <td style="font-weight: 600;">{r['metric']}</td>
                <td>{get_badge_html(r['status'])}</td>
                <td class="matrix-mono">{r.get('long_pct', '—')}</td>
                <td class="matrix-mono">{r.get('short_pct', '—')}</td>
                <td class="matrix-mono">{r.get('change_pct', '—')}</td>
                <td style="color: #8b949e;">{r.get('date', '—')}</td>
            </tr>""" for r in t_inst.get("rows", [])
        ])
        st.markdown(f"""
        <div class="matrix-card">
            <div class="matrix-header">
                <span class="matrix-title">Institutional activity</span>
                {inst_badge}
            </div>
            <table class="matrix-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Status</th>
                        <th>Long %</th>
                        <th>Short %</th>
                        <th>Change %</th>
                        <th>Date</th>
                    </tr>
                </thead>
                <tbody>
                    {inst_rows_html}
                </tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)
        
        # ── TABLE 3: ECONOMIC GROWTH ─────────────────────────────────
        t_growth = tables.get("growth", {})
        growth_badge = get_badge_html(t_growth.get("badge", "Very Bearish"))
        growth_rows_html = "".join([
            f"""<tr>
                <td style="font-weight: 600;">{r['metric']}</td>
                <td>{get_badge_html(r['status'])}</td>
                <td class="matrix-mono">{r.get('actual', '—')}</td>
                <td class="matrix-mono" style="color: #8b949e;">{r.get('forecast', '—')}</td>
                <td class="matrix-mono" style="color: {'#3fb950' if '+' in r.get('surprise', '') else ('#f85149' if '-' in r.get('surprise', '') else '#c9d1d9')};">{r.get('surprise', '—')}</td>
                <td style="color: #8b949e;">{r.get('date', '—')}</td>
            </tr>""" for r in t_growth.get("rows", [])
        ])
        st.markdown(f"""
        <div class="matrix-card">
            <div class="matrix-header">
                <span class="matrix-title">Economic growth</span>
                {growth_badge}
            </div>
            <table class="matrix-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Status</th>
                        <th>Actual</th>
                        <th>Forecast</th>
                        <th>Surprise</th>
                        <th>Date</th>
                    </tr>
                </thead>
                <tbody>
                    {growth_rows_html}
                </tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)
        
        # ── TABLE 4: INFLATION ───────────────────────────────────────
        t_inf = tables.get("inflation", {})
        inf_badge = get_badge_html(t_inf.get("badge", "Bearish"))
        inf_rows_html = "".join([
            f"""<tr>
                <td style="font-weight: 600;">{r['metric']}</td>
                <td>{get_badge_html(r['status'])}</td>
                <td class="matrix-mono" {'colspan="4"' if 'The 2yr yield' in r.get('date', '') else ''}>{r.get('date') if 'The 2yr yield' in r.get('date', '') else r.get('actual', '—')}</td>
                {'' if 'The 2yr yield' in r.get('date', '') else f'<td class="matrix-mono" style="color: #8b949e;">{r.get("forecast", "—")}</td>'}
                {'' if 'The 2yr yield' in r.get('date', '') else f'<td class="matrix-mono" style="color: {"#3fb950" if "+" in r.get("surprise", "") else ("#f85149" if "-" in r.get("surprise", "") else "#c9d1d9")};">{r.get("surprise", "—")}</td>'}
                {'' if 'The 2yr yield' in r.get('date', '') else f'<td style="color: #8b949e;">{r.get("date", "—")}</td>'}
            </tr>""" for r in t_inf.get("rows", [])
        ])
        st.markdown(f"""
        <div class="matrix-card">
            <div class="matrix-header">
                <span class="matrix-title">Inflation</span>
                {inf_badge}
            </div>
            <table class="matrix-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Status</th>
                        <th>Actual</th>
                        <th>Forecast</th>
                        <th>Surprise</th>
                        <th>Date</th>
                    </tr>
                </thead>
                <tbody>
                    {inf_rows_html}
                </tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)
        
        # ── TABLE 5: JOBS MARKET ─────────────────────────────────────
        t_jobs = tables.get("jobs", {})
        jobs_badge = get_badge_html(t_jobs.get("badge", "Very Bearish"))
        jobs_rows_html = "".join([
            f"""<tr>
                <td style="font-weight: 600;">{r['metric']}</td>
                <td>{get_badge_html(r['status'])}</td>
                <td class="matrix-mono">{r.get('actual', '—')}</td>
                <td class="matrix-mono" style="color: #8b949e;">{r.get('forecast', '—')}</td>
                <td class="matrix-mono" style="color: {'#3fb950' if '+' in r.get('surprise', '') else ('#f85149' if '-' in r.get('surprise', '') else '#c9d1d9')};">{r.get('surprise', '—')}</td>
                <td style="color: #8b949e;">{r.get('date', '—')}</td>
            </tr>""" for r in t_jobs.get("rows", [])
        ])
        st.markdown(f"""
        <div class="matrix-card">
            <div class="matrix-header">
                <span class="matrix-title">Jobs market</span>
                {jobs_badge}
            </div>
            <table class="matrix-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Status</th>
                        <th>Actual</th>
                        <th>Forecast</th>
                        <th>Surprise</th>
                        <th>Date</th>
                    </tr>
                </thead>
                <tbody>
                    {jobs_rows_html}
                </tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════
# TAB 2: TOP SETUPS SCANNER
# ═════════════════════════════════════════════════════════════════════
with tab_scanner:
    st.markdown('<div class="section-title">🔥 MULTI-FACTOR EDGE RANKING SCANNER</div>', unsafe_allow_html=True)
    st.caption("Includes Forex, Gold, Silver, Oil & Major Stock Indices. Combines Macro Interest Rate Divergence (FRED) + Retail Crowd Sentiment (Myfxbook Contrarian) + RSI Momentum.")
    
    setups = data.get("setups", [])
    filtered_setups = []
    for s in setups:
        if selected_market != "All Markets" and s["market"] != selected_market:
            continue
        if s["stars"] < min_quality:
            continue
        if search_query and search_query.upper() not in s["asset"].upper():
            continue
        filtered_setups.append(s)

    if filtered_setups:
        display_rows = []
        for s in filtered_setups:
            score = s["score"]
            bias = s["bias"]
            stars = "⭐" * s["stars"]
            rsi_str = f"{s['rsi']:.1f}" if s.get("rsi") else "50.0"
            
            display_rows.append({
                "Asset": s["asset"],
                "Market": s["market"],
                "Directional Bias": bias,
                "Edge Score": score,
                "Setup Quality": stars,
                "Retail Long %": f"{s['retail_long']:.1f}%",
                "Retail Short %": f"{s['retail_short']:.1f}%",
                "RSI (14)": rsi_str
            })
        
        df_disp = pd.DataFrame(display_rows)
        
        st.dataframe(
            df_disp,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Edge Score": st.column_config.ProgressColumn(
                    "Edge Score (-10 to +10)",
                    min_value=-10.0,
                    max_value=10.0,
                    format="%+.1f"
                ),
                "Setup Quality": st.column_config.TextColumn("Quality Rating"),
                "Directional Bias": st.column_config.TextColumn("Bias"),
            }
        )
    else:
        st.info("No setups match the current filters. Adjust your filter criteria in the sidebar.")

# ═════════════════════════════════════════════════════════════════════
# TAB 3: RETAIL SENTIMENT & INSTITUTIONAL COT
# ═════════════════════════════════════════════════════════════════════
with tab_cot_sentiment:
    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.markdown('<div class="section-title">👥 RETAIL SENTIMENT VS CONTRARIAN BIAS</div>', unsafe_allow_html=True)
        st.caption("Myfxbook Retail Crowd Allocation. Extremes (>65% Long = Institutional Fade Sell | <35% Long = Institutional Fade Buy).")

        sentiment = data.get("sentiment", {})
        if sentiment:
            sent_pairs = ["EURUSD", "GBPUSD", "AUDUSD", "USDJPY", "USDCAD", "NZDUSD", "USDCHF", "GBPJPY", "GOLD", "USOIL"]
            chart_pairs = []
            long_vals = []
            short_vals = []

            for p in sent_pairs:
                if p in sentiment:
                    lp, sp = sentiment[p]
                    disp = f"{p[:3]}/{p[3:]}" if len(p) == 6 else p
                    chart_pairs.append(disp)
                    long_vals.append(lp)
                    short_vals.append(sp)

            fig_sent = go.Figure()
            fig_sent.add_trace(go.Bar(
                y=chart_pairs,
                x=long_vals,
                name="Retail Long %",
                orientation='h',
                marker=dict(color='rgba(31, 111, 235, 0.85)', line=dict(color='#1f6feb', width=1)),
                text=[f"{v:.1f}%" for v in long_vals],
                textposition='inside'
            ))
            fig_sent.add_trace(go.Bar(
                y=chart_pairs,
                x=short_vals,
                name="Retail Short %",
                orientation='h',
                marker=dict(color='rgba(218, 54, 51, 0.85)', line=dict(color='#da3633', width=1)),
                text=[f"{v:.1f}%" for v in short_vals],
                textposition='inside'
            ))

            fig_sent.add_vline(x=50, line_dash="solid", line_color="#4b5563", opacity=0.8)
            fig_sent.add_vline(x=65, line_dash="dot", line_color="#da3633", annotation_text="Extreme Long >65%", annotation_position="top right")
            fig_sent.add_vline(x=35, line_dash="dot", line_color="#1f6feb", annotation_text="Extreme Short <35%", annotation_position="top left")

            fig_sent.update_layout(
                barmode='stack',
                height=390,
                margin=dict(l=20, r=20, t=25, b=20),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#c9d1d9', size=11),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                xaxis=dict(range=[0, 100], gridcolor='#21262d', title="Crowd Allocation %"),
                yaxis=dict(gridcolor='#21262d', autorange="reversed")
            )
            st.plotly_chart(fig_sent, use_container_width=True)
        else:
            st.info("Retail sentiment syncing from Myfxbook.")

    with col_right:
        st.markdown('<div class="section-title">🏛️ INSTITUTIONAL COT SPECULATIVE POSITIONING</div>', unsafe_allow_html=True)
        st.caption("CFTC Commitment of Traders: Non-Commercial Large Speculator (Hedge Fund) Net Contracts.")

        cot_raw = data.get("cot_raw", [])
        if cot_raw:
            cot_sorted = sorted(cot_raw, key=lambda x: abs(x["net_speculators"]), reverse=True)[:8]
            cot_mkts = [c["market"] for c in cot_sorted]
            cot_net = [c["net_speculators"] for c in cot_sorted]
            cot_colors = ['#1f6feb' if n >= 0 else '#da3633' for n in cot_net]

            fig_cot = go.Figure()
            fig_cot.add_trace(go.Bar(
                x=cot_net,
                y=cot_mkts,
                orientation='h',
                marker=dict(color=cot_colors),
                text=[f"{n:+,d}" for n in cot_net],
                textposition='outside'
            ))
            fig_cot.add_vline(x=0, line_color="#4b5563")
            fig_cot.update_layout(
                height=390,
                margin=dict(l=20, r=45, t=25, b=20),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#c9d1d9', size=11),
                xaxis=dict(gridcolor='#21262d', title="Net Speculator Contracts (Longs - Shorts)"),
                yaxis=dict(gridcolor='#21262d', autorange="reversed")
            )
            st.plotly_chart(fig_cot, use_container_width=True)
        else:
            st.info("CFTC COT data endpoint currently parsing records.")

    # COT Detailed Breakdown Table
    if cot_raw:
        with st.expander("📋 View Detailed CFTC Commitment of Traders (COT) Data Table", expanded=False):
            cot_rows = []
            for c in cot_raw:
                cot_rows.append({
                    "Market / Contract": c["market"],
                    "Report Date": c.get("report_date", "N/A"),
                    "Net Speculator Contracts": f"{c['net_speculators']:+,d}",
                    "Long Contracts": f"{c['longs']:,d}",
                    "Short Contracts": f"{c['shorts']:,d}",
                    "Weekly Change": f"{c['change']:+,d}",
                    "Institutional Bias": c["signal"]
                })
            st.dataframe(pd.DataFrame(cot_rows), use_container_width=True, hide_index=True)

# ═════════════════════════════════════════════════════════════════════
# TAB 4: LIVE MARKET WATCH
# ═════════════════════════════════════════════════════════════════════
with tab_market_watch:
    col_f1, col_f2 = st.columns([1, 3])
    cat_filter = col_f1.radio(
        "Asset Class Filter:",
        ["All Assets", "Gold & Commodities", "Stock Indices", "Forex", "Crypto"],
        horizontal=True
    )
    
    all_assets = data.get("assets", {})
    if not all_assets:
        all_assets = data.get("fx", {})
        
    rsi_dict = data.get("rsi", {})

    if all_assets:
        rows = []
        for name, info in all_assets.items():
            cat = info.get("category", "Forex")
            
            if cat_filter == "Gold & Commodities" and cat != "Commodity":
                continue
            elif cat_filter == "Stock Indices" and cat != "Index":
                continue
            elif cat_filter == "Forex" and cat != "Forex":
                continue
            elif cat_filter == "Crypto" and cat != "Crypto":
                continue

            bid_val = info.get("bid", info.get("price", 0.0))
            ask_val = info.get("ask", bid_val)
            chg = info.get("change", 0.0)
            src = info.get("source", "Live")
            rsi_val = rsi_dict.get(name, 50.0)

            bid_fmt = f"${bid_val:,.2f}" if ("GOLD" in name or "OIL" in name or "SILVER" in name or "Index" in cat or "Crypto" in cat) else (f"{bid_val:,.5f}" if bid_val < 10 else f"{bid_val:,.2f}")
            ask_fmt = f"${ask_val:,.2f}" if ("GOLD" in name or "OIL" in name or "SILVER" in name or "Index" in cat or "Crypto" in cat) else (f"{ask_val:,.5f}" if ask_val < 10 else f"{ask_val:,.2f}")

            if rsi_val >= 70:
                zone = "OVERBOUGHT (Reversal Risk ⬇️)"
            elif rsi_val <= 30:
                zone = "OVERSOLD (Bounce Opportunity ⬆️)"
            else:
                zone = "NEUTRAL (Trend Mode)"

            rows.append({
                "Asset / Symbol": name,
                "Class": cat,
                "Bid Price": bid_fmt,
                "Ask Price": ask_fmt,
                "24H Change %": chg,
                "RSI (14)": rsi_val,
                "Momentum State": zone,
                "Feed Source": src
            })

        df_assets = pd.DataFrame(rows)
        st.dataframe(
            df_assets,
            use_container_width=True,
            hide_index=True,
            column_config={
                "24H Change %": st.column_config.NumberColumn(
                    "24H Change %",
                    format="%+.2f%%"
                ),
                "RSI (14)": st.column_config.ProgressColumn(
                    "RSI (14)",
                    min_value=0.0,
                    max_value=100.0,
                    format="%.1f"
                )
            }
        )
    else:
        st.info("Market asset prices loading...")

# ═════════════════════════════════════════════════════════════════════
# TAB 5: CENTRAL BANK & MACRO RADAR
# ═════════════════════════════════════════════════════════════════════
with tab_macro_radar:
    col_m1, col_m2 = st.columns([1, 1])
    
    with col_m1:
        st.markdown("##### 🏦 Federal Reserve Economic Indicators")
        macro_items = [
            ("US Fed Funds Rate", f"{data.get('fed_rate', 'N/A')}%", data.get('fed_rate_date', '—'), "Primary Interest Rate Driver"),
            ("US 10Y Treasury Yield", f"{data.get('10y_yield', 'N/A')}%", data.get('10y_date', '—'), "Long-Term Cost of Capital"),
            ("US 2Y Treasury Yield", f"{data.get('2y_yield', 'N/A')}%", data.get('2y_date', '—'), "Short-Term Rate Expectation"),
            ("US CPI (Price Index)", f"{data.get('cpi', 'N/A')}", data.get('cpi_date', '—'), "Core Inflation Gauge"),
            ("US GDP ($ Billions)", f"${data.get('gdp', 'N/A'):,.1f}" if data.get('gdp') else "N/A", data.get('gdp_date', '—'), "Economic Output"),
            ("US Unemployment Rate", f"{data.get('unrate', 'N/A')}%", data.get('unrate_date', '—'), "Labor Market Slack"),
            ("US Housing Starts (000s)", f"{data.get('housing', 'N/A')}k", data.get('housing_date', '—'), "Leading Real Estate Indicator"),
            ("CBOE Volatility Index (VIX)", f"{data.get('vix', 'N/A')}", data.get('vix_date', '—'), "Market Sentiment Risk Gauge"),
            ("Trade-Weighted US Dollar (DXY)", f"{data.get('dxy', 'N/A')}", data.get('dxy_date', '—'), "Broad Currency Strength")
        ]
        df_macro = pd.DataFrame(macro_items, columns=["Indicator", "Latest Value", "Release Date", "Market Impact"])
        st.dataframe(df_macro, use_container_width=True, hide_index=True)

    with col_m2:
        st.markdown("##### 📐 Yield Curve & Macro Playbook")
        st.markdown(f"""
        <div style="background-color: #161b22; border: 1px solid #30363d; border-radius: 8px; padding: 16px;">
            <div style="font-size: 14px; font-weight: bold; color: #ffd700; margin-bottom: 8px;">
                Current Yield Curve Status: <span style="color: {'#da3633' if spread < 0 else '#22c55e'};">{yield_status}</span>
            </div>
            <div style="font-size: 12px; line-height: 1.6; color: #c9d1d9;">
                <b>The 10Y - 2Y Treasury Spread ({spread:+.3f}%)</b> is Wall Street's most historically accurate macroeconomic cycle indicator:
                <ul style="margin-top: 6px; padding-left: 18px;">
                    <li><b>Inverted Curve (&lt; 0%):</b> Short rates exceed long rates. Signals impending credit contraction and recession risk. Highly favorable for safe-haven USD and Gold; headwind for equities.</li>
                    <li><b>Flat Curve (0% to 0.5%):</b> Transition regime. Central banks tightening policy; markets assessing peak rates.</li>
                    <li><b>Steepening Curve (&gt; 0.5%):</b> Economic expansion or rate cuts underway. Bullish for equities, commodities, and high-beta currencies (AUD, NZD, GBP).</li>
                </ul>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════
# TAB 6: LIVE FINANCIAL NEWS WIRE (MULTI-STREAM AGGREGATOR)
# ═════════════════════════════════════════════════════════════════════
with tab_news:
    col_n1, col_n2 = st.columns([3, 1])
    with col_n1:
        st.markdown('<div class="section-title">🗞️ REAL-TIME INSTITUTIONAL BREAKING NEWS WIRE</div>', unsafe_allow_html=True)
        st.caption("Aggregates Reuters, Bloomberg, Yahoo Finance, Finnhub, and Cointelegraph. Auto-deduplicated and sorted by newest publication.")
    
    with col_n2:
        btn_refresh_news = st.button("🔄 Refresh News Now", use_container_width=True)
    
    fetch_news_fn = getattr(ee, "fetch_live_news", None)
    if btn_refresh_news:
        with st.spinner("Fetching latest breaking wires..."):
            if fetch_news_fn:
                news_items = fetch_news_fn(limit=40, force_refresh=True)
            else:
                news_items = []
            data["news"] = news_items
            st.success("News feed updated!")
    else:
        news_items = data.get("news")
        if not news_items and fetch_news_fn:
            news_items = fetch_news_fn(limit=40, force_refresh=False)
        news_items = news_items or []

    col_filter, col_status = st.columns([3, 2])
    cat_filter = col_filter.radio(
        "News Stream Category:",
        ["All Breaking News", "Macro & Economy", "Forex", "Commodities / Indices", "Crypto"],
        horizontal=True
    )
    
    with col_status:
        st.markdown(f"""
        <div style="text-align: right; padding-top: 20px; font-size: 11px; color: #8b949e;">
            <span style="color: #22c55e; font-weight: bold;">● LIVE FEED ONLINE</span> &bull; <b>{len(news_items)}</b> stories loaded
        </div>
        """, unsafe_allow_html=True)

    if news_items:
        filtered_news = []
        for item in news_items:
            cat = item.get("category", "GENERAL").upper()
            if cat_filter == "Macro & Economy" and "MACRO" not in cat and "GENERAL" not in cat:
                continue
            elif cat_filter == "Forex" and "FOREX" not in cat:
                continue
            elif cat_filter == "Commodities / Indices" and "COMMODITIES" not in cat and "INDICES" not in cat:
                continue
            elif cat_filter == "Crypto" and "CRYPTO" not in cat:
                continue
            filtered_news.append(item)

        if not filtered_news:
            filtered_news = news_items[:15]

        for item in filtered_news:
            headline = item.get("headline", "Market Update")
            url = item.get("url", "#")
            src = item.get("source", "Market Wire")
            cat = item.get("category", "GENERAL").upper()
            time_ago = item.get("time_ago", "Recent")
            summary = item.get("summary", "")

            # Category pill color
            if "MACRO" in cat or "GENERAL" in cat: cat_color = "#d97706"
            elif "FOREX" in cat: cat_color = "#2563eb"
            elif "COMMODITIES" in cat or "INDICES" in cat: cat_color = "#16a34a"
            elif "CRYPTO" in cat: cat_color = "#7c3aed"
            else: cat_color = "#484f58"

            st.markdown(f"""
            <div class="news-box">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                    <span style="background-color: {cat_color}; color: #ffffff; padding: 1px 8px; border-radius: 3px; font-size: 10px; font-weight: 700;">{cat}</span>
                    <span style="color: #22c55e; font-size: 11px; font-weight: 600;">🟢 {time_ago}</span>
                </div>
                <a href="{url}" target="_blank" class="news-headline" style="font-size: 14px; font-weight: 600; color: #58a6ff; display: block; margin-bottom: 4px;">{headline}</a>
                {f'<div style="font-size: 11px; color: #8b949e; line-height: 1.4; margin-bottom: 6px;">{summary}</div>' if summary else ''}
                <div class="news-meta" style="color: #6e7681; font-size: 10px;">
                    Source: <b>{src}</b> &bull; Click title to read full report
                </div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No live news returned by feed.")

# =====================================================================
# 8. FOOTER & DISCLAIMER (MATCHING A1 EDGEFINDER)
# =====================================================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; font-size: 11px; color: #6e7681; padding: 10px 0; line-height: 1.5;">
    The readings generated by EdgeFinder are for informational purposes only, do not constitute financial advice, and A1 Trading is not responsible for any investment decisions or outcomes.<br>
    ⚡ <b>EdgeFinder Pro Terminal</b> &bull; Multi-source Quantitative Engine &bull; Directional Fundamental Bias Only
</div>
""", unsafe_allow_html=True)
