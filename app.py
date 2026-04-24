import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import calendar
import datetime

from extractor import (
    extract_from_files, get_top_defects,
    get_rolling_stats, get_leak_trend, get_tip_trend,
    get_leak_trend_by_product, get_leak_valve_by_product,
    get_leak_bond_by_product
)

# ── Product color scheme ──────────────────────────────────────────────────────
PRODUCT_COLORS = {
    "BLN":   "#1565C0",   # blue
    "SMG":   "#424242",   # black (dark grey — visible on dark bg)
    "CTI":   "#2e7d32",   # green
    "MP3-M": "#e65100",   # orange
    "MP3":   "#e65100",   # orange (combined MP3)
    "MP3-S": "#bdbdbd",   # white / light grey
    "BMD":   "#f9a825",   # yellow
    "NFE":   "#e91e63",   # pink
    "NFS":   "#e91e63",   # pink (alias)
    "BLT":   "#fdd835",   # yellow (distinct from BMD)
    "BBB":   "#6d4c41",   # brown
}

def prod_color(prod):
    """Return the standard color for a product code."""
    return PRODUCT_COLORS.get(str(prod).upper(), "#8fa3b8")

st.set_page_config(page_title="SSPC Quality Dashboard", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.kpi-card{background:linear-gradient(135deg,#1e2a3a 0%,#162032 100%);
  border-radius:12px;padding:16px 20px;border-left:4px solid #2e75b6;margin-bottom:8px;}
.kpi-label{color:#8fa3b8;font-size:11px;font-weight:700;letter-spacing:.6px;
  text-transform:uppercase;margin-bottom:4px;}
.kpi-value{color:#fff;font-size:24px;font-weight:700;line-height:1.1;}
.kpi-delta-good{color:#00c853;font-size:11px;font-weight:600;margin-top:4px;}
.kpi-delta-bad{color:#ff3d3d;font-size:11px;font-weight:600;margin-top:4px;}
.kpi-delta-neutral{color:#8fa3b8;font-size:11px;margin-top:4px;}
.sec{color:#8fa3b8;font-size:10px;font-weight:700;letter-spacing:1.5px;
  text-transform:uppercase;margin:16px 0 8px;
  border-bottom:1px solid #1e2a3a;padding-bottom:3px;}
.note-box{background:#1e2a3a;border-radius:8px;padding:12px 16px;
  border-left:3px solid #ffc000;margin-bottom:8px;font-size:13px;color:#c0cfe0;}
.upload-box{background:#1e2a3a;border-radius:12px;padding:32px;
  border:1px dashed #2e4a6a;text-align:center;}
.tag-open{background:#c00000;color:#fff;border-radius:4px;
  padding:2px 8px;font-size:11px;font-weight:700;}
.tag-done{background:#00c853;color:#fff;border-radius:4px;
  padding:2px 8px;font-size:11px;font-weight:700;}
.tag-prog{background:#ffc000;color:#000;border-radius:4px;
  padding:2px 8px;font-size:11px;font-weight:700;}
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_c(v):     return f"${v:,.0f}"    if v is not None else "—"
def fmt_p(v):     return f"{v:.2%}"      if v is not None else "—"
def fmt_n(v, d=2):return f"{v:.{d}f}"   if v is not None else "—"

def delta_tag(curr, prev, higher_good=True, fmt="currency"):
    if prev is None or curr is None or prev == 0: return ""
    diff = curr - prev; pct = diff / prev
    good = (diff >= 0) == higher_good
    cls  = "kpi-delta-good" if good else "kpi-delta-bad"
    a    = "▲" if diff > 0 else "▼"
    if fmt == "currency": s = f"{a} ${abs(diff):,.0f} ({pct:+.1%})"
    elif fmt == "pct":    s = f"{a} {abs(diff):.2%} ({pct:+.1%})"
    else:                 s = f"{a} {abs(diff):.2f} ({pct:+.1%})"
    return f'<div class="{cls}">{s} vs prev</div>'

def kpi(label, value, delta="", accent="#2e75b6"):
    st.markdown(f"""<div class="kpi-card" style="border-left-color:{accent}">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>{delta}
    </div>""", unsafe_allow_html=True)

def sec(t):
    st.markdown(f'<div class="sec">{t}</div>', unsafe_allow_html=True)

PLOT = dict(
    plot_bgcolor="#1e2a3a", paper_bgcolor="#1e2a3a",
    font=dict(color="white", family="Arial"),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="white")),
)
PLOT_MARGIN = dict(t=40, b=30, l=10, r=10)  # default margin, override per chart

def line_fig(df_t, title, ytitle, cols_colors, yfmt=".2%", height=320):
    import numpy as np
    fig = go.Figure()
    for col, color in cols_colors:
        if col not in df_t.columns: continue
        # Fill NaN with 0 so missing months show as 0, not gaps
        vals = df_t[col].fillna(0).tolist()
        avg  = float(np.mean(vals))
        std  = float(np.std(vals)) if len(vals) > 1 else 0
        # Std dev band
        fig.add_trace(go.Scatter(
            x=list(df_t["label"]) + list(df_t["label"])[::-1],
            y=[avg+std]*len(df_t) + [avg-std]*len(df_t),
            fill="toself",
            fillcolor=color.replace(")", ",0.12)").replace("rgb","rgba")
                if color.startswith("rgb") else
                f"rgba({int(color[1:3],16)},{int(color[3:5],16)},{int(color[5:7],16)},0.12)",
            line=dict(color="rgba(0,0,0,0)"),
            showlegend=False, hoverinfo="skip",
            name=f"{col} ±1σ"))
        # Avg line
        fig.add_trace(go.Scatter(
            x=df_t["label"], y=[avg]*len(df_t),
            mode="lines", line=dict(color=color, width=1, dash="dot"),
            showlegend=False, hoverinfo="skip", name=f"{col} avg"))
        # Main line
        fig.add_trace(go.Scatter(
            x=df_t["label"], y=vals, name=col,
            mode="lines+markers",
            line=dict(color=color, width=3),
            marker=dict(size=8, color=color),
            text=[f"{v:{yfmt}}" for v in vals],
            hovertemplate="%{x}<br>" + col + ": %{text}<extra></extra>"))
    fig.update_layout(**PLOT, height=height, margin=PLOT_MARGIN,
        title=dict(text=title, font=dict(color="white", size=13)),
        yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                   tickformat=yfmt, title=ytitle),
        xaxis=dict(showgrid=False))
    return fig


# ── Session state ─────────────────────────────────────────────────────────────
if "data" not in st.session_state:
    st.session_state.data = None
if "manual_actions" not in st.session_state:
    st.session_state.manual_actions = []
if "manual_notes" not in st.session_state:
    st.session_state.manual_notes = []


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📊 SSPC Quality")
    st.markdown("---")
    st.markdown("### Upload DOR File")
    dor_file = st.file_uploader("DOR File (.xlsm / .xlsx)",
                                  type=["xlsm","xlsx"], key="dor")

    if dor_file:
        if st.button("▶  Process File", use_container_width=True, type="primary"):
            with st.spinner("Reading file and computing COPQ…"):
                file_bytes = dor_file.read()
                data, errors = extract_from_files(file_bytes, file_bytes)
            if errors:
                for e in errors: st.error(e)
            else:
                st.session_state.data = data
                st.success(f"✅ {len(data['months'])} months loaded")
    else:
        st.info("Upload the DOR file above, then click Process.")


    st.markdown("---")

    months_list = st.session_state.data["months"] \
                  if st.session_state.data else []

    if months_list:
        month_labels = [f"{calendar.month_abbr[m]} {y}"
                        for y, m in months_list]
        sel_idx = st.selectbox("Select Month",
                                range(len(month_labels)),
                                format_func=lambda i: month_labels[i],
                                index=len(month_labels) - 1)
        sel_year, sel_month = months_list[sel_idx]
        sel_label = month_labels[sel_idx]
        prev_tuple = months_list[sel_idx - 1] if sel_idx > 0 else None
    else:
        sel_year = sel_month = sel_idx = None
        sel_label = "—"; prev_tuple = None; month_labels = []

    st.markdown("---")
    st.markdown("**Facilities**")
    st.markdown("🔵 AuST Manufacturing")
    st.markdown("🟢 CenterPoint")


# ══════════════════════════════════════════════════════════════════════════════
# EMPTY STATE
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.data is None:
    st.markdown("## 🏭 SSPC Quality Dashboard")
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        st.markdown("""<div class="upload-box">
          <div style="font-size:48px;margin-bottom:12px">📂</div>
          <h3 style="color:#8fa3b8;margin:0 0 8px">Upload Your DOR Files</h3>
          <p style="color:#6a7f94;margin:0 0 16px">
            Upload the DOR file (contains both AuST &amp; CenterPoint)<br><br>
            then click <b>Process Files</b>.
          </p>
          <p style="color:#4a5a6a;font-size:13px;margin:0">
            All months load automatically · Supports .xlsm and .xlsx
          </p>
        </div>""", unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# DATA REFS
# ══════════════════════════════════════════════════════════════════════════════
D        = st.session_state.data
df_scrap = D["scrap"]
df_copq  = D["copq"]
df_prod  = D["prod"]
df_demand= D["demand"]
df_att   = D["att"]
df_actions = D["actions"]

curr_r = df_copq[(df_copq["year"] == sel_year) &
                  (df_copq["month"] == sel_month)] \
         if sel_year else pd.DataFrame()
prev_r = df_copq[(df_copq["year"] == prev_tuple[0]) &
                  (df_copq["month"] == prev_tuple[1])] \
         if prev_tuple else pd.DataFrame()

def cv(col):
    if curr_r.empty or col not in curr_r.columns: return None
    v = curr_r[col].iloc[0]
    return float(v) if pd.notna(v) else None

def pv(col):
    if prev_r.empty or col not in prev_r.columns: return None
    v = prev_r[col].iloc[0]
    return float(v) if pd.notna(v) else None


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs(["🏭 Overview", "💰 COPQ",
                "📅 Daily Tracker", "📋 Actions", "⏱ Hour by Hour", "📉 Downtime", "🎯 Goals"])


# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 1 — OVERVIEW
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown(f"### SSPC Overview — {sel_label}")
    st.markdown(
        f"<span style='color:#8fa3b8;font-size:13px'>"
        f"{len(months_list)} months loaded</span>",
        unsafe_allow_html=True)

    # KPIs
    sec("Key Performance Indicators")
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        v = (cv("aust_scrap") or 0) + (cv("cp_scrap") or 0)
        p = ((pv("aust_scrap") or 0) + (pv("cp_scrap") or 0)) \
            if prev_tuple else None
        kpi("Total Scrap Cost", fmt_c(v),
            delta_tag(v, p, False), "#e05c5c")
    with k2:
        cpp = cv("copq_per_part")
        kpi("CoPQ / Part",
            f"${cpp:.2f}" if cpp is not None else "—",
            delta_tag(cpp, pv("copq_per_part"), False, fmt="num"),
            "#e05c5c")
    with k3:
        kpi("Costed Yield", fmt_p(cv("costed_yield")),
            delta_tag(cv("costed_yield"), pv("costed_yield"),
                      True, fmt="pct"), "#00c853")
    with k4:
        kpi("AuST Good Cost", fmt_c(cv("aust_good")),
            delta_tag(cv("aust_good"), pv("aust_good"), True),
            "#2e75b6")
    with k5:
        kpi("CP Good Cost", fmt_c(cv("cp_good")),
            delta_tag(cv("cp_good"), pv("cp_good"), True),
            "#70ad47")

    st.markdown("")

    # Yield by product + rolling CoPQ/Part
    c1, c2 = st.columns(2)
    with c1:
        sec("Costed Yield by Product")
        prods  = ["BLN", "SMG", "CTI", "BMD", "MP3"]
        yields = [cv(f"{p.lower()}_yield") for p in prods]
        pcols  = [prod_color(p) for p in prods]
        fig2   = go.Figure()
        for p, y, c in zip(prods, yields, pcols):
            if y and y > 0:
                fig2.add_trace(go.Bar(
                    name=p, x=[p], y=[y], marker_color=c,
                    text=[fmt_p(y)], textposition="outside",
                    textfont=dict(color="white", size=12)))
        fig2.add_hline(y=0.90, line_dash="dash", line_color="#8fa3b8",
                       annotation_text="90% target",
                       annotation_font_color="#8fa3b8")
        fig2.update_layout(**PLOT, margin=PLOT_MARGIN, height=300, showlegend=False,
            yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                       tickformat=".0%", range=[0, 1.08]),
            xaxis=dict(showgrid=False))
        st.plotly_chart(fig2, use_container_width=True)

    with c2:
        sec("CoPQ/Part — Rolling 12 Months with Avg ± 1 Std Dev")
        df_roll, avg_cpp, std_cpp = get_rolling_stats(
            df_copq, "copq_per_part", 12)
        if not df_roll.empty:
            x_vals = list(df_roll["label"])
            y_vals = df_roll["copq_per_part"].fillna(0).tolist()
            fig_r = go.Figure()
            fig_r.add_trace(go.Scatter(
                x=x_vals + x_vals[::-1],
                y=[avg_cpp + std_cpp]*len(x_vals) +
                  [avg_cpp - std_cpp]*len(x_vals),
                fill="toself",
                fillcolor="rgba(46,117,182,0.15)",
                line=dict(color="rgba(0,0,0,0)"),
                showlegend=True, name="±1 Std Dev"))
            fig_r.add_trace(go.Scatter(
                x=x_vals, y=y_vals, name="CoPQ/Part",
                mode="lines+markers",
                line=dict(color="#ffc000", width=3),
                marker=dict(size=8),
                text=[f"${v:.2f}" for v in y_vals],
                hovertemplate="%{x}<br>%{text}<extra></extra>"))
            fig_r.add_hline(y=avg_cpp, line_dash="dash",
                            line_color="#8fa3b8",
                            annotation_text=f"Avg ${avg_cpp:.2f}",
                            annotation_font_color="#8fa3b8")
            fig_r.update_layout(**PLOT, margin=PLOT_MARGIN, height=300,
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                           tickprefix="$"),
                xaxis=dict(showgrid=False))
            st.plotly_chart(fig_r, use_container_width=True)

    # Defect Analysis
    sec(f"Defect Analysis — {sel_label}")
    top10 = get_top_defects(df_scrap, sel_year, sel_month)
    mdf   = df_scrap[(df_scrap["year"] == sel_year) &
                      (df_scrap["month"] == sel_month)]

    DC_AUST = {
        "Leak - valve":16.43,"Leak - bond":16.43,"Scratch":16.43,"Dirty":16.43,
        "Wire at Tip":11.08,"Destroyed Tip":11.08,
        "Fiber/ Embedded Particulate":13.38,"Skive":9.21,"Kink":16.43,
        "Burn":16.43,"Damaged/ Melted Hub":16.43,"Unknown Reflow Time":9.21,
        "Cut Short":16.43,"Pin Holes in Heat Shrink":9.21,"Tip Bleed":11.08,
        "Irregular Braid":1.82,"Hole at Tip":11.08,"Marker":16.43,
        "Wrong Valve Position":16.43,"Flash OD":11.08,"Stress Marks (Hub)":16.43,
        "Destructive Test":16.43,"Extrusions in Wrong Order":9.21,"Other":16.43,
        "Gap during Reflow":9.21,"Irregular Liner at Tip":16.43,
    }
    DC_CP = {k: 22.08 for k in DC_AUST}
    DC_CP.update({"Leak - valve":22.51,"Leak - bond":22.51})

    def _defect_cost(name):
        aq = float(mdf[mdf["entity"]=="AuST"][name].sum()) if name in mdf.columns else 0
        cq = float(mdf[mdf["entity"]=="CenterPoint"][name].sum()) if name in mdf.columns else 0
        return aq * DC_AUST.get(name,16.43) + cq * DC_CP.get(name,22.08)

    d1, d2 = st.columns([3, 2])
    with d1:
        if top10:
            names  = [x[0] for x in top10]
            values = [x[1] for x in top10]
            pcts   = [x[2] for x in top10]
            costs  = [_defect_cost(n) for n in names]
            bcols  = ["#c00000" if i==0 else "#2e75b6" if i<3 else "#4472c4"
                      for i in range(len(names))]
            fig3 = go.Figure(go.Bar(
                x=names, y=values,
                marker_color=bcols,
                text=[f"{int(v)}<br>({p:.1%})<br>${c:,.0f}"
                      for v,p,c in zip(values, pcts, costs)],
                textposition="outside",
                textfont=dict(color="white", size=10),
                hovertemplate="%{x}: %{y} units<extra></extra>"))
            fig3.update_layout(**PLOT, margin=PLOT_MARGIN, height=440,
                title=dict(text="Top 10 Failure Modes — Count · % · Cost",
                           font=dict(color="white", size=13)),
                xaxis=dict(showgrid=False, tickangle=-30, tickfont=dict(size=10)),
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a", title="Units"))
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No defect data for this month.")

    with d2:
        lv = float(mdf["Leak - valve"].sum())
        lb = float(mdf["Leak - bond"].sum())
        tl = lv + lb
        if tl > 0:
            st.markdown(f'''<div style="background:#1e2a3a;border-radius:10px;
                padding:16px;margin-bottom:12px;border:1px solid #2e4a6a">
              <div style="color:#8fa3b8;font-size:11px;font-weight:700;
                letter-spacing:1px;text-transform:uppercase;margin-bottom:10px">
                Leak Breakdown</div>
              <div style="display:flex;justify-content:space-between;margin-bottom:8px">
                <span style="color:#fff">Total Leak</span>
                <span style="color:#ffc000;font-weight:700;font-size:18px">{int(tl)}</span>
              </div>
              <div style="display:flex;justify-content:space-between;margin-bottom:6px">
                <span style="color:#8fa3b8;font-size:12px">↳ Leak — Valve</span>
                <span style="color:#e05c5c;font-weight:600">{int(lv)} ({lv/tl:.0%})</span>
              </div>
              <div style="display:flex;justify-content:space-between">
                <span style="color:#8fa3b8;font-size:12px">↳ Leak — Bond</span>
                <span style="color:#ff7c00;font-weight:600">{int(lb)} ({lb/tl:.0%})</span>
              </div>
            </div>''', unsafe_allow_html=True)
        if top10:
            df_t = pd.DataFrame(
                [(x[0], int(x[1]), f"{x[2]:.1%}", f"${_defect_cost(x[0]):,.0f}")
                 for x in top10],
                columns=["Defect Mode","Count","% of Total","Est. Cost"])
            df_t.index = range(1, len(df_t)+1)
            st.dataframe(df_t, use_container_width=True, height=340)

    # Product breakdown per failure mode
    if top10:
        sec("Scrap by Product — Top Failure Modes")
        products = ["BLN","SMG","CTI","BMD","MP3","NFE","BLT"]
        fig_prod = go.Figure()
        for prod in products:
            pm = mdf[mdf["product"]==prod]
            if pm.empty: continue
            yvals = [float(pm[n].sum()) if n in pm.columns else 0 for n,_,_ in top10]
            if sum(yvals)==0: continue
            fig_prod.add_trace(go.Bar(
                name=prod,
                x=[x[0] for x in top10],
                y=yvals,
                marker_color=prod_color(prod),
                hovertemplate=f"<b>{prod}</b><br>%{{x}}: %{{y:.0f}} units<extra></extra>",
            ))
        fig_prod.update_layout(**PLOT, margin=PLOT_MARGIN, height=360,
            barmode="stack",
            title=dict(text="Which Product Drives Each Failure Mode",
                       font=dict(color="white", size=13)),
            xaxis=dict(showgrid=False, tickangle=-25, tickfont=dict(size=10)),
            yaxis=dict(showgrid=True, gridcolor="#2a3a4a", title="Units"))
        st.plotly_chart(fig_prod, use_container_width=True)

    # 12-Month Trends
    sec("12-Month Trend Analysis")
    leak_trend    = get_leak_trend(df_scrap, df_copq)
    leak_valve_df = get_leak_valve_by_product(df_scrap, df_copq)
    leak_bond_df  = get_leak_bond_by_product(df_scrap, df_copq)

    t1, t2 = st.columns(2)
    with t1:
        if not leak_trend.empty:
            st.plotly_chart(
                line_fig(leak_trend, "Leak Rate — AuST vs CenterPoint (12 Months)",
                         "Leak Rate",
                         [("AuST","#2e75b6"),("CenterPoint","#70ad47")]),
                use_container_width=True)
        else:
            st.info("Need at least 2 months of data.")

    with t2:
        if not leak_valve_df.empty:
            prods = [c for c in ["BLN","SMG","CTI","BMD","MP3","NFE","BLT"]
                     if c in leak_valve_df.columns
                     and leak_valve_df[c].notna().any()]
            if prods:
                st.plotly_chart(
                    line_fig(leak_valve_df,
                             "Leak — Valve Rate by Product (12 Months)",
                             "Valve Leak Rate",
                             [(p, prod_color(p)) for p in prods]),
                    use_container_width=True)

    t3, t4 = st.columns(2)
    with t3:
        if not leak_bond_df.empty:
            prods = [c for c in ["BLN","SMG","CTI","BMD","MP3","NFE","BLT"]
                     if c in leak_bond_df.columns
                     and leak_bond_df[c].notna().any()]
            if prods:
                st.plotly_chart(
                    line_fig(leak_bond_df,
                             "Leak — Bond Rate by Product (12 Months)",
                             "Bond Leak Rate",
                             [(p, prod_color(p)) for p in prods]),
                    use_container_width=True)
    with t4:
        pass  # placeholder for future chart

    # Leak Summary KPIs
    sec("Leak Rate Summary")
    l1, l2, l3, l4 = st.columns(4)
    with l1:
        kpi("Total Leak — AuST",
            f"{int(cv('leak_aust') or 0)}",
            delta_tag(cv("leak_aust"), pv("leak_aust"),
                      False, fmt="num"), "#2e75b6")
    with l2:
        kpi("Total Leak — CP",
            f"{int(cv('leak_cp') or 0)}",
            delta_tag(cv("leak_cp"), pv("leak_cp"),
                      False, fmt="num"), "#70ad47")
    with l3:
        kpi("Leak Rate — AuST", fmt_p(cv("leak_rate_aust")),
            delta_tag(cv("leak_rate_aust"), pv("leak_rate_aust"),
                      False, fmt="pct"), "#2e75b6")
    with l4:
        kpi("Cumulative Leak Rate", fmt_p(cv("cumul_leak")),
            delta_tag(cv("cumul_leak"), pv("cumul_leak"),
                      False, fmt="pct"), "#e05c5c")


# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 2 — COPQ
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown(f"### Cost of Poor Quality — {sel_label}")

    sec("COPQ Summary")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi("AuST Scrap Cost", fmt_c(cv("aust_scrap")),
            delta_tag(cv("aust_scrap"), pv("aust_scrap"), False),
            "#e05c5c")
    with c2:
        kpi("CP Scrap Cost", fmt_c(cv("cp_scrap")),
            delta_tag(cv("cp_scrap"), pv("cp_scrap"), False),
            "#e05c5c")
    with c3:
        total_good = (cv("aust_good") or 0) + (cv("cp_good") or 0)
        prev_good  = (pv("aust_good") or 0) + (pv("cp_good") or 0)
        kpi("Total Good Cost", fmt_c(total_good),
            delta_tag(total_good, prev_good, True), "#00c853")
    with c4:
        kpi("Costed Yield", fmt_p(cv("costed_yield")),
            delta_tag(cv("costed_yield"), pv("costed_yield"),
                      True, fmt="pct"), "#00c853")

    st.markdown("")
    c1, c2 = st.columns(2)

    with c1:
        sec("Good vs Scrap Cost — AuST vs CP")
        cats   = ["AuST Good","AuST Scrap","CP Good","CP Scrap"]
        vals   = [cv("aust_good"), cv("aust_scrap"),
                  cv("cp_good"),   cv("cp_scrap")]
        colors = ["#2e75b6","#c00000","#70ad47","#ff7c00"]
        fig = go.Figure(go.Bar(
            x=cats, y=vals, marker_color=colors,
            text=[fmt_c(v) for v in vals],
            textposition="outside",
            textfont=dict(color="white", size=11)))
        fig.update_layout(**PLOT, margin=PLOT_MARGIN, height=320, showlegend=False,
            yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                       tickprefix="$", tickformat=",.0f"),
            xaxis=dict(showgrid=False))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        sec("CoPQ/Part — All Time with Avg ± Std Dev")
        df_all = df_copq.sort_values(["year","month"]).copy()
        df_all["label"] = df_all.apply(
            lambda r: f"{calendar.month_abbr[int(r.month)]} {int(r.year)}",
            axis=1)
        vals_all  = df_all["copq_per_part"].fillna(0).tolist()
        avg_all   = sum(vals_all)/len(vals_all) if vals_all else 0
        std_all   = pd.Series(vals_all).std() if len(vals_all) > 1 else 0
        bar_cols  = [
            "#ffc000" if lbl == sel_label else
            "#c00000" if v == max(vals_all) else
            "#00c853" if v == min(vals_all) else "#2e75b6"
            for lbl, v in zip(df_all["label"], vals_all)]
        fig_c = go.Figure()
        fig_c.add_trace(go.Bar(
            x=df_all["label"], y=vals_all,
            marker_color=bar_cols,
            text=[f"${v:.2f}" for v in vals_all],
            textposition="outside",
            textfont=dict(color="white", size=10),
            hovertemplate="%{x}<br>$%{y:.2f}<extra></extra>"))
        fig_c.add_hrect(
            y0=avg_all - std_all, y1=avg_all + std_all,
            fillcolor="rgba(46,117,182,0.12)", line_width=0,
            annotation_text="±1σ",
            annotation_font_color="#8fa3b8")
        fig_c.add_hline(y=avg_all, line_dash="dash",
                        line_color="#8fa3b8",
                        annotation_text=f"Avg ${avg_all:.2f}",
                        annotation_font_color="#8fa3b8")
        fig_c.update_layout(**PLOT, margin=PLOT_MARGIN, height=320, showlegend=False,
            yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                       tickprefix="$"),
            xaxis=dict(showgrid=False, tickangle=-35))
        st.plotly_chart(fig_c, use_container_width=True)

    # Scrap cost by product
    sec("Scrap Cost Breakdown by Product")
    mdf_s = df_scrap[(df_scrap["year"] == sel_year) &
                      (df_scrap["month"] == sel_month)]
    DC_AUST = {
        "Leak - valve":16.43,"Leak - bond":16.43,"Scratch":16.43,
        "Dirty":16.43,"Wire at Tip":11.08,"Destroyed Tip":11.08,
        "Fiber/ Embedded Particulate":13.38,"Skive":9.21,"Kink":16.43,
        "Burn":16.43,"Damaged/ Melted Hub":16.43,
        "Unknown Reflow Time":9.21,"Cut Short":16.43,
        "Pin Holes in Heat Shrink":9.21,"Tip Bleed":11.08,
        "Irregular Braid":1.82,"Hole at Tip":11.08,"Marker":16.43,
        "Wrong Valve Position":16.43,"Flash OD":11.08,
        "Stress Marks (Hub)":16.43,"Other":16.43,
    }
    DC_CP = {k: 22.08 for k in DC_AUST}
    DC_CP.update({"Leak - valve":22.51,"Leak - bond":22.51})
    prod_costs = {}
    for prod in ["BLN","SMG","CTI","BMD","MP3"]:
        prows = mdf_s[mdf_s["product"] == prod]
        cost  = 0.0
        for _, row in prows.iterrows():
            dc = DC_AUST if row["entity"] == "AuST" else DC_CP
            for col, rate in dc.items():
                q = float(row.get(col, 0) or 0)
                if q > 0: cost += q * rate
        prod_costs[prod] = cost
    fig_prod = go.Figure(go.Bar(
        x=list(prod_costs.keys()),
        y=list(prod_costs.values()),
        marker_color=[prod_color(p) for p in ["BLN","SMG","CTI","BMD","MP3"]],
        text=[fmt_c(v) for v in prod_costs.values()],
        textposition="outside",
        textfont=dict(color="white", size=12)))
    fig_prod.update_layout(**PLOT, margin=PLOT_MARGIN, height=300, showlegend=False,
        yaxis=dict(showgrid=True, gridcolor="#2a3a4a", tickprefix="$"),
        xaxis=dict(showgrid=False))
    st.plotly_chart(fig_prod, use_container_width=True)

    # Weighted Average COPQ View
    sec("COPQ Normalised View — Weighted by Volume")
    st.markdown("""<div style="color:#8fa3b8;font-size:12px;margin-bottom:12px">
      Because production volume varies month to month, simple averages overstate
      low-volume months. The charts below weight each month's COPQ by the number
      of good CP units built that month, giving a volume-adjusted picture.
    </div>""", unsafe_allow_html=True)

    df_wt = df_copq.copy()
    df_wt["label"] = df_wt.apply(
        lambda r: f"{calendar.month_abbr[int(r.month)]} {int(r.year)}", axis=1)
    df_wt = df_wt[df_wt["good_cp"].notna() & (df_wt["good_cp"] > 0)].copy()

    if not df_wt.empty:
        # Weighted CoPQ/Part = total_scrap_cost / total_good_cp_units
        total_scrap  = df_wt["total_scrap"].fillna(0).sum()
        total_good   = df_wt["good_cp"].sum()
        weighted_cpp = total_scrap / total_good if total_good > 0 else 0

        # Simple avg for comparison
        simple_avg   = df_wt["copq_per_part"].dropna().mean()
        diff         = weighted_cpp - simple_avg

        wk1, wk2, wk3 = st.columns(3)
        with wk1:
            kpi("Weighted CoPQ/Part",
                f"${weighted_cpp:.2f}",
                f'<div class="kpi-delta-neutral">Weighted by CP good units</div>',
                "#ffc000")
        with wk2:
            kpi("Simple Average CoPQ/Part",
                f"${simple_avg:.2f}",
                f'<div class="kpi-delta-neutral">Unweighted monthly avg</div>',
                "#2e75b6")
        with wk3:
            col = "#00c853" if diff < 0 else "#e05c5c"
            diff_label = ("Weighted is lower — high-volume months are better quality"
                          if diff < 0 else
                          "Weighted is higher — high-volume months have worse quality")
            kpi("Difference",
                f"${abs(diff):.2f}",
                f'<div style="color:{col};font-size:11px;margin-top:4px">'
                f'{diff_label}</div>',
                col)

        # Scrap cost per good unit over time (normalised)
        df_wt["scrap_per_unit"] = df_wt["total_scrap"] / df_wt["good_cp"]
        import numpy as np
        vals = df_wt["scrap_per_unit"].fillna(0).tolist()
        avg  = float(np.mean(vals)); std = float(np.std(vals))

        wc1, wc2 = st.columns(2)
        with wc1:
            sec("Scrap Cost per Good Unit — All Months")
            fig_wt = go.Figure()
            bar_cols = ["#ffc000" if lbl == sel_label else
                        "#c00000" if v == max(vals) else
                        "#00c853" if v == min(vals) else "#2e75b6"
                        for lbl, v in zip(df_wt["label"], vals)]
            fig_wt.add_trace(go.Bar(
                x=df_wt["label"], y=vals,
                marker_color=bar_cols,
                text=[f"${v:.2f}" for v in vals],
                textposition="outside",
                textfont=dict(color="white", size=9),
                hovertemplate="%{x}<br>$%{y:.2f}/unit<extra></extra>"))
            fig_wt.add_hrect(y0=avg-std, y1=avg+std,
                             fillcolor="rgba(46,117,182,0.12)", line_width=0)
            fig_wt.add_hline(y=avg, line_dash="dash", line_color="#8fa3b8",
                             annotation_text=f"Avg ${avg:.2f}",
                             annotation_font_color="#8fa3b8")
            fig_wt.update_layout(**PLOT, margin=PLOT_MARGIN, height=320,
                showlegend=False,
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a", tickprefix="$"),
                xaxis=dict(showgrid=False, tickangle=-40,
                           tickfont=dict(size=9)))
            st.plotly_chart(fig_wt, use_container_width=True)

        with wc2:
            sec("Total Scrap Cost vs Good Units Built")
            fig_sc = go.Figure()
            fig_sc.add_trace(go.Bar(
                name="Total Scrap $",
                x=df_wt["label"], y=df_wt["total_scrap"].fillna(0),
                marker_color="#e05c5c",
                hovertemplate="%{x}<br>Scrap: $%{y:,.0f}<extra></extra>",
                yaxis="y"))
            fig_sc.add_trace(go.Scatter(
                name="Good CP Units",
                x=df_wt["label"], y=df_wt["good_cp"],
                mode="lines+markers",
                line=dict(color="#70ad47", width=2),
                marker=dict(size=6),
                hovertemplate="%{x}<br>Good Units: %{y:,.0f}<extra></extra>",
                yaxis="y2"))
            fig_sc.update_layout(**PLOT, margin=PLOT_MARGIN, height=320,
                barmode="group",
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                           tickprefix="$", title="Scrap Cost"),
                yaxis2=dict(overlaying="y", side="right",
                            showgrid=False, title="Good Units"),
                xaxis=dict(showgrid=False, tickangle=-40,
                           tickfont=dict(size=9)))
            st.plotly_chart(fig_sc, use_container_width=True)

        # Scrap cost breakdown table — normalised
        sec("Monthly COPQ — Normalised Detail")
        df_tbl = df_wt[["label","total_scrap","good_cp","scrap_per_unit",
                         "costed_yield","cumul_leak"]].copy()
        df_tbl.columns = ["Month","Total Scrap","Good CP Units",
                          "Scrap $/Unit","Costed Yield","Cumul. Leak"]
        df_tbl["Total Scrap"]   = df_tbl["Total Scrap"].apply(fmt_c)
        df_tbl["Good CP Units"] = df_tbl["Good CP Units"].apply(
            lambda v: f"{int(v):,}" if pd.notna(v) else "—")
        df_tbl["Scrap $/Unit"]  = df_tbl["Scrap $/Unit"].apply(
            lambda v: f"${v:.2f}" if pd.notna(v) else "—")
        df_tbl["Costed Yield"]  = df_tbl["Costed Yield"].apply(fmt_p)
        df_tbl["Cumul. Leak"]   = df_tbl["Cumul. Leak"].apply(fmt_p)
        df_tbl = df_tbl.sort_values("Month", ascending=False)
        df_tbl.index = range(1, len(df_tbl)+1)
        st.dataframe(df_tbl, use_container_width=True, hide_index=True)
    else:
        st.info("Not enough data for weighted analysis.")

# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 3 — DAILY TRACKER
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown(f"### Daily Production Tracker — {sel_label}")

    if df_prod.empty:
        st.info("No daily production data found in the uploaded file.")
    else:
        mdf_prod = df_prod[(df_prod["year"] == sel_year) &
                            (df_prod["month"] == sel_month)]

        # Shot to target
        sec("Shot to Target — MTD Plan vs Actual")
        if not df_demand.empty:
            demand_m = df_demand[
                (df_demand["year"] == sel_year) &
                (df_demand["month"] == sel_month)]
            for facility in ["AuST","CenterPoint"]:
                fac_prod = mdf_prod[mdf_prod["facility"] == facility]
                fac_dem  = demand_m[demand_m["facility"] == facility]
                if fac_prod.empty: continue
                actual_by_prod = fac_prod.groupby("product")["qty"].sum()
                total_actual   = int(actual_by_prod.sum())
                total_plan_row = fac_dem["plan"].sum()
                total_plan     = int(total_plan_row) if total_plan_row > 0 else None

                # Header row with facility name and totals
                total_pct = total_actual/total_plan if total_plan else None
                total_col = ("#00c853" if total_pct and total_pct >= 1 else
                             "#ff7c00" if total_pct and total_pct >= 0.9 else
                             "#e05c5c")
                st.markdown(
                    f'''<div style="display:flex;align-items:center;
                        justify-content:space-between;margin-bottom:8px">
                      <b style="color:#fff;font-size:14px">{facility}</b>
                      <span style="color:{total_col};font-size:13px;font-weight:600">
                        Total: {total_actual:,} built
                        {f"/ {total_plan:,} planned ({total_pct:.1%})"
                         if total_plan else "/ no plan set"}
                      </span>
                    </div>''', unsafe_allow_html=True)

                col_list = list(actual_by_prod.items())
                if col_list:
                    cols = st.columns(len(col_list))
                    for i, (prod, actual) in enumerate(col_list):
                        plan_row = fac_dem[fac_dem["product"] == prod]["plan"]
                        plan = float(plan_row.iloc[0])                                if not plan_row.empty and float(plan_row.iloc[0]) > 0                                else None
                        pct  = actual/plan if plan else None
                        # Color: green=at/above plan, orange=90-99%, red=below 90%
                        # If no plan: use product color
                        perf_color = ("#00c853" if pct and pct >= 1 else
                                      "#ff7c00" if pct and pct >= 0.9 else
                                      "#e05c5c")
                        accent = perf_color if plan else prod_color(prod)
                        if plan:
                            delta = (f'<div style="color:{perf_color};'
                                     f'font-size:11px;margin-top:4px">'
                                     f'{"✅" if pct >= 1 else "⚠️"} '
                                     f'{pct:.1%} of {int(plan):,} planned'
                                     f'</div>')
                        else:
                            delta = (f'<div style="color:{prod_color(prod)};'
                                     f'font-size:11px;margin-top:4px">'
                                     f'No plan — {actual/total_actual:.1%} of total'
                                     f'</div>')
                        with cols[i]:
                            kpi(prod, f"{int(actual):,}", delta, accent=accent)
                st.markdown("<div style='margin-bottom:16px'></div>",
                            unsafe_allow_html=True)
        else:
            st.info("No demand plan data found.")
        # Day vs Swing production
        sec("Daily Production — Day vs Swing Shift")
        daily = mdf_prod.groupby(["date","facility","shift"]).agg(
            qty=("qty","sum")).reset_index()

        for facility in ["AuST","CenterPoint"]:
            fac = daily[daily["facility"] == facility]
            if fac.empty: continue
            fig_daily = go.Figure()
            for shift, color in [("Day","#2e75b6"),("Swing","#70ad47")]:
                s = fac[fac["shift"] == shift].sort_values("date")
                if s.empty: continue
                fig_daily.add_trace(go.Bar(
                    x=s["date"].apply(
                        lambda d: f"{d.month}/{d.day}"),
                    y=s["qty"], name=shift,
                    marker_color=color,
                    hovertemplate=(
                        f"{shift}<br>%{{x}}: %{{y:,}} units"
                        "<extra></extra>")))
            fig_daily.update_layout(**PLOT, margin=PLOT_MARGIN, height=280,
                barmode="group",
                title=dict(
                    text=f"{facility} — Daily Production by Shift",
                    font=dict(color="white", size=13)),
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                           title="Units"),
                xaxis=dict(showgrid=False, title="Date"))
            st.plotly_chart(fig_daily, use_container_width=True)

        # Reject rate
        sec("Daily Reject Rate — Day vs Swing")
        daily_rej = mdf_prod.groupby(["date","facility","shift"]).agg(
            qty=("qty","sum"),
            rejects=("rejects","sum")).reset_index()
        daily_rej["reject_rate"] = daily_rej.apply(
            lambda r: r["rejects"]/r["qty"] if r["qty"] > 0 else 0,
            axis=1)
        for facility in ["AuST","CenterPoint"]:
            fac = daily_rej[daily_rej["facility"] == facility]
            if fac.empty: continue
            fig_rr = go.Figure()
            for shift, color in [("Day","#2e75b6"),("Swing","#70ad47")]:
                s = fac[fac["shift"] == shift].sort_values("date")
                if s.empty: continue
                fig_rr.add_trace(go.Scatter(
                    x=s["date"].apply(
                        lambda d: f"{d.month}/{d.day}"),
                    y=s["reject_rate"], name=shift,
                    mode="lines+markers",
                    line=dict(color=color, width=2),
                    marker=dict(size=6),
                    text=[f"{v:.1%}" for v in s["reject_rate"]],
                    hovertemplate=(
                        f"{shift}<br>%{{x}}: %{{text}}"
                        "<extra></extra>")))
            fig_rr.update_layout(**PLOT, margin=PLOT_MARGIN, height=240,
                title=dict(
                    text=f"{facility} — Reject Rate by Shift",
                    font=dict(color="white", size=13)),
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                           tickformat=".1%"),
                xaxis=dict(showgrid=False))
            st.plotly_chart(fig_rr, use_container_width=True)

        # Attendance
        sec("Attendance — " + sel_label)
        if not df_att.empty:
            att_m = df_att[(df_att["year"] == sel_year) &
                            (df_att["month"] == sel_month)]
            if not att_m.empty:
                pivot = att_m.pivot_table(
                    index="team", columns="day",
                    values="status", aggfunc="first")
                pivot.columns = [f"Day {c}" for c in pivot.columns]
                st.dataframe(pivot, use_container_width=True)
                counts = att_m.groupby("status").size().reset_index(
                    name="count")
                total  = counts["count"].sum()
                c1, c2, c3 = st.columns(3)
                for i, (status, color) in enumerate([
                    ("On Time","#00c853"),
                    ("Late","#ffc000"),
                    ("No Show","#e05c5c")]):
                    row = counts[counts["status"] == status]
                    cnt = int(row["count"].iloc[0]) \
                          if not row.empty else 0
                    with [c1, c2, c3][i]:
                        kpi(status, str(cnt),
                            f'<div class="kpi-delta-neutral">'
                            f'{cnt/total:.0%} of records</div>'
                            if total > 0 else "",
                            accent=color)
            else:
                st.info("No attendance data for this month.")
        else:
            st.info("No attendance sheet found.")

        # Shift notes
        sec("Shift Notes")
        notes_df = mdf_prod[mdf_prod["notes"].str.len() > 0][
            ["date","facility","shift","people","hours","notes"]
        ].drop_duplicates(
            subset=["date","facility","shift"]
        ).sort_values(["date","facility","shift"])

        if not notes_df.empty:
            for _, row in notes_df.iterrows():
                dt = row["date"]
                dt_str = (f"{dt.month}/{dt.day}"
                          if hasattr(dt, "month") else str(dt))
                st.markdown(f"""<div class="note-box">
                  <b>{row['facility']} — {row['shift']} — {dt_str}</b>
                  &nbsp;|&nbsp;
                  👥 {int(row['people']) if row['people'] else '?'} people
                  &nbsp;|&nbsp;
                  ⏰ {int(row['hours']) if row['hours'] else '?'} hrs<br>
                  <span style="color:#8fa3b8">{row['notes']}</span>
                </div>""", unsafe_allow_html=True)
        else:
            st.info("No shift notes for this month.")

        # Manual log form
        sec("Log a Safety or Quality Issue")
        with st.form("daily_log_form", clear_on_submit=True):
            cols = st.columns([1,1,2,2,1])
            log_date     = cols[0].date_input("Date",
                            value=datetime.date.today())
            log_type     = cols[1].selectbox(
                "Type",["Safety","Quality","Service","Cost","People"])
            log_desc     = cols[2].text_input("Description")
            log_action   = cols[3].text_input("Action / Task")
            log_assigned = cols[4].text_input("Assigned To")
            submitted    = st.form_submit_button(
                "➕ Add Entry", use_container_width=True)
            if submitted and log_desc:
                st.session_state.manual_notes.append({
                    "date":str(log_date),"type":log_type,
                    "description":log_desc,"action":log_action,
                    "assigned":log_assigned,"status":"Open"})
                st.success("Entry added!")

        if st.session_state.manual_notes:
            st.dataframe(
                pd.DataFrame(st.session_state.manual_notes),
                use_container_width=True, hide_index=True)


# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 4 — ACTIONS
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown("### 📋 Actions Log")

    # Add new action
    sec("Add New Morning Meeting Action")
    with st.form("action_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        a_date   = c1.date_input("Date", value=datetime.date.today())
        a_cat    = c2.selectbox("Category",
            ["Quality","Safety","Service","Cost","People","Projects"])
        a_urg    = c3.selectbox("Urgency",
            ["Low","Medium","High","Critical"])
        c4, c5   = st.columns(2)
        a_desc   = c4.text_input("Description")
        a_task   = c5.text_input("Task / Action Required")
        c6, c7, c8 = st.columns(3)
        a_assign = c6.text_input("Assigned To")
        a_due    = c7.date_input("Due Date",
                    value=datetime.date.today())
        a_status = c8.selectbox("Status",
            ["Open","In Progress","Complete"])
        add_btn  = st.form_submit_button(
            "➕ Add Action", type="primary",
            use_container_width=True)
        if add_btn and a_desc:
            st.session_state.manual_actions.append({
                "date":str(a_date),"category":a_cat,
                "urgency":a_urg,"description":a_desc,
                "task":a_task,"assigned_to":a_assign,
                "due_date":str(a_due),"status":a_status,
                "source":"Manual"})
            st.success(f"Action added: {a_desc}")

    # Filters
    sec("View & Filter Actions")
    fc1, fc2, fc3 = st.columns(3)
    f_cat    = fc1.multiselect("Category",
        ["Quality","Safety","Service","Cost","People","Projects"],
        default=[])
    f_status = fc2.multiselect("Status",
        ["Open","In Progress","Complete"],
        default=["Open","In Progress"])
    f_search = fc3.text_input("Search description")

    # Combine Excel + manual actions
    all_actions = []
    if not df_actions.empty:
        for _, row in df_actions.iterrows():
            all_actions.append({
                "date": str(row["date"])[:10]
                        if row["date"] else "",
                "category":    row["category"],
                "urgency":     row.get("urgency",""),
                "description": row["description"],
                "task":        row["task"],
                "assigned_to": row["assigned_to"],
                "due_date":    str(row["due_date"])[:10]
                               if row["due_date"] else "",
                "status":      row["status"],
                "source":      "Excel"})
    all_actions += st.session_state.manual_actions
    df_all_act = pd.DataFrame(all_actions) \
                 if all_actions else pd.DataFrame()

    if not df_all_act.empty:
        filtered = df_all_act.copy()
        if f_cat:
            filtered = filtered[filtered["category"].isin(f_cat)]
        if f_status:
            filtered = filtered[
                filtered["status"].str.lower().isin(
                    [s.lower() for s in f_status])]
        if f_search:
            filtered = filtered[
                filtered["description"].str.contains(
                    f_search, case=False, na=False) |
                filtered["task"].str.contains(
                    f_search, case=False, na=False)]

        st.markdown(f"**{len(filtered)} actions** shown")

        for _, row in filtered.sort_values(
                "date", ascending=False).iterrows():
            status = str(row.get("status",""))
            if "complete" in status.lower():
                tag = '<span class="tag-done">Complete</span>'
            elif "progress" in status.lower():
                tag = '<span class="tag-prog">In Progress</span>'
            else:
                tag = '<span class="tag-open">Open</span>'
            urg_color = {
                "Critical":"#e05c5c","High":"#ff7c00",
                "Medium":"#ffc000","Low":"#8fa3b8"
            }.get(str(row.get("urgency","")), "#8fa3b8")

            st.markdown(f"""<div style="background:#1e2a3a;
                border-radius:8px;padding:12px 16px;margin-bottom:8px;
                border-left:3px solid {urg_color}">
              <div style="display:flex;justify-content:space-between;
                align-items:center;margin-bottom:4px">
                <span style="color:#fff;font-weight:600">
                  {row.get('description','')}</span>
                <span>{tag}</span>
              </div>
              <div style="color:#8fa3b8;font-size:12px">
                📅 {row.get('date','')} &nbsp;|&nbsp;
                🏷️ {row.get('category','')} &nbsp;|&nbsp;
                👤 {row.get('assigned_to','')} &nbsp;|&nbsp;
                🗓️ Due: {row.get('due_date','')}
              </div>
              <div style="color:#c0cfe0;font-size:12px;margin-top:4px">
                → {row.get('task','')}
              </div>
            </div>""", unsafe_allow_html=True)

        sec("Action Summary")
        s1, s2, s3 = st.columns(3)
        open_ct = len(df_all_act[
            df_all_act["status"].str.lower() == "open"])
        prog_ct = len(df_all_act[
            df_all_act["status"].str.lower().str.contains(
                "progress", na=False)])
        done_ct = len(df_all_act[
            df_all_act["status"].str.lower().str.contains(
                "complete", na=False)])
        with s1: kpi("Open", str(open_ct), accent="#e05c5c")
        with s2: kpi("In Progress", str(prog_ct), accent="#ffc000")
        with s3: kpi("Complete", str(done_ct), accent="#00c853")
    else:
        st.info("No actions found. Add one above.")


# ╔══════════════════════════════════════════════════════════════════════════════
# ╔══════════════════════════════════════════════════════════════════════════════
# ║  SHARED HxH PARSER
# ╚══════════════════════════════════════════════════════════════════════════════
def _parse_hxh_bytes(content_bytes):
    import io as _io, zipfile as _zf
    def _parse_one(raw_bytes):
        df = pd.read_csv(_io.BytesIO(raw_bytes), encoding="latin1", header=None)
        def sf(v):
            if pd.isna(v): return None
            s = str(v).strip()
            if s in ("","nan","#REF!","#VALUE!"): return None
            try: return float(s)
            except: return s
        r2 = df.iloc[1]
        meta = {
            "date":str(sf(r2[8]) or ""),"lot":str(sf(r2[13]) or ""),
            "product":str(sf(r2[19]) or ""),"shift":str(sf(r2[23]) or ""),
            "duration":str(sf(r2[28]) or ""),
        }
        is_aust = any("Liner Prep" in str(df.iloc[r][1]) for r in range(24,35) if pd.notna(df.iloc[r][1]))
        facility = "AuST" if is_aust else "CenterPoint"
        present = absent = 0
        for ridx in range(6,12):
            row = df.iloc[ridx]
            for cidx,val in enumerate(row):
                if str(val).strip() == "Total Operators Present":
                    nrow = df.iloc[ridx+1] if ridx+1<len(df) else None
                    if nrow is not None:
                        v = sf(nrow[cidx])
                        try: present = int(float(v)) if v else present
                        except: pass
                if str(val).strip().startswith("Total Operators Absent"):
                    nrow = df.iloc[ridx+1] if ridx+1<len(df) else None
                    if nrow is not None:
                        v = sf(nrow[cidx])
                        try: absent = int(float(v)) if v else absent
                        except: pass
        # Downtime code legend
        dt_legend = {}
        desc_col = 15 if is_aust else 18
        code_col = 19 if is_aust else 22
        for ridx in range(47,90):
            row = df.iloc[ridx]
            desc = str(sf(row[desc_col]) or "") if desc_col < len(row) else ""
            code = str(sf(row[code_col]) or "") if code_col < len(row) else ""
            if desc and code and desc != "nan" and code != "nan":
                dt_legend[code.strip()] = desc.strip()
        operators = []
        for ridx in range(4,18):
            row = df.iloc[ridx]
            for cidx,val in enumerate(row):
                cert = str(val).strip()
                if cert in ("Certified","Training") and cidx+1<len(row):
                    station = str(sf(row[cidx+1]) or "")
                    if not station or station=="nan": continue
                    nrow = df.iloc[ridx+1] if ridx+1<len(df) else None
                    name = str(sf(nrow[cidx+1]) or "") if nrow is not None else ""
                    if name and name!="nan":
                        operators.append({"station":station,"name":name,"status":cert})
        hour_slots = list(range(5,90,6))
        stations = []
        for ridx in range(24,36):
            row = df.iloc[ridx]
            initials = str(sf(row[0]) or "")
            station  = str(sf(row[1]) or "")
            if not initials or not station: continue
            if initials in ("nan","Final","Total","Rejects","Reject"): continue
            total   = sf(row[90]); rej_tot = sf(row[92]); oee = sf(row[97])
            try: total   = int(float(total))   if total   else 0
            except: total = 0
            try: rej_tot = int(float(rej_tot)) if rej_tot else 0
            except: rej_tot = 0
            try: oee = float(oee) if oee is not None else None
            except: oee = None
            hours = []
            hn = 1
            for h in hour_slots:
                actual = sf(row[h]) if h<len(row) else None
                if actual is None: continue
                try: actual = int(float(actual))
                except: continue
                if actual <= 0: continue
                rejects  = sf(row[h+1]) if h+1<len(row) else None
                downtime = sf(row[h+3]) if h+3<len(row) else None
                code     = sf(row[h+4]) if h+4<len(row) else None
                try: rejects  = int(float(rejects))  if rejects  else 0
                except: rejects = 0
                try: downtime = int(float(downtime)) if downtime else 0
                except: downtime = 0
                code_str = str(code).strip() if code else ""
                hours.append({"hour":hn,"actual":actual,"rejects":rejects,
                              "downtime_min":downtime,"downtime_code":code_str,
                              "downtime_label":dt_legend.get(code_str,code_str)})
                hn += 1
            if not hours and total==0: continue
            stations.append({"initials":initials,"station":station,"total":total,
                              "total_rejects":rej_tot,"oee":oee,"hours":hours})
        failure_start = total_col = None
        for ridx in range(45,90):
            row = df.iloc[ridx]
            for cidx,val in enumerate(row):
                if str(val).strip()=="Failure Mode": failure_start=ridx+2
                if str(val).strip()=="Total Rejects by Failure Code": total_col=cidx
            if failure_start and total_col: break
        failure_modes = {}
        if failure_start and total_col:
            for ridx in range(failure_start,min(failure_start+40,len(df))):
                row = df.iloc[ridx]
                mode = str(sf(row[0]) or "")
                if not mode or mode in ("nan","Daily Total Rejects"): continue
                v = sf(row[total_col])
                if v is None: continue
                try:
                    n = int(float(v))
                    if n>0: failure_modes[mode]=n
                except: pass
        notes = []
        for ridx in range(80,len(df)):
            if str(sf(df.iloc[ridx][0]) or "").strip()=="Hour By Hour Notes":
                for r2i in range(ridx+1,len(df)):
                    r2row = df.iloc[r2i]
                    t = str(sf(r2row[0]) or "")
                    c = str(sf(r2row[5]) or "") if len(r2row)>5 else ""
                    if t and c and t!="nan" and c!="nan":
                        notes.append({"time":t,"note":c})
                break
        return {"facility":facility,**meta,"present":present,"absent":absent,
                "dt_legend":dt_legend,"operators":operators,"stations":stations,
                "failure_modes":failure_modes,"notes":notes}

    # Handle zip or single CSV
    results = []
    try:
        with _zf.ZipFile(_io.BytesIO(content_bytes)) as zf:
            csv_names = sorted([n for n in zf.namelist() if n.lower().endswith(".csv")])
            for name in csv_names:
                try:
                    raw = zf.read(name)
                    p = _parse_one(raw)
                    p["filename"] = name.split("/")[-1]
                    results.append(p)
                except Exception: pass
    except _zf.BadZipFile:
        try:
            p = _parse_one(content_bytes)
            p["filename"] = ""
            results.append(p)
        except Exception as e:
            pass
    return results


# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 5 — HOUR BY HOUR
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown("### ⏱ Hour by Hour Tracker")
    st.markdown("""<div class="note-box">
      📂 Upload the <b>AuST zip</b> and <b>CP zip</b> separately below — one zip per facility,
      each containing all shift CSVs for the month.
      &nbsp;|&nbsp; The DOR <code>.xlsm</code> file goes in the sidebar.
    </div>""", unsafe_allow_html=True)

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown("**🔵 AuST HxH Zip**")
        hxh_aust_f = st.file_uploader("AuST zip folder",
                                       type=["zip","csv"], key="hxh_aust")
    with col_r:
        st.markdown("**🟢 CenterPoint HxH Zip**")
        hxh_cp_f   = st.file_uploader("CP zip folder",
                                       type=["zip","csv"], key="hxh_cp")

    if "hxh_parsed" not in st.session_state:
        st.session_state.hxh_parsed = []

    for f_obj in [hxh_aust_f, hxh_cp_f]:
        if f_obj:
            results = _parse_hxh_bytes(f_obj.read())
            added = 0
            for r in results:
                key = (r["facility"], r["date"], r["shift"])
                if not any((p["facility"],p["date"],p["shift"])==key
                           for p in st.session_state.hxh_parsed):
                    st.session_state.hxh_parsed.append(r)
                    added += 1
            if added:
                st.success(f"✅ {f_obj.name} — loaded {added} new shift(s)")
            else:
                st.info(f"{f_obj.name} — all shifts already loaded.")

    parsed_all = st.session_state.hxh_parsed

    if not parsed_all:
        st.info("Upload AuST and/or CP HxH files above to view shift analysis.")
    else:
        # ── Filters ───────────────────────────────────────────────────────────
        fc1, fc2, fc3, fc4 = st.columns(4)
        all_facilities = sorted(set(p["facility"] for p in parsed_all))
        all_dates      = sorted(set(p["date"]     for p in parsed_all))
        all_shifts     = sorted(set(p["shift"]     for p in parsed_all))

        f_fac   = fc1.multiselect("Facility",  all_facilities, default=all_facilities, key="hxh_fac")
        f_date  = fc2.multiselect("Date",       all_dates,      default=all_dates,      key="hxh_date")
        f_shift = fc3.multiselect("Shift",      all_shifts,     default=all_shifts,     key="hxh_shift")
        f_prod  = fc4.text_input("Product filter", placeholder="e.g. BLN", key="hxh_prod")

        filtered_all = [p for p in parsed_all
                        if p["facility"] in f_fac
                        and p["date"]     in f_date
                        and p["shift"]    in f_shift
                        and (not f_prod or f_prod.upper() in p["product"].upper())]

        if not filtered_all:
            st.info("No shifts match the current filters.")
        else:
            labels = [f"{p['facility']} | {p['date']} | {p['shift']} | {p['product']}"
                      for p in filtered_all]
            sel = st.selectbox("Select Shift to View", range(len(labels)),
                               format_func=lambda i: labels[i],
                               index=len(labels)-1, key="hxh_sel")
            p = filtered_all[sel]

            shift_color = {"Days":"#2e75b6","Swings":"#70ad47","Weekends":"#9b59b6"}
            color = shift_color.get(p["shift"], "#2e75b6")
            st.markdown(f"""<div style="background:#1e2a3a;border-radius:10px;
                padding:14px 18px;margin-bottom:16px;border-left:4px solid {color}">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <span style="color:#fff;font-weight:700;font-size:16px">
                  {p["facility"]} — {p["date"]} &nbsp;|&nbsp;
                  <span style="color:{color}">{p["shift"]} Shift</span>
                </span>
                <span style="color:#8fa3b8;font-size:13px">
                  {p["product"]} &nbsp;|&nbsp; Lot: {p["lot"]} &nbsp;|&nbsp; {p["duration"]}
                </span>
              </div>
              <div style="color:#8fa3b8;font-size:12px;margin-top:6px">
                👥 {p["present"]} present &nbsp;|&nbsp; ❌ {p["absent"]} absent
              </div>
            </div>""", unsafe_allow_html=True)

            if p["stations"]:
                # ── Hour-by-hour production ───────────────────────────────────
                sec(f"Hour-by-Hour Production — {p['facility']} {p['shift']} Shift")
                fig_hxh = go.Figure()
                for s in p["stations"]:
                    if not s["hours"]: continue
                    fig_hxh.add_trace(go.Bar(
                        name=s["station"],
                        x=[f"Hour {h['hour']}" for h in s["hours"]],
                        y=[h["actual"] for h in s["hours"]],
                        hovertemplate=(f"<b>{s['station']}</b><br>"
                                       "%{x}<br>Actual: %{y}<extra></extra>"),
                    ))
                fig_hxh.update_layout(**PLOT, margin=PLOT_MARGIN, height=350, barmode="group",
                    yaxis=dict(showgrid=True, gridcolor="#2a3a4a", title="Units"),
                    xaxis=dict(showgrid=False),
                    title=dict(text="Units Produced per Hour by Station",
                               font=dict(color="white", size=13)))
                st.plotly_chart(fig_hxh, use_container_width=True)

                # ── Rejects per hour ──────────────────────────────────────────
                has_rejects = any(any(h["rejects"]>0 for h in s["hours"])
                                  for s in p["stations"])
                if has_rejects:
                    sec("Rejects per Hour by Station")
                    fig_rej = go.Figure()
                    for s in p["stations"]:
                        if not any(h["rejects"]>0 for h in s["hours"]): continue
                        fig_rej.add_trace(go.Bar(
                            name=s["station"],
                            x=[f"Hour {h['hour']}" for h in s["hours"]],
                            y=[h["rejects"] for h in s["hours"]],
                            hovertemplate=(f"<b>{s['station']}</b><br>"
                                           "%{x}<br>Rejects: %{y}<extra></extra>"),
                        ))
                    fig_rej.update_layout(**PLOT, margin=PLOT_MARGIN, height=280, barmode="stack",
                        yaxis=dict(showgrid=True, gridcolor="#2a3a4a", title="Rejects"),
                        xaxis=dict(showgrid=False))
                    st.plotly_chart(fig_rej, use_container_width=True)

                # ── Defects by product (HxH) ──────────────────────────────────
                if p["failure_modes"]:
                    sec("Defects by Product — This Shift")
                    c1, c2 = st.columns(2)
                    with c1:
                        # Failure mode breakdown table
                        df_fm = pd.DataFrame(
                            sorted(p["failure_modes"].items(), key=lambda x:-x[1]),
                            columns=["Failure Mode","Count"])
                        total_fm = df_fm["Count"].sum()
                        df_fm["% of Rejects"] = df_fm["Count"].apply(
                            lambda v: f"{v/total_fm:.1%}" if total_fm>0 else "—")
                        df_fm.index = range(1, len(df_fm)+1)
                        st.markdown(f'<div class="sec">Failure Modes ({total_fm} total rejects)</div>',
                                    unsafe_allow_html=True)
                        st.dataframe(df_fm, use_container_width=True,
                                     height=min(40*len(df_fm)+50, 350))

                    with c2:
                        # Pie chart of failure modes
                        modes = [k for k,v in sorted(p["failure_modes"].items(),
                                                      key=lambda x:-x[1])]
                        counts = [p["failure_modes"][m] for m in modes]
                        fig_pie = go.Figure(go.Pie(
                            labels=modes, values=counts,
                            hole=0.4,
                            textfont=dict(color="white", size=11),
                            hovertemplate="%{label}: %{value} (%{percent})<extra></extra>",
                        ))
                        fig_pie.update_layout(**PLOT, height=300,
                            margin=dict(t=20,b=10,l=10,r=10),
                            title=dict(text="Defect Mix",
                                       font=dict(color="white",size=12)))
                        st.plotly_chart(fig_pie, use_container_width=True)

                    # Defects by product — per hour stacked bars
                    # HxH files log which station the defect came from,
                    # and station → product mapping from the lot info
                    # Show per-hour defect totals with product from lot field
                    prod_from_lot = p["product"]  # primary product this shift
                    sec(f"Defects per Hour — {prod_from_lot} ({p['shift']} Shift)")
                    # Aggregate rejects per hour across all stations
                    hour_rejects = {}
                    for s in p["stations"]:
                        for h in s["hours"]:
                            hr = h["hour"]
                            if hr not in hour_rejects:
                                hour_rejects[hr] = {"station_data":{}}
                            hour_rejects[hr]["station_data"][s["station"]] = h["rejects"]

                    if hour_rejects:
                        hrs_sorted = sorted(hour_rejects.keys())
                        all_stations = list({s["station"] for s in p["stations"]
                                             if any(h["rejects"]>0 for h in s["hours"])})
                        fig_hr = go.Figure()
                        for station in all_stations:
                            yvals = [hour_rejects.get(hr,{}).get("station_data",{}).get(station,0)
                                     for hr in hrs_sorted]
                            if sum(yvals)==0: continue
                            fig_hr.add_trace(go.Bar(
                                name=station,
                                x=[f"Hour {h}" for h in hrs_sorted],
                                y=yvals,
                                hovertemplate=(f"<b>{station}</b><br>"
                                               "%{x}: %{y} rejects<extra></extra>"),
                            ))
                        fig_hr.update_layout(**PLOT, margin=PLOT_MARGIN, height=300, barmode="stack",
                            title=dict(text=f"Rejects per Hour by Station — {prod_from_lot}",
                                       font=dict(color="white",size=13)),
                            xaxis=dict(showgrid=False),
                            yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                                       title="Rejects"))
                        st.plotly_chart(fig_hr, use_container_width=True)

                # ── Station summary ───────────────────────────────────────────
                sec("Station Summary")
                df_s = pd.DataFrame([{
                    "Station":    s["station"],
                    "Total":      s["total"],
                    "Rejects":    s["total_rejects"],
                    "Reject %":   f"{s['total_rejects']/s['total']:.1%}"
                                  if s["total"]>0 else "—",
                    "OEE":        f"{s['oee']:.1%}" if s["oee"] else "—",
                    "Hours Run":  len(s["hours"]),
                } for s in p["stations"]])
                if not df_s.empty:
                    df_s.index = range(1, len(df_s)+1)
                    st.dataframe(df_s, use_container_width=True)

                # ── Operator assignments ──────────────────────────────────────
                if p["operators"]:
                    sec("Operator Station Assignments")
                    df_ops = pd.DataFrame(p["operators"]).rename(columns={
                        "station":"Station","name":"Operator","status":"Status"})
                    df_ops = df_ops.drop_duplicates(subset=["Station","Operator"])
                    df_ops.index = range(1, len(df_ops)+1)
                    st.dataframe(df_ops, use_container_width=True)

                # ── Shift notes ───────────────────────────────────────────────
                if p["notes"]:
                    sec("Shift Notes")
                    for note in p["notes"]:
                        st.markdown(f'''<div class="note-box">
                          <b>{note["time"]}</b><br>
                          <span style="color:#c0cfe0">{note["note"]}</span>
                        </div>''', unsafe_allow_html=True)

        if st.button("🗑 Clear all loaded HxH shifts", key="clear_hxh"):
            st.session_state.hxh_parsed = []
            st.rerun()


# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 6 — DOWNTIME
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown("### 📉 Downtime Analysis")

    parsed_all_dt = st.session_state.get("hxh_parsed", [])
    if not parsed_all_dt:
        st.info("Upload HxH files in the **⏱ Hour by Hour** tab first — downtime data comes from the same files.")
    else:
        # Aggregate across all loaded shifts
        from collections import defaultdict as _dd

        # Filter controls
        fc1, fc2, fc3 = st.columns(3)
        all_facilities = sorted(set(p["facility"] for p in parsed_all_dt))
        all_dates      = sorted(set(p["date"]     for p in parsed_all_dt))
        f_fac  = fc1.multiselect("Facility", all_facilities, default=all_facilities)
        f_date = fc2.multiselect("Date",     all_dates,      default=all_dates)
        f_shift= fc3.multiselect("Shift",    ["Days","Swing","Weekend"],
                                  default=["Days","Swing","Weekend"])

        filtered = [p for p in parsed_all_dt
                    if p["facility"] in f_fac
                    and p["date"] in f_date
                    and any(s.lower() in p["shift"].lower() for s in f_shift)]

        if not filtered:
            st.info("No shifts match the selected filters.")
        else:
            # Aggregate all downtime events
            dt_by_station = _dd(int)
            dt_by_code    = _dd(int)
            dt_events     = []
            code_to_label = {}

            for p in filtered:
                # Merge legend
                code_to_label.update(p.get("dt_legend", {}))
                for s in p["stations"]:
                    for h in s["hours"]:
                        mins = h["downtime_min"]
                        code = h["downtime_code"]
                        if mins > 0:
                            dt_by_station[s["station"]] += mins
                            dt_by_code[code] += mins
                            label = p["dt_legend"].get(code, code) if code else "Unknown"
                            dt_events.append({
                                "date":     p["date"],
                                "facility": p["facility"],
                                "shift":    p["shift"],
                                "station":  s["station"],
                                "hour":     h["hour"],
                                "minutes":  mins,
                                "code":     code,
                                "reason":   label,
                            })

            total_dt_mins = sum(dt_by_station.values())

            # ── KPIs ──────────────────────────────────────────────────────────
            sec("Downtime Summary")
            dk1, dk2, dk3, dk4 = st.columns(4)
            with dk1: kpi("Total Downtime",
                          f"{total_dt_mins:,} min",
                          f'<div class="kpi-delta-neutral">{total_dt_mins/60:.1f} hrs</div>',
                          "#e05c5c")
            with dk2: kpi("Shifts Analysed", str(len(filtered)), accent="#2e75b6")
            with dk3: kpi("Avg per Shift",
                          f"{total_dt_mins//len(filtered)} min" if filtered else "—",
                          accent="#ffc000")
            with dk4: kpi("Unique DT Codes", str(len(dt_by_code)), accent="#9b59b6")

            c1, c2 = st.columns(2)

            # Downtime by station
            with c1:
                sec("Downtime by Station (minutes)")
                if dt_by_station:
                    sorted_s = sorted(dt_by_station.items(), key=lambda x:-x[1])
                    fig_dts = go.Figure(go.Bar(
                        y=[x[0] for x in sorted_s][::-1],
                        x=[x[1] for x in sorted_s][::-1],
                        orientation="h",
                        marker_color="#e05c5c",
                        text=[f"{x[1]:,} min ({x[1]/total_dt_mins:.0%})"
                              for x in sorted_s][::-1],
                        textposition="outside",
                        textfont=dict(color="white",size=10),
                    ))
                    fig_dts.update_layout(**PLOT, margin=PLOT_MARGIN, height=max(300,len(dt_by_station)*40),
                        xaxis=dict(showgrid=True,gridcolor="#2a3a4a",title="Minutes"),
                        yaxis=dict(showgrid=False,tickfont=dict(size=11)))
                    st.plotly_chart(fig_dts, use_container_width=True)

            # Downtime by reason code
            with c2:
                sec("Downtime by Reason Code")
                if dt_by_code:
                    # Categorise: planned vs unplanned
                    PLANNED = {"P","Q","K","O","N"}  # Break, Lunch, Scheduled Maint, Production Complete, Materials Filled
                    sorted_c = sorted(dt_by_code.items(), key=lambda x:-x[1])
                    labels_c = [f"{k}: {code_to_label.get(k,k)}" for k,v in sorted_c]
                    values_c = [v for k,v in sorted_c]
                    colors_c = ["#4472c4" if k in PLANNED else "#e05c5c"
                                for k,v in sorted_c]
                    fig_dtc = go.Figure(go.Bar(
                        y=labels_c[::-1], x=values_c[::-1],
                        orientation="h",
                        marker_color=colors_c[::-1],
                        text=[f"{v:,} min" for v in values_c][::-1],
                        textposition="outside",
                        textfont=dict(color="white",size=10),
                    ))
                    fig_dtc.update_layout(**PLOT, margin=PLOT_MARGIN, height=max(300,len(dt_by_code)*36),
                        xaxis=dict(showgrid=True,gridcolor="#2a3a4a",title="Minutes"),
                        yaxis=dict(showgrid=False,tickfont=dict(size=10)),
                        title=dict(text="🔴 Unplanned  🔵 Planned",
                                   font=dict(color="#8fa3b8",size=11)))
                    st.plotly_chart(fig_dtc, use_container_width=True)

            # Downtime events table
            if dt_events:
                sec("All Downtime Events")
                df_dte = pd.DataFrame(dt_events).sort_values(
                    ["date","facility","shift","station","hour"])
                df_dte["Planned"] = df_dte["code"].apply(
                    lambda c: "✅ Planned" if c in {"P","Q","K","O","N"} else "🔴 Unplanned")
                df_dte = df_dte.rename(columns={
                    "date":"Date","facility":"Facility","shift":"Shift",
                    "station":"Station","hour":"Hour",
                    "minutes":"Minutes","code":"Code","reason":"Reason"})
                df_dte = df_dte[["Date","Facility","Shift","Station",
                                  "Hour","Code","Reason","Minutes","Planned"]]
                df_dte.index = range(1,len(df_dte)+1)
                st.dataframe(df_dte, use_container_width=True, height=400)

            # Planned vs unplanned summary
            sec("Planned vs Unplanned Downtime")
            PLANNED_CODES = {"P","Q","K","O","N"}
            planned_mins   = sum(v for k,v in dt_by_code.items() if k in PLANNED_CODES)
            unplanned_mins = total_dt_mins - planned_mins
            if total_dt_mins > 0:
                p1, p2, p3 = st.columns(3)
                with p1: kpi("Planned DT",   f"{planned_mins:,} min",
                             f'<div class="kpi-delta-neutral">{planned_mins/total_dt_mins:.0%} of total</div>',
                             "#2e75b6")
                with p2: kpi("Unplanned DT", f"{unplanned_mins:,} min",
                             f'<div class="kpi-delta-bad">{unplanned_mins/total_dt_mins:.0%} of total</div>',
                             "#e05c5c")
                with p3:
                    fig_pie = go.Figure(go.Pie(
                        values=[planned_mins, unplanned_mins],
                        labels=["Planned","Unplanned"],
                        marker_colors=["#2e75b6","#e05c5c"],
                        hole=0.5,
                        textfont=dict(color="white"),
                    ))
                    fig_pie.update_layout(**PLOT, height=180,
                        margin=dict(t=10,b=10,l=10,r=10),
                        showlegend=False)
                    st.plotly_chart(fig_pie, use_container_width=True)


# ╔══════════════════════════════════════════════════════════════════════════════
# ║  TAB 7 — GOALS
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.markdown("### 🎯 Goals & Targets")
    st.markdown("""<div class="note-box">
      ⚠️ <b>Goals not yet set.</b> This section is ready to populate
      once targets are confirmed by the quality team.
    </div>""", unsafe_allow_html=True)

    sec("Pending from Quality Team")
    goals_needed = [
        ("Costed Yield Target","e.g. ≥ 95%",
         "Monthly yield we're aiming for"),
        ("CoPQ / Part Target","e.g. < $3.00",
         "Max acceptable cost of poor quality per good unit"),
        ("Cumulative Leak Rate Target","e.g. < 2%",
         "Acceptable leak rate across AuST + CP"),
        ("AuST Leak Rate Target","e.g. < 2%",
         "AuST-specific leak target"),
        ("CP Leak Rate Target","e.g. < 1.5%",
         "CenterPoint-specific leak target"),
        ("Destroyed Tip Rate Target","e.g. < 1%",
         "Acceptable destroyed tip rate"),
        ("Monthly Production Target — AuST","Total units/month",
         "From planning team"),
        ("Monthly Production Target — CP","Total units/month",
         "From planning team"),
        ("Reject Rate Target per Shift","e.g. < 8%",
         "Shift-level quality target"),
    ]
    for goal, example, note in goals_needed:
        st.markdown(f"""<div style="background:#1e2a3a;border-radius:8px;
            padding:12px 16px;margin-bottom:6px;
            border-left:3px solid #2e4a6a">
          <div style="display:flex;justify-content:space-between">
            <span style="color:#fff;font-weight:600">{goal}</span>
            <span style="color:#8fa3b8;font-size:12px">{example}</span>
          </div>
          <div style="color:#6a7f94;font-size:12px;margin-top:3px">
            {note}</div>
        </div>""", unsafe_allow_html=True)

    sec("How to Add Goals Once Confirmed")
    st.markdown("""<div style="color:#8fa3b8;font-size:13px;
        line-height:1.8">
      Once you have the targets from the quality team:<br>
      1. Open <code>app.py</code> in GitHub<br>
      2. Search for <code>GOALS_CONFIG</code><br>
      3. Fill in the target values — the dashboard will automatically
         show green/red indicators on every KPI card.<br>
      4. Commit and push — live in 30 seconds.
    </div>""", unsafe_allow_html=True)

st.markdown(
    "<br><div style='color:#2a3a4a;text-align:center;font-size:11px'>"
    "SSPC Quality Dashboard</div>",
    unsafe_allow_html=True)
