import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import calendar
import datetime

from extractor import (
    extract_from_files, get_top_defects,
    get_rolling_stats, get_leak_trend, get_tip_trend
)

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
    margin=dict(t=40, b=30, l=10, r=10),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="white")),
)

def line_fig(df_t, title, ytitle, cols_colors, yfmt=".2%", height=320):
    fig = go.Figure()
    for col, color in cols_colors:
        if col not in df_t.columns: continue
        fig.add_trace(go.Scatter(
            x=df_t["label"], y=df_t[col], name=col,
            mode="lines+markers",
            line=dict(color=color, width=3),
            marker=dict(size=8, color=color),
            text=[f"{v:{yfmt}}" for v in df_t[col]],
            hovertemplate="%{x}<br>" + col + ": %{text}<extra></extra>"))
    fig.update_layout(**PLOT, height=height,
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
    st.markdown("### Upload DOR Files")
    aust_file = st.file_uploader("AuST DOR (.xlsm / .xlsx)",
                                  type=["xlsm","xlsx"], key="aust")
    cp_file   = st.file_uploader("CenterPoint DOR (.xlsm / .xlsx)",
                                  type=["xlsm","xlsx"], key="cp")

    if aust_file and cp_file:
        if st.button("▶  Process Files", use_container_width=True, type="primary"):
            with st.spinner("Reading files and computing COPQ…"):
                data, errors = extract_from_files(
                    aust_file.read(), cp_file.read()
                )
            if errors:
                for e in errors: st.error(e)
            else:
                st.session_state.data = data
                st.success(f"✅ {len(data['months'])} months loaded")
    else:
        st.info("Upload both files above, then click Process.")

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
            Upload the AuST and CenterPoint DOR files in the sidebar,<br>
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
                "📅 Daily Tracker", "📋 Actions", "🎯 Goals"])


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
        pcols  = ["#2e75b6","#70ad47","#ffc000","#e05c5c","#9b59b6"]
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
        fig2.update_layout(**PLOT, height=300, showlegend=False,
            yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                       tickformat=".0%", range=[0, 1.08]),
            xaxis=dict(showgrid=False))
        st.plotly_chart(fig2, use_container_width=True)

    with c2:
        sec("CoPQ/Part — Rolling 6 Months with Avg ± 1 Std Dev")
        df_roll, avg_cpp, std_cpp = get_rolling_stats(
            df_copq, "copq_per_part", 6)
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
            fig_r.update_layout(**PLOT, height=300,
                yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                           tickprefix="$"),
                xaxis=dict(showgrid=False))
            st.plotly_chart(fig_r, use_container_width=True)

    # Defect Analysis
    sec(f"Defect Analysis — {sel_label}")
    top10 = get_top_defects(df_scrap, sel_year, sel_month)
    d1, d2 = st.columns([3, 2])

    with d1:
        if top10:
            names  = [x[0] for x in top10]
            values = [x[1] for x in top10]
            pcts   = [x[2] for x in top10]
            bcols  = ["#c00000" if i == 0 else
                      "#2e75b6" if i < 3 else "#4472c4"
                      for i in range(len(names))]
            fig3 = go.Figure(go.Bar(
                y=names[::-1], x=values[::-1],
                orientation="h", marker_color=bcols[::-1],
                text=[f"{int(v)} ({p:.1%})"
                      for v, p in zip(values[::-1], pcts[::-1])],
                textposition="outside",
                textfont=dict(color="white", size=11),
                hovertemplate="%{y}: %{x}<extra></extra>"))
            fig3.update_layout(**PLOT, height=400,
                title=dict(text="Top 10 Failure Modes — Count & % of Total",
                           font=dict(color="white", size=13)),
                xaxis=dict(showgrid=True, gridcolor="#2a3a4a"),
                yaxis=dict(showgrid=False, tickfont=dict(size=11)))
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No defect data for this month.")

    with d2:
        mdf = df_scrap[(df_scrap["year"] == sel_year) &
                        (df_scrap["month"] == sel_month)]
        lv = float(mdf["Leak - valve"].sum())
        lb = float(mdf["Leak - bond"].sum())
        tl = lv + lb
        if tl > 0:
            st.markdown(f"""<div style="background:#1e2a3a;border-radius:10px;
                padding:16px;margin-bottom:12px;border:1px solid #2e4a6a">
              <div style="color:#8fa3b8;font-size:11px;font-weight:700;
                letter-spacing:1px;text-transform:uppercase;margin-bottom:10px">
                Leak Breakdown</div>
              <div style="display:flex;justify-content:space-between;
                margin-bottom:8px">
                <span style="color:#fff">Total Leak</span>
                <span style="color:#ffc000;font-weight:700;font-size:18px">
                  {int(tl)}</span>
              </div>
              <div style="display:flex;justify-content:space-between;
                margin-bottom:6px">
                <span style="color:#8fa3b8;font-size:12px">
                  ↳ Leak — Valve</span>
                <span style="color:#e05c5c;font-weight:600">
                  {int(lv)} ({lv/tl:.0%})</span>
              </div>
              <div style="display:flex;justify-content:space-between">
                <span style="color:#8fa3b8;font-size:12px">
                  ↳ Leak — Bond</span>
                <span style="color:#ff7c00;font-weight:600">
                  {int(lb)} ({lb/tl:.0%})</span>
              </div>
            </div>""", unsafe_allow_html=True)
        else:
            combined = float(mdf["Leak"].sum())
            if combined > 0:
                st.markdown(f"""<div style="background:#1e2a3a;
                    border-radius:10px;padding:16px;margin-bottom:12px;
                    border:1px solid #2e4a6a">
                  <div style="color:#8fa3b8;font-size:11px;font-weight:700;
                    letter-spacing:1px;text-transform:uppercase;
                    margin-bottom:8px">Leak</div>
                  <div style="display:flex;justify-content:space-between">
                    <span style="color:#fff">Total Leak</span>
                    <span style="color:#ffc000;font-weight:700;
                      font-size:18px">{int(combined)}</span>
                  </div>
                </div>""", unsafe_allow_html=True)
        if top10:
            df_t = pd.DataFrame(
                [(x[0], int(x[1]), f"{x[2]:.1%}") for x in top10],
                columns=["Defect Mode", "Count", "% of Total"])
            df_t.index = range(1, len(df_t) + 1)
            st.dataframe(df_t, use_container_width=True, height=300)

    # 6-Month Trends
    sec("6-Month Trend Analysis")
    t1, t2 = st.columns(2)
    leak_trend = get_leak_trend(df_scrap, df_copq)
    tip_trend  = get_tip_trend(df_scrap, df_copq)
    with t1:
        if not leak_trend.empty:
            st.plotly_chart(
                line_fig(leak_trend, "Leak Rate Trend (Last 6 Months)",
                         "Leak Rate",
                         [("AuST","#2e75b6"),("CenterPoint","#70ad47")]),
                use_container_width=True)
        else:
            st.info("Need at least 2 months of data.")
    with t2:
        if not tip_trend.empty:
            st.plotly_chart(
                line_fig(tip_trend,
                         "Destroyed Tip Rate Trend (Last 6 Months)",
                         "Tip Rate",
                         [("AuST","#2e75b6"),("CenterPoint","#e05c5c")]),
                use_container_width=True)
        else:
            st.info("Need at least 2 months of data.")

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
        fig.update_layout(**PLOT, height=320, showlegend=False,
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
        fig_c.update_layout(**PLOT, height=320, showlegend=False,
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
        marker_color=["#2e75b6","#70ad47","#ffc000","#e05c5c","#9b59b6"],
        text=[fmt_c(v) for v in prod_costs.values()],
        textposition="outside",
        textfont=dict(color="white", size=12)))
    fig_prod.update_layout(**PLOT, height=300, showlegend=False,
        yaxis=dict(showgrid=True, gridcolor="#2a3a4a", tickprefix="$"),
        xaxis=dict(showgrid=False))
    st.plotly_chart(fig_prod, use_container_width=True)

    # Historical table
    sec("Monthly COPQ History")
    df_hist = df_copq.copy()
    df_hist["Month"] = df_hist.apply(
        lambda r: f"{calendar.month_abbr[int(r.month)]} {int(r.year)}",
        axis=1)
    df_hist["AuST Scrap"]   = df_hist["aust_scrap"].apply(fmt_c)
    df_hist["CP Scrap"]     = df_hist["cp_scrap"].apply(fmt_c)
    df_hist["Total Scrap"]  = df_hist["total_scrap"].apply(fmt_c)
    df_hist["Costed Yield"] = df_hist["costed_yield"].apply(fmt_p)
    df_hist["CoPQ/Part"]    = df_hist["copq_per_part"].apply(
        lambda v: f"${v:.2f}" if pd.notna(v) else "—")
    df_hist["Cumul. Leak"]  = df_hist["cumul_leak"].apply(fmt_p)
    st.dataframe(
        df_hist[["Month","AuST Scrap","CP Scrap","Total Scrap",
                 "Costed Yield","CoPQ/Part","Cumul. Leak"]]
        .sort_values("Month", ascending=False),
        use_container_width=True, hide_index=True)


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
                st.markdown(f"**{facility}**")
                col_list = list(actual_by_prod.items())
                if col_list:
                    cols = st.columns(len(col_list))
                    for i, (prod, actual) in enumerate(col_list):
                        plan_row = fac_dem[fac_dem["product"] == prod]["plan"]
                        plan = float(plan_row.iloc[0]) \
                               if not plan_row.empty else None
                        pct  = actual/plan if plan and plan > 0 else None
                        color = ("#00c853" if pct and pct >= 1 else
                                 "#ff7c00" if pct and pct >= 0.9 else
                                 "#e05c5c")
                        with cols[i]:
                            kpi(prod, f"{int(actual):,}",
                                f'<div style="color:{color};font-size:11px;'
                                f'margin-top:4px">'
                                f'{"✅" if pct and pct>=1 else "⚠️"} '
                                f'{fmt_p(pct)} of {int(plan):,} planned'
                                f'</div>'
                                if plan else
                                '<div class="kpi-delta-neutral">'
                                'No plan set</div>',
                                accent=color)
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
            fig_daily.update_layout(**PLOT, height=280,
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
            fig_rr.update_layout(**PLOT, height=240,
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
# ║  TAB 5 — GOALS
# ╚══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
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
