import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import calendar

from extractor import (
    extract_from_files, get_top_defects,
    get_leak_trend, get_tip_trend
)

st.set_page_config(
    page_title="SSPC Quality Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.kpi-card {
    background:linear-gradient(135deg,#1e2a3a 0%,#162032 100%);
    border-radius:12px; padding:18px 22px;
    border-left:4px solid #2e75b6; margin-bottom:8px;
}
.kpi-label { color:#8fa3b8; font-size:12px; font-weight:700;
             letter-spacing:.6px; text-transform:uppercase; margin-bottom:5px; }
.kpi-value { color:#fff; font-size:26px; font-weight:700; line-height:1; }
.kpi-delta-good { color:#00c853; font-size:12px; font-weight:600; margin-top:5px; }
.kpi-delta-bad  { color:#ff3d3d; font-size:12px; font-weight:600; margin-top:5px; }
.section-hdr { color:#8fa3b8; font-size:11px; font-weight:700;
               letter-spacing:1.5px; text-transform:uppercase;
               margin:18px 0 10px; border-bottom:1px solid #1e2a3a; padding-bottom:4px; }
.upload-box { background:#1e2a3a; border-radius:12px; padding:32px;
              border:1px dashed #2e4a6a; text-align:center; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_c(v):  return f"${v:,.0f}" if v is not None else "—"
def fmt_p(v):  return f"{v:.2%}"   if v is not None else "—"

def delta_tag(curr, prev, higher_good=True, fmt="currency"):
    if prev is None or curr is None or prev == 0: return ""
    diff = curr - prev; pct = diff / prev
    good = (diff >= 0) == higher_good
    cls  = "kpi-delta-good" if good else "kpi-delta-bad"
    arrow = "▲" if diff > 0 else "▼"
    if fmt == "currency": s = f"{arrow} ${abs(diff):,.0f} ({pct:+.1%})"
    elif fmt == "pct":    s = f"{arrow} {abs(diff):.2%} ({pct:+.1%})"
    else:                 s = f"{arrow} {abs(diff):.2f} ({pct:+.1%})"
    return f'<div class="{cls}">{s} vs prev</div>'

def kpi(label, value, delta="", accent="#2e75b6"):
    st.markdown(f"""
    <div class="kpi-card" style="border-left-color:{accent}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {delta}
    </div>""", unsafe_allow_html=True)

PLOT = dict(
    plot_bgcolor="#1e2a3a", paper_bgcolor="#1e2a3a",
    font=dict(color="white", family="Arial"),
    margin=dict(t=40, b=30, l=10, r=10),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="white")),
)


# ── Session state ─────────────────────────────────────────────────────────────
for key in ["df_rows","df_copq","months"]:
    if key not in st.session_state:
        st.session_state[key] = None if key != "months" else []


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

    process_ready = aust_file is not None and cp_file is not None

    if process_ready:
        if st.button("▶  Process Files", use_container_width=True, type="primary"):
            with st.spinner("Reading files and computing COPQ…"):
                df_rows, df_copq, months, errors = extract_from_files(
                    aust_file.read(), cp_file.read()
                )
            if errors:
                for e in errors: st.error(e)
            else:
                st.session_state.df_rows = df_rows
                st.session_state.df_copq = df_copq
                st.session_state.months  = months
                st.success(f"✅ {len(months)} months loaded")
    else:
        st.info("Upload both files above, then click Process.")

    st.markdown("---")

    if st.session_state.months:
        month_labels = [f"{calendar.month_abbr[m]} {y}"
                        for y, m in st.session_state.months]
        sel_idx = st.selectbox("Select Month",
                                range(len(month_labels)),
                                format_func=lambda i: month_labels[i],
                                index=len(month_labels)-1)
        sel_year, sel_month = st.session_state.months[sel_idx]
        sel_label = month_labels[sel_idx]
        prev_tuple = st.session_state.months[sel_idx-1] if sel_idx > 0 else None
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
if st.session_state.df_copq is None:
    st.markdown("## 🏭 SSPC Quality Dashboard")
    st.markdown("")
    _, mid, _ = st.columns([1,2,1])
    with mid:
        st.markdown("""
        <div class="upload-box">
            <div style="font-size:48px;margin-bottom:12px">📂</div>
            <h3 style="color:#8fa3b8;margin:0 0 8px">Upload Your DOR Files</h3>
            <p style="color:#6a7f94;margin:0 0 16px">
                Use the sidebar to upload the AuST and CenterPoint<br>
                DOR files, then click <b>Process Files</b>.
            </p>
            <p style="color:#4a5a6a;font-size:13px;margin:0">
                All months in the files load automatically.<br>
                Supports .xlsm and .xlsx
            </p>
        </div>
        """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
df_rows = st.session_state.df_rows
df_copq = st.session_state.df_copq
months  = st.session_state.months

curr_r = df_copq[(df_copq["year"]==sel_year) & (df_copq["month"]==sel_month)]
prev_r = df_copq[(df_copq["year"]==prev_tuple[0]) & (df_copq["month"]==prev_tuple[1])] \
         if prev_tuple else pd.DataFrame()

def cv(col):
    return float(curr_r[col].iloc[0]) if not curr_r.empty and col in curr_r.columns \
           and pd.notna(curr_r[col].iloc[0]) else None
def pv(col):
    return float(prev_r[col].iloc[0]) if not prev_r.empty and col in prev_r.columns \
           and pd.notna(prev_r[col].iloc[0]) else None

# Header
st.markdown("## 🏭 SSPC Quality Dashboard")
st.markdown(f"<span style='color:#8fa3b8;font-size:14px'>Showing: <b>{sel_label}</b>"
            f"&nbsp;|&nbsp;{len(months)} months loaded</span>", unsafe_allow_html=True)
st.markdown("---")

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">Key Performance Indicators</div>',
            unsafe_allow_html=True)
k1,k2,k3,k4,k5 = st.columns(5)
with k1:
    v = (cv("aust_scrap") or 0) + (cv("cp_scrap") or 0)
    p = ((pv("aust_scrap") or 0) + (pv("cp_scrap") or 0)) if prev_tuple else None
    kpi("Total Scrap Cost", fmt_c(v),
        delta_tag(v, p, higher_good=False), "#e05c5c")
with k2:
    cpp = cv("copq_per_part")
    kpi("CoPQ / Part", f"${cpp:.2f}" if cpp else "—",
        delta_tag(cpp, pv("copq_per_part"), higher_good=False, fmt="num"), "#e05c5c")
with k3:
    kpi("Costed Yield", fmt_p(cv("costed_yield")),
        delta_tag(cv("costed_yield"), pv("costed_yield"), higher_good=True, fmt="pct"),
        "#00c853")
with k4:
    kpi("AuST Good Cost", fmt_c(cv("aust_good")),
        delta_tag(cv("aust_good"), pv("aust_good"), higher_good=True), "#2e75b6")
with k5:
    kpi("CP Good Cost", fmt_c(cv("cp_good")),
        delta_tag(cv("cp_good"), pv("cp_good"), higher_good=True), "#70ad47")

st.markdown("")

# ── Cost Breakdown & Yield by Product ────────────────────────────────────────
c1, c2 = st.columns(2)
with c1:
    st.markdown('<div class="section-hdr">Cost Breakdown</div>', unsafe_allow_html=True)
    cats   = ["AuST Good","AuST Scrap","CP Good","CP Scrap"]
    vals   = [cv("aust_good"), cv("aust_scrap"), cv("cp_good"), cv("cp_scrap")]
    colors = ["#2e75b6","#c00000","#70ad47","#ff7c00"]
    fig = go.Figure(go.Bar(
        x=cats, y=vals, marker_color=colors,
        text=[fmt_c(v) for v in vals], textposition="outside",
        textfont=dict(color="white", size=11),
        hovertemplate="%{x}<br>%{text}<extra></extra>"
    ))
    fig.update_layout(**PLOT, height=300, showlegend=False,
        yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                   tickprefix="$", tickformat=",.0f"),
        xaxis=dict(showgrid=False))
    st.plotly_chart(fig, use_container_width=True)

with c2:
    st.markdown('<div class="section-hdr">Costed Yield by Product</div>',
                unsafe_allow_html=True)
    prods  = ["BLN","SMG","CTI","BMD","MP3"]
    yields = [cv(f"{p.lower()}_yield") for p in prods]
    pcols  = ["#2e75b6","#70ad47","#ffc000","#e05c5c","#9b59b6"]
    fig2   = go.Figure()
    for p, y, c in zip(prods, yields, pcols):
        if y and y > 0:
            fig2.add_trace(go.Bar(
                name=p, x=[p], y=[y], marker_color=c,
                text=[fmt_p(y)], textposition="outside",
                textfont=dict(color="white", size=12)
            ))
    fig2.add_hline(y=0.9, line_dash="dash", line_color="#8fa3b8",
                   annotation_text="90% target",
                   annotation_font_color="#8fa3b8")
    fig2.update_layout(**PLOT, height=300, showlegend=False,
        yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                   tickformat=".0%", range=[0, 1.08]),
        xaxis=dict(showgrid=False))
    st.plotly_chart(fig2, use_container_width=True)

# ── Defect Analysis ───────────────────────────────────────────────────────────
st.markdown(f'<div class="section-hdr">Defect Analysis — {sel_label}</div>',
            unsafe_allow_html=True)
d1, d2 = st.columns([3, 2])
top10 = get_top_defects(df_rows, sel_year, sel_month)

with d1:
    if top10:
        names  = [x[0] for x in top10]
        values = [x[1] for x in top10]
        bcols  = ["#c00000" if i==0 else "#2e75b6" if i<3 else "#4472c4"
                  for i in range(len(names))]
        fig3 = go.Figure(go.Bar(
            y=names[::-1], x=values[::-1], orientation="h",
            marker_color=bcols[::-1],
            text=[str(int(v)) for v in values[::-1]],
            textposition="outside", textfont=dict(color="white", size=11),
            hovertemplate="%{y}: %{x}<extra></extra>"
        ))
        fig3.update_layout(**PLOT, height=400,
            title=dict(text="Top 10 Failure Modes by Quantity",
                       font=dict(color="white", size=13)),
            xaxis=dict(showgrid=True, gridcolor="#2a3a4a"),
            yaxis=dict(showgrid=False, tickfont=dict(size=11)))
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("No defect data for this month.")

with d2:
    mdf = df_rows[(df_rows["year"]==sel_year) & (df_rows["month"]==sel_month)]
    lv  = float(mdf["Leak - valve"].sum())
    lb  = float(mdf["Leak - bond"].sum())
    tl  = lv + lb

    if tl > 0:
        st.markdown(f"""
        <div style="background:#1e2a3a;border-radius:10px;padding:16px;
                    margin-bottom:12px;border:1px solid #2e4a6a">
          <div style="color:#8fa3b8;font-size:11px;font-weight:700;
                      letter-spacing:1px;text-transform:uppercase;
                      margin-bottom:10px">Leak Breakdown</div>
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
        </div>""", unsafe_allow_html=True)
    else:
        combined = float(mdf["Leak"].sum())
        if combined > 0:
            st.markdown(f"""
            <div style="background:#1e2a3a;border-radius:10px;padding:16px;
                        margin-bottom:12px;border:1px solid #2e4a6a">
              <div style="color:#8fa3b8;font-size:11px;font-weight:700;
                          letter-spacing:1px;text-transform:uppercase;margin-bottom:8px">
                Leak</div>
              <div style="display:flex;justify-content:space-between">
                <span style="color:#fff">Total Leak</span>
                <span style="color:#ffc000;font-weight:700;font-size:18px">{int(combined)}</span>
              </div>
            </div>""", unsafe_allow_html=True)

    if top10:
        df_t = pd.DataFrame(top10, columns=["Defect Mode","Count"])
        df_t["Count"] = df_t["Count"].astype(int)
        df_t.index    = range(1, len(df_t)+1)
        st.dataframe(df_t, use_container_width=True, height=300)

# ── 6-Month Trends ────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">6-Month Trend Analysis</div>',
            unsafe_allow_html=True)
t1, t2 = st.columns(2)

def line_chart(df_t, title, ytitle, color_b="#70ad47", yfmt=".2%"):
    fig = go.Figure()
    for col, color in [("AuST","#2e75b6"), ("CenterPoint", color_b)]:
        fig.add_trace(go.Scatter(
            x=df_t["label"], y=df_t[col], name=col,
            mode="lines+markers",
            line=dict(color=color, width=3),
            marker=dict(size=8, color=color),
            text=[f"{v:{yfmt}}" for v in df_t[col]],
            hovertemplate="%{x}<br>" + col + ": %{text}<extra></extra>"
        ))
    fig.update_layout(**PLOT, height=320,
        title=dict(text=title, font=dict(color="white", size=13)),
        yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
                   tickformat=yfmt, title=ytitle),
        xaxis=dict(showgrid=False))
    return fig

leak_trend = get_leak_trend(df_rows, df_copq)
tip_trend  = get_tip_trend(df_rows, df_copq)

with t1:
    if not leak_trend.empty:
        st.plotly_chart(
            line_chart(leak_trend, "Leak Rate Trend (Last 6 Months)", "Leak Rate"),
            use_container_width=True)
    else:
        st.info("Need at least 2 months of data for trend.")

with t2:
    if not tip_trend.empty:
        st.plotly_chart(
            line_chart(tip_trend, "Destroyed Tip Rate Trend (Last 6 Months)",
                       "Destroyed Tip Rate", color_b="#e05c5c"),
            use_container_width=True)
    else:
        st.info("Need at least 2 months of data for trend.")

# ── CoPQ/Part All-Time ────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">CoPQ per Part — All Time</div>',
            unsafe_allow_html=True)

all_labels = [f"{calendar.month_abbr[m]} {y}" for y, m in months]
all_copq   = []
for y, m in months:
    r = df_copq[(df_copq["year"]==y) & (df_copq["month"]==m)]
    v = float(r["copq_per_part"].iloc[0]) if not r.empty and pd.notna(r["copq_per_part"].iloc[0]) else 0
    all_copq.append(v)

avg_copq   = sum(all_copq) / len(all_copq) if all_copq else 0
bar_colors = [
    "#ffc000" if lbl == sel_label else
    "#c00000" if v == max(all_copq) else
    "#00c853" if v == min(all_copq) else "#2e75b6"
    for lbl, v in zip(all_labels, all_copq)
]

fig_c = go.Figure(go.Bar(
    x=all_labels, y=all_copq, marker_color=bar_colors,
    text=[f"${v:.2f}" for v in all_copq],
    textposition="outside", textfont=dict(color="white", size=10),
    hovertemplate="%{x}<br>CoPQ/Part: %{text}<extra></extra>"
))
fig_c.add_hline(y=avg_copq, line_dash="dash", line_color="#8fa3b8",
                annotation_text=f"Avg ${avg_copq:.2f}",
                annotation_font_color="#8fa3b8")
fig_c.update_layout(**PLOT, height=340, showlegend=False,
    yaxis=dict(showgrid=True, gridcolor="#2a3a4a",
               tickprefix="$", title="CoPQ / Part"),
    xaxis=dict(showgrid=False, tickangle=-35))
st.plotly_chart(fig_c, use_container_width=True)

# ── Leak Summary ──────────────────────────────────────────────────────────────
st.markdown('<div class="section-hdr">Leak Rate Summary</div>', unsafe_allow_html=True)
l1,l2,l3,l4 = st.columns(4)
with l1:
    kpi("Total Leak — AuST", f"{int(cv('leak_aust') or 0)}",
        delta_tag(cv("leak_aust"), pv("leak_aust"), higher_good=False, fmt="num"), "#2e75b6")
with l2:
    kpi("Total Leak — CP", f"{int(cv('leak_cp') or 0)}",
        delta_tag(cv("leak_cp"), pv("leak_cp"), higher_good=False, fmt="num"), "#70ad47")
with l3:
    kpi("Leak Rate — AuST", fmt_p(cv("leak_rate_aust")),
        delta_tag(cv("leak_rate_aust"), pv("leak_rate_aust"), higher_good=False, fmt="pct"),
        "#2e75b6")
with l4:
    kpi("Cumulative Leak Rate", fmt_p(cv("cumul_leak")),
        delta_tag(cv("cumul_leak"), pv("cumul_leak"), higher_good=False, fmt="pct"),
        "#e05c5c")

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    f"<div style='color:#3a4a5a;text-align:center;font-size:11px'>"
    f"SSPC Quality Dashboard &nbsp;•&nbsp; {len(months)} months loaded</div>",
    unsafe_allow_html=True)
