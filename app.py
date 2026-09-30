import pandas as pd, numpy as np, plotly.express as px, plotly.graph_objects as go, streamlit as st
import model as M

st.set_page_config(page_title="Gas Allocation LP", layout="wide")
st.title("⛽ Gas Allocation over 12 Months — LP Dashboard (PuLP)")
inr = lambda v: f"₹{v:,.0f}"

# ---------------------------------------------------------------- sidebar: data
sb = st.sidebar
sb.header("1 · Data")
mode = sb.radio("Source", ["Demo data (corridors bind)", "Original image data", "Upload Excel"])
sb.download_button("⬇ Download sample-data template (.xlsx)", M.sample_excel_bytes(), "gas_sample_data.xlsx")
if mode == "Upload Excel":
    up = sb.file_uploader("Upload filled template", type="xlsx")
    if up is None:
        st.info("Upload a workbook (use the template above) or switch to Sample data."); st.stop()
    try: base, key = M.read_excel_data(up), up.name
    except Exception as e: st.error(f"Could not read file: {e}"); st.stop()
else:
    base, key = M.sample_data("image" if mode.startswith("Original") else "demo"), mode

sb.header("2 · Scenario levers")
sc = dict(M.DEFAULT_SC)
sc["supply"] = sb.slider("Gas supply ×", 0.5, 1.5, 1.0, 0.05)
sc["demand"] = sb.slider("Demand ×", 0.5, 1.5, 1.0, 0.05)
sc["cap"] = sb.slider("All corridor capacity ×", 0.5, 1.5, 1.0, 0.05)
sc["cost"] = sb.slider("Transport cost ×", 0.5, 2.0, 1.0, 0.05)
sc["pen"] = sb.slider("Shortage penalty ×", 0.5, 2.0, 1.0, 0.05)
sc["svc_shift"] = sb.slider("Service-level shift (pts)", -0.30, 0.10, 0.0, 0.01,
                            help="Added to every tier's minimum service level (e.g. 0.90 → 0.85 with −0.05).")
_o = sb.selectbox("Uniform service floor (all nodes)", ["Off - use tier floors", 0.70, 0.75, 0.80, 0.85, 0.90, 0.95],
                  format_func=lambda v: v if isinstance(v, str) else f"{v:.0%}")
sc["svc_all"] = None if isinstance(_o, str) else _o
sc["floor_pen"] = sb.number_input("Penalty for breaching a service floor (₹/SCM)", 0.0, 1e5, 500.0,
                                  help="Floors are soft so the model always solves; breaches are reported.")
corrs = sorted(base["corridors"]["Corridor"].unique())
with sb.expander("Per-node demand ×"):
    sc["dem_node"] = {n: st.slider(f"Node {n}", 0.5, 1.5, 1.0, 0.05) for n in base["nodes"]["Node"]}
with sb.expander("Per-corridor capacity ×"):
    sc["cap_corr"] = {k: st.slider(f"Corridor {k}", 0.5, 1.5, 1.0, 0.05) for k in corrs}

# ---------------------------------------------------------------- inputs (editable)
EX = {}
tabs = st.tabs(["📥 Inputs", "📊 Results", "🌐 Network diagram", "🛣 Corridors", "💲 Shadow prices", "🎚 Service level", "🎯 Sensitivity", "💾 Export"])
with tabs[0]:
    st.caption("Edit any table directly; results update immediately.")
    data = {}
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Tiers (service floors & penalties)")
        data["tiers"] = st.data_editor(base["tiers"], key=f"t{key}", width="stretch", hide_index=True)
        st.subheader("Nodes")
        data["nodes"] = st.data_editor(base["nodes"], key=f"n{key}", width="stretch", hide_index=True)
        st.subheader("Source availability (SCM/month)")
        data["supply"] = st.data_editor(base["supply"], key=f"s{key}", width="stretch", hide_index=True)
    with c2:
        st.subheader("Corridors (cost ₹/SCM, capacity SCM/month)")
        data["corridors"] = st.data_editor(base["corridors"], key=f"c{key}", width="stretch", hide_index=True)
    st.subheader("Demand (SCM/month)")
    data["demand"] = st.data_editor(base["demand"], key=f"d{key}", width="stretch", hide_index=True)

try: res = M.solve(data, sc)
except Exception as e: st.error(f"Model failed - check inputs: {e}"); st.stop()
S = res["summary"]; dv = res["deliv"]

# ---------------------------------------------------------------- results
with tabs[1]:
    st.caption(f"Solver status: **{S['Status']}**")
    k = st.columns(5)
    k[0].metric("Total cost", inr(S["Total_Cost"])); k[1].metric("Transport", inr(S["Transport_Cost"]))
    k[2].metric("Shortage penalty", inr(S["Shortage_Penalty"])); k[3].metric("Total shortage (SCM)", f"{S['Total_Shortage']:,.0f}")
    k[4].metric("Fill rate", f"{S['Fill_Rate']:.1%}")
    if S["Floor_Violation_SCM"] > 1:
        st.warning(f"Service floors cannot all be met with the available gas/capacity: shortfall below the floor totals "
                   f"**{S['Floor_Violation_SCM']:,.0f} SCM**. Raise supply/capacity or lower floors (Scenario levers).")
    a, b = st.columns(2)
    m_sh = dv.groupby("Month", as_index=False)[["Demand", "Delivered", "Shortage"]].sum()
    f = go.Figure([go.Bar(x=m_sh.Month, y=m_sh.Delivered, name="Delivered"), go.Bar(x=m_sh.Month, y=m_sh.Shortage, name="Shortage")])
    f.add_trace(go.Scatter(x=m_sh.Month, y=m_sh.Demand, name="Demand", mode="lines+markers"))
    f.add_trace(go.Scatter(x=list(M.MONTHS), y=[float(data["supply"].set_index("Month").loc[t, "Available"]) * sc["supply"] for t in M.MONTHS],
                           name="Supply", mode="lines", line=dict(dash="dash")))
    f.update_layout(barmode="stack", title="Monthly delivery vs demand & supply (SCM)", xaxis_title="Month")
    a.plotly_chart(f, width="stretch")
    ty = dv.groupby(["Month", "Type"], as_index=False)["Shortage"].sum()
    b.plotly_chart(px.bar(ty, x="Month", y="Shortage", color="Type", title="Shortage by customer type (SCM)"), width="stretch")
    pv = dv.pivot(index="Node", columns="Month", values="Fill_Rate")
    st.plotly_chart(px.imshow(pv, text_auto=".0%", aspect="auto", color_continuous_scale="RdYlGn", zmin=0.6, zmax=1,
                              title="Fill rate by node and month"), width="stretch")
    st.subheader("Delivery detail"); st.dataframe(dv.round(1), width="stretch", hide_index=True)

# ---------------------------------------------------------------- network diagram
with tabs[2]:
    mo = st.selectbox("Month", ["Average of all months"] + M.MONTHS)
    ms = M.MONTHS if isinstance(mo, str) else [mo]
    P = M.layout(data); fl, cu = res["flows"][res["flows"].Month.isin(ms)], res["corr"][res["corr"].Month.isin(ms)]
    fig = go.Figure(); mf = max(fl.groupby(["From", "To"]).Flow.mean().max(), 1)
    for r in data["corridors"].itertuples():
        f = fl[(fl.From == r.From) & (fl.To == r.To)].Flow.sum() / len(ms); u = cu[cu.Corridor == r.Corridor].Utilization.mean()
        (x1, y1), (x2, y2) = P[r.From], P[r.To]; col = f"hsl({120 - 120 * min(1, u):.0f},70%,42%)"
        fig.add_annotation(x=x2 - 38, y=y2, ax=x1 + 38, ay=y1, xref="x", yref="y", axref="x", ayref="y", showarrow=True,
                           arrowhead=2, arrowwidth=1.5 + 8 * f / mf, arrowcolor=col, text="")
        fig.add_annotation(x=x1 + .55 * (x2 - x1), y=y1 + .55 * (y2 - y1), showarrow=False, font=dict(size=10),
                           text=f"<b>{f:,.0f}</b> SCM<br>C{r.Corridor} · {u:.0%}", bgcolor="rgba(255,255,255,.85)")
    dv_ = res["deliv"][res["deliv"].Month.isin(ms)].groupby("Node")[["Demand", "Delivered", "Shortage"]].sum() / len(ms)
    nx = ["G"] + list(data["nodes"].Node)
    fr = [1.0] + [1 - dv_.loc[n, "Shortage"] / dv_.loc[n, "Demand"] for n in nx[1:]]
    txt = ["<b>City Gate</b>"] + [f"<b>{n}</b><br>{dv_.loc[n,'Delivered']:,.0f}/{dv_.loc[n,'Demand']:,.0f}" for n in nx[1:]]
    fig.add_trace(go.Scatter(x=[P[n][0] for n in nx], y=[P[n][1] for n in nx], mode="markers+text", text=txt, textfont=dict(color="black", size=11),
                             marker=dict(symbol="square", size=62, color=fr, colorscale="RdYlGn", cmin=.6, cmax=1, line=dict(width=1, color="#334"),
                                         colorbar=dict(title="Fill rate")), hoverinfo="skip"))
    fig.update_layout(height=560, showlegend=False, xaxis=dict(visible=False, range=[0, 740]), yaxis=dict(visible=False, autorange="reversed"),
                      title="Line width = flow · line colour = corridor utilisation (green→red) · box = delivered/demand (SCM)")
    st.plotly_chart(fig, width="stretch")

# ---------------------------------------------------------------- corridors
with tabs[3]:
    cu = res["corr"]
    st.plotly_chart(px.imshow(cu.pivot(index="Corridor", columns="Month", values="Utilization"), text_auto=".0%", aspect="auto",
                              color_continuous_scale="OrRd", zmin=0, zmax=1, title="Corridor utilisation (flow / capacity)"), width="stretch")

# ---------------------------------------------------------------- shadow prices
with tabs[4]:
    pct = st.selectbox("Alter capacity / supply / demand by", [-0.20, -0.10, -0.05, 0.05, 0.10, 0.20], index=4, format_func=lambda v: f"{v:+.0%}")
    si = M.shadow_impact(data, sc, res, pct); EX["Shadow_Impact"] = si
    if S["Floor_Violation_SCM"] > 1:
        st.warning(f"Service floors are being breached, so shadow prices include the ₹{sc['floor_pen']:.0f}/SCM breach penalty. "
                   "Use the Service level tab to find the highest feasible floor and select it in the sidebar.")
    top = si.sort_values("Actual objective saving", ascending=False).iloc[0]
    k = st.columns(3); k[0].metric("Highest-value constraint", top["Constraint"]); k[1].metric("Avg shadow price", f"₹{top['Avg shadow (Rs/SCM)']:.2f}/SCM")
    k[2].metric(f"Objective saving from {pct:+.0%}", inr(top["Actual objective saving"]))
    st.caption("Estimated saving = Σ shadow price × extra units (linear). Actual = full re-solve. A corridor shows ₹0 while gas supply, not pipeline capacity, is the "
               "limit and its peak utilisation is below 100% - lower that corridor's capacity in the sidebar to see it become a bottleneck.")
    st.dataframe(si.round(2), hide_index=True, width="stretch")
    l, r = st.columns(2)
    l.plotly_chart(px.bar(si.melt("Constraint", ["Est. objective saving (linear)", "Actual objective saving"]), x="Constraint", y="value", color="variable",
                          barmode="group", title=f"Objective saving from {pct:+.0%} (₹)"), width="stretch")
    st.subheader(f"Demand alteration - {pct:+.0%} demand at each node")
    di = M.demand_impact(data, sc, res, pct); EX["Demand_Impact"] = di
    st.caption("Marginal cost = extra real cost (transport + shortage penalty) per extra SCM demanded; nodes behind a bottleneck corridor cost the most.")
    st.dataframe(di.round(2), hide_index=True, width="stretch")
    du = res["duals"]; du = du[du.Type.isin(["Source availability", "Corridor capacity"])].assign(Constraint=lambda d: np.where(d.Type.str.startswith("Source"), "Source availability", "Corridor " + d.Item.astype(str)))
    r.plotly_chart(px.line(du, x="Month", y="Shadow_Price", color="Constraint", markers=True, title="Shadow price by month (₹/SCM)"), width="stretch")
    st.plotly_chart(px.imshow(du.pivot(index="Constraint", columns="Month", values="Shadow_Price"), text_auto=".1f", aspect="auto",
                              color_continuous_scale="YlOrRd", title="Shadow price heat-map (₹/SCM)"), width="stretch")

# ---------------------------------------------------------------- service level
with tabs[5]:
    st.caption("Uniform minimum service floor for every node, 70%→95% in 5% steps (other levers as set in the sidebar). Select a floor in the sidebar to apply it everywhere.")
    sv = M.service_levels(data, sc); EX["Service_Level"] = sv
    st.dataframe(sv.style.format({"Service floor": "{:.0%}", "Fill rate": "{:.1%}", **{c: "{:.1%}" for c in sv.columns if c.startswith("Fill tier")},
                                  **{c: "{:,.0f}" for c in ["Total cost", "Transport", "Shortage penalty", "Shortage (SCM)", "Floor breach (SCM)", "Cost vs previous level"]}}),
                 hide_index=True, width="stretch")
    l, r = st.columns(2)
    f2 = go.Figure([go.Scatter(x=sv["Service floor"], y=sv["Total cost"], name="Total cost (₹)"),
                    go.Scatter(x=sv["Service floor"], y=sv["Floor breach (SCM)"], name="Floor breach (SCM)", yaxis="y2"),
                    go.Scatter(x=sv["Service floor"], y=sv["Shortage (SCM)"], name="Shortage (SCM)", yaxis="y2")])
    f2.update_layout(title="Cost & shortage vs service floor", yaxis2=dict(overlaying="y", side="right"), xaxis=dict(tickformat=".0%"))
    l.plotly_chart(f2, width="stretch")
    tc = [c for c in sv.columns if c.startswith("Fill tier")]
    r.plotly_chart(px.line(sv.melt("Service floor", tc), x="Service floor", y="value", color="variable", markers=True, title="Fill rate by tier vs floor")
                   .update_layout(xaxis=dict(tickformat=".0%"), yaxis=dict(tickformat=".0%")), width="stretch")

# ---------------------------------------------------------------- sensitivity
with tabs[6]:
    st.caption("Each run re-solves the LP around your **data** (scenario sliders are ignored here so results are comparable).")
    P = ["Gas supply", "Demand", "All corridor capacity"] + [f"Corridor {k}" for k in corrs] + \
        ["Transport cost", "Shortage penalty", "Service level (pts shift)"]
    st.session_state.setdefault("sens", {})
    st.subheader("Tornado - impact of ±x% change on total cost")
    pct = st.slider("Change (%)", 5, 30, 10, 5) / 100
    if st.button("Run tornado"):
        with st.spinner("Solving..."): st.session_state.sens["Tornado"] = M.tornado(data, P, pct, svc_pts=pct / 2)
    if "Tornado" in st.session_state.sens:
        t = st.session_state.sens["Tornado"].sort_values("Swing")
        fig = go.Figure([go.Bar(y=t.Parameter, x=t.Delta_Low, orientation="h", name="Low"),
                         go.Bar(y=t.Parameter, x=t.Delta_High, orientation="h", name="High")])
        fig.update_layout(barmode="relative", xaxis_title="Δ total cost (₹) vs base", title="Tornado (service level: ±half of the % as points)")
        st.plotly_chart(fig, width="stretch"); st.dataframe(t.round(0), width="stretch", hide_index=True)
    st.divider(); st.subheader("One-way sweep")
    p = st.selectbox("Parameter", P)
    lo, hi = (-0.30, 0.10) if p.startswith("Service") else (0.5, 1.5)
    rng = st.slider("Range", lo, hi, (lo, hi), 0.05 if not p.startswith("Service") else 0.01)
    if st.button("Run sweep"):
        with st.spinner("Solving..."):
            st.session_state.sens[f"Sweep_{p}"[:31]] = M.sweep(data, p, np.round(np.linspace(rng[0], rng[1], 11), 3))
    for nm, df in list(st.session_state.sens.items()):
        if nm.startswith("Sweep_"):
            if nm != f"Sweep_{p}"[:31]: continue
            l, r = st.columns(2)
            l.plotly_chart(px.line(df, x="Value", y=["Total_Cost", "Transport_Cost", "Shortage_Penalty"], markers=True, title="Cost"), width="stretch")
            r.plotly_chart(px.line(df, x="Value", y=["Total_Shortage", "Floor_Violation_SCM"], markers=True, title="Shortage / floor breach (SCM)"), width="stretch")
    st.divider(); st.subheader("Two-way: supply × demand → total cost")
    if st.button("Run heat-map (25 solves)"):
        with st.spinner("Solving..."): st.session_state.sens["Heatmap"] = M.heatmap(data)
    if "Heatmap" in st.session_state.sens:
        h = st.session_state.sens["Heatmap"]
        st.plotly_chart(px.imshow(h, text_auto=",.0f", color_continuous_scale="YlOrRd", aspect="auto"), width="stretch")

# ---------------------------------------------------------------- export
with tabs[7]:
    st.write("Workbook contains: summary, scenario settings, deliveries, shortage matrix, flows, corridor utilisation, "
             "shadow prices, any sensitivity runs you executed, and the input tables.")
    st.download_button("⬇ Download results (.xlsx)", M.results_excel(res, data, {**st.session_state.get("sens", {}), **EX}),
                       "gas_allocation_results.xlsx", type="primary")
