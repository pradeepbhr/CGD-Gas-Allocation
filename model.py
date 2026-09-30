"""Gas allocation LP (PuLP) - model, sample data, Excel I/O. No Streamlit dependency."""
import io
import numpy as np
import pandas as pd
import pulp as pl
from openpyxl.styles import Font, PatternFill, Alignment

MONTHS = list(range(1, 13))
MCOLS = [f"M{m}" for m in MONTHS]
DEFAULT_SC = dict(supply=1.0, demand=1.0, cap=1.0, cost=1.0, pen=1.0,
                  cap_corr={}, svc_shift=0.0, svc_all=None, floor_pen=500.0, dem_node={})

# ------------------------------------------------------------------ sample data
def sample_data(variant="demo"):
    rng = {"A": ("CNG Stations", 1, 2500, 2700), "B": ("CNG Stations", 1, 2000, 2200),
           "P1": ("PNG Households", 1, 1800, 2000), "P2": ("PNG Households", 1, 1600, 1800),
           "P3": ("PNG Households", 1, 1700, 1900), "P4": ("PNG Households", 1, 1400, 1500),
           "I1": ("Industries", 1, 1500, 1700), "I2": ("Industries", 1, 1000, 1100)}
    tier = {"PNG Households": 1, "CNG Stations": 2, "Industries": 3}
    rows, nodes = [], []
    for n, (typ, _, a, b) in rng.items():
        base = np.linspace(a, b, 12) * (1 + 0.03 * np.sin(2 * np.pi * (np.arange(12)) / 12))
        rows.append([n] + list(np.round(base, -1)))
        nodes.append([n, typ, tier[typ]])
    demand = pd.DataFrame(rows, columns=["Node"] + MCOLS)
    nodes = pd.DataFrame(nodes, columns=["Node", "Type", "Tier"])
    tiers = pd.DataFrame({"Tier": [1, 2, 3], "Customer": ["PNG Households", "CNG Stations", "Industries"],
                          "MinService": [0.95, 0.90, 0.80], "Penalty": [20, 15, 10]})
    img = variant == "image"   # "image" = exactly the figures in the problem picture; "demo" = corridors bind
    supply = pd.DataFrame({"Month": MONTHS, "Available": [11000, 11500, 11200, 11800, 12000, 12500, 12300, 11900, 11600, 11500, 11000, 10800]
                           if img else [14800, 14900, 15000, 14700, 14000, 13950, 14600, 14800, 14900, 14900, 14300, 14100]})
    CP = {1: 4700, 2: 2080, 3: 6850, 4: 3200, 5: 2650}
    arcs = [(1, "G", "A", 2.0, 10000), (2, "A", "B", 1.5, 8000), (3, "G", "P1", 1.8, 12000),
            (3, "G", "P2", 1.8, 12000), (4, "P1", "P3", 1.2, 9000), (4, "P1", "P4", 1.2, 9000),
            (4, "P2", "P3", 1.2, 9000), (4, "P2", "P4", 1.2, 9000), (5, "G", "I1", 2.5, 6000),
            (5, "G", "I2", 2.5, 6000)]
    corr = pd.DataFrame(arcs, columns=["Corridor", "From", "To", "Cost", "Capacity"])
    if not img: corr["Capacity"] = corr["Corridor"].map(CP)
    return dict(nodes=nodes, tiers=tiers, demand=demand, supply=supply, corridors=corr)

README = [
    ["Gas allocation - input template"], [""],
    ["Edit blue-header sheets; keep sheet names and column headers unchanged."],
    ["Nodes", "Node, Type, Tier - each demand node and its priority tier (Tier must exist in Tiers)."],
    ["Tiers", "Tier, Customer, MinService (fraction of demand, e.g. 0.95), Penalty (Rs/SCM shortage)."],
    ["Demand", "One row per node; M1..M12 = demand in SCM for each month."],
    ["Supply", "Month (1-12), Available = gas at city gate in SCM."],
    ["Corridors", "One row per arc: Corridor id, From, To, Cost (Rs/SCM), Capacity (SCM/month, shared by all arcs with the same Corridor id)."],
    ["Source node is always 'G'. Values in the sheets are the sample (image) data."],
]

def _style(ws, header_fill="1F3864"):
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=header_fill)
        c.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial")
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = max(12, min(45, max(len(str(c.value or "")) for c in col) + 2))

def sample_excel_bytes(data=None):
    data = data or sample_data()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(README).to_excel(xw, sheet_name="README", index=False, header=False)
        for name, key in [("Nodes", "nodes"), ("Tiers", "tiers"), ("Demand", "demand"),
                          ("Supply", "supply"), ("Corridors", "corridors")]:
            data[key].to_excel(xw, sheet_name=name, index=False); _style(xw.sheets[name])
        xw.sheets["README"].column_dimensions["A"].width = 14
        for r in xw.sheets["README"].iter_rows():
            for c in r: c.font = Font(name="Arial", bold=(c.row == 1))
    return buf.getvalue()

def read_excel_data(f):
    x = pd.ExcelFile(f)
    need = {"Nodes": "nodes", "Tiers": "tiers", "Demand": "demand", "Supply": "supply", "Corridors": "corridors"}
    miss = [s for s in need if s not in x.sheet_names]
    if miss: raise ValueError(f"Missing sheet(s): {miss}")
    d = {k: x.parse(s) for s, k in need.items()}
    for k, cols in dict(nodes=["Node", "Type", "Tier"], tiers=["Tier", "MinService", "Penalty"],
                        demand=["Node"] + MCOLS, supply=["Month", "Available"],
                        corridors=["Corridor", "From", "To", "Cost", "Capacity"]).items():
        m = [c for c in cols if c not in d[k].columns]
        if m: raise ValueError(f"Sheet for '{k}' is missing column(s): {m}")
    return d

# ------------------------------------------------------------------ model
def solve(data, sc=None):
    sc = {**DEFAULT_SC, **(sc or {})}
    nodes = data["nodes"].merge(data["tiers"][["Tier", "MinService", "Penalty"]], on="Tier", how="left")
    N = list(nodes["Node"]); dem = data["demand"].set_index("Node")
    D = {(n, t): float(dem.loc[n, f"M{t}"]) * sc["demand"] * sc["dem_node"].get(n, 1.0) for n in N for t in MONTHS}
    A = {t: float(data["supply"].set_index("Month").loc[t, "Available"]) * sc["supply"] for t in MONTHS}
    alpha = {r.Node: (float(sc["svc_all"]) if sc["svc_all"] is not None else float(np.clip(r.MinService + sc["svc_shift"], 0, 1)))
             for r in nodes.itertuples()}
    pen = {r.Node: float(r.Penalty) * sc["pen"] for r in nodes.itertuples()}
    arcs = data["corridors"].reset_index(drop=True)
    ids = list(arcs.index)
    cap = {k: float(g["Capacity"].iloc[0]) * sc["cap"] * sc["cap_corr"].get(k, 1.0)
           for k, g in arcs.groupby("Corridor")}
    cst = {i: float(arcs.loc[i, "Cost"]) * sc["cost"] for i in ids}

    P = pl.LpProblem("GasAllocation", pl.LpMinimize)
    mk = lambda p, keys: {a: {t: pl.LpVariable(f"{p}_{a}_{t}", lowBound=0) for t in MONTHS} for a in keys}
    x = mk("x", ids)          # flow on arc
    y = mk("y", N)            # unmet demand (shortage)
    z = mk("z", N)            # service-floor slack (soft floor)
    P += (pl.lpSum(cst[i] * x[i][t] for i in ids for t in MONTHS)
          + pl.lpSum(pen[n] * y[n][t] for n in N for t in MONTHS)
          + sc["floor_pen"] * pl.lpSum(z[n][t] for n in N for t in MONTHS))
    C = {}
    def net(n, t):
        return (pl.lpSum(x[i][t] for i in ids if arcs.loc[i, "To"] == n)
                - pl.lpSum(x[i][t] for i in ids if arcs.loc[i, "From"] == n))
    for t in MONTHS:
        C[("SRC", "G", t)] = pl.lpSum(x[i][t] for i in ids if arcs.loc[i, "From"] == "G") <= A[t]
        for k in cap:
            C[("CAP", k, t)] = pl.lpSum(x[i][t] for i in ids if arcs.loc[i, "Corridor"] == k) <= cap[k]
        for n in N:
            C[("BAL", n, t)] = net(n, t) + y[n][t] == D[n, t]
            C[("SVC", n, t)] = net(n, t) + z[n][t] >= alpha[n] * D[n, t]
            C[("YUB", n, t)] = y[n][t] <= D[n, t]
    for key, c in C.items(): P += c, "_".join(map(str, key))
    P.solve(pl.PULP_CBC_CMD(msg=False))
    status = pl.LpStatus[P.status]
    v = lambda var: var.value() or 0.0

    flows = pd.DataFrame([dict(Month=t, Corridor=arcs.loc[i, "Corridor"], From=arcs.loc[i, "From"],
                               To=arcs.loc[i, "To"], Flow=v(x[i][t]), Cost=cst[i], Transport_Cost=cst[i] * v(x[i][t]))
                          for i in ids for t in MONTHS])
    tp = nodes.set_index("Node")
    deliv = pd.DataFrame([dict(Month=t, Node=n, Type=tp.loc[n, "Type"], Tier=tp.loc[n, "Tier"], Demand=D[n, t],
                               Delivered=D[n, t] - v(y[n][t]), Shortage=v(y[n][t]), Required=alpha[n] * D[n, t],
                               Floor_Violation=v(z[n][t]), Fill_Rate=1 - v(y[n][t]) / D[n, t] if D[n, t] else 1.0)
                          for n in N for t in MONTHS])
    cu = flows.groupby(["Month", "Corridor"], as_index=False)["Flow"].sum()
    cu["Capacity"] = cu["Corridor"].map(cap); cu["Utilization"] = cu["Flow"] / cu["Capacity"]
    lab = dict(SRC="Source availability", CAP="Corridor capacity", BAL="Node balance (demand)", SVC="Service-level floor")
    du = []
    for (typ, item, t), c in C.items():
        if typ == "YUB": continue
        pi = c.pi or 0.0
        val = -pi if typ in ("SRC", "CAP") else pi   # Rs saved per +1 unit capacity / Rs per +1 SCM required
        du.append(dict(Type=lab[typ], Item=item, Month=t, Shadow_Price=val, Slack=c.slack,
                       Binding=abs(c.slack) < 1e-6 and abs(pi) > 1e-6))
    duals = pd.DataFrame(du)
    tr = flows["Transport_Cost"].sum()
    sp = float((deliv["Shortage"] * deliv["Node"].map(pen)).sum())
    fv = deliv["Floor_Violation"].sum()
    tr, sp, fv = float(tr), float(sp), float(fv)
    sh, dm = float(deliv["Shortage"].sum()), float(deliv["Demand"].sum())
    summary = dict(Status=status, Total_Cost=tr + sp, Transport_Cost=tr, Shortage_Penalty=sp,
                   Total_Shortage=sh, Total_Demand=dm, Fill_Rate=1 - sh / max(dm, 1e-9),
                   Floor_Violation_SCM=fv, Floor_Violation_Penalty=fv * sc["floor_pen"],
                   Objective_incl_floor_penalty=float(pl.value(P.objective) or 0.0))
    return dict(summary=summary, flows=flows, deliv=deliv, corr=cu, duals=duals, sc=sc)

# ------------------------------------------------------------------ sensitivity
def make_sc(param, v):
    sc = {**DEFAULT_SC, "cap_corr": {}}
    if param.startswith("Corridor "): sc["cap_corr"] = {int(param.split()[1]): v}
    else: sc[{"Gas supply": "supply", "Demand": "demand", "All corridor capacity": "cap",
              "Transport cost": "cost", "Shortage penalty": "pen", "Service level (pts shift)": "svc_shift"}[param]] = v
    return sc

def sweep(data, param, values):
    out = []
    for v in values:
        s = solve(data, make_sc(param, v))["summary"]
        out.append(dict(Value=v, **{k: s[k] for k in ["Total_Cost", "Transport_Cost", "Shortage_Penalty",
                                                      "Total_Shortage", "Floor_Violation_SCM", "Fill_Rate"]}))
    return pd.DataFrame(out)

def tornado(data, params, pct=0.10, svc_pts=0.05):
    base = solve(data)["summary"]["Total_Cost"]; rows = []
    for p in params:
        lo, hi = (-svc_pts, svc_pts) if p.startswith("Service") else (1 - pct, 1 + pct)
        cl = solve(data, make_sc(p, lo))["summary"]["Total_Cost"]; ch = solve(data, make_sc(p, hi))["summary"]["Total_Cost"]
        rows.append(dict(Parameter=p, Low_Cost=cl, High_Cost=ch, Delta_Low=cl - base, Delta_High=ch - base,
                         Swing=abs(ch - cl)))
    return pd.DataFrame(rows).sort_values("Swing", ascending=False)

def heatmap(data, sv=(0.8, 0.9, 1.0, 1.1, 1.2), dv=(0.8, 0.9, 1.0, 1.1, 1.2)):
    z = [[solve(data, {**DEFAULT_SC, "supply": s, "demand": d})["summary"]["Total_Cost"] for s in sv] for d in dv]
    return pd.DataFrame(z, index=[f"Demand x{d}" for d in dv], columns=[f"Supply x{s}" for s in sv])

# ------------------------------------------------------------------ diagram, shadow impact, service levels
def layout(data):
    arcs, depth, q = data["corridors"], {"G": 0}, ["G"]
    while q:
        u = q.pop(0)
        for t in arcs[arcs["From"] == u]["To"]:
            if t not in depth: depth[t] = depth[u] + 1; q.append(t)
    for n in data["nodes"]["Node"]: depth.setdefault(n, 1)
    mx = max(depth.values()) or 1; H = 460; P = {"G": (70, H / 2)}
    for d in range(1, mx + 1):
        L = [n for n in depth if depth[n] == d]
        tg = []
        for i, n in enumerate(L):
            par = [P[f][1] for f in arcs[arcs["To"] == n]["From"] if f in P]
            tg.append((i + .5) * H / len(L) if d == 1 or not par else float(np.mean(par)))
        last = -1e9
        for i in np.argsort(tg):
            y = max(tg[i], last + 66); last = y; P[L[i]] = (70 + d * 570 / mx, y)
    return P

def shadow_impact(data, sc, res, pct=0.10):
    sc = {**DEFAULT_SC, **sc}; du, cu = res["duals"], res["corr"]; base = res["summary"]
    cap0 = data["corridors"].groupby("Corridor")["Capacity"].first()
    A = data["supply"].sort_values("Month")["Available"].values * sc["supply"]
    dvt = res["deliv"]["Delivered"].sum(); rows = []
    for item in ["Source availability"] + [f"Corridor {k}" for k in cap0.index]:
        k = None if item.startswith("Source") else int(item.split()[1])
        s2 = {**sc, "cap_corr": dict(sc["cap_corr"])}
        if k is None:
            s2["supply"] = sc["supply"] * (1 + pct); extra = A * pct
            d = du[du.Type == "Source availability"].sort_values("Month")
        else:
            s2["cap_corr"][k] = sc["cap_corr"].get(k, 1) * (1 + pct)
            extra = np.full(12, cap0[k] * sc["cap"] * sc["cap_corr"].get(k, 1) * pct)
            d = du[(du.Type == "Corridor capacity") & (du.Item == k)].sort_values("Month")
        r2 = solve(data, s2); b2 = r2["summary"]
        cf = lambda r: r["flows"][r["flows"].Corridor == k]["Flow"].sum()
        u = cu[cu.Corridor == k]["Utilization"] if k is not None else None
        rows.append({"Constraint": item, "Avg shadow (Rs/SCM)": d.Shadow_Price.mean(), "Max shadow (Rs/SCM)": d.Shadow_Price.max(),
                     "Months binding": int((d.Shadow_Price > 1e-6).sum()),
                     "Avg util": None if u is None else u.mean(), "Peak util": None if u is None else u.max(),
                     f"Extra SCM/month ({pct:+.0%})": extra.mean(),
                     "Est. objective saving (linear)": float((d.Shadow_Price.values * extra).sum()),
                     "Actual objective saving": base["Objective_incl_floor_penalty"] - b2["Objective_incl_floor_penalty"],
                     "Change in real cost": b2["Total_Cost"] - base["Total_Cost"],
                     "Change in shortage (SCM)": b2["Total_Shortage"] - base["Total_Shortage"],
                     "Change in flow (SCM)": (r2["deliv"]["Delivered"].sum() - dvt) if k is None else cf(r2) - cf(res)})
    return pd.DataFrame(rows)

def demand_impact(data, sc, res, pct=0.10):
    sc = {**DEFAULT_SC, **sc}; b = res["summary"]; rows = []
    for r in data["nodes"].itertuples():
        s2 = {**sc, "dem_node": dict(sc["dem_node"])}; s2["dem_node"][r.Node] = sc["dem_node"].get(r.Node, 1.0) * (1 + pct)
        d0 = res["deliv"][res["deliv"].Node == r.Node]["Demand"].mean(); ex = d0 * pct; b2 = solve(data, s2)["summary"]
        rows.append({"Node": r.Node, "Type": r.Type, "Base demand SCM/mo": d0, "Extra SCM/mo": ex,
                     "Change in objective": b2["Objective_incl_floor_penalty"] - b["Objective_incl_floor_penalty"],
                     "Change in real cost": b2["Total_Cost"] - b["Total_Cost"], "Change in shortage (SCM)": b2["Total_Shortage"] - b["Total_Shortage"],
                     "Marginal Rs per extra SCM": (b2["Total_Cost"] - b["Total_Cost"]) / (ex * 12) if ex else 0.0})
    return pd.DataFrame(rows).sort_values("Marginal Rs per extra SCM", ascending=False)

def service_levels(data, sc, levels=(0.70, 0.75, 0.80, 0.85, 0.90, 0.95)):
    rows, prev = [], None
    for l in levels:
        r = solve(data, {**sc, "svc_all": l}); s = r["summary"]
        row = {"Service floor": l, "Total cost": s["Total_Cost"], "Transport": s["Transport_Cost"], "Shortage penalty": s["Shortage_Penalty"],
               "Shortage (SCM)": s["Total_Shortage"], "Fill rate": s["Fill_Rate"], "Floor breach (SCM)": s["Floor_Violation_SCM"],
               "Feasible": "Yes" if s["Floor_Violation_SCM"] < 1 else "No", "Cost vs previous level": 0 if prev is None else s["Total_Cost"] - prev}
        for t, g in r["deliv"].groupby("Tier"):
            row[f"Fill tier {t}"] = 1 - g.Shortage.sum() / g.Demand.sum()
        prev = s["Total_Cost"]; rows.append(row)
    return pd.DataFrame(rows)

# ------------------------------------------------------------------ export
def results_excel(res, data, sens=None):
    buf = io.BytesIO(); s = res["summary"]
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(list(s.items()), columns=["Metric", "Value"]).to_excel(xw, sheet_name="Summary", index=False)
        pd.DataFrame([{k: v for k, v in res["sc"].items() if k != "cap_corr"} | {"cap_corr": str(res["sc"]["cap_corr"])}]
                     ).T.reset_index().set_axis(["Scenario setting", "Value"], axis=1).to_excel(xw, sheet_name="Scenario", index=False)
        res["deliv"].to_excel(xw, sheet_name="Deliveries", index=False)
        res["deliv"].pivot(index="Node", columns="Month", values="Shortage").reset_index().to_excel(xw, sheet_name="Shortage_Matrix", index=False)
        res["flows"].to_excel(xw, sheet_name="Flows", index=False)
        res["corr"].to_excel(xw, sheet_name="Corridor_Utilization", index=False)
        res["duals"].to_excel(xw, sheet_name="Shadow_Prices", index=False)
        for name, df in (sens or {}).items():
            df.to_excel(xw, sheet_name=name[:31], index=isinstance(df.index, pd.Index) and df.index.name is not None or name == "Heatmap")
        for k in ["nodes", "tiers", "demand", "supply", "corridors"]:
            data[k].to_excel(xw, sheet_name=f"Input_{k}", index=False)
        for ws in xw.book.worksheets: _style(ws, "375623")
    return buf.getvalue()
