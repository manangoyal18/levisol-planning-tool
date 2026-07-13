"""
Levisol Monthly Planning Tool - web front-end.
Runs the SAME tested engine as the Excel tool (engine/planning_engine.py) via
engine/scenario_runner.py, and adds: in-browser input editing, KPI dashboard,
charts, a network routing map, and side-by-side scenario comparison.

Launch:  double-click "RUN WEB APP.bat"  (or:  py -m streamlit run app.py)
"""
import os
import sys
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "engine"))
import scenario_runner as sr

# ---------------- palette ----------------
# Castrol brand accents (matches the presentation deck)
CGREEN, CDARK, CRED = "#00843D", "#005228", "#E4002B"
# chart series (validated dataviz reference palette, slot 1 swapped to Castrol green)
BLUE, AQUA, YELLOW, GREEN, VIOLET, RED = "#2a78d6", "#1baf7a", "#eda100", "#008300", "#4a3aa7", "#e34948"
SERIES = [CGREEN, BLUE, YELLOW, VIOLET, AQUA, RED]
INK, INK2, MUTED, GRID, SURFACE = "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#fcfcfb"

st.set_page_config(page_title="Levisol Planning Tool", page_icon="🛢️", layout="wide")

# banner + light brand styling (headings deep green, red hairline under the banner, grey sidebar edge)
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_banner = os.path.join(_APP_DIR, "castrol_image.webp")
if os.path.exists(_banner):
    st.image(_banner, width="stretch")
st.markdown(
    f"""
    <div style="height:4px;background:{CRED};margin:-10px 0 14px 0;border-radius:2px;"></div>
    <style>
      h1, h2, h3 {{ color: {CDARK}; }}
      [data-testid="stSidebar"] {{ border-right: 3px solid {CGREEN}; }}
      [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 {{ color: {CDARK}; }}
      [data-testid="stMetricValue"] {{ color: {CDARK}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

CITY_LATLON = {
    "Guwahati CFA": (26.14, 91.74), "Kolkata CFA": (22.57, 88.36), "Jamshedpur CFA": (22.80, 86.18),
    "Kanpur CFA": (26.45, 80.33), "Haryana CFA": (28.46, 76.99), "Rajpura CFA": (30.48, 76.59),
    "Bhiwandi CFA": (19.30, 73.06), "Bangalore CFA": (12.97, 77.59), "Ahmedabad CFA": (23.02, 72.57),
    "Hyderabad CFA": (17.38, 78.49),
    "MHW": (19.60, 73.20), "MHE": (22.90, 88.20),
    "BOM": (19.08, 72.88), "AHM": (23.02, 72.57), "KOL": (22.57, 88.36),
}

def base_fig(height=380):
    fig = go.Figure()
    fig.update_layout(
        height=height, plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
        font=dict(family='system-ui, "Segoe UI", sans-serif', color=INK2, size=13),
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=dict(gridcolor=GRID, zerolinecolor=GRID, tickfont=dict(color=MUTED)),
        yaxis=dict(gridcolor=GRID, zerolinecolor=GRID, tickfont=dict(color=MUTED)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    return fig

# ---------------- state ----------------
if "inputs" not in st.session_state:
    st.session_state.inputs = sr.load_inputs()
if "scenarios" not in st.session_state:
    st.session_state.scenarios = {}   # name -> result dict from scenario_runner
if "inputs_version" not in st.session_state:
    st.session_state.inputs_version = 0   # bumped on upload/reset so data editors refresh cleanly
if "upload_names" not in st.session_state:
    st.session_state.upload_names = {}    # sheet -> filename of the upload currently applied

# ---------------- generic sheet upload ----------------
# Required columns per uploadable sheet (mirrors the engine's validation layer).
MONTH_COLS = [m + " (kL)" for m in ["Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25"]]
SHEET_REQUIRED = {
    "Jan Demand Forecast": ["Product Name", "CFA", "Jan-26 Forecast (kL)"],
    "Plants & Capacity": ["Plant Code", "Capacity <=1.5 LT (kL)", "Capacity 3-5 LT (kL)",
                           "Capacity 7-20 LT (kL)", "Capacity 50 LT (kL)",
                           "Capacity 180-210 LT (kL)", "Production Cost (Rs per kL)"],
    "Transport Plant-Hub": ["From Plant Code", "To MHW (Rs per kL)", "To MHE (Rs per kL)"],
    "Transport Hub-CFA": ["CFA", "From MHW (Rs per kL)", "From MHE (Rs per kL)"],
    "SKU Master": ["Product Name", "Capacity Line", "Penalty Cost (Rs per kL)", "Contractual (Yes/No)"],
    "Service Levels (Tiers)": ["Tier", "Volume Share", "Target Fill Rate"],
    "Sales History": ["Product Name", "CFA"] + MONTH_COLS,
    "Lead Times": ["Product Name", "CFA", "Source Hub", "Plant to Hub LT (days)",
                    "Hub to CFA LT (days)", "Production LT (days)",
                    "Production Variability (days)", "Transit Variability (days)"],
    "Opening Inventory CFA": ["Product Name", "CFA", "Opening Inventory (kL)"],
    "Opening Inventory Hub": ["Product Name", "Hub", "Opening Inventory (kL)"],
    "Hub SS Override": ["Product Name", "Hub (MHW/MHE)", "Override Safety Stock (kL)"],
}
# Optional columns preserved if the upload has them
SHEET_OPTIONAL = {
    "Jan Demand Forecast": ["Pack size", "CFA region"],
    "Plants & Capacity": ["Location"],
    "Transport Hub-CFA": ["Region"],
    "SKU Master": ["Pack size"],
}

import re as _re

def _canon(s):
    """Normalise a header for matching: lowercase, drop parenthetical units, drop punctuation.
    'Capacity <=1.5 LT (kL)' and 'capacity 1.5 lt' both become 'capacity15lt'."""
    s = str(s).lower()
    s = _re.sub(r"\(.*?\)", "", s)
    return _re.sub(r"[^a-z0-9]", "", s)

# alternative header spellings accepted per canonical column (canon-form)
_ALIASES = {
    "productname": {"product", "sku", "skuname", "skucode", "productname", "item"},
    "cfa": {"cfa", "cfaname", "warehouse", "depot"},
    "hub": {"hub", "hubcode", "hubname"},
    "hubmhwmhe": {"hub", "hubcode", "hubname", "hubmhwmhe"},
    "openinginventory": {"openinginventory", "openingstock", "opening", "openinginv", "stock"},
    "jan26forecast": {"jan26forecast", "forecast", "demand", "jan26", "forecastkl", "demandkl", "janforecast"},
}

def parse_uploaded_sheet(file, sheet_name):
    """Read an uploaded .xlsx (first sheet) or .csv and normalise it to the columns the engine
    expects for `sheet_name`. Flexible about header spellings. Returns (df, error_message)."""
    try:
        if file.name.lower().endswith(".csv"):
            df = pd.read_csv(file)
        else:
            xls = pd.ExcelFile(file)
            df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    except Exception as ex:
        return None, f"could not read the file ({ex})."
    if df.empty:
        return None, "the uploaded file has no data rows."

    uploaded = {_canon(c): c for c in df.columns}
    rename, missing = {}, []
    for expected in SHEET_REQUIRED[sheet_name] + SHEET_OPTIONAL.get(sheet_name, []):
        exp_c = _canon(expected)
        found = uploaded.get(exp_c)
        if found is None:
            for alias in _ALIASES.get(exp_c, ()):  # accepted alternative spellings
                if alias in uploaded:
                    found = uploaded[alias]
                    break
        if found is None and exp_c == "jan26forecast":
            # last resort for the forecast column: any remaining numeric column
            numeric = [c for c in df.columns
                       if c not in rename and pd.to_numeric(df[c], errors="coerce").notna().any()]
            if numeric:
                found = numeric[0]
        if found is not None:
            rename[found] = expected
        elif expected in SHEET_REQUIRED[sheet_name]:
            missing.append(expected)
    if missing:
        return None, (f"required column(s) not found: {missing}. "
                      f"Columns in your file: {list(df.columns)}.")

    df = df.rename(columns=rename)
    keep = [c for c in SHEET_REQUIRED[sheet_name] + SHEET_OPTIONAL.get(sheet_name, []) if c in df.columns]
    df = df[keep].copy()
    for key_col in ("Product Name", "CFA", "Hub", "Plant Code", "From Plant Code"):
        if key_col in df.columns:
            df[key_col] = df[key_col].astype(str).str.strip()
    if "CFA" in df.columns:  # tolerate CFA names without the ' CFA' suffix
        df.loc[~df["CFA"].str.upper().str.endswith(" CFA"), "CFA"] = df["CFA"] + " CFA"
    return df.reset_index(drop=True), None

st.title("Levisol Monthly Planning Tool")
st.caption("Edit inputs, run, read the plan. Same verified engine as the Excel tool; "
           "every scenario you run is kept for side-by-side comparison.")

# ---------------- sidebar: settings + run ----------------
with st.sidebar:
    st.header("Run a scenario")
    scen_name = st.text_input("Scenario name", value=f"Scenario {len(st.session_state.scenarios) + 1}")
    time_limit = st.select_slider("Solver time limit (seconds)", options=[60, 150, 300, 600], value=60,
                                   help="60s is within a few % of optimal; use 600s for a final run.")
    basis = st.radio("Replenish CFAs up to...", ["ROP", "SS"], horizontal=True,
                     help="ROP = reorder point (primary, per assignment). SS = safety stock only (leaner).")
    contr_mult = st.number_input("Contractual penalty multiplier", 1.0, 100.0, 5.0, 1.0)
    hub_factor = st.number_input("Hub shortfall penalty factor", 0.0, 10.0, 0.5, 0.1)
    force_contr = st.toggle("Force contractual supply (hard rule)", value=False)
    run_btn = st.button("Run plan", type="primary", width="stretch")
    if st.session_state.scenarios:
        st.divider()
        if st.button("Clear all scenarios", width="stretch"):
            st.session_state.scenarios = {}
            st.rerun()

if run_btn:
    inputs = {k: v.copy() for k, v in st.session_state.inputs.items()}
    sr.apply_settings(
        inputs,
        **{"Solver time limit (seconds)": time_limit,
           "Requirement basis (ROP or SS)": basis,
           "Contractual penalty multiplier": contr_mult,
           "Hub shortfall penalty factor": hub_factor,
           "Force contractual supply (Yes/No)": "Yes" if force_contr else "No"})
    with st.status(f"Running '{scen_name}' (up to {time_limit}s solver + overhead)...", expanded=True) as status:
        st.write("Validating inputs, computing norms, optimising...")
        result = sr.run_scenario(inputs)
        if result["ok"]:
            status.update(label=f"'{scen_name}' complete in {result['seconds']:.0f}s", state="complete")
        else:
            status.update(label=f"'{scen_name}' failed input validation", state="error")
    st.session_state.scenarios[scen_name] = result
    st.rerun()

tab_inputs, tab_results, tab_map, tab_compare = st.tabs(
    ["Inputs", "Plan Results", "Network Map", "Compare Scenarios"])

# ---------------- inputs tab ----------------
with tab_inputs:
    left, right = st.columns([3, 1])
    with left:
        st.markdown("Edit any table below - changes apply to the **next run**. "
                    "The original case data is never modified on disk.")
    with right:
        if st.button("Reset all inputs to case data", width="stretch"):
            st.session_state.inputs = sr.load_inputs()
            st.session_state.inputs_version += 1
            st.session_state.upload_names = {}
            st.rerun()

    def sheet_section(name, height=300):
        """One input sheet: an optional file upload that replaces the table, plus the editable grid."""
        ver = st.session_state.inputs_version
        up = st.file_uploader(
            f"Optional: upload a file to replace this table (.xlsx first sheet, or .csv). "
            f"Needs columns: {', '.join(SHEET_REQUIRED[name])} (flexible spellings accepted).",
            type=["xlsx", "csv"], key=f"up_{name}_v{ver}")
        if up is not None:
            new_df, err = parse_uploaded_sheet(up, name)
            if err:
                st.error(f"Upload not applied - {err} The existing table is unchanged.")
            else:
                st.session_state.inputs[name] = new_df
                st.session_state.upload_names[name] = f"{up.name} ({len(new_df)} rows)"
                st.session_state.inputs_version += 1
                st.rerun()
        if name in st.session_state.upload_names:
            st.success(f"Currently using uploaded file: {st.session_state.upload_names[name]}. "
                       "Rows are still editable below; 'Reset all inputs' restores the case data.")
        st.session_state.inputs[name] = st.data_editor(
            st.session_state.inputs[name], num_rows="dynamic", width="stretch",
            key=f"ed_{name}_v{ver}", height=height)

    main_sheets = ["Jan Demand Forecast", "Plants & Capacity", "Transport Plant-Hub",
                   "Transport Hub-CFA", "SKU Master"]
    ref_sheets = ["Service Levels (Tiers)", "Sales History", "Lead Times",
                  "Opening Inventory CFA", "Opening Inventory Hub", "Hub SS Override"]
    for name in main_sheets:
        with st.expander(f"**{name}**", expanded=(name == "Jan Demand Forecast")):
            if name == "Jan Demand Forecast":
                st.caption("Note: an upload replaces the entire table - only the SKU-CFA rows in "
                           "your file will be planned.")
            sheet_section(name, height=300)
    st.markdown("**Reference inputs** (used for norms & opening positions - also editable)")
    for name in ref_sheets:
        with st.expander(name):
            sheet_section(name, height=260)

# ---------------- helpers for results ----------------
def pick_scenario(key):
    names = list(st.session_state.scenarios.keys())
    if not names:
        st.info("No scenarios yet - set up a run in the sidebar and press **Run plan**.")
        return None, None
    name = st.selectbox("Scenario", names, index=len(names) - 1, key=key)
    return name, st.session_state.scenarios[name]

# ---------------- results tab ----------------
with tab_results:
    name, res = pick_scenario("res_pick")
    if res is not None:
        if not res["ok"]:
            st.error("This run failed input validation - the engine stopped safely. Details below.")
            if "Run Log" in res["outputs"]:
                st.dataframe(res["outputs"]["Run Log"], width="stretch")
            st.code(res["stdout"] or "(no console output)")
        else:
            out = res["outputs"]
            k = sr.kpis(out)
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total cost (Rs)", k.get("Total cost (Rs)", "-"))
            c2.metric("Fill rate", k.get("Fill rate", "-"))
            c3.metric("Unmet (kL)", k.get("Unmet (kL)", "-"))
            c4.metric("Production (kL)", k.get("Total production (kL)", "-"))
            c5.metric("Contractual unmet (kL)", k.get("Contractual unmet (kL)", "-"))

            colA, colB = st.columns(2)
            with colA:
                st.subheader("Cost breakdown")
                cost = out["Cost Summary"]
                cost = cost[cost["Cost Component"] != "TOTAL"]
                fig = base_fig(340)
                fig.add_bar(y=cost["Cost Component"], x=cost["Rs"], orientation="h",
                            marker=dict(color=CGREEN, cornerradius=4), width=0.55,
                            text=[f"Rs {v:,.0f}" for v in cost["Rs"]], textposition="outside",
                            textfont=dict(color=INK2),
                            hovertemplate="%{y}: Rs %{x:,.0f}<extra></extra>")
                fig.update_layout(xaxis_title=None, yaxis=dict(autorange="reversed"), showlegend=False)
                fig.update_xaxes(showticklabels=False, showgrid=False)
                st.plotly_chart(fig, width="stretch")
            with colB:
                st.subheader("Plant line utilisation")
                util = out["Capacity Utilisation"].dropna(subset=["Utilisation"])
                fig = base_fig(340)
                for i, p in enumerate(util["Plant"].unique()):
                    sub = util[util["Plant"] == p]
                    fig.add_bar(x=sub["Capacity Line"], y=sub["Utilisation"] * 100, name=p,
                                marker=dict(color=SERIES[i % len(SERIES)], cornerradius=4), width=0.22,
                                text=[f"{v * 100:.0f}%" for v in sub["Utilisation"]],
                                textposition="outside", textfont=dict(color=INK2, size=11),
                                hovertemplate=p + " | %{x}: %{y:.1f}%<extra></extra>")
                fig.update_layout(barmode="group", bargap=0.25, yaxis_title="% of capacity",
                                  yaxis=dict(range=[0, 118]))
                st.plotly_chart(fig, width="stretch")

            st.markdown(f'<h3 style="color:{CRED};">What could not be supplied</h3>', unsafe_allow_html=True)
            unmet = out["Unmet Demand"]
            if "Unmet (kL)" in unmet.columns and len(unmet):
                st.dataframe(unmet, width="stretch", height=240)
            else:
                st.success("Nothing - all net requirements fully supplied.")
            hubshort = out["Hub Stock Position"]
            if "Shortfall (kL)" in hubshort.columns:
                short = hubshort[hubshort["Shortfall (kL)"] > 1e-4]
                if len(short):
                    with st.expander(f"Hub safety-stock shortfalls ({len(short)} positions, "
                                     f"{short['Shortfall (kL)'].sum():,.1f} kL)"):
                        st.dataframe(short, width="stretch")

            with st.expander("Production plan (SKU x plant)"):
                st.dataframe(out["Production Plan"], width="stretch")
            with st.expander("Routing: plant to hub"):
                st.dataframe(out["Routing Plant-Hub"], width="stretch")
            with st.expander("Routing: hub to CFA"):
                st.dataframe(out["Routing Hub-CFA"], width="stretch")
            with st.expander("Inventory norms (CFA & hub)"):
                st.dataframe(out["Inventory Norms (CFA)"], width="stretch")
                st.dataframe(out["Inventory Norms (Hub)"], width="stretch")
            with st.expander("Run log"):
                st.dataframe(out["Run Log"], width="stretch")
            if res["output_path"] and os.path.exists(res["output_path"]):
                with open(res["output_path"], "rb") as f:
                    st.download_button("Download this plan as Excel", f.read(),
                                       file_name=f"Levisol Plan - {name}.xlsx")

# ---------------- map tab ----------------
with tab_map:
    name, res = pick_scenario("map_pick")
    if res is not None and res["ok"]:
        out = res["outputs"]
        ph, hc = out["Routing Plant-Hub"], out["Routing Hub-CFA"]
        fig = go.Figure()
        if len(ph) and "Qty (kL)" in ph.columns:
            agg = ph.groupby(["From Plant", "To Hub"])["Qty (kL)"].sum().reset_index()
            maxq = max(agg["Qty (kL)"].max(), 1)
            for _, r in agg.iterrows():
                a, b = CITY_LATLON.get(r["From Plant"]), CITY_LATLON.get(r["To Hub"])
                if a and b:
                    fig.add_trace(go.Scattergeo(
                        lat=[a[0], b[0]], lon=[a[1], b[1]], mode="lines",
                        line=dict(width=1 + 6 * r["Qty (kL)"] / maxq, color=VIOLET), opacity=0.55,
                        name="Plant to Hub", legendgroup="ph", showlegend=False,
                        hoverinfo="text", text=f"{r['From Plant']} -> {r['To Hub']}: {r['Qty (kL)']:,.0f} kL"))
        if len(hc) and "Qty (kL)" in hc.columns:
            agg = hc.groupby(["From Hub", "To CFA"])["Qty (kL)"].sum().reset_index()
            maxq = max(agg["Qty (kL)"].max(), 1)
            for _, r in agg.iterrows():
                a, b = CITY_LATLON.get(r["From Hub"]), CITY_LATLON.get(r["To CFA"])
                if a and b:
                    fig.add_trace(go.Scattergeo(
                        lat=[a[0], b[0]], lon=[a[1], b[1]], mode="lines",
                        line=dict(width=0.8 + 5 * r["Qty (kL)"] / maxq, color=BLUE), opacity=0.45,
                        name="Hub to CFA", legendgroup="hc", showlegend=False,
                        hoverinfo="text", text=f"{r['From Hub']} -> {r['To CFA']}: {r['Qty (kL)']:,.0f} kL"))
        groups = [("Plant", ["BOM", "AHM", "KOL"], VIOLET, "square"),
                  ("Hub", ["MHW", "MHE"], BLUE, "diamond"),
                  ("CFA", [c for c in CITY_LATLON if c.endswith(" CFA")], AQUA, "circle")]
        for label, nodes, colour, symbol in groups:
            fig.add_trace(go.Scattergeo(
                lat=[CITY_LATLON[n][0] for n in nodes], lon=[CITY_LATLON[n][1] for n in nodes],
                mode="markers+text", text=nodes, textposition="top center",
                textfont=dict(color=INK2, size=11),
                marker=dict(size=13, color=colour, symbol=symbol,
                            line=dict(width=2, color=SURFACE)),
                name=label, hoverinfo="text"))
        fig.update_layout(
            height=620, paper_bgcolor=SURFACE, margin=dict(l=0, r=0, t=10, b=0),
            legend=dict(orientation="h", y=0.02, x=0.02, bgcolor="rgba(252,252,251,0.8)"),
            geo=dict(scope="asia", projection_type="mercator",
                     lataxis_range=[6, 34], lonaxis_range=[66, 96],
                     showland=True, landcolor="#f0efec", countrycolor="#c3c2b7",
                     showcountries=True, bgcolor=SURFACE))
        st.plotly_chart(fig, width="stretch")
        st.caption("Line width is proportional to volume shipped. Hover any line or node for details. "
                   "Violet = plant-to-hub legs, blue = hub-to-CFA legs.")

# ---------------- compare tab ----------------
with tab_compare:
    ok_scen = {n: r for n, r in st.session_state.scenarios.items() if r["ok"]}
    if len(ok_scen) < 2:
        st.info("Run at least two scenarios (e.g. baseline, then change a capacity or cost) to compare them here.")
    else:
        rows = []
        for n, r in ok_scen.items():
            k = sr.kpis(r["outputs"])
            cost = r["outputs"]["Cost Summary"]
            total = float(cost.loc[cost["Cost Component"] == "TOTAL", "Rs"].values[0])
            rows.append({"Scenario": n, "Total cost (Rs)": total,
                         "Fill rate": k.get("Fill rate"), "Unmet (kL)": k.get("Unmet (kL)"),
                         "Production (kL)": k.get("Total production (kL)"),
                         "Contractual unmet (kL)": k.get("Contractual unmet (kL)")})
        cmp_df = pd.DataFrame(rows)
        base = cmp_df["Total cost (Rs)"].iloc[0]
        cmp_df["Delta cost vs first (Rs)"] = cmp_df["Total cost (Rs)"] - base
        st.dataframe(cmp_df.style.format({"Total cost (Rs)": "Rs {:,.0f}", "Delta cost vs first (Rs)": "Rs {:+,.0f}"}),
                     width="stretch")

        st.subheader("Cost components by scenario")
        fig = base_fig(400)
        for i, (n, r) in enumerate(ok_scen.items()):
            cost = r["outputs"]["Cost Summary"]
            cost = cost[cost["Cost Component"] != "TOTAL"]
            fig.add_bar(x=cost["Cost Component"], y=cost["Rs"], name=n,
                        marker=dict(color=SERIES[i % len(SERIES)], cornerradius=4), width=0.18,
                        hovertemplate=n + " | %{x}: Rs %{y:,.0f}<extra></extra>")
        fig.update_layout(barmode="group", bargap=0.3, yaxis_title="Rs")
        st.plotly_chart(fig, width="stretch")
