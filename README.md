# Levisol Monthly Planning Tool

A supply-chain planning tool for the Castrol "Power Up 4.0" case competition. A planner edits
inputs (demand, capacities, costs), clicks Run, and gets a cost-optimal production & distribution
plan for the Levisol network (3 plants → 2 hubs → 10 CFAs, 100 SKUs), including inventory norms,
routing, a network map, and side-by-side scenario comparison.

## What's inside

| Path | Purpose |
|---|---|
| `app.py` | Streamlit web app (4 tabs: Inputs / Plan Results / Network Map / Compare Scenarios) |
| `engine/planning_engine.py` | The analytical core: input validation → inventory norms → net requirement → MILP optimisation (PuLP + CBC) → results. Never crashes on bad input; writes a Run Log instead. |
| `engine/scenario_runner.py` | Bridge between the app and the engine |
| `engine/build_inputs_workbook.py` | Regenerates the INPUTS workbook from the original case data file |
| `Levisol Planning Tool - INPUTS.xlsx` | Base input data (editable in the app; original case data is never modified) |
| `.streamlit/config.toml` | Theme |
| `requirements.txt` | Python dependencies |

## Run locally

Requires Python 3.11+.

```
pip install -r requirements.txt
python -m streamlit run app.py
```

(On Windows you can also just double-click `RUN WEB APP.bat`.)

## Deploy to Streamlit Community Cloud (free)

1. Push this folder to a GitHub repository (this folder as the repo root).
2. Go to https://share.streamlit.io → "Create app" → pick the repo, branch, and `app.py`.
3. Deploy. First build takes a few minutes; the app then gets a public URL you can share/submit.

Notes for cloud use:
- The CBC solver ships inside the `pulp` package (Linux binary included) — no extra install.
- Free-tier cloud CPUs are slower than a laptop; prefer the 60s or 150s solver setting there.
  For the live demo, running locally is the most reliable option (no venue-WiFi dependency).

## Using the tool

1. **Inputs tab** — edit any table (demand forecast, plant capacities, transport costs, SKU
   penalties, lead times, opening inventory). "Reset all inputs" restores the case data.
2. **Sidebar** — name the scenario, pick solver time / requirement basis (ROP = conservative,
   SS = leaner) / contractual policy, press **Run plan**.
3. **Plan Results** — KPIs, cost breakdown, plant utilisation, what could not be supplied (and
   what it costs), full plan tables, downloadable Excel output.
4. **Network Map** — plants/hubs/CFAs on an India map with flow lines proportional to volume.
5. **Compare Scenarios** — every run is kept; compare costs and service side by side.
