"""
Thin, UI-free layer between a front-end (Streamlit app or anything else) and the tested
planning engine. The engine itself is reused as-is via subprocess, so both front-ends
(Excel .bat flow and the web app) are guaranteed to produce identical results.
"""
import os
import subprocess
import sys
import tempfile
import time
import pandas as pd

ENGINE_DIR = os.path.dirname(os.path.abspath(__file__))
TOOL_DIR = os.path.dirname(ENGINE_DIR)
ENGINE = os.path.join(ENGINE_DIR, "planning_engine.py")
BASE_INPUTS = os.path.join(TOOL_DIR, "Levisol Planning Tool - INPUTS.xlsx")

INPUT_SHEETS = [
    "Settings", "Jan Demand Forecast", "Plants & Capacity", "Transport Plant-Hub",
    "Transport Hub-CFA", "SKU Master", "Service Levels (Tiers)", "Sales History",
    "Lead Times", "Opening Inventory CFA", "Opening Inventory Hub", "Hub SS Override",
]
OUTPUT_SHEETS = [
    "Run Log", "Dashboard", "Cost Summary", "Production Plan", "Capacity Utilisation",
    "Routing Plant-Hub", "Routing Hub-CFA", "Hub Stock Position", "Unmet Demand",
    "Inventory Norms (CFA)", "Inventory Norms (Hub)", "Net Requirement",
]


def load_inputs(path=BASE_INPUTS):
    """Read all editable input sheets into a dict of DataFrames."""
    xls = pd.ExcelFile(path)
    return {s: pd.read_excel(xls, sheet_name=s) for s in INPUT_SHEETS if s in xls.sheet_names}


def save_inputs(inputs, path):
    """Write a dict of input DataFrames to an INPUTS-format workbook."""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in inputs.items():
            df.to_excel(w, sheet_name=name, index=False)


def apply_settings(inputs, **overrides):
    """Override Settings rows by name, e.g. apply_settings(inp, **{'Solver time limit (seconds)': 60})."""
    s = inputs["Settings"].copy()
    for name, value in overrides.items():
        mask = s["Setting"].astype(str).str.strip() == name
        s.loc[mask, "Value"] = value
    inputs["Settings"] = s
    return inputs


def run_scenario(inputs, workdir=None):
    """Write inputs to a temp workbook, run the tested engine on it, read the output back.
    Returns dict with: ok (bool), outputs (dict of DataFrames), stdout (str), seconds (float)."""
    workdir = workdir or tempfile.mkdtemp(prefix="levisol_")
    in_path = os.path.join(workdir, "scenario_inputs.xlsx")
    out_path = os.path.join(workdir, "scenario_output.xlsx")
    save_inputs(inputs, in_path)
    t0 = time.time()
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    r = subprocess.run([sys.executable, ENGINE, in_path, out_path],
                       capture_output=True, text=True, env=env, timeout=3900)
    secs = time.time() - t0
    outputs = {}
    if os.path.exists(out_path):
        xls = pd.ExcelFile(out_path)
        outputs = {s: pd.read_excel(xls, sheet_name=s) for s in xls.sheet_names}
    return {"ok": r.returncode == 0, "outputs": outputs, "stdout": r.stdout,
            "seconds": secs, "output_path": out_path if os.path.exists(out_path) else None}


def kpis(outputs):
    """Dashboard sheet as a {KPI: Value} dict."""
    if "Dashboard" not in outputs:
        return {}
    d = outputs["Dashboard"]
    return dict(zip(d["KPI"], d["Value"]))
