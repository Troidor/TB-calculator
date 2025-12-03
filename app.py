# app.py - TB Calculator (explicit converted core formulas)
import streamlit as st
import math
from copy import deepcopy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import re, os

# ---- load workbook once (the uploaded workbook used to generate constants) ----
WB_PATH = "TB G5 - G6 Calculator for SK.xlsx"  # place the uploaded workbook in repo root for reference
# If you deploy without the .xlsx (preferred), the VALUES dict below will be used instead.
try:
    wb = load_workbook(WB_PATH, data_only=False)
    ws = wb.active
    def cell(r,c):
        return ws.cell(row=r, column=c).value
    # read the constants for the rows we need into a dictionary
    VALUES = {}
    # we'll read columns A..L for rows 1..30 to have the labels and constants
    for r in range(1, 40):
        for c in range(1, 16):
            coord = get_column_letter(c) + str(r)
            VALUES[coord] = ws.cell(row=r, column=c).value
except Exception:
    # If workbook not present, fall back to empty dict (app will still run if VALUES filled)
    VALUES = {}

# If you prefer to embed constants directly, replace/extend this VALUES dict with the saved values.
# Example: VALUES['H8'] = 20

# ==== helper functions ====
def safe_float(x, default=0.0):
    try:
        if x is None: return default
        return float(x)
    except:
        return default

def rounddown_to_tens(x):
    # implements Excel's ROUNDDOWN(x, -1) -> rounds DOWN to nearest 10
    try:
        x = float(x)
    except:
        return None
    return math.floor(x/10.0) * 10

# ---- rows we will compute explicitly ----
NORMAL_ROWS = list(range(8, 16))   # 8..15
MERC_ROWS   = list(range(19, 27))  # 19..26

st.set_page_config(page_title="TB Calculator (Converted)", layout="wide")
st.title("TB Calculator — Inputs & Results (explicit Python conversion)")

st.markdown("**Instructions:** Enter B2 and percentages for the listed units (enter `100` = 100%). Click **Calculate**.")

# ---- Inputs ----
inputs = {}

# B2 numeric input
label_b2 = VALUES.get("A2", "Target Leadership")
default_b2 = safe_float(VALUES.get("B2", 0))
inputs["B2"] = st.number_input(f"{label_b2} (B2)", value=default_b2, format="%.6g")

st.subheader("Unit percentage inputs (G8–G15 and G19–G26)")
# build inputs for G8..G15 and G19..G26
for r in NORMAL_ROWS + MERC_ROWS:
    coordG = f"G{r}"
    label = VALUES.get(f"A{r}", coordG)
    default = safe_float(VALUES.get(coordG, 0))
    v = st.number_input(f"{label} ({coordG}) - %", value=default, format="%.6g")
    inputs[coordG] = v / 100.0   # convert percent (100 -> 1.0)

# ---- Calculation (explicit formulas) ----
if st.button("Calculate"):
    vm = deepcopy(VALUES)   # start from constants (strings, numbers)
    # inject user inputs into vm
    for k,v in inputs.items():
        vm[k] = v

    # We'll compute Army (col B) and Leadership (col C) explicitly
    # Rows 8..15 (Normal Troops)
    for r in NORMAL_ROWS:
        # read needed constants:
        I = safe_float(vm.get(f"I{r}", 0))   # I_r
        H = safe_float(vm.get(f"H{r}", 1))   # H_r
        G = safe_float(vm.get(f"G{r}", 0))   # input percent as decimal
        E = safe_float(vm.get(f"E{r}", 0))   # leadership per active troop (E)
        # Army formula: ROUNDDOWN( (I / H) * G , -1 )
        if H == 0:
            army = 0
        else:
            raw = (I / H) * G
            # Excel's order: often I/H -> number of stacks, * G (percentage), then ROUNDDOWN to tens
            # multiply raw by 1 (we already used G as decimal)
            army = rounddown_to_tens(raw)
        vm[f"B{r}"] = army
        # Leadership: = B_r * E_r
        vm[f"C{r}"] = safe_float(army) * E

    # Totals for block 1
    total_army_1 = sum(safe_float(vm.get(f"B{r}", 0)) for r in NORMAL_ROWS)
    total_lead_1 = sum(safe_float(vm.get(f"C{r}", 0)) for r in NORMAL_ROWS)
    vm["B16"] = total_army_1
    vm["C16"] = total_lead_1

    # Rows 19..26 (Mercs / second block)
    for r in MERC_ROWS:
        I = safe_float(vm.get(f"I{r}", 0))
        H = safe_float(vm.get(f"H{r}", 1))
        G = safe_float(vm.get(f"G{r}", 0))
        # leadership multiplier for mercs section assumed in column F (based on your screenshots)
        Fcol = safe_float(vm.get(f"F{r}", 0))
        if H == 0:
            army = 0
        else:
            raw = (I / H) * G
            army = rounddown_to_tens(raw)
        vm[f"B{r}"] = army
        vm[f"C{r}"] = safe_float(army) * Fcol

    # Totals for merc block
    total_army_2 = sum(safe_float(vm.get(f"B{r}", 0)) for r in MERC_ROWS)
    total_lead_2 = sum(safe_float(vm.get(f"C{r}", 0)) for r in MERC_ROWS)
    vm["B27"] = total_army_2
    vm["C27"] = total_lead_2

    # ---- Prepare results tables for display ----
    def table_for_range(start_row, end_row):
        rows = []
        for r in range(start_row, end_row + 1):
            name = vm.get(f"A{r}", "")
            colB = vm.get(f"B{r}", "")
            colC = vm.get(f"C{r}", "")
            rows.append([name, colB, colC])
        return rows

    st.header("Results — Block 1 (A8:C16)")
    st.table(table_for_range(8, 16))

    st.header("Results — Block 2 (A19:C26)")
    st.table(table_for_range(19, 26))

    st.markdown("**Notes:**\n- Army numbers are rounded down to the nearest 10 (Excel's ROUNDDOWN(...,-1)).\n- If any row's leadership multiplier column differs from the assumed column (E for block1, F for block2) tell me and I'll adjust.\n")

