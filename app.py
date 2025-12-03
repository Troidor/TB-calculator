import streamlit as st
import pandas as pd
import re

# ---------------------------------------------------------
# LOAD EXCEL FILE
# ---------------------------------------------------------
EXCEL_FILE = "TB.xlsx"     # <---- rename your file to TB.xlsx
SHEET_NAME = None          # load all sheets

df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

sheet = list(df.values())[0]       # your main sheet as df
cells = {}                         # store values by Excel name


# ---------------------------------------------------------
# BUILD CELL REFERENCE MAP
# ---------------------------------------------------------
cols = sheet.columns.tolist()

for r in range(len(sheet)):
    for c in range(len(cols)):
        col_letter = chr(ord('A') + c)
        cell_name = f"{col_letter}{r+1}"
        cells[cell_name] = sheet.iloc[r, c]


# ---------------------------------------------------------
# SUMIF IMPLEMENTATION
# ---------------------------------------------------------
def __SUMIF__(range_ref, criteria, sum_ref):
    # convert A1:A10 to list of values
    def get_range(r):
        if ":" in r:
            start, end = r.split(":")
            start = start.strip()
            end = end.strip()

            col1 = ord(start[0]) - 65
            row1 = int(start[1:]) - 1
            col2 = ord(end[0]) - 65
            row2 = int(end[1:]) - 1

            values = []
            for rr in range(row1, row2 + 1):
                values.append(sheet.iloc[rr, col1])
            return values
        else:
            col = ord(r[0]) - 65
            row = int(r[1:]) - 1
            return [sheet.iloc[row, col]]

    # Load ranges
    range_vals = get_range(range_ref)
    sum_vals = get_range(sum_ref)

    # Parse criteria
    crit = criteria.replace('"', "").replace("'", "")

    total = 0
    for i, val in enumerate(range_vals):
        ok = False
        if crit.startswith(">="):
            ok = val >= float(crit[2:])
        elif crit.startswith("<="):
            ok = val <= float(crit[2:])
        elif crit.startswith(">"):
            ok = val > float(crit[1:])
        elif crit.startswith("<"):
            ok = val < float(crit[1:])
        elif crit.startswith("="):
            ok = str(val) == crit[1:]
        else:
            ok = str(val) == crit

        if ok:
            total += float(sum_vals[i])

    return total


# ---------------------------------------------------------
# EXCEL FORMULA PARSER
# ---------------------------------------------------------
def eval_formula(expr):

    # SUMIF
    sumif_pattern = r"SUMIF\s*\(\s*([^,]+)\s*,\s*([^,]_
