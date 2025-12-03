# app.py
import streamlit as st
import math, re
from copy import deepcopy
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter

# ---------- CONFIG ----------
EXCEL_FILENAME = "TB.xlsx"     # <-- put your file in repo root and name it TB.xlsx
SHEET_NAME = "Φύλλο1"          # sheet name in your file (Greek). Change if different.

# Inputs specification
INPUT_CELLS = ["B2"] + \
    [f"G{r}" for r in range(8, 16)] + \
    [f"G{r}" for r in range(19, 22)] + \
    [f"G{r}" for r in range(24, 26)]

# Output ranges to display
OUTPUT_RANGES = [("A8", "C16"), ("A18", "C21"), ("A23", "C24")]

st.set_page_config(page_title="TB Calculator", layout="wide")


# ---------- Utility functions (same evaluator logic as discussed) ----------
def load_sheet_models(path=EXCEL_FILENAME, sheet=SHEET_NAME):
    wb = load_workbook(path, data_only=False)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.active

    values = {}
    formulas = {}
    labels = {}

    for row in ws.iter_rows():
        for cell in row:
            coord = cell.coordinate
            val = cell.value
            if isinstance(val, str) and val.startswith("="):
                formulas[coord] = val
            else:
                values[coord] = val
            if cell.column_letter == "A":
                labels[coord] = val
    return values, formulas, labels

def range_to_coords(range_text):
    rt = range_text.replace("$", "")
    if ":" in rt:
        min_col, min_row, max_col, max_row = range_boundaries(rt)
        coords = []
        for r in range(min_row, max_row+1):
            for c in range(min_col, max_col+1):
                coords.append(f"{get_column_letter(c)}{r}")
        return coords
    else:
        return [rt]

def py_SUM(range_text, values_map):
    coords = range_to_coords(range_text)
    s = 0.0
    for c in coords:
        v = values_map.get(c, None)
        try:
            if v is None:
                continue
            s += float(v)
        except:
            pass
    return s

def py_SUMIF(range_text, criteria, sum_range_text, values_map):
    coords = range_to_coords(range_text)
    sum_coords = range_to_coords(sum_range_text)
    total = 0.0
    def make_pred(c):
        if isinstance(c, str) and re.match(r'^[A-Za-z]+\d+$', c):
            return lambda val: val == values_map.get(c)
        if isinstance(c, str):
            c = c.strip()
            m = re.match(r'^(>=|<=|>|<|=)(.*)$', c)
            if m:
                op = m.group(1); rhs = m.group(2).strip()
                try:
                    rhs_num = float(rhs)
                    if op == ">": return lambda x: (x is not None) and float(x) > rhs_num
                    if op == "<": return lambda x: (x is not None) and float(x) < rhs_num
                    if op == ">=": return lambda x: (x is not None) and float(x) >= rhs_num
                    if op == "<=": return lambda x: (x is not None) and float(x) <= rhs_num
                    if op == "=": return lambda x: (x is not None) and float(x) == rhs_num
                except:
                    if op == "=": return lambda x: str(x) == rhs
            try:
                rr = float(c); return lambda x: (x is not None) and float(x) == rr
            except:
                return lambda x: str(x) == c
        else:
            return lambda x: (x is not None) and float(x) == float(c)
    pred = make_pred(criteria)
    for i, rc in enumerate(coords):
        val = values_map.get(rc, None)
        try:
            if pred(val):
                sc = sum_coords[i] if i < len(sum_coords) else rc
                s_val = values_map.get(sc, None)
                if s_val is None: continue
                total += float(s_val)
        except:
            pass
    return total

def py_AVERAGE(range_text, values_map):
    coords = range_to_coords(range_text)
    s = 0.0; n = 0
    for c in coords:
        v = values_map.get(c, None)
        if v is None: continue
        try:
            s += float(v); n += 1
        except: pass
    return s / n if n>0 else None

def py_MAX(range_text, values_map):
    coords = range_to_coords(range_text)
    vals=[]
    for c in coords:
        v = values_map.get(c, None)
        try: vals.append(float(v))
        except: pass
    return max(vals) if vals else None

def py_MIN(range_text, values_map):
    coords = range_to_coords(range_text)
    vals=[]
    for c in coords:
        v = values_map.get(c, None)
        try: vals.append(float(v))
        except: pass
    return min(vals) if vals else None

def py_IF(cond_bool, a, b):
    return a if cond_bool else b

def py_ROUNDDOWN(x, n):
    try: x = float(x)
    except: return None
    if n >= 0:
        factor = 10**n; return math.floor(x*factor)/factor
    else:
        factor = 10**(-n); return math.floor(x/factor)*factor

def py_RANK(value, ref_range, values_map):
    coords = range_to_coords(ref_range)
    vals=[]
    for c in coords:
        v = values_map.get(c, None)
        try: vals.append(float(v))
        except: pass
    if not vals: return None
    sorted_vals = sorted(vals, reverse=True)
    try: v = float(value)
    except: return None
    return sorted_vals.index(v)+1 if v in sorted_vals else None

def evaluate_all(values_map, formulas_map, max_passes=15):
    values = deepcopy(values_map)
    for _ in range(max_passes):
        changed = False
        for cell, formula in formulas_map.items():
            f = formula.lstrip("=+").strip()
            try:
                expr = f
                expr = re.sub(r"SUMIF\(([^,]+),([^,]+),([^)]+)\)", 
                              lambda m: f"__SUMIF__('{m.group(1).strip()}', {m.group(2).strip()}, '{m.group(3).strip()}')",
                              expr, flags=re.IGNORECASE)
                expr = re.sub(r"AVERAGE\(([^)]+)\)", lambda m: f"__AVERAGE__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"MAX\(([^)]+)\)", lambda m: f"__MAX__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"MIN\(([^)]+)\)", lambda m: f"__MIN__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"IF\(([^,]+),([^,]+),([^)]+)\)", lambda m: f"__IF__({m.group(1).strip()},{m.group(2).strip()},{m.group(3).strip()})", expr, flags=re.IGNORECASE)
                expr = re.sub(r"ROUNDDOWN\(([^,]+),([^)]+)\)", lambda m: f"__ROUNDDOWN__({m.group(1).strip()},{m.group(2).strip()})", expr, flags=re.IGNORECASE)
                expr = re.sub(r"RANK\(([^,]+),([^)]+)\)", lambda m: f"__RANK__({m.group(1).strip()},'{m.group(2).strip()}')", expr, flags=re.IGNORECASE)

                def repl_cellref(m):
                    ref = m.group(0)
                    return f\"__VAL__('{ref}')\"
                expr = re.sub(r\"\\b([A-Za-z]{1,3}\\d{1,4})\\b\", repl_cellref, expr)
                expr = re.sub(r\"([0-9]*\\.?[0-9]+)\\%\", lambda m: f\"({m.group(1)}/100)\", expr)

                def __VAL__(ref):
                    return values.get(ref, None)
                def __SUM__(range_str):
                    return py_SUM(range_str, values)
                def __SUMIF__(range_str, crit_raw, sum_range_str):
                    try:
                        crit_eval = eval(str(crit_raw), {\"__VAL__\": __VAL__})
                    except:
                        crit_eval = crit_raw
                    return py_SUMIF(range_str, crit_eval, sum_range_str, values)
                def __AVERAGE__(range_str): return py_AVERAGE(range_str, values)
                def __MAX__(range_str): return py_MAX(range_str, values)
                def __MIN__(range_str): return py_MIN(range_str, values)
                def __IF__(condexpr, aexpr, bexpr):
                    try:
                        cond_val = eval(str(condexpr), {\"__VAL__\": __VAL__})
                        cond_bool = bool(cond_val)
                    except:
                        cond_bool = False
                    try: aval = eval(str(aexpr), {\"__VAL__\": __VAL__})
                    except: aval = aexpr
                    try: bval = eval(str(bexpr), {\"__VAL__\": __VAL__})
                    except: bval = bexpr
                    return py_IF(cond_bool, aval, bval)
                def __ROUNDDOWN__(xexpr, nexpr):
                    try: x = float(eval(str(xexpr), {\"__VAL__\": __VAL__}))
                    except: x = xexpr
                    try: n = int(eval(str(nexpr), {\"__VAL__\": __VAL__}))
                    except: n = int(nexpr)
                    return py_ROUNDDOWN(x, n)
                def __RANK__(valexpr, range_str):
                    try: v = float(eval(str(valexpr), {\"__VAL__\": __VAL__}))
                    except: v = valexpr
                    return py_RANK(v, range_str, values)

                safe_env = {
                    \"__VAL__\": __VAL__, \"__SUM__\": __SUM__, \"__SUMIF__\": __SUMIF__,
                    \"__AVERAGE__\": __AVERAGE__, \"__MAX__\": __MAX__, \"__MIN__\": __MIN__,
                    \"__IF__\": __IF__, \"__ROUNDDOWN__\": __ROUNDDOWN__, \"__RANK__\": __RANK__,
                    \"math\": math
                }

                result = eval(expr, {}, safe_env)

                if isinstance(result, bool):
                    values[cell] = result
                else:
                    try:
                        if result is None:
                            values[cell] = None
                        elif isinstance(result, float) and result.is_integer():
                            values[cell] = int(result)
                        else:
                            values[cell] = result
                    except:
                        values[cell] = result
                changed = True
            except Exception:
                pass
        if not changed:
            break
    return values

# ---------- Streamlit UI ----------
st.title("TB Calculator")

# Load workbook values & formulas
try:
    values_map, formulas_map, labels = load_sheet_models()
except FileNotFoundError:
    st.error(f"Excel file '{EXCEL_FILENAME}' not found. Upload it to repo root and name it {EXCEL_FILENAME}.")
    st.stop()

st.write("Sheet:", SHEET_NAME)

# display and edit inputs inside an expander form
with st.form("inputs_form"):
    st.markdown("### Inputs — enter values for the highlighted Column G cells")
    # B2 single input (label use A2)
    label_B2 = labels.get("A2", "B2")
    val_B2 = values_map.get("B2", None)
    b2 = st.number_input(f"{label_B2} (B2)", value=float(val_B2) if val_B2 is not None else 0.0, format="%.6g", step=1.0)

    # groups of G cells: show Column A label and allow number input
    def column_inputs(start_row, end_row):
        inputs = {}
        for r in range(start_row, end_row+1):
            coord = f"G{r}"
            label = labels.get(f"A{r}", f"A{r}")
            default = values_map.get(coord, None)
            if default is None:
                default = 0.0
            else:
                try:
                    default = float(default)
                except:
                    default = 0.0
            inputs[coord] = st.number_input(f"{label} (G{r})", value=default, format="%.6g", step=1.0)
        return inputs

    g1 = column_inputs(8, 15)
    g2 = column_inputs(19, 21)
    g3 = column_inputs(24, 25)

    submitted = st.form_submit_button("Calculate")

if submitted:
    # override input values_map
    values_map["B2"] = b2
    for k,v in {**g1, **g2, **g3}.items():
        values_map[k] = v

    # evaluate formulas
    computed = evaluate_all(values_map, formulas_map, max_passes=25)

    # Render the three output blocks
    st.markdown("## Results")
    for start, end in OUTPUT_RANGES:
        min_col, min_row, max_col, max_row = range_boundaries(f"{start}:{end}")
        # build simple table list
        rows = []
        headers = []
        for r in range(min_row, max_row+1):
            row_list = []
            for c in range(min_col, max_col+1):
                coord = f\"{get_column_letter(c)}{r}\"
                row_list.append(computed.get(coord, "" ))
            rows.append(row_list)
        st.table(rows)
    st.success("Calculation complete.")
