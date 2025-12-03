import streamlit as st
import math, re
from copy import deepcopy

# Embedded workbook data (auto-generated)

VALUES = {'A1': 'Target Leadership', 'B1': None, 'C1': None, 'D1': 'HEALTH RATIO (MAX/MIN)', 'E1': None, 'F1': None, 'G1': None, 'H1': None, 'I1': None, 'J1': None, 'K1': None, 'L1': None, 'M1': None, 'A2': 'Target Leadership:', 'B2': 46800, 'C2': None, 'E2': None, 'F2': None, 'G2': None, 'H2': None, 'I2': None, 'J2': None, 'K2': None, 'L2': None, 'M2': None, 'A3': 'Calculated Leadership:', 'D3': None, 'E3': None, 'F3': None, 'G3': None, 'H3': None, 'I3': None, 'J3': None, 'K3': None, 'L3': None, 'M3': None, 'A4': 'Ratio', 'C4': None, 'D4': None, 'E4': None, 'F4': None, 'G4': None, 'H4': None, 'I4': None, 'J4': None, 'K4': None, 'L4': None, 'M4': None, 'A5': 'Target Max Health', 'C5': None, 'D5': None, 'E5': None, 'F5': None, 'G5': None, 'H5': None, 'I5': None, 'J5': None, 'K5': None, 'L5': None, 'M5': None, 'A6': None, 'B6': None, 'C6': None, 'D6': None, 'E6': None, 'F6': None, 'G6': None, 'H6': None, 'I6': None, 'J6': None, 'K6': None, 'L6': None, 'M6': None, 'A7': 'Normal Troops', 'B7': 'Army', 'C7': 'Leadership', 'D7': 'Rank', 'E7': 'Leadership per Active Troop', 'F7': 'Troop Base Health', 'G7': 'INPUT Troop Health Bonus', 'H7': 'Total Health Bonus', 'I7': 'Target Health per Stack', 'J7': 'Actual Health per Stack', 'K7': 'Troop Level', 'L7': None, 'M7': None, 'A8': 'G6 Flyings', 'E8': 20, 'F8': 57000, 'G8': 0, 'K8': 6, 'L8': None, 'M8': None, 'A9': 'G5 Flyings', 'E9': 20, 'F9': 30000, 'G9': 0, 'K9': 5, 'L9': None, 'M9': None, 'A10': 'G6 Mounted', 'E10': 2, 'F10': 5700, 'G10': 0, 'K10': 6, 'L10': None, 'M10': None, 'A11': 'G5 Mounted', 'E11': 2, 'F11': 3150, 'G11': 0, 'K11': 5, 'L11': None, 'M11': None, 'A12': 'G6 Ranged', 'E12': 1, 'F12': 2820, 'G12': 0, 'K12': 6, 'L12': None, 'M12': None, 'A13': 'G6 Melee', 'E13': 1, 'F13': 2820, 'G13': 0, 'K13': 6, 'L13': None, 'M13': None, 'A14': 'G5 Ranged', 'E14': 1, 'F14': 1560, 'G14': 0, 'K14': 5, 'L14': None, 'M14': None, 'A15': 'G5 Melee', 'E15': 1, 'F15': 1560, 'G15': 0, 'K15': 5, 'L15': None, 'M15': None, 'A16': None, 'D16': None, 'E16': None, 'F16': None, 'G16': None, 'H16': None, 'I16': None, 'J16': None, 'K16': None, 'L16': None, 'M16': None, 'A17': None, 'B17': None, 'C17': None, 'D17': None, 'E17': None, 'F17': None, 'G17': None, 'H17': None, 'I17': None, 'J17': None, 'K17': None, 'L17': None, 'M17': None, 'A18': None, 'L18': None, 'M18': None, 'A19': 'M6 Jungle Melee', 'D19': None, 'E19': 34, 'F19': 390000, 'G19': 0, 'K19': 6, 'L19': None, 'M19': None, 'A20': 'M5 Ettin Melee', 'D20': None, 'E20': 23, 'F20': 144000, 'G20': 0, 'K20': 5, 'L20': None, 'M20': None, 'A21': 'M4 Phoenix Flyings', 'D21': None, 'E21': 15, 'F21': 51000, 'G21': 0, 'K21': 4, 'L21': None, 'M21': None, 'A22': None, 'B22': None, 'C22': None, 'D22': None, 'E22': None, 'F22': None, 'G22': None, 'H22': None, 'I22': None, 'J22': None, 'K22': None, 'L22': None, 'M22': None, 'A23': 'Mercs', 'C23': 'Authority', 'L23': None, 'M23': None, 'A24': 'Epic Monster Hunter (Green)', 'C24': None, 'D24': None, 'E24': 1, 'F24': 75000, 'G24': 0, 'K24': None, 'L24': None, 'M24': None, 'A25': 'Epic Monster Hunter (Yellow)', 'C25': None, 'D25': None, 'E25': 1, 'F25': 11250, 'G25': 0, 'K25': None, 'L25': None, 'M25': None}

FORMULAS = {'D2': '=+MAX(J8:J15)/MIN(J8:J15)', 'B3': '=+C16', 'C3': '=+B2-B3', 'B4': '=+IF(G15<=G14,16.1%,16.1%/(G15/AVERAGE(G8:G14)))', 'B5': '=+B4*B2*H15/E15', 'B8': '=+ROUNDDOWN(I8/H8,-1)', 'C8': '=+B8*E8', 'D8': '=+RANK(J8,$J$8:$J$15)', 'H8': '=+F8*(1+G8)', 'I8': '=+J10*(1-0.01%)', 'J8': '=+B8*H8', 'B9': '=+ROUNDDOWN(I9/H9,-1)', 'C9': '=+B9*E9', 'D9': '=+RANK(J9,$J$8:$J$15)', 'H9': '=+F9*(1+G9)', 'I9': '=+J11*(1-0.01%)', 'J9': '=+B9*H9', 'B10': '=+ROUNDDOWN(I10/H10,-1)', 'C10': '=+B10*E10', 'D10': '=+RANK(J10,$J$8:$J$15)', 'H10': '=+F10*(1+G10)', 'I10': '=+J13*(1-0.01%)', 'J10': '=+B10*H10', 'B11': '=+ROUNDDOWN(I11/H11,-1)', 'C11': '=+B11*E11', 'D11': '=+RANK(J11,$J$8:$J$15)', 'H11': '=+F11*(1+G11)', 'I11': '=+J15*(1-0.01%)', 'J11': '=+B11*H11', 'B12': '=+ROUNDDOWN(I12/H12,-1)', 'C12': '=+B12*E12', 'D12': '=+RANK(J12,$J$8:$J$15)', 'H12': '=+F12*(1+G12)', 'I12': '=+J8*(1-0.01%)', 'J12': '=+B12*H12', 'B13': '=+ROUNDDOWN(I13/H13,-1)', 'C13': '=+B13*E13', 'D13': '=+RANK(J13,$J$8:$J$15)', 'H13': '=+F13*(1+G13)', 'I13': '=+J14*(1-0.01%)', 'J13': '=+B13*H13', 'B14': '=+ROUNDDOWN(I14/H14,-1)', 'C14': '=+B14*E14', 'D14': '=+RANK(J14,$J$8:$J$15)', 'H14': '=+F14*(1+G14)', 'I14': '=+J9*(1-0.01%)', 'J14': '=+B14*H14', 'B15': '=+ROUNDDOWN(I15/H15,-1)', 'C15': '=+B15*E15', 'D15': '=+RANK(J15,$J$8:$J$15)', 'H15': '=+F15*(1+G15)', 'I15': '=+B5', 'J15': '=+B15*H15', 'B16': '=SUM(B8:B15)', 'C16': '=SUM(C8:C15)', 'B18': '=+B7', 'C18': '=+C7', 'D18': '=+D7', 'E18': '=+E7', 'F18': '=+F7', 'G18': '=+G7', 'H18': '=+H7', 'I18': '=+I7', 'J18': '=+J7', 'K18': '=+K7', 'B19': '=+ROUNDDOWN(I19/H19,-1)', 'C19': '=+B19*F19', 'H19': '=+F19*(1+G19)', 'I19': '=+J13*(1-0.01%)', 'J19': '=+B19*H19', 'B20': '=+ROUNDDOWN(I20/H20,-1)', 'C20': '=+B20*F20', 'H20': '=+F20*(1+G20)', 'I20': '=+J19*(1-0.01%)', 'J20': '=+B20*H20', 'B21': '=+ROUNDDOWN(I21/H21,-1)', 'C21': '=+B21*F21', 'H21': '=+F21*(1+G21)', 'I21': '=+J20*(1-0.01%)', 'J21': '=+B21*H21', 'B23': '=+B7', 'D23': '=+D7', 'E23': '=+E7', 'F23': '=+F7', 'G23': '=+G7', 'H23': '=+H7', 'I23': '=+I7', 'J23': '=+J7', 'K23': '=+K7', 'B24': '=+ROUNDDOWN(I24/H24,-1)', 'H24': '=+F24*(1+G24)', 'I24': '=+J19*(1-0.01%)', 'J24': '=+B24*H24', 'B25': '=+ROUNDDOWN(I25/H25,-1)', 'H25': '=+F25*(1+G25)', 'I25': '=+J24*(1-0.01%)', 'J25': '=+B25*H25'}

LABELS = {'A1': 'Target Leadership', 'A2': 'Target Leadership:', 'A3': 'Calculated Leadership:', 'A4': 'Ratio', 'A5': 'Target Max Health', 'A6': None, 'A7': 'Normal Troops', 'A8': 'G6 Flyings', 'A9': 'G5 Flyings', 'A10': 'G6 Mounted', 'A11': 'G5 Mounted', 'A12': 'G6 Ranged', 'A13': 'G6 Melee', 'A14': 'G5 Ranged', 'A15': 'G5 Melee', 'A16': None, 'A17': None, 'A18': None, 'A19': 'M6 Jungle Melee', 'A20': 'M5 Ettin Melee', 'A21': 'M4 Phoenix Flyings', 'A22': None, 'A23': 'Mercs', 'A24': 'Epic Monster Hunter (Green)', 'A25': 'Epic Monster Hunter (Yellow)'}

INPUT_CELLS = ['B2', 'G8', 'G9', 'G10', 'G11', 'G12', 'G13', 'G14', 'G15', 'G19', 'G20', 'G21', 'G22', 'G23', 'G24', 'G25', 'G26']
OUTPUT_RANGES = [('A8', 'C16'), ('A18', 'C21'), ('A23', 'C24')]

st.set_page_config(page_title='TB Calculator (Standalone)', layout='wide')
st.title('TB Calculator — Inputs & Results (no Excel required)')


def range_to_coords_simple(range_text):
    rt = range_text.replace("$","")
    if ":" in rt:
        left, right = rt.split(":")
        colL = ''.join([c for c in left if c.isalpha()])
        rowL = int(''.join([c for c in left if c.isdigit()]))
        colR = ''.join([c for c in right if c.isalpha()])
        rowR = int(''.join([c for c in right if c.isdigit()]))
        coords = [f"{colL}{r}" for r in range(rowL, rowR+1)]
        return coords
    else:
        return [rt.replace("$","")]

def py_SUM(range_text, vm):
    coords = range_to_coords_simple(range_text)
    s = 0.0
    for c in coords:
        v = vm.get(c, None)
        try: s += float(v)
        except: pass
    return s

def py_SUMIF(range_text, criteria, sum_range_text, vm):
    coords = range_to_coords_simple(range_text)
    sum_coords = range_to_coords_simple(sum_range_text)
    total = 0.0
    crit = criteria
    if isinstance(crit, str):
        crit = crit.strip().strip('"').strip("'")
    def pred(val):
        try:
            if isinstance(crit, str) and crit.startswith('>='): return float(val) >= float(crit[2:])
            if isinstance(crit, str) and crit.startswith('<='): return float(val) <= float(crit[2:])
            if isinstance(crit, str) and crit.startswith('>'): return float(val) > float(crit[1:])
            if isinstance(crit, str) and crit.startswith('<'): return float(val) < float(crit[1:])
            if isinstance(crit, str) and crit.startswith('='): return str(val) == crit[1:]
            try: return float(val) == float(crit)
            except: return str(val) == str(crit)
        except: return False
    for i, rc in enumerate(coords):
        val = vm.get(rc, None)
        try:
            if pred(val):
                sc = sum_coords[i] if i < len(sum_coords) else rc
                s_val = vm.get(sc, None)
                if s_val is None: continue
                total += float(s_val)
        except: pass
    return total

def py_AVERAGE(range_text, vm):
    coords = range_to_coords_simple(range_text)
    vals = []
    for c in coords:
        try: vals.append(float(vm.get(c)))
        except: pass
    return sum(vals)/len(vals) if vals else None

def py_MAX(range_text, vm):
    coords = range_to_coords_simple(range_text)
    vals = []
    for c in coords:
        try: vals.append(float(vm.get(c)))
        except: pass
    return max(vals) if vals else None

def py_MIN(range_text, vm):
    coords = range_to_coords_simple(range_text)
    vals = []
    for c in coords:
        try: vals.append(float(vm.get(c)))
        except: pass
    return min(vals) if vals else None

def py_ROUNDDOWN(x, n):
    try: x = float(x)
    except: return None
    if n >= 0:
        factor = 10 ** n
        return math.floor(x * factor) / factor
    else:
        factor = 10 ** (-n)
        return math.floor(x / factor) * factor

def py_RANK(value, ref_range, vm):
    coords = range_to_coords_simple(ref_range)
    vals = []
    for c in coords:
        try: vals.append(float(vm.get(c)))
        except: pass
    if not vals: return None
    sorted_vals = sorted(vals, reverse=True)
    try: v = float(value)
    except: return None
    return sorted_vals.index(v) + 1 if v in sorted_vals else None

def evaluate_all(vm, fm, max_passes=30):
    values = deepcopy(vm)
    for _ in range(max_passes):
        changed = False
        for cell, formula in fm.items():
            f = formula.lstrip('=+').strip()
            try:
                expr = f
                expr = re.sub(r"SUMIF\(([^,]+),([^,]+),([^)]+)\)", lambda m: f"__SUMIF__('{m.group(1).strip()}','{m.group(2).strip()}','{m.group(3).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"SUM\(([^)]+)\)", lambda m: f"__SUM__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"AVERAGE\(([^)]+)\)", lambda m: f"__AVERAGE__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"MAX\(([^)]+)\)", lambda m: f"__MAX__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"MIN\(([^)]+)\)", lambda m: f"__MIN__('{m.group(1).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"IF\(([^,]+),([^,]+),([^)]+)\)", lambda m: f"__IF__({m.group(1).strip()},{m.group(2).strip()},{m.group(3).strip()})", expr, flags=re.IGNORECASE)
                expr = re.sub(r"ROUNDDOWN\(([^,]+),([^)]+)\)", lambda m: f"__ROUNDDOWN__({m.group(1).strip()},{m.group(2).strip()})", expr, flags=re.IGNORECASE)
                expr = re.sub(r"RANK\(([^,]+),([^)]+)\)", lambda m: f"__RANK__({m.group(1).strip()},'{m.group(2).strip()}')", expr, flags=re.IGNORECASE)
                expr = re.sub(r"\b([A-Za-z]{1,3}\d{1,4})\b", lambda m: f"__VAL__('{m.group(1)}')", expr)
                expr = re.sub(r"([0-9]*\.?[0-9]+)%", lambda m: f"({m.group(1)}/100)", expr)
                def __VAL__(ref): return values.get(ref, None)
                def __SUM__(r): return py_SUM(r, values)
                def __SUMIF__(r, crit, s): return py_SUMIF(r, crit, s, values)
                def __AVERAGE__(r): return py_AVERAGE(r, values)
                def __MAX__(r): return py_MAX(r, values)
                def __MIN__(r): return py_MIN(r, values)
                def __IF__(cond,a,b):
                    try:
                        cond_eval = eval(str(cond), {'__VAL__': __VAL__})
                        cond_bool = bool(cond_eval)
                    except:
                        cond_bool = False
                    try: a_eval = eval(str(a), {'__VAL__': __VAL__})
                    except: a_eval = a
                    try: b_eval = eval(str(b), {'__VAL__': __VAL__})
                    except: b_eval = b
                    return a_eval if cond_bool else b_eval
                def __ROUNDDOWN__(x_expr,n_expr):
                    try: xv = float(eval(str(x_expr), {'__VAL__': __VAL__}))
                    except: xv = x_expr
                    try: nv = int(eval(str(n_expr), {'__VAL__': __VAL__}))
                    except: nv = int(n_expr)
                    return py_ROUNDDOWN(xv, nv)
                def __RANK__(valexpr, range_str): return py_RANK(valexpr, range_str, values)
                safe_env = {'__VAL__': __VAL__, '__SUM__': __SUM__, '__SUMIF__': __SUMIF__, '__AVERAGE__': __AVERAGE__, '__MAX__': __MAX__, '__MIN__': __MIN__, '__IF__': __IF__, '__ROUNDDOWN__': __ROUNDDOWN__, '__RANK__': __RANK__, 'math': math}
                result = eval(expr, { }, safe_env)
                if isinstance(result, float) and result.is_integer(): result = int(result)
                values[cell] = result
                changed = True
            except Exception:
                pass
        if not changed:
            break
    return values

st.header('Inputs (enter percentages — 100 = 100%) — labels from Column A')
inputs = {}
for coord in INPUT_CELLS:
    row = int(re.findall(r'\d+', coord)[0])
    label = LABELS.get(f'A{row}', coord)
    default = VALUES.get(coord, 0) or 0
    try: default = float(default)
    except: default = 0.0
    v = st.number_input(f"{label} ({coord}) - %", value=default, format="%.6g")
    inputs[coord] = v / 100.0

if st.button('Calculate'):
    vm = deepcopy(VALUES)
    for k,v in inputs.items():
        vm[k] = v
    computed = evaluate_all(vm, FORMULAS, max_passes=40)
    st.header('Results (final output blocks)')
    for start,end in OUTPUT_RANGES:
        srow = int(start[1:]); erow = int(end[1:])
        cols = ['A','B','C']
        rows = []
        for r in range(srow, erow+1):
            rowvals = []
            for c in cols:
                coord = f"{c}{r}"
                rowvals.append(computed.get(coord, ''))
            rows.append(rowvals)
        st.table(rows)
