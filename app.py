import streamlit as st
import pandas as pd
import re

# -----------------------------
# Helpers for Excel-like engine
# -----------------------------

def col_to_index(col):
    """Convert Excel column letters (A,B,AA) → 0-based index."""
    col = col.upper()
    num = 0
    for c in col:
        num = num * 26 + (ord(c) - 64)
    return num - 1


def parse_cell_reference(ref):
    """Convert 'A1' → (row_index, col_index)."""
    match = re.match(r"([A-Za-z]+)(\d+)$", ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    col = match.group(1)
    row = int(match.group(2))
    return row - 1, col_to_index(col)


def get_cell_value(df, ref):
    """Return the numeric value of a cell."""
    r, c = parse_cell_reference(ref)
    val = df.iat[r, c]
    try:
        return float(val)
    except:
        return 0.0


def get_range_values(df, range_ref):
    """Return list of values from a range like A1:A10 or A1:C5."""
    if ":" not in range_ref:
        return [get_cell_value(df, range_ref)]

    start, end = range_ref.split(":")
    r1, c1 = parse_cell_reference(start.strip())
    r2, c2 = parse_cell_reference(end.strip())

    values = []
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(c1, c2), max(c1, c2) + 1):
            try:
                values.append(float(df.iat[r, c]))
            except:
                values.append(0)
    return values


# -----------------------------
# Excel-like functions
# -----------------------------

def __SUM__(values):
    return sum(values)


def __SUMIF__(range_vals, criteria, sum_vals):
    """Excel-style SUMIF implementation."""
    crit = criteria.strip()

    if crit.startswith('"') and crit.endswith('"'):
        crit = crit[1:-1]

    filtered_sum = 0

    for cond, val in zip(range_vals, sum_vals):

        ok = False

        # Numeric comparisons
        if re.match(r"^[<>]=?\d+(\.\d+)?$", crit):
            ok = eval(f"{cond}{crit}")

        # Equality comparison like =5
        elif re.match(r"^=\d+(\.\d+)?$", crit):
            ok = cond == float(crit[1:])

        # String match
        elif crit == str(cond):
            ok = True

        if ok:
            filtered_sum += val

    return filtered_sum


# -----------------------------
# Formula parser
# -----------------------------

def evaluate_formula(df, formula):

    expr = formula

    # ---- SUM(range) ----
    sum_pattern = r"SUM\s*\(\s*([A-Za-z0-9:]+)\s*\)"

    def repl_sum(m):
        range_ref = m.group(1)
        vals = get_range_values(df, range_ref)
        return str(__SUM__(vals))

    expr = re.sub(sum_pattern, repl_sum, expr)

    # ---- SUMIF(range, criteria, sum_range) ----
    sumif_pattern = r"SUMIF\s*\(\s*([^,]+)\s*,\s*([^,]+)\s*,\s*([^)]+)\s*\)"

    def repl_sumif(m):
        range_ref = m.group(1).strip()
        criteria = m.group(2).strip()
        sum_ref = m.group(3).strip()

        range_vals = get_range_values(df, range_ref)
        sum_vals = get_range_values(df, sum_ref)

        return str(__SUMIF__(range_vals, criteria, sum_vals))

    expr = re.sub(sumif_pattern, repl_sumif, expr)

    # ---- Cell references (A1, B12, etc.) ----
    cellref_pattern = r"\b([A-Za-z]{1,3}\d{1,4})\b"

    def repl_cellref(m):
        ref = m.group(1)
        return str(get_cell_value(df, ref))

    expr = re.sub(cellref_pattern, repl_cellref, expr)

    # ---- Final safe eval (numbers + operators only) ----
    allowed = re.fullmatch(r"[0-9\.\+\-\*\/\(\)\s]+", expr)
    if not allowed:
        raise ValueError("Invalid characters in expression after parsing.")

    return eval(expr)


# -----------------------------
# Streamlit interface
# -----------------------------

st.title("Excel Formula Calculator (Python Engine)")

uploaded = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.write("### Loaded Data")
    st.dataframe(df)

    formula = st.text_input("Enter Excel formula, e.g. `SUM(A1:A10)` or `SUMIF(A1:A10, \">5\", B1:B10)`")

    if st.button("Calculate"):
        try:
            result = evaluate_formula(df, formula)
            st.success(f"Result: {result}")
        except Exception as e:
            st.error(str(e))
else:
    st.info("Upload an Excel file to begin.")
