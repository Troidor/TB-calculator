import os
from flask import Flask, render_template, request, redirect, url_for, session, flash
from openpyxl import load_workbook
from xlcalculator import ModelCompiler, Evaluator
from xlcalculator import reader as xlreader

# --- Config ---
APP_PASSWORD = os.environ.get("APP_PASSWORD", None)
SECRET_KEY = os.environ.get("SECRET_KEY", "change-this-secret")  # set in Render
EXCEL_PATHS = [
    "excel/TB_v18.xlsx",
    "excel/workbook.xlsx",
]

def find_excel_path():
    for p in EXCEL_PATHS:
        if os.path.exists(p):
            return p
    raise FileNotFoundError("Excel file not found. Put your file as excel/TB_v18.xlsx (or excel/workbook.xlsx)")

def is_yellow_fill(cell):
    """Return True if an openpyxl cell has a yellow-ish solid fill."""
    try:
        fill = cell.fill
        if fill is None:
            return False
        # Look for solid pattern
        pt = getattr(fill, 'patternType', None)
        if pt is None or pt.lower() != 'solid':
            return False
        fg = getattr(fill, 'fgColor', None)
        if fg is None:
            return False
        rgb = getattr(fg, 'rgb', None)
        if rgb:
            # rgb often comes as ARGB (8 chars) or RGB (6 chars). take last 6 chars.
            rgb = rgb.upper()
            hex6 = rgb[-6:]
            try:
                r = int(hex6[0:2], 16)
                g = int(hex6[2:4], 16)
                b = int(hex6[4:6], 16)
                # Yellow-ish: high red & green, low blue
                if r >= 200 and g >= 200 and b <= 120:
                    return True
            except Exception:
                pass
        # Fallback: check indexed or theme (not robust)
        # If no rgb available, return False
        return False
    except Exception:
        return False

def scan_calcs_sheet_for_yellow_inputs(ws):
    """Detect inputs in CALCS sheet as yellow-shaded cells in column C.
    Outputs: formula cells anywhere in used range on Calcs sheet.
    Returns inputs, outputs lists."""
    inputs = []
    outputs = []
    min_row, max_row = ws.min_row, ws.max_row
    # Column C index = 3
    col_idx = 3
    for r in range(min_row, max_row + 1):
        cell = ws.cell(row=r, column=col_idx)
        val = cell.value
        coord = cell.coordinate
        # detect yellow fill
        if is_yellow_fill(cell):
            # guess label from column B or above
            label = None
            try:
                left = ws.cell(row=r, column=2).value
                if left and isinstance(left, str):
                    label = left.strip()
                if not label:
                    up = ws.cell(row=r-1, column=col_idx).value if r-1>=min_row else None
                    if up and isinstance(up, str):
                        label = up.strip()
            except Exception:
                label = None
            inputs.append({"coord": coord, "value": val, "label": label})
    # outputs: any formula cell in used range
    for r in range(min_row, max_row + 1):
        for c in range(ws.min_column, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            is_formula = isinstance(val, str) and val.startswith('=')
            if is_formula:
                coord = cell.coordinate
                label = None
                try:
                    left = ws.cell(row=r, column=c-1).value if c-1>=ws.min_column else None
                    if left and isinstance(left, str):
                        label = left.strip()
                except Exception:
                    label = None
                outputs.append({"coord": coord, "formula": val, "label": label})
    return inputs, outputs

def build_model_from_file(path):
    wb_reader = xlreader.Reader()
    wb_reader.load_workbook(path)
    model = wb_reader.read()
    compiler = ModelCompiler()
    compiled_model = compiler.read_and_parse_archive(model)
    evaluator = Evaluator(compiled_model)
    return evaluator, model

def evaluate_with_inputs(evaluator, model, inputs_map):
    for sheet_name, coords in inputs_map.items():
        for coord, value in coords.items():
            ref = f"{sheet_name}!{coord}"
            try:
                evaluator.set_cell_value(ref, value)
            except Exception:
                evaluator.set_cell_value(ref, value)
    results = {}
    for sheet in model.sheets:
        sheetname = sheet.name
        results[sheetname] = {}
        try:
            for cell in sheet.cells:
                coord = cell.address
                if cell.formula is not None:
                    ref = f"{sheetname}!{coord}"
                    try:
                        val = evaluator.evaluate(ref)
                    except Exception:
                        val = None
                    results[sheetname][coord] = val
        except Exception:
            pass
    return results

app = Flask(__name__)
app.secret_key = SECRET_KEY

EXCEL_PATH = None
EVALUATOR = None
MODEL = None
SCANNED_INPUTS = []
SCANNED_OUTPUTS = []

def initialize():
    global EXCEL_PATH, EVALUATOR, MODEL, SCANNED_INPUTS, SCANNED_OUTPUTS
    EXCEL_PATH = find_excel_path()
    wb = load_workbook(EXCEL_PATH, data_only=False)
    if 'Calcs' not in wb.sheetnames:
        raise RuntimeError("Worksheet 'Calcs' not found in workbook.")
    ws = wb['Calcs']
    SCANNED_INPUTS, SCANNED_OUTPUTS = scan_calcs_sheet_for_yellow_inputs(ws)
    EVALUATOR, MODEL = build_model_from_file(EXCEL_PATH)

initialize()

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if APP_PASSWORD is None:
            return f(*args, **kwargs)
        if session.get('logged_in'):
            return f(*args, **kwargs)
        return redirect(url_for('login', next=request.path))
    return decorated

@app.route('/login', methods=['GET','POST'])
def login():
    if APP_PASSWORD is None:
        flash('Warning: APP_PASSWORD not set. The app is running without password protection.', 'warning')
        return redirect(url_for('index'))
    if request.method == 'POST':
        pw = request.form.get('password','')
        if pw == APP_PASSWORD:
            session['logged_in'] = True
            next_url = request.args.get('next') or url_for('index')
            return redirect(next_url)
        flash('Incorrect password', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route('/', methods=['GET'])
@login_required
def index():
    return render_template('form.html', inputs=SCANNED_INPUTS, sheet_name='Calcs')

@app.route('/calculate', methods=['POST'])
@login_required
def calculate():
    sheet_name = request.form.get('sheet_name', 'Calcs')
    inputs_map = {sheet_name: {}}
    for item in SCANNED_INPUTS:
        coord = item['coord']
        field_name = f'input__{coord}'
        if field_name in request.form:
            raw_val = request.form.get(field_name)
            val = None
            if raw_val is None or raw_val == '':
                val = None
            else:
                try:
                    if '.' in raw_val:
                        val = float(raw_val)
                    else:
                        val = int(raw_val)
                except Exception:
                    val = raw_val
            inputs_map[sheet_name][coord] = val
    try:
        results = evaluate_with_inputs(EVALUATOR, MODEL, inputs_map)
    except Exception as e:
        flash(f'Calculation error: {e}', 'danger')
        results = {}
    calcs_results = results.get('Calcs', {})
    outputs = []
    coord_to_val = calcs_results
    for out in SCANNED_OUTPUTS:
        coord = out['coord']
        outputs.append({'coord': coord, 'label': out.get('label'), 'value': coord_to_val.get(coord)})
    return render_template('result.html', outputs=outputs)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
