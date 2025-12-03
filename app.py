import os
import tempfile
import subprocess
from flask import Flask, render_template, request, redirect, url_for, session, flash
from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple

# --- Config ---
APP_PASSWORD = os.environ.get("APP_PASSWORD")  # REQUIRED: set on Fly as secret
SECRET_KEY = os.environ.get("SECRET_KEY", "change-me")
EXCEL_PATH = os.path.join("excel", "TB_v18.xlsx")
SHEET_NAME = "Calcs"

app = Flask(__name__)
app.secret_key = SECRET_KEY


# ---------- Helpers ----------
def is_yellow_cell(cell):
    """Return True if openpyxl cell has a yellow-ish fill (common patterns)."""
    try:
        fill = cell.fill
        if not fill:
            return False
        # Check start_color or fgColor for rgb value
        color = None
        if hasattr(fill, "start_color") and fill.start_color:
            color = fill.start_color
        elif hasattr(fill, "fgColor") and fill.fgColor:
            color = fill.fgColor

        if color is None:
            return False

        rgb = getattr(color, "rgb", None)
        if not rgb:
            return False
        rgb = rgb.upper()
        # Many yellow fills start with "FFFF" / ARGB or are like FFFFEB9C / FFFF... etc.
        return rgb.startswith("FFFF") and (int(rgb[-6:0+2], 16) >= 0 or True)
    except Exception:
        return False


def scan_yellow_inputs():
    """Return list of coordinates (e.g. C3) that are yellow in SHEET_NAME."""
    wb = load_workbook(EXCEL_PATH, data_only=False)
    if SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f"Sheet {SHEET_NAME} not found in workbook.")
    ws = wb[SHEET_NAME]

    yellow_coords = []
    for row in ws.iter_rows():
        for cell in row:
            try:
                if is_yellow_cell(cell):
                    yellow_coords.append(cell.coordinate)
            except Exception:
                pass
    return yellow_coords


def find_formula_cells():
    """Return list of coordinates of formula cells on SHEET_NAME."""
    wb = load_workbook(EXCEL_PATH, data_only=False)
    ws = wb[SHEET_NAME]
    formula_coords = []
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, str) and val.startswith("="):
                formula_coords.append(cell.coordinate)
    return formula_coords


def recalc_with_libreoffice(input_path, output_dir):
    """
    Run LibreOffice headless to open and save the workbook, causing recalculation.
    Writes the recalculated workbook into output_dir and returns the output path.
    """
    # LibreOffice convert-to will open and save; use --headless to avoid GUI.
    # We will convert to xlsx and save into output_dir.
    cmd = [
        "soffice",
        "--headless",
        "--convert-to", "xlsx",
        "--outdir", output_dir,
        input_path
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice failed: stdout={proc.stdout.decode('utf8',errors='ignore')}\nstderr={proc.stderr.decode('utf8',errors='ignore')}")
    # LibreOffice names the output file with the same basename but .xlsx
    base = os.path.basename(input_path)
    name_without_ext = os.path.splitext(base)[0]
    out_path = os.path.join(output_dir, f"{name_without_ext}.xlsx")
    if not os.path.exists(out_path):
        # in some LibreOffice installs it preserves the full name; try to locate any xlsx in output_dir
        candidates = [os.path.join(output_dir, f) for f in os.listdir(output_dir) if f.lower().endswith(".xlsx")]
        if candidates:
            out_path = candidates[0]
        else:
            raise FileNotFoundError("LibreOffice did not create output xlsx.")
    return out_path


# ---------- Authentication ----------
def logged_in():
    if APP_PASSWORD is None:
        # if no password set, run in open mode (not recommended)
        return True
    return session.get("logged_in") is True


@app.route("/login", methods=["GET", "POST"])
def login():
    if APP_PASSWORD is None:
        flash("APP_PASSWORD not set on server; running without authentication", "warning")
        session["logged_in"] = True
        return redirect(url_for("index"))
    if request.method == "POST":
        pw = request.form.get("password", "")
        if pw == APP_PASSWORD:
            session["logged_in"] = True
            return redirect(url_for("index"))
        flash("Incorrect password", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    return redirect(url_for("login"))


# ---------- Routes ----------
@app.route("/", methods=["GET"])
def index():
    if not logged_in():
        return redirect(url_for("login"))
    try:
        inputs = scan_yellow_inputs()
    except Exception as e:
        return f"Error scanning workbook: {e}", 500
    # Render a form with one field per yellow cell (name uses the coordinate)
    return render_template("form.html", inputs=inputs)


@app.route("/calculate", methods=["POST"])
def calculate():
    if not logged_in():
        return redirect(url_for("login"))
    # Get all yellow input coordinates
    inputs = scan_yellow_inputs()

    # Create temp file copy of workbook, write user-provided values into same coords,
    # run LibreOffice to recalc, then read formula outputs.
    try:
        tmpdir = tempfile.mkdtemp(prefix="tbcalc_")
        # copy original workbook to temp
        tmp_input = os.path.join(tmpdir, "input.xlsx")
        wb = load_workbook(EXCEL_PATH, data_only=False)
        wb.save(tmp_input)

        # Load the temp workbook and write values
        wb2 = load_workbook(tmp_input, data_only=False)
        ws = wb2[SHEET_NAME]

        # For each expected input coordinate, read posted value
        for coord in inputs:
            field_name = coord
            if field_name in request.form:
                raw = request.form.get(field_name)
                # attempt numeric conversion
                v = None
                if raw is None or raw == "":
                    v = None
                else:
                    try:
                        if "." in raw:
                            v = float(raw)
                        else:
                            v = int(raw)
                    except Exception:
                        v = raw
                ws[coord].value = v

        # save modified workbook to temp file
        tmp_modified = os.path.join(tmpdir, "modified.xls")  # use xls or xlsx; LibreOffice will convert
        wb2.save(tmp_modified)

        # Run LibreOffice to convert/save (this will recalc formulas)
        out_path = recalc_with_libreoffice(tmp_modified, tmpdir)

        # Now read recalculated workbook (data_only=True to read cached computed values)
        wb_final = load_workbook(out_path, data_only=True)
        ws_final = wb_final[SHEET_NAME]

        # Collect results: we will return all formula cells (their computed values)
        outputs = []
        for row in ws_final.iter_rows():
            for cell in row:
                # Identify cells that were formulas in the original (we check original file)
                orig = load_workbook(EXCEL_PATH, data_only=False)[SHEET_NAME]
                orig_cell = orig[cell.coordinate]
                if isinstance(orig_cell.value, str) and orig_cell.value.startswith("="):
                    outputs.append({
                        "coord": cell.coordinate,
                        "label": None,  # could guess from left/up cells if desired
                        "value": cell.value
                    })

        # Sort outputs by coordinate for stable ordering
        outputs = sorted(outputs, key=lambda x: (coordinate_to_tuple(x["coord"])[0], coordinate_to_tuple(x["coord"])[1]))

        return render_template("result.html", outputs=outputs)

    except Exception as e:
        # Surface helpful error
        return f"Calculation failed: {e}", 500


if __name__ == "__main__":
    # PORT will be provided by Fly (via $PORT)
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
