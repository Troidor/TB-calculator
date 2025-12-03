from flask import Flask, render_template, request
from openpyxl import load_workbook
from pycel.excelcompiler import ExcelCompiler
import os

app = Flask(__name__)

EXCEL_PATH = os.path.join("excel", "TB_v18.xlsx")
SHEET = "Calcs"

def find_yellow_cells():
    """Scan Excel for yellow fill input cells."""
    wb = load_workbook(EXCEL_PATH, data_only=False)
    ws = wb[SHEET]

    yellow_cells = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                if cell.fill.start_color.rgb.upper() in ("FFFFFF00", "FFFFEB9C"):
                    yellow_cells.append(cell.coordinate)

    return yellow_cells


@app.route("/", methods=["GET"])
def home():
    inputs = find_yellow_cells()
    return render_template("form.html", inputs=inputs)


@app.route("/calculate", methods=["POST"])
def calculate():
    user_inputs = request.form.to_dict()

    # Load workbook (not needed for calculation but for value injection)
    wb = load_workbook(EXCEL_PATH, data_only=False)
    ws = wb[SHEET]

    # Insert user inputs into yellow cells
    for cell, value in user_inputs.items():
        try:
            ws[cell].value = float(value)
        except:
            ws[cell].value = value

    # Save temp workbook
    temp_path = "temp_calc.xlsx"
    wb.save(temp_path)

    # Recalculate using pycel
    comp = ExcelCompiler(temp_path)
    comp.calculate()

    # Example: return 10 result cells (you can change)
    output_cells = ["H20", "H21", "H22", "H23", "H24", "H25", "H26", "H27", "H28", "H29"]
    results = {}

    for cell in output_cells:
        try:
            results[cell] = comp.evaluate(f"{SHEET}!{cell}")
        except:
            results[cell] = "ERR"

    return render_template("result.html", results=results)
    

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
