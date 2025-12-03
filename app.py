
from flask import Flask, render_template, request

app = Flask(__name__)

def calculate(B2, inputs):
    # TODO: Replace with real Excel-converted formulas
    results = {k: v*2 for k,v in inputs.items()}
    results["B2_effect"] = B2 * 10
    return results

@app.route("/", methods=["GET","POST"])
def index():
    result=None
    if request.method=="POST":
        B2 = float(request.form.get("B2",0))
        inputs = {f"G{i}": float(request.form.get(f"G{i}",0)) 
                  for i in list(range(8,16)) + list(range(19,27))}
        result = calculate(B2, inputs)
    return render_template("index.html", result=result)

if __name__=="__main__":
    app.run()
