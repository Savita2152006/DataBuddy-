from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import pandas as pd
import matplotlib.pyplot as plt
import os
import io
import contextlib

# Flask App Setup
app = Flask(__name__)
CORS(app)

# Configurations
UPLOAD_FOLDER = "uploads"
STATIC_FOLDER = "static"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(STATIC_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Globals
df = pd.DataFrame()
command_history = []
plt.switch_backend("Agg")
pd.options.mode.use_inf_as_na = True  # Treat inf as NA

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/download', methods=['GET'])
def download_file():
    path = os.path.join(STATIC_FOLDER, "output.xlsx")
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    else:
        return jsonify({"error": "❌ No Excel file available yet."})

@app.route('/upload', methods=['POST'])
def upload_file():
    global df
    if 'file' not in request.files:
        return jsonify({"error": "No file part in request."})

    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({"error": "No file selected."})

    if file.filename.endswith('.csv'):
        path = os.path.join(app.config["UPLOAD_FOLDER"], "uploaded.csv")
        file.save(path)
        df = pd.read_csv(path)
        df.columns = [col.strip().lower() for col in df.columns]
        return jsonify({"message": "✅ File uploaded successfully."})
    return jsonify({"error": "❌ Only .csv files are allowed."})

@app.route('/analyze', methods=['POST'])
def analyze():
    global df

    if df.empty:
        return jsonify({"error": "❌ Please upload a CSV file first."})

    user_input = request.json.get("command", "").lower().strip()
    command_history.append(user_input)
    output = None
    code = ""

    try:
        # Step 1: Export Command
        if "export" in user_input:
            output_path = os.path.join(STATIC_FOLDER, "output.xlsx")
            df.to_excel(output_path, index=False)
            return jsonify({"file_url": f"/{output_path}"})

        # Step 2: NLP-based Command Matching
        if "print all" in user_input or "show all" in user_input or "display file" in user_input:
            code = "output = df"

        elif "first" in user_input and "rows" in user_input:
            nums = [int(s) for s in user_input.split() if s.isdigit()]
            n = nums[0] if nums else 5
            code = f"output = df.head({n})"

        elif "last" in user_input and "rows" in user_input:
            nums = [int(s) for s in user_input.split() if s.isdigit()]
            n = nums[0] if nums else 5
            code = f"output = df.tail({n})"

        elif "index list" in user_input or "get index" in user_input:
            code = "output = df.index.tolist()"

        elif "info" in user_input:
            code = (
                "buffer = io.StringIO(); "
                "df.info(buf=buffer); "
                "output = buffer.getvalue()"
            )
        elif "unique count" in user_input or "nunique" in user_input or "how many unique" in user_input:
             code = "output = df.nunique()"

        elif "describe" in user_input:
            code = "output = df.describe()"

        elif "null" in user_input and "remove" in user_input:
            code = "df.dropna(inplace=True); output = '✅ Null values removed.'"

        elif "column names" in user_input:
            code = "output = ', '.join(df.columns)"

        elif "average of" in user_input:
            col = user_input.split("average of")[1].strip()
            code = f"output = df['{col}'].mean()"

        elif "max of" in user_input:
            col = user_input.split("max of")[1].strip()
            code = f"output = df['{col}'].max()"

        elif "min of" in user_input:
            col = user_input.split("min of")[1].strip()
            code = f"output = df['{col}'].min()"

        elif "sort by" in user_input:
            col = user_input.split("sort by")[1].strip()
            code = f"output = df.sort_values(by='{col}').head()"

        elif "scatter" in user_input and "vs" in user_input:
            parts = user_input.replace("scatter", "").replace("of", "").strip().split("vs")
            if len(parts) == 2:
                x = parts[0].strip()
                y = parts[1].strip()
                code = (
                    f"plt.clf(); df.plot.scatter(x='{x}', y='{y}'); "
                    f"plt.savefig('{STATIC_FOLDER}/output.png'); output = 'plot'"
                )

        elif "bar chart" in user_input or "bargraph" in user_input:
            col = user_input.split("of")[1].strip()
            code = (
                f"plt.clf(); df['{col}'].value_counts().plot(kind='bar'); "
                f"plt.title('Bar chart of {col}'); plt.savefig('{STATIC_FOLDER}/output.png'); output = 'plot'"
            )

        elif "histogram of" in user_input:
            col = user_input.split("of")[1].strip()
            code = (
                f"plt.clf(); df['{col}'].hist(); "
                f"plt.title('Histogram of {col}'); plt.savefig('{STATIC_FOLDER}/output.png'); output = 'plot'"
            )

        else:
            # Freeform code
            code = user_input

        # Step 3: Column Validation (basic)
        if "df['" in code:
            try:
                col_name = code.split("df['")[1].split("']")[0]
                if col_name not in df.columns:
                    return jsonify({"error": f"❌ Column '{col_name}' not found in dataset."})
            except:
                return jsonify({"error": "❌ Column parsing error."})

        # Step 4: Code Execution
        local_vars = {}
        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            exec(code, {"df": df, "plt": plt, "pd": pd, "io": io}, local_vars)

        output = local_vars.get("output")
        if output is None:
            output = stdout.getvalue() or "✅ Done."

        # Save latest DataFrame
        df.to_excel(os.path.join(STATIC_FOLDER, "output.xlsx"), index=False)

        # Step 5: Return Response
        if isinstance(output, str) and output == "plot":
            return jsonify({"plot_url": "/static/output.png"})
        elif isinstance(output, pd.DataFrame):
            return jsonify({"result": output.to_html()})
        elif hasattr(output, "to_html") and not isinstance(output, pd.DataFrame):
            return jsonify({"result": output.to_html()})
        elif isinstance(output, str):
            # Display plain string like df.info() using <pre> tag for alignment
            return jsonify({"result": f"<pre>{output}</pre>"})
        elif isinstance(output, (list, dict, int, float)):
            return jsonify({"result": str(output)})
        else:
            return jsonify({"result":str(output)})

    except Exception as e:
        return jsonify({"error": f"❌ Error: {str(e)}"})

@app.route('/history', methods=['GET'])
def history():
    return jsonify({"history": command_history})

if __name__ == '__main__':
    app.run(debug=False)