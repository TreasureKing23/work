from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from io import BytesIO
import mapping  # the module above
import os
from datetime import datetime
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = "temp_folder"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024 

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        exam_name = request.form.get("exam")
        mode = request.form.get("mode")
        exam_period = request.form.get("period")
        inp = request.files.get("input_spec")
        data = request.files.get("data_wb")

        if not exam_name or not inp or not data:
            flash("ALL FIELDS REQUIRED.", "error")
            return redirect(url_for("index"))
        
        inp_stream = BytesIO(inp.read())
        data_stream = BytesIO(data.read())

        out_stream = mapping.formatting(inp_stream, data_stream, exam_name, exam_period, mode)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = secure_filename(f"{exam_name}_{exam_period}_{mode}_formatted_{timestamp}.xlsx")

        return send_file(out_stream,
                         as_attachment=True,
                         download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
