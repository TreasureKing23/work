from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from io import BytesIO
import mapping  # the module above
import os
from datetime import datetime
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = "uploads"   # simple temp folder
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = "replace-with-a-secret"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024   # 20 MB


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        exam = request.form.get("exam")
        mode = request.form.get("mode")
        inp = request.files.get("input_spec")
        data = request.files.get("data_file")

        if not exam or not inp or not data:
            flash("All fields are required.", "error")
            return redirect(url_for("index"))

        # Secure filenames and save to disk (or keep in memory)
        inp_stream = BytesIO(inp.read())
        data_stream = BytesIO(data.read())

        # Run formatter
        out_stream = mapping.format_workbook(inp_stream, data_stream, exam)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = secure_filename(f"{exam}_{mode}_formatted_{timestamp}.xlsx")

        return send_file(out_stream,
                         as_attachment=True,
                         download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)   # set False in production
