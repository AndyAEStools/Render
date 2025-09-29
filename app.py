import os
import tempfile
import zipfile
import hmac
from pathlib import Path
from flask import Flask, render_template, request, send_file, abort

from SAPXMLTool import process_xmls  # expects (excel_file, xml_folder, output_folder)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

def ensure_ext(filename, allowed):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed

def check_password(pw_input: str) -> bool:
    expected = os.environ.get("APP_PASSWORD", "")
    # Allow empty expected to behave as "no password set"
    if expected == "":
        return True
    return hmac.compare_digest(pw_input or "", expected)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", error=None)

@app.route("/process", methods=["POST"])
def process_route():
    # 1) Password check
    if not check_password(request.form.get("password", "")):
        return render_template("index.html", error="Incorrect password."), 401

    # 2) Uploaded files
    excel = request.files.get("excel")
    xml_zip = request.files.get("xmlzip")
    if not excel or not xml_zip:
        return render_template("index.html", error="Please upload BOTH the Excel file and the ZIP of XMLs."), 400

    if not ensure_ext(excel.filename, {"xlsx", "xlsm", "xltx", "xltm", "xls"}):
        return render_template("index.html", error="Excel must be .xls/.xlsx/.xlsm/.xltx/.xltm"), 400

    if not ensure_ext(xml_zip.filename, {"zip"}):
        return render_template("index.html", error="XMLs must be provided as a .zip"), 400

    # 3) Work in a temporary sandbox
    with tempfile.TemporaryDirectory() as workdir:
        workdir = Path(workdir)
        inputs_dir = workdir / "inputs"
        xml_in_dir = inputs_dir / "xmls"
        out_dir = workdir / "out"
        inputs_dir.mkdir(parents=True, exist_ok=True)
        xml_in_dir.mkdir(parents=True, exist_ok=True)
        out_dir.mkdir(parents=True, exist_ok=True)

        # Save inputs
        excel_path = inputs_dir / excel.filename
        excel.save(str(excel_path))

        zip_path = inputs_dir / xml_zip.filename
        xml_zip.save(str(zip_path))

        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(xml_in_dir)

        # 4) Run your existing logic
        process_xmls(str(excel_path), str(xml_in_dir), str(out_dir))

        # 5) Zip outputs and return
        out_zip_path = workdir / "edited_xmls.zip"
        with zipfile.ZipFile(out_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for p in out_dir.rglob("*"):
                if p.is_file():
                    zf.write(p, p.relative_to(out_dir))

        return send_file(
            out_zip_path,
            mimetype="application/zip",
            as_attachment=True,
            download_name="edited_xmls.zip",
        )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
