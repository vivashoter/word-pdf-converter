from flask import Flask, request, send_file, send_from_directory, jsonify
import os
import subprocess
import tempfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)


@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/style.css")
def styles():
    return send_from_directory(BASE_DIR, "style.css")


@app.route("/script.js")
def scripts():
    return send_from_directory(BASE_DIR, "script.js")


@app.route("/convert/word-to-pdf", methods=["POST"])
def word_to_pdf():

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    uploaded_file = request.files["file"]

    if uploaded_file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not uploaded_file.filename.lower().endswith((".doc", ".docx")):
        return jsonify({"error": "Please upload a Word document"}), 400

    with tempfile.TemporaryDirectory() as temp_dir:

        input_path = os.path.join(
            temp_dir,
            uploaded_file.filename
        )

        uploaded_file.save(input_path)

        try:
            subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    temp_dir,
                    input_path
                ],
                check=True,
                timeout=120
            )

        except Exception as error:
            print(error)

            return jsonify({
                "error": "The document could not be converted."
            }), 500

        output_name = (
            os.path.splitext(uploaded_file.filename)[0]
            + ".pdf"
        )

        output_path = os.path.join(
            temp_dir,
            output_name
        )

        if not os.path.exists(output_path):

            return jsonify({
                "error": "The PDF was not created."
            }), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/pdf"
        )


if __name__ == "__main__":

    port = int(
        os.environ.get("PORT", 10000)
    )

    app.run(
        host="0.0.0.0",
        port=port
    )