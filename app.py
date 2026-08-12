from flask import Flask, request, send_file, send_from_directory, jsonify
import os
import subprocess
import tempfile

app = Flask(__name__, static_folder=".", static_url_path="")


@app.route("/")
def home():
    return send_from_directory(".", "index.html")


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
        input_path = os.path.join(temp_dir, uploaded_file.filename)
        uploaded_file.save(input_path)

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
            check=True
        )

        output_name = os.path.splitext(uploaded_file.filename)[0] + ".pdf"
        output_path = os.path.join(temp_dir, output_name)

        if not os.path.exists(output_path):
            return jsonify({"error": "Conversion failed"}), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
