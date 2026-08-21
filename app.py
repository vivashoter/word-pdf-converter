from flask import (
    Flask,
    request,
    send_file,
    send_from_directory,
    jsonify
)

from werkzeug.utils import secure_filename

import os
import subprocess
import tempfile


# =========================================
# BASIC CONFIG
# =========================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

app = Flask(__name__)

MAX_FILE_SIZE = 25 * 1024 * 1024

app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE


# =========================================
# WEBSITE ROUTES
# =========================================

@app.route("/")
def home():
    return send_from_directory(
        BASE_DIR,
        "index.html"
    )


@app.route("/privacy.html")
def privacy():
    return send_from_directory(
        BASE_DIR,
        "privacy.html"
    )


@app.route("/terms.html")
def terms():
    return send_from_directory(
        BASE_DIR,
        "terms.html"
    )


@app.route("/about.html")
def about():
    return send_from_directory(
        BASE_DIR,
        "about.html"
    )


@app.route("/contact.html")
def contact():
    return send_from_directory(
        BASE_DIR,
        "contact.html"
    )


@app.route("/style.css")
def styles():
    return send_from_directory(
        BASE_DIR,
        "style.css"
    )


@app.route("/script.js")
def scripts():
    return send_from_directory(
        BASE_DIR,
        "script.js"
    )


@app.route("/convertdocgoose-logo.png")
def logo():
    return send_from_directory(
        BASE_DIR,
        "convertdocgoose-logo.png"
    )


@app.route("/favicon.png")
def favicon():
    return send_from_directory(
        BASE_DIR,
        "favicon.png"
    )


# =========================================
# FILE TOO LARGE
# =========================================

@app.errorhandler(413)
def file_too_large(error):

    return jsonify({
        "error":
        "File is too large. Maximum size is 25 MB."
    }), 413


# =========================================
# WORD TO PDF
# =========================================

@app.route(
    "/convert/word-to-pdf",
    methods=["POST"]
)
def word_to_pdf():

    if "file" not in request.files:
        return jsonify({
            "error": "No file uploaded."
        }), 400


    uploaded_file = request.files["file"]


    if uploaded_file.filename == "":
        return jsonify({
            "error": "No file selected."
        }), 400


    filename = secure_filename(
        uploaded_file.filename
    )


    if not filename.lower().endswith(
        (".doc", ".docx")
    ):
        return jsonify({
            "error":
            "Please upload a DOC or DOCX file."
        }), 400


    with tempfile.TemporaryDirectory() as temp_dir:

        input_path = os.path.join(
            temp_dir,
            filename
        )

        uploaded_file.save(
            input_path
        )


        try:

            result = subprocess.run(
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
                timeout=120,
                capture_output=True,
                text=True
            )


            if result.stdout:
                print(
                    "LibreOffice output:",
                    result.stdout
                )


            if result.stderr:
                print(
                    "LibreOffice warnings:",
                    result.stderr
                )


        except subprocess.TimeoutExpired:

            return jsonify({
                "error":
                "The conversion took too long."
            }), 504


        except subprocess.CalledProcessError as error:

            print(
                "LibreOffice error:",
                error.stderr
            )

            return jsonify({
                "error":
                "The Word document could not be converted."
            }), 500


        except Exception as error:

            print(
                "Word to PDF error:",
                repr(error)
            )

            return jsonify({
                "error":
                "The Word document could not be converted."
            }), 500


        output_name = (
            os.path.splitext(filename)[0]
            + ".pdf"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        if not os.path.exists(output_path):
            return jsonify({
                "error":
                "The PDF could not be created."
            }), 500


        if os.path.getsize(output_path) == 0:
            return jsonify({
                "error":
                "The PDF was created but was empty."
            }), 500


        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/pdf"
        )


# =========================================
# EXACTDOC HELPER
# =========================================

def run_exactdoc(
    input_path,
    output_path
):

    return subprocess.run(
        [
            "exactdoc",
            input_path,
            "-o",
            output_path
        ],
        check=True,
        timeout=180,
        capture_output=True,
        text=True
    )


# =========================================
# FLATTEN INTERACTIVE PDF
# =========================================

def flatten_pdf_form(
    input_path,
    output_path
):

    result = subprocess.run(
        [
            "qpdf",

            input_path,

            output_path,

            "--generate-appearances",

            "--flatten-annotations=all"
        ],

        check=True,
        timeout=120,
        capture_output=True,
        text=True
    )


    if result.stdout:
        print(
            "QPDF output:",
            result.stdout
        )


    if result.stderr:
        print(
            "QPDF warnings:",
            result.stderr
        )


    return (
        os.path.exists(output_path)
        and
        os.path.getsize(output_path) > 0
    )


# =========================================
# PDF TO WORD
# EXACTDOC + FORM FLATTEN RETRY
# =========================================

@app.route(
    "/convert/pdf-to-word",
    methods=["POST"]
)
def pdf_to_word():

    if "file" not in request.files:
        return jsonify({
            "error":
            "No file uploaded."
        }), 400


    uploaded_file = request.files["file"]


    if uploaded_file.filename == "":
        return jsonify({
            "error":
            "No file selected."
        }), 400


    filename = secure_filename(
        uploaded_file.filename
    )


    if not filename.lower().endswith(
        ".pdf"
    ):
        return jsonify({
            "error":
            "Please upload a PDF file."
        }), 400


    with tempfile.TemporaryDirectory() as temp_dir:

        input_path = os.path.join(
            temp_dir,
            filename
        )

        uploaded_file.save(
            input_path
        )


        base_name = os.path.splitext(
            filename
        )[0]


        output_name = (
            base_name +
            ".docx"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        flattened_path = os.path.join(
            temp_dir,
            base_name +
            "_flattened.pdf"
        )


        try:

            print(
                "Starting ExactDoc conversion:",
                filename
            )


            try:

                result = run_exactdoc(
                    input_path,
                    output_path
                )


                if result.stdout:
                    print(
                        "ExactDoc output:",
                        result.stdout
                    )


                if result.stderr:
                    print(
                        "ExactDoc warnings:",
                        result.stderr
                    )


            except subprocess.CalledProcessError as error:

                print(
                    "ExactDoc first attempt exit code:",
                    error.returncode
                )

                print(
                    "ExactDoc first attempt stdout:",
                    error.stdout
                )

                print(
                    "ExactDoc first attempt stderr:",
                    error.stderr
                )


                # =================================
                # INTERACTIVE FORM
                # =================================

                if error.returncode == 19:

                    print(
                        "Interactive PDF detected."
                    )

                    print(
                        "Flattening temporary copy "
                        "and retrying ExactDoc."
                    )


                    flatten_success = flatten_pdf_form(
                        input_path,
                        flattened_path
                    )


                    if not flatten_success:

                        return jsonify({
                            "error":
                            "The interactive PDF could not "
                            "be prepared for Word conversion."
                        }), 500


                    # Remove any partial DOCX
                    if os.path.exists(
                        output_path
                    ):

                        os.remove(
                            output_path
                        )


                    retry_result = run_exactdoc(
                        flattened_path,
                        output_path
                    )


                    if retry_result.stdout:
                        print(
                            "ExactDoc retry output:",
                            retry_result.stdout
                        )


                    if retry_result.stderr:
                        print(
                            "ExactDoc retry warnings:",
                            retry_result.stderr
                        )


                else:

                    raise


        except subprocess.TimeoutExpired:

            print(
                "PDF conversion timed out."
            )

            return jsonify({
                "error":
                "The PDF conversion took too long."
            }), 504


        except subprocess.CalledProcessError as error:

            print(
                "Final conversion exit code:",
                error.returncode
            )

            print(
                "Final conversion stdout:",
                error.stdout
            )

            print(
                "Final conversion stderr:",
                error.stderr
            )


            if error.returncode == 17:

                return jsonify({
                    "error":
                    "This appears to be a scanned PDF. "
                    "OCR support will be added next."
                }), 422


            if error.returncode == 19:

                return jsonify({
                    "error":
                    "This interactive PDF still could not "
                    "be converted reliably after flattening."
                }), 422


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        except FileNotFoundError as error:

            print(
                "Missing server command:",
                repr(error)
            )

            return jsonify({
                "error":
                "The PDF conversion tool is not available "
                "on the server."
            }), 500


        except Exception as error:

            print(
                "PDF to Word unexpected error:",
                repr(error)
            )

            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        # =================================
        # VERIFY OUTPUT
        # =================================

        if not os.path.exists(
            output_path
        ):

            return jsonify({
                "error":
                "The Word document could not be created."
            }), 500


        if os.path.getsize(
            output_path
        ) == 0:

            return jsonify({
                "error":
                "The Word document was created but was empty."
            }), 500


        print(
            "PDF to Word conversion complete:",
            output_name
        )


        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype=(
                "application/vnd."
                "openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )


# =========================================
# START SERVER
# =========================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )


    app.run(
        host="0.0.0.0",
        port=port
    )
