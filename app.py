from flask import (
    Flask,
    request,
    send_file,
    send_from_directory,
    jsonify
)

from pdf2docx import Converter
from werkzeug.utils import secure_filename

import os
import subprocess
import tempfile


BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)


app = Flask(__name__)


# ---------------------------------
# FILE SIZE LIMIT
# ---------------------------------

MAX_FILE_SIZE = 25 * 1024 * 1024

app.config["MAX_CONTENT_LENGTH"] = (
    MAX_FILE_SIZE
)


# ---------------------------------
# HOME PAGE
# ---------------------------------

@app.route("/")
def home():

    return send_from_directory(
        BASE_DIR,
        "index.html"
    )

# ---------------------------------
# INFORMATION PAGES
# ---------------------------------

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


# ---------------------------------
# WEBSITE FILES
# ---------------------------------

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


# ---------------------------------
# LOGO
# ---------------------------------

@app.route("/convertdocgoose-logo.png")
def logo():

    return send_from_directory(
        BASE_DIR,
        "convertdocgoose-logo.png"
    )


# ---------------------------------
# FILE TOO LARGE
# ---------------------------------

@app.errorhandler(413)
def file_too_large(error):

    return jsonify({
        "error":
        "File is too large. Maximum size is 25 MB."
    }), 413


# ---------------------------------
# WORD TO PDF
# ---------------------------------

@app.route(
    "/convert/word-to-pdf",
    methods=["POST"]
)
def word_to_pdf():

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
                "The conversion took too long. Please try again."
            }), 504


        except subprocess.CalledProcessError as error:

            print(
                "LibreOffice conversion error:",
                error.stderr
            )


            return jsonify({
                "error":
                "The Word document could not be converted."
            }), 500


        except Exception as error:

            print(
                "Word to PDF error:",
                error
            )


            return jsonify({
                "error":
                "The Word document could not be converted."
            }), 500


        output_name = (
            os.path.splitext(
                filename
            )[0]
            + ".pdf"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        if not os.path.exists(
            output_path
        ):

            return jsonify({
                "error":
                "The PDF could not be created."
            }), 500


        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/pdf"
        )


# ---------------------------------
# OCR PDF
# ---------------------------------

def create_ocr_pdf(
    input_path,
    output_path
):

    command = [
        "ocrmypdf",

        "--skip-text",

        "--rotate-pages",

        "--deskew",

        "--optimize",
        "0",

        "--output-type",
        "pdf",

        "--language",
        "eng",

        input_path,

        output_path
    ]


    result = subprocess.run(
        command,
        check=True,
        timeout=150,
        capture_output=True,
        text=True
    )


    if result.stdout:

        print(
            "OCRmyPDF output:",
            result.stdout
        )


    if result.stderr:

        print(
            "OCRmyPDF warnings:",
            result.stderr
        )


    return (
        os.path.exists(
            output_path
        )
        and
        os.path.getsize(
            output_path
        ) > 0
    )


# ---------------------------------
# PDF TO WORD
# ---------------------------------

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


        ocr_pdf_path = os.path.join(
            temp_dir,
            base_name
            + "_ocr.pdf"
        )


        output_name = (
            base_name
            + ".docx"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        pdf_for_conversion = (
            input_path
        )


        # --------------------------
        # OCR PASS
        # --------------------------

        try:

            ocr_success = create_ocr_pdf(
                input_path,
                ocr_pdf_path
            )


            if ocr_success:

                pdf_for_conversion = (
                    ocr_pdf_path
                )


                print(
                    "OCR-enhanced PDF will be used."
                )


        except subprocess.TimeoutExpired:

            print(
                "OCR timed out. "
                "Falling back to original PDF."
            )


        except subprocess.CalledProcessError as error:

            print(
                "OCRmyPDF could not process this PDF."
            )


            print(
                error.stderr
            )


            print(
                "Falling back to original PDF."
            )


        except Exception as error:

            print(
                "OCR error:",
                error
            )


            print(
                "Falling back to original PDF."
            )


        # --------------------------
        # PDF TO DOCX
        # --------------------------

        converter = None


        try:

            converter = Converter(
                pdf_for_conversion
            )


            converter.convert(
                output_path,
                start=0,
                end=None
            )


        except Exception as error:

            print(
                "PDF to Word error:",
                error
            )


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        finally:

            if converter is not None:

                converter.close()


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


# ---------------------------------
# START SERVER
# ---------------------------------

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
