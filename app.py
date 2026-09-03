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


            if result.stdout:

                print(
                    "LibreOffice output:",
                    result.stdout,
                    flush=True
                )


            if result.stderr:

                print(
                    "LibreOffice warnings:",
                    result.stderr,
                    flush=True
                )


        except subprocess.TimeoutExpired:

            return jsonify({
                "error":
                "The conversion took too long."
            }), 504


        except subprocess.CalledProcessError as error:

            print(
                "LibreOffice error:",
                error.stderr,
                flush=True
            )

            return jsonify({
                "error":
                "The Word document could not be converted."
            }), 500


        except Exception as error:

            print(
                "Word to PDF error:",
                repr(error),
                flush=True
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


        if os.path.getsize(
            output_path
        ) == 0:

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
# EXACTDOC
# =========================================

def run_exactdoc(
    input_path,
    output_path
):

    print(
        "Running ExactDoc on:",
        input_path,
        flush=True
    )


    result = subprocess.run(
        [
            "exactdoc",
            input_path,
            "-o",
            output_path
        ],
        check=True,
        timeout=300,
        capture_output=True,
        text=True
    )


    if result.stdout:

        print(
            "ExactDoc output:",
            result.stdout,
            flush=True
        )


    if result.stderr:

        print(
            "ExactDoc warnings:",
            result.stderr,
            flush=True
        )


    return result


# =========================================
# FORCE OCR / FULL RASTERIZE
# =========================================

def create_clean_ocr_pdf(
    input_path,
    output_path
):

    print(
        "Creating completely flattened OCR version of PDF...",
        flush=True
    )


    result = subprocess.run(
        [
            "ocrmypdf",

            "--mode",
            "force",

            "--output-type",
            "pdf",

            "--optimize",
            "0",

            "--rotate-pages",

            "--deskew",

            "--oversample",
            "300",

            "--language",
            "eng",

            "--jobs",
            "1",

            "--tesseract-timeout",
            "120",

            input_path,

            output_path
        ],
        check=True,
        timeout=240,
        capture_output=True,
        text=True
    )


    if result.stdout:

        print(
            "OCRmyPDF output:",
            result.stdout,
            flush=True
        )


    if result.stderr:

        print(
            "OCRmyPDF warnings:",
            result.stderr,
            flush=True
        )


    if not os.path.exists(
        output_path
    ):

        print(
            "OCRmyPDF did not create output.",
            flush=True
        )

        return False


    if os.path.getsize(
        output_path
    ) == 0:

        print(
            "OCRmyPDF output is empty.",
            flush=True
        )

        return False


    print(
        "Clean OCR PDF created.",
        flush=True
    )


    return True


# =========================================
# CONVERT AFTER OCR
# =========================================

def convert_after_ocr(
    input_path,
    output_path,
    clean_pdf_path
):

    print(
        "Rasterizing PDF and rebuilding text layer...",
        flush=True
    )


    clean_success = create_clean_ocr_pdf(
        input_path,
        clean_pdf_path
    )


    if not clean_success:

        raise RuntimeError(
            "OCRmyPDF did not create a usable clean PDF."
        )


    if os.path.exists(
        output_path
    ):

        os.remove(
            output_path
        )


    print(
        "Sending clean OCR PDF to ExactDoc...",
        flush=True
    )


    return run_exactdoc(
        clean_pdf_path,
        output_path
    )


# =========================================
# PDF TO WORD
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
            base_name
            + ".docx"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        clean_pdf_path = os.path.join(
            temp_dir,
            base_name
            + "_clean_ocr.pdf"
        )


        print(
            "=================================",
            flush=True
        )

        print(
            "PDF TO WORD START:",
            filename,
            flush=True
        )

        print(
            "=================================",
            flush=True
        )


        try:

            # =================================
            # FIRST ATTEMPT:
            # ORIGINAL PDF THROUGH EXACTDOC
            # =================================

            try:

                run_exactdoc(
                    input_path,
                    output_path
                )


            except subprocess.CalledProcessError as first_error:

                print(
                    "ExactDoc first attempt exit code:",
                    first_error.returncode,
                    flush=True
                )


                print(
                    "ExactDoc first attempt stdout:",
                    first_error.stdout,
                    flush=True
                )


                print(
                    "ExactDoc first attempt stderr:",
                    first_error.stderr,
                    flush=True
                )


                # =================================
                # EXACTDOC 17:
                # OCR REQUIRED
                #
                # EXACTDOC 19:
                # INTERACTIVE FORM
                # =================================

                if first_error.returncode in (
                    17,
                    19
                ):

                    if (
                        first_error.returncode
                        == 19
                    ):

                        print(
                            "Interactive PDF form detected.",
                            flush=True
                        )

                    else:

                        print(
                            "Scanned PDF detected.",
                            flush=True
                        )


                    print(
                        "Switching to FORCE OCR conversion pipeline.",
                        flush=True
                    )


                    try:

                        convert_after_ocr(
                            input_path,
                            output_path,
                            clean_pdf_path
                        )


                    except subprocess.CalledProcessError as second_error:

                        print(
                            "ExactDoc after OCR exit code:",
                            second_error.returncode,
                            flush=True
                        )


                        print(
                            "ExactDoc after OCR stdout:",
                            second_error.stdout,
                            flush=True
                        )


                        print(
                            "ExactDoc after OCR stderr:",
                            second_error.stderr,
                            flush=True
                        )


                        return jsonify({
                            "error":
                            "The PDF was cleaned successfully, "
                            "but its layout could not be converted "
                            "reliably to editable Word."
                        }), 422


                else:

                    raise first_error


        except subprocess.TimeoutExpired:

            print(
                "PDF conversion timed out.",
                flush=True
            )


            return jsonify({
                "error":
                "The PDF conversion took too long."
            }), 504


        except subprocess.CalledProcessError as error:

            print(
                "PDF conversion command failed.",
                flush=True
            )


            print(
                "Exit code:",
                error.returncode,
                flush=True
            )


            print(
                "STDOUT:",
                error.stdout,
                flush=True
            )


            print(
                "STDERR:",
                error.stderr,
                flush=True
            )


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        except FileNotFoundError as error:

            print(
                "Required conversion command missing:",
                repr(error),
                flush=True
            )


            return jsonify({
                "error":
                "A required PDF conversion tool "
                "is unavailable on the server."
            }), 500


        except RuntimeError as error:

            print(
                "PDF conversion runtime error:",
                repr(error),
                flush=True
            )


            return jsonify({
                "error":
                str(error)
            }), 500


        except Exception as error:

            print(
                "Unexpected PDF to Word error:",
                repr(error),
                flush=True
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

            print(
                "No DOCX file was produced.",
                flush=True
            )


            return jsonify({
                "error":
                "The Word document could not be created."
            }), 500


        if os.path.getsize(
            output_path
        ) == 0:

            print(
                "DOCX file is empty.",
                flush=True
            )


            return jsonify({
                "error":
                "The Word document was created "
                "but was empty."
            }), 500


        print(
            "=================================",
            flush=True
        )

        print(
            "PDF TO WORD SUCCESS:",
            output_name,
            flush=True
        )

        print(
            "=================================",
            flush=True
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
