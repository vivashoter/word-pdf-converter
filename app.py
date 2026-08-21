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

import pikepdf


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

    print(
        "Running ExactDoc on:",
        input_path,
        flush=True
    )


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
# OCR HELPER
# =========================================

def create_ocr_pdf(
    input_path,
    output_path
):

    print(
        "Running OCRmyPDF...",
        flush=True
    )


    result = subprocess.run(
        [
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
        ],
        check=True,
        timeout=150,
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


    return (
        os.path.exists(output_path)
        and
        os.path.getsize(output_path) > 0
    )


# =========================================
# FLATTEN INTERACTIVE PDF
# USING PIKEPDF
# =========================================

def flatten_pdf_form(
    input_path,
    output_path
):

    print(
        "Flattening PDF form with pikepdf...",
        flush=True
    )


    try:

        with pikepdf.Pdf.open(
            input_path
        ) as pdf:

            # Create appearance streams for fields
            # so their visible values remain on the page.
            pdf.generate_appearance_streams()


            # Burn annotations/form appearances
            # into normal PDF page content.
            pdf.flatten_annotations(
                "all"
            )


            pdf.save(
                output_path
            )


        if not os.path.exists(
            output_path
        ):

            print(
                "pikepdf did not create output.",
                flush=True
            )

            return False


        if os.path.getsize(
            output_path
        ) == 0:

            print(
                "pikepdf output is empty.",
                flush=True
            )

            return False


        print(
            "pikepdf form flatten complete.",
            flush=True
        )


        return True


    except Exception as error:

        print(
            "pikepdf flatten error:",
            repr(error),
            flush=True
        )


        return False


# =========================================
# PDF TO WORD
# EXACTDOC + PIKEPDF + OCR FALLBACKS
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


        ocr_path = os.path.join(
            temp_dir,
            base_name +
            "_ocr.pdf"
        )


        try:

            print(
                "=================================",
                flush=True
            )

            print(
                "Starting PDF to Word:",
                filename,
                flush=True
            )

            print(
                "=================================",
                flush=True
            )


            # =================================
            # FIRST EXACTDOC ATTEMPT
            # =================================

            try:

                result = run_exactdoc(
                    input_path,
                    output_path
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
                # INTERACTIVE FORM
                # =================================

                if first_error.returncode == 19:

                    print(
                        "Interactive PDF detected.",
                        flush=True
                    )

                    print(
                        "Flattening with pikepdf "
                        "and retrying ExactDoc.",
                        flush=True
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


                    if os.path.exists(
                        output_path
                    ):

                        os.remove(
                            output_path
                        )


                    try:

                        retry_result = run_exactdoc(
                            flattened_path,
                            output_path
                        )


                        if retry_result.stdout:

                            print(
                                "ExactDoc flattened retry output:",
                                retry_result.stdout,
                                flush=True
                            )


                        if retry_result.stderr:

                            print(
                                "ExactDoc flattened retry warnings:",
                                retry_result.stderr,
                                flush=True
                            )


                    except subprocess.CalledProcessError as retry_error:

                        print(
                            "Flattened ExactDoc exit code:",
                            retry_error.returncode,
                            flush=True
                        )

                        print(
                            "Flattened ExactDoc stdout:",
                            retry_error.stdout,
                            flush=True
                        )

                        print(
                            "Flattened ExactDoc stderr:",
                            retry_error.stderr,
                            flush=True
                        )


                        if retry_error.returncode == 17:

                            print(
                                "Flattened PDF requires OCR.",
                                flush=True
                            )


                            ocr_success = create_ocr_pdf(
                                flattened_path,
                                ocr_path
                            )


                            if not ocr_success:

                                return jsonify({
                                    "error":
                                    "OCR could not prepare this PDF "
                                    "for Word conversion."
                                }), 500


                            if os.path.exists(
                                output_path
                            ):

                                os.remove(
                                    output_path
                                )


                            final_result = run_exactdoc(
                                ocr_path,
                                output_path
                            )


                            if final_result.stdout:

                                print(
                                    "OCR ExactDoc output:",
                                    final_result.stdout,
                                    flush=True
                                )


                            if final_result.stderr:

                                print(
                                    "OCR ExactDoc warnings:",
                                    final_result.stderr,
                                    flush=True
                                )


                        elif retry_error.returncode == 19:

                            return jsonify({
                                "error":
                                "This PDF still contains form "
                                "structures that ExactDoc cannot "
                                "convert reliably."
                            }), 422


                        else:

                            raise retry_error


                # =================================
                # OCR REQUIRED
                # =================================

                elif first_error.returncode == 17:

                    print(
                        "PDF requires OCR.",
                        flush=True
                    )


                    ocr_success = create_ocr_pdf(
                        input_path,
                        ocr_path
                    )


                    if not ocr_success:

                        return jsonify({
                            "error":
                            "OCR could not prepare this PDF "
                            "for Word conversion."
                        }), 500


                    if os.path.exists(
                        output_path
                    ):

                        os.remove(
                            output_path
                        )


                    ocr_result = run_exactdoc(
                        ocr_path,
                        output_path
                    )


                    if ocr_result.stdout:

                        print(
                            "OCR ExactDoc output:",
                            ocr_result.stdout,
                            flush=True
                        )


                    if ocr_result.stderr:

                        print(
                            "OCR ExactDoc warnings:",
                            ocr_result.stderr,
                            flush=True
                        )


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
                "Final conversion exit code:",
                error.returncode,
                flush=True
            )

            print(
                "Final conversion stdout:",
                error.stdout,
                flush=True
            )

            print(
                "Final conversion stderr:",
                error.stderr,
                flush=True
            )


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


        except FileNotFoundError as error:

            print(
                "Missing server command:",
                repr(error),
                flush=True
            )

            return jsonify({
                "error":
                "The PDF conversion tool is not "
                "available on the server."
            }), 500


        except Exception as error:

            print(
                "PDF to Word unexpected error:",
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
                "No DOCX output was created.",
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
                "DOCX output was empty.",
                flush=True
            )

            return jsonify({
                "error":
                "The Word document was created "
                "but was empty."
            }), 500


        print(
            "PDF to Word conversion complete:",
            output_name,
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
