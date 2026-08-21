from flask import (
    Flask,
    request,
    send_file,
    send_from_directory,
    jsonify
)

from pdf2docx import Converter
from werkzeug.utils import secure_filename

import fitz
import os
import subprocess
import tempfile


# ---------------------------------
# BASE DIRECTORY
# ---------------------------------

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)


# ---------------------------------
# APP
# ---------------------------------

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
# FAVICON
# ---------------------------------

@app.route("/favicon.png")
def favicon():

    return send_from_directory(
        BASE_DIR,
        "favicon.png"
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


# =================================
# WORD TO PDF
# =================================

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


# =================================
# PDF ANALYSIS
# =================================

def pdf_needs_ocr(
    input_path
):

    """
    Decide whether a PDF is primarily scanned/image-based.

    Digital PDFs already containing usable text should NOT
    be OCRed because OCR can damage the original positioning
    and make PDF-to-DOCX layout reconstruction worse.
    """

    document = None


    try:

        document = fitz.open(
            input_path
        )


        if document.page_count == 0:

            return False


        total_characters = 0

        pages_with_text = 0


        for page_number in range(
            document.page_count
        ):

            page = document.load_page(
                page_number
            )


            text = page.get_text(
                "text"
            ).strip()


            character_count = len(
                text
            )


            total_characters += (
                character_count
            )


            if character_count >= 20:

                pages_with_text += 1


        average_characters = (
            total_characters /
            document.page_count
        )


        text_page_ratio = (
            pages_with_text /
            document.page_count
        )


        print(
            "PDF text analysis:",
            {
                "pages":
                    document.page_count,

                "characters":
                    total_characters,

                "average_characters":
                    round(
                        average_characters,
                        2
                    ),

                "text_page_ratio":
                    round(
                        text_page_ratio,
                        2
                    )
            }
        )


        # If most pages contain real text,
        # preserve the original PDF.
        if (
            average_characters >= 50
            or
            text_page_ratio >= 0.50
        ):

            return False


        # Very little extracted text means
        # this is probably scanned.
        return True


    except Exception as error:

        print(
            "PDF text analysis error:",
            error
        )


        # Safest fallback:
        # do NOT modify the original PDF.
        return False


    finally:

        if document is not None:

            document.close()


# =================================
# OCR PDF
# =================================

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


# =================================
# PDF TO WORD
# EDITABLE MODE
# =================================

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


        ocr_pdf_path = os.path.join(
            temp_dir,
            base_name
            + "_ocr.pdf"
        )


        pdf_for_conversion = (
            input_path
        )


        # ---------------------------------
        # DETECT WHETHER OCR IS NEEDED
        # ---------------------------------

        try:

            needs_ocr = pdf_needs_ocr(
                input_path
            )


        except Exception as error:

            print(
                "OCR detection error:",
                error
            )

            needs_ocr = False


        # ---------------------------------
        # OCR ONLY SCANNED PDFs
        # ---------------------------------

        if needs_ocr:

            print(
                "Scanned PDF detected. "
                "Running OCR before Word conversion."
            )


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


                else:

                    print(
                        "OCR did not create a usable PDF. "
                        "Using original PDF."
                    )


            except subprocess.TimeoutExpired:

                print(
                    "OCR timed out. "
                    "Using original PDF."
                )


            except subprocess.CalledProcessError as error:

                print(
                    "OCRmyPDF could not process this PDF."
                )


                if error.stderr:

                    print(
                        error.stderr
                    )


                print(
                    "Using original PDF."
                )


            except Exception as error:

                print(
                    "OCR error:",
                    error
                )


                print(
                    "Using original PDF."
                )


        else:

            print(
                "Digital/selectable-text PDF detected. "
                "Skipping OCR to preserve layout."
            )


        # ---------------------------------
        # PDF TO DOCX
        # ---------------------------------

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

                try:

                    converter.close()

                except Exception as error:

                    print(
                        "Converter close warning:",
                        error
                    )


        # ---------------------------------
        # VERIFY OUTPUT
        # ---------------------------------

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


# =================================
# START SERVER
# =================================

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
