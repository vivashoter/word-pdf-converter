from flask import (
    Flask,
    request,
    send_file,
    send_from_directory,
    jsonify
)

from pdf2docx import Converter

from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from werkzeug.utils import secure_filename

import fitz
import os
import re
import subprocess
import tempfile


# =========================================
# BASIC CONFIG
# =========================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

app = Flask(__name__)

MAX_FILE_SIZE = (
    25 *
    1024 *
    1024
)

app.config[
    "MAX_CONTENT_LENGTH"
] = MAX_FILE_SIZE


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


@app.route(
    "/convertdocgoose-logo.png"
)
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


    uploaded_file = (
        request.files["file"]
    )


    if uploaded_file.filename == "":

        return jsonify({
            "error":
            "No file selected."
        }), 400


    filename = secure_filename(
        uploaded_file.filename
    )


    if not filename.lower().endswith(
        (
            ".doc",
            ".docx"
        )
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
                    "LibreOffice:",
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
            +
            ".pdf"
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

            download_name=(
                output_name
            ),

            mimetype=(
                "application/pdf"
            )
        )


# =========================================
# PDF TEXT CHECK
# =========================================

def pdf_has_real_text(
    input_path
):

    document = None


    try:

        document = fitz.open(
            input_path
        )


        if document.page_count == 0:

            return False


        total_characters = 0

        pages_with_text = 0


        for page in document:

            text = page.get_text(
                "text"
            ).strip()


            count = len(
                text
            )


            total_characters += (
                count
            )


            if count >= 30:

                pages_with_text += 1


        average = (
            total_characters /
            document.page_count
        )


        ratio = (
            pages_with_text /
            document.page_count
        )


        print(
            "PDF analysis:",
            {
                "pages":
                    document.page_count,

                "characters":
                    total_characters,

                "average":
                    round(
                        average,
                        2
                    ),

                "text_page_ratio":
                    round(
                        ratio,
                        2
                    )
            }
        )


        return (
            average >= 40
            or
            ratio >= 0.50
        )


    except Exception as error:

        print(
            "PDF analysis error:",
            error
        )

        return False


    finally:

        if document is not None:

            document.close()


# =========================================
# OCR
# =========================================

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
            "OCRmyPDF:",
            result.stdout
        )


    if result.stderr:

        print(
            "OCR warnings:",
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


# =========================================
# FONT HELPERS
# =========================================

def clean_font_name(
    name
):

    if not name:

        return "Arial"


    # Remove embedded PDF subset prefix,
    # for example ABCDEF+Arial
    name = re.sub(
        r"^[A-Z]{6}\+",
        "",
        name
    )


    replacements = {
        "Helvetica":
            "Arial",

        "Helvetica-Bold":
            "Arial",

        "Times-Roman":
            "Times New Roman",

        "Times-Bold":
            "Times New Roman",

        "Times-Italic":
            "Times New Roman",

        "Courier":
            "Courier New"
    }


    return replacements.get(
        name,
        name
    )


def pdf_color_to_hex(
    color_value
):

    try:

        value = int(
            color_value
        )


        red = (
            value >>
            16
        ) & 255


        green = (
            value >>
            8
        ) & 255


        blue = (
            value
        ) & 255


        return (
            f"{red:02X}"
            f"{green:02X}"
            f"{blue:02X}"
        )


    except Exception:

        return "000000"


# =========================================
# EDITABLE WORD TEXTBOX
# =========================================

def add_editable_textbox(
    paragraph,
    text,
    x,
    y,
    width,
    height,
    font_name,
    font_size,
    color,
    bold=False,
    italic=False
):

    if not text:

        return


    # Give text a little extra room.
    # This reduces unwanted Word wrapping.

    width = max(
        width + 3,
        4
    )


    height = max(
        height + 2,
        font_size * 1.2
    )


    pict = OxmlElement(
        "w:pict"
    )


    shape = OxmlElement(
        "v:rect"
    )


    shape.set(
        "style",

        (
            "position:absolute;"
            f"margin-left:{x}pt;"
            f"margin-top:{y}pt;"
            f"width:{width}pt;"
            f"height:{height}pt;"
            "z-index:10;"
            "mso-position-horizontal-relative:page;"
            "mso-position-vertical-relative:page;"
            "mso-wrap-style:none;"
        )
    )


    shape.set(
        "stroked",
        "f"
    )


    shape.set(
        "filled",
        "f"
    )


    textbox = OxmlElement(
        "v:textbox"
    )


    textbox.set(
        "inset",
        "0,0,0,0"
    )


    text_content = OxmlElement(
        "w:txbxContent"
    )


    word_paragraph = OxmlElement(
        "w:p"
    )


    paragraph_properties = (
        OxmlElement(
            "w:pPr"
        )
    )


    spacing = OxmlElement(
        "w:spacing"
    )


    spacing.set(
        qn("w:before"),
        "0"
    )


    spacing.set(
        qn("w:after"),
        "0"
    )


    spacing.set(
        qn("w:line"),
        "200"
    )


    spacing.set(
        qn("w:lineRule"),
        "auto"
    )


    paragraph_properties.append(
        spacing
    )


    word_paragraph.append(
        paragraph_properties
    )


    run = OxmlElement(
        "w:r"
    )


    run_properties = (
        OxmlElement(
            "w:rPr"
        )
    )


    fonts = OxmlElement(
        "w:rFonts"
    )


    fonts.set(
        qn("w:ascii"),
        font_name
    )


    fonts.set(
        qn("w:hAnsi"),
        font_name
    )


    fonts.set(
        qn("w:eastAsia"),
        font_name
    )


    run_properties.append(
        fonts
    )


    size = OxmlElement(
        "w:sz"
    )


    size.set(
        qn("w:val"),
        str(
            max(
                2,
                round(
                    font_size *
                    2
                )
            )
        )
    )


    run_properties.append(
        size
    )


    color_element = (
        OxmlElement(
            "w:color"
        )
    )


    color_element.set(
        qn("w:val"),
        color
    )


    run_properties.append(
        color_element
    )


    if bold:

        run_properties.append(
            OxmlElement(
                "w:b"
            )
        )


    if italic:

        run_properties.append(
            OxmlElement(
                "w:i"
            )
        )


    run.append(
        run_properties
    )


    text_element = OxmlElement(
        "w:t"
    )


    text_element.set(
        qn("xml:space"),
        "preserve"
    )


    text_element.text = text


    run.append(
        text_element
    )


    word_paragraph.append(
        run
    )


    text_content.append(
        word_paragraph
    )


    textbox.append(
        text_content
    )


    shape.append(
        textbox
    )


    pict.append(
        shape
    )


    paragraph._p.append(
        pict
    )


# =========================================
# CREATE TEXT-FREE PDF BACKGROUND
# =========================================

def create_background_pdf(
    input_path,
    background_path
):

    source = fitz.open(
        input_path
    )


    # Save a working copy.
    source.save(
        background_path
    )

    source.close()


    background = fitz.open(
        background_path
    )


    for page in background:

        text_data = page.get_text(
            "dict"
        )


        found_text = False


        for block in text_data[
            "blocks"
        ]:

            if block.get(
                "type"
            ) != 0:

                continue


            for line in block.get(
                "lines",
                []
            ):

                for span in line.get(
                    "spans",
                    []
                ):

                    text = span.get(
                        "text",
                        ""
                    )


                    if not text.strip():

                        continue


                    bbox = span.get(
                        "bbox"
                    )


                    if not bbox:

                        continue


                    rectangle = fitz.Rect(
                        bbox
                    )


                    # Slightly shrink vertically
                    # so nearby form lines stay visible.

                    inset = min(
                        0.5,
                        rectangle.height *
                        0.05
                    )


                    rectangle.y0 += (
                        inset
                    )

                    rectangle.y1 -= (
                        inset
                    )


                    page.add_redact_annot(
                        rectangle,

                        fill=(
                            1,
                            1,
                            1
                        ),

                        cross_out=False
                    )


                    found_text = True


        if found_text:

            # Remove text only.
            # Keep graphics and images.
            page.apply_redactions(
                images=0,
                graphics=0,
                text=0
            )


    temp_output = (
        background_path +
        ".clean.pdf"
    )


    background.save(
        temp_output,

        garbage=4,

        deflate=True
    )


    background.close()


    os.replace(
        temp_output,
        background_path
    )


# =========================================
# HIGH-FIDELITY EDITABLE PDF TO WORD
# =========================================

def convert_positioned_pdf_to_docx(
    input_path,
    output_path,
    temp_dir
):

    source = fitz.open(
        input_path
    )


    background_pdf_path = (
        os.path.join(
            temp_dir,
            "background.pdf"
        )
    )


    create_background_pdf(
        input_path,
        background_pdf_path
    )


    backgrounds = fitz.open(
        background_pdf_path
    )


    word = Document()


    # Use existing first paragraph.
    paragraph = (
        word.paragraphs[0]
    )


    paragraph.paragraph_format.space_before = (
        Pt(0)
    )

    paragraph.paragraph_format.space_after = (
        Pt(0)
    )


    for page_number in range(
        source.page_count
    ):

        source_page = (
            source.load_page(
                page_number
            )
        )


        background_page = (
            backgrounds.load_page(
                page_number
            )
        )


        page_width = (
            source_page.rect.width
        )


        page_height = (
            source_page.rect.height
        )


        # ---------------------------------
        # PAGE SECTION
        # ---------------------------------

        if page_number > 0:

            word.add_section(
                WD_SECTION.NEW_PAGE
            )


            paragraph = (
                word.add_paragraph()
            )


            paragraph.paragraph_format.space_before = (
                Pt(0)
            )

            paragraph.paragraph_format.space_after = (
                Pt(0)
            )


        section = (
            word.sections[-1]
        )


        section.page_width = (
            Pt(
                page_width
            )
        )


        section.page_height = (
            Pt(
                page_height
            )
        )


        section.top_margin = Pt(0)
        section.bottom_margin = Pt(0)

        section.left_margin = Pt(0)
        section.right_margin = Pt(0)

        section.header_distance = Pt(0)
        section.footer_distance = Pt(0)


        # ---------------------------------
        # RENDER TEXT-FREE BACKGROUND
        # ---------------------------------

        image_path = os.path.join(
            temp_dir,
            (
                f"page_"
                f"{page_number + 1}.png"
            )
        )


        pixmap = (
            background_page.get_pixmap(
                matrix=fitz.Matrix(
                    2,
                    2
                ),

                alpha=False
            )
        )


        pixmap.save(
            image_path
        )


        run = paragraph.add_run()


        # Leave a tiny amount of room
        # for Word's paragraph mark.

        run.add_picture(
            image_path,

            width=Pt(
                page_width
            ),

            height=Pt(
                max(
                    1,
                    page_height -
                    1
                )
            )
        )


        # ---------------------------------
        # EXTRACT ORIGINAL EDITABLE TEXT
        # ---------------------------------

        page_data = source_page.get_text(
            "dict"
        )


        for block in page_data[
            "blocks"
        ]:

            if block.get(
                "type"
            ) != 0:

                continue


            for line in block.get(
                "lines",
                []
            ):

                for span in line.get(
                    "spans",
                    []
                ):

                    text = span.get(
                        "text",
                        ""
                    )


                    if not text:

                        continue


                    bbox = span.get(
                        "bbox"
                    )


                    if not bbox:

                        continue


                    x0 = bbox[0]

                    y0 = bbox[1]

                    x1 = bbox[2]

                    y1 = bbox[3]


                    width = (
                        x1 -
                        x0
                    )


                    height = (
                        y1 -
                        y0
                    )


                    size = float(
                        span.get(
                            "size",
                            10
                        )
                    )


                    raw_font = span.get(
                        "font",
                        "Arial"
                    )


                    font_name = (
                        clean_font_name(
                            raw_font
                        )
                    )


                    font_lower = (
                        raw_font.lower()
                    )


                    bold = (
                        "bold"
                        in
                        font_lower
                    )


                    italic = (
                        "italic"
                        in
                        font_lower
                        or
                        "oblique"
                        in
                        font_lower
                    )


                    color = (
                        pdf_color_to_hex(
                            span.get(
                                "color",
                                0
                            )
                        )
                    )


                    # Word's text baseline differs
                    # slightly from PDF positioning.

                    y_position = max(
                        0,
                        y0 -
                        0.5
                    )


                    add_editable_textbox(
                        paragraph=paragraph,

                        text=text,

                        x=x0,

                        y=y_position,

                        width=width,

                        height=height,

                        font_name=font_name,

                        font_size=size,

                        color=color,

                        bold=bold,

                        italic=italic
                    )


    source.close()

    backgrounds.close()


    word.save(
        output_path
    )


# =========================================
# SCANNED PDF FALLBACK
# =========================================

def convert_scanned_pdf_to_docx(
    input_path,
    output_path,
    temp_dir
):

    ocr_path = os.path.join(
        temp_dir,
        "ocr.pdf"
    )


    pdf_to_convert = (
        input_path
    )


    try:

        if create_ocr_pdf(
            input_path,
            ocr_path
        ):

            pdf_to_convert = (
                ocr_path
            )


    except Exception as error:

        print(
            "OCR fallback error:",
            error
        )


    converter = None


    try:

        converter = Converter(
            pdf_to_convert
        )


        converter.convert(
            output_path,
            start=0,
            end=None
        )


    finally:

        if converter is not None:

            converter.close()


# =========================================
# PDF TO WORD ROUTE
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


    uploaded_file = (
        request.files["file"]
    )


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


        base_name = (
            os.path.splitext(
                filename
            )[0]
        )


        output_name = (
            base_name +
            ".docx"
        )


        output_path = os.path.join(
            temp_dir,
            output_name
        )


        try:

            if pdf_has_real_text(
                input_path
            ):

                print(
                    "Digital PDF detected."
                )

                print(
                    "Using positioned editable "
                    "layout conversion."
                )


                convert_positioned_pdf_to_docx(
                    input_path,
                    output_path,
                    temp_dir
                )


            else:

                print(
                    "Scanned PDF detected."
                )

                print(
                    "Using OCR fallback."
                )


                convert_scanned_pdf_to_docx(
                    input_path,
                    output_path,
                    temp_dir
                )


        except Exception as error:

            print(
                "PDF to Word error:",
                repr(
                    error
                )
            )


            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500


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
            "Created:",
            output_name
        )


        return send_file(
            output_path,

            as_attachment=True,

            download_name=(
                output_name
            ),

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
