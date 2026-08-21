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
from docx.oxml import parse_xml

from werkzeug.utils import secure_filename

from xml.sax.saxutils import escape

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
                repr(error)
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
# DETECT DIGITAL VS SCANNED PDF
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


            character_count = len(
                text
            )


            total_characters += (
                character_count
            )


            if character_count >= 30:

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
            "PDF analysis:",
            {
                "pages":
                    document.page_count,

                "characters":
                    total_characters,

                "average":
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


        return (
            average_characters >= 40
            or
            text_page_ratio >= 0.50
        )


    except Exception as error:

        print(
            "PDF analysis error:",
            repr(error)
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

        "Helvetica-Oblique":
            "Arial",

        "Helvetica-BoldOblique":
            "Arial",

        "Times-Roman":
            "Times New Roman",

        "Times-Bold":
            "Times New Roman",

        "Times-Italic":
            "Times New Roman",

        "Times-BoldItalic":
            "Times New Roman",

        "Courier":
            "Courier New",

        "Courier-Bold":
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
# EDITABLE POSITIONED TEXTBOX
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


    width = max(
        width + 5,
        8
    )


    height = max(
        height + 4,
        font_size * 1.45
    )


    safe_text = escape(
        str(text)
    )


    safe_font = escape(
        str(font_name),
        {
            '"':
                "&quot;"
        }
    )


    safe_color = str(
        color
    ).replace(
        '"',
        ""
    )


    bold_xml = ""

    italic_xml = ""


    if bold:

        bold_xml = "<w:b/>"


    if italic:

        italic_xml = "<w:i/>"


    font_size_half_points = max(
        2,
        round(
            font_size *
            2
        )
    )


    textbox_xml = f"""
    <w:pict
        xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:o="urn:schemas-microsoft-com:office:office">

        <v:rect
            stroked="f"
            filled="f"
            style="
                position:absolute;
                margin-left:{x}pt;
                margin-top:{y}pt;
                width:{width}pt;
                height:{height}pt;
                z-index:10;
                mso-position-horizontal-relative:page;
                mso-position-vertical-relative:page;
                mso-wrap-distance-left:0;
                mso-wrap-distance-right:0;
                mso-wrap-distance-top:0;
                mso-wrap-distance-bottom:0;
            ">

            <v:textbox
                inset="0,0,0,0"
            >

                <w:txbxContent>

                    <w:p>

                        <w:pPr>

                            <w:spacing
                                w:before="0"
                                w:after="0"
                                w:line="200"
                                w:lineRule="auto"
                            />

                        </w:pPr>

                        <w:r>

                            <w:rPr>

                                <w:rFonts
                                    w:ascii="{safe_font}"
                                    w:hAnsi="{safe_font}"
                                    w:eastAsia="{safe_font}"
                                />

                                <w:sz
                                    w:val="{font_size_half_points}"
                                />

                                <w:szCs
                                    w:val="{font_size_half_points}"
                                />

                                <w:color
                                    w:val="{safe_color}"
                                />

                                {bold_xml}

                                {italic_xml}

                            </w:rPr>

                            <w:t
                                xml:space="preserve"
                            >{safe_text}</w:t>

                        </w:r>

                    </w:p>

                </w:txbxContent>

            </v:textbox>

        </v:rect>

    </w:pict>
    """


    pict = parse_xml(
        textbox_xml
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


        for block in text_data.get(
            "blocks",
            []
        ):

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


                    # Slightly shrink the redaction
                    # so nearby form lines remain visible.

                    vertical_inset = min(
                        0.35,
                        rectangle.height *
                        0.04
                    )


                    horizontal_inset = min(
                        0.20,
                        rectangle.width *
                        0.01
                    )


                    rectangle.x0 += (
                        horizontal_inset
                    )


                    rectangle.x1 -= (
                        horizontal_inset
                    )


                    rectangle.y0 += (
                        vertical_inset
                    )


                    rectangle.y1 -= (
                        vertical_inset
                    )


                    if (
                        rectangle.width <= 0
                        or
                        rectangle.height <= 0
                    ):

                        continue


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

            page.apply_redactions(
                images=0,
                graphics=0
            )


    cleaned_path = (
        background_path +
        ".clean.pdf"
    )


    background.save(
        cleaned_path,

        garbage=4,

        deflate=True
    )


    background.close()


    os.replace(
        cleaned_path,
        background_path
    )


# =========================================
# POSITIONED EDITABLE PDF -> DOCX
# =========================================

def convert_positioned_pdf_to_docx(
    input_path,
    output_path,
    temp_dir
):

    source = fitz.open(
        input_path
    )


    background_pdf_path = os.path.join(
        temp_dir,
        "background.pdf"
    )


    create_background_pdf(
        input_path,
        background_pdf_path
    )


    backgrounds = fitz.open(
        background_pdf_path
    )


    word = Document()


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


        page_width = float(
            source_page.rect.width
        )


        page_height = float(
            source_page.rect.height
        )


        # ---------------------------------
        # SECTION
        # ---------------------------------

        if page_number == 0:

            section = (
                word.sections[0]
            )

        else:

            section = (
                word.add_section(
                    WD_SECTION.NEW_PAGE
                )
            )


        section.page_width = Pt(
            page_width
        )


        section.page_height = Pt(
            page_height
        )


        section.top_margin = Pt(0)

        section.bottom_margin = Pt(0)

        section.left_margin = Pt(0)

        section.right_margin = Pt(0)

        section.header_distance = Pt(0)

        section.footer_distance = Pt(0)


        # ---------------------------------
        # CREATE PAGE PARAGRAPH
        # ---------------------------------

        paragraph = (
            word.add_paragraph()
        )


        paragraph.paragraph_format.space_before = (
            Pt(0)
        )


        paragraph.paragraph_format.space_after = (
            Pt(0)
        )


        paragraph.paragraph_format.line_spacing = (
            1
        )


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


        background_run = (
            paragraph.add_run()
        )


        background_run.add_picture(
            image_path,

            width=Pt(
                page_width
            ),

            height=Pt(
                max(
                    1,
                    page_height - 2
                )
            )
        )


        # ---------------------------------
        # ORIGINAL EDITABLE TEXT
        # ---------------------------------

        page_data = source_page.get_text(
            "dict"
        )


        for block in page_data.get(
            "blocks",
            []
        ):

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


                    x0 = float(
                        bbox[0]
                    )


                    y0 = float(
                        bbox[1]
                    )


                    x1 = float(
                        bbox[2]
                    )


                    y1 = float(
                        bbox[3]
                    )


                    width = max(
                        1,
                        x1 - x0
                    )


                    height = max(
                        1,
                        y1 - y0
                    )


                    font_size = float(
                        span.get(
                            "size",
                            10
                        )
                    )


                    raw_font = str(
                        span.get(
                            "font",
                            "Arial"
                        )
                    )


                    font_name = (
                        clean_font_name(
                            raw_font
                        )
                    )


                    lower_font = (
                        raw_font.lower()
                    )


                    bold = (
                        "bold"
                        in
                        lower_font
                    )


                    italic = (
                        "italic"
                        in
                        lower_font
                        or
                        "oblique"
                        in
                        lower_font
                    )


                    color = (
                        pdf_color_to_hex(
                            span.get(
                                "color",
                                0
                            )
                        )
                    )


                    # PDF and Word text origins differ
                    # slightly. This compensates for it.

                    y_position = max(
                        0,
                        y0 - 1.0
                    )


                    add_editable_textbox(
                        paragraph=paragraph,

                        text=text,

                        x=x0,

                        y=y_position,

                        width=width,

                        height=height,

                        font_name=font_name,

                        font_size=font_size,

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

        ocr_success = create_ocr_pdf(
            input_path,
            ocr_path
        )


        if ocr_success:

            pdf_to_convert = (
                ocr_path
            )


    except subprocess.TimeoutExpired:

        print(
            "OCR timed out. "
            "Using original PDF."
        )


    except subprocess.CalledProcessError as error:

        print(
            "OCR failed:",
            error.stderr
        )


    except Exception as error:

        print(
            "OCR fallback error:",
            repr(error)
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

            try:

                converter.close()

            except Exception:

                pass


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
                repr(error)
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
            "PDF to Word created:",
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
