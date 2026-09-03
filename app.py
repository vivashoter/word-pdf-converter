from flask import (
    Flask,
    request,
    send_file,
    send_from_directory,
    jsonify
)

from werkzeug.utils import secure_filename

from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml

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
# EXACTDOC
# NORMAL PDF CONVERSION
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
# INTERACTIVE FORM DETECTION
# =========================================

def pdf_has_interactive_form(
    input_path
):

    document = None

    try:

        document = fitz.open(
            input_path
        )

        widget_count = 0

        for page in document:

            widgets = page.widgets()

            if widgets:

                widget_count += len(
                    list(widgets)
                )

        print(
            "Interactive widget count:",
            widget_count,
            flush=True
        )

        return (
            widget_count > 0
        )

    except Exception as error:

        print(
            "Form detection error:",
            repr(error),
            flush=True
        )

        return False

    finally:

        if document is not None:

            document.close()


# =========================================
# FONT HELPERS
# =========================================

def clean_font_name(
    font_name
):

    if not font_name:

        return "Arial"

    font_name = re.sub(
        r"^[A-Z]{6}\+",
        "",
        font_name
    )

    lower_name = (
        font_name.lower()
    )

    if "helvetica" in lower_name:

        return "Arial"

    if "times" in lower_name:

        return "Times New Roman"

    if "courier" in lower_name:

        return "Courier New"

    return font_name


def pdf_color_to_hex(
    value
):

    try:

        value = int(
            value
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
# WORD TEXT BOX
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
    color="000000",
    bold=False,
    italic=False
):

    if not text:

        return

    text = str(
        text
    )

    if not text.strip():

        return

    safe_text = escape(
        text
    )

    safe_font = escape(
        str(font_name),
        {
            '"': "&quot;"
        }
    )

    safe_color = (
        str(color)
        .replace(
            '"',
            ""
        )
    )

    width = max(
        float(width) + 3,
        5
    )

    height = max(
        float(height) + 2,
        float(font_size) * 1.25
    )

    font_half_points = max(
        2,
        round(
            float(font_size) *
            2
        )
    )

    bold_xml = (
        "<w:b/>"
        if bold
        else ""
    )

    italic_xml = (
        "<w:i/>"
        if italic
        else ""
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
                inset="0,0,0,0">

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
                                    w:val="{font_half_points}"
                                />

                                <w:szCs
                                    w:val="{font_half_points}"
                                />

                                <w:color
                                    w:val="{safe_color}"
                                />

                                {bold_xml}

                                {italic_xml}

                            </w:rPr>

                            <w:t xml:space="preserve">{safe_text}</w:t>

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
# BUILD FORM BACKGROUND IMAGE
# =========================================

def create_form_background(
    document,
    page_number,
    temp_dir
):

    original_page = (
        document.load_page(
            page_number
        )
    )

    page_width = (
        original_page.rect.width
    )

    page_height = (
        original_page.rect.height
    )

    # Render the page including form fields.
    original_pixmap = (
        original_page.get_pixmap(
            matrix=fitz.Matrix(
                2,
                2
            ),
            alpha=False,
            annots=True
        )
    )

    original_image_path = (
        os.path.join(
            temp_dir,
            f"form_original_{page_number}.png"
        )
    )

    original_pixmap.save(
        original_image_path
    )

    # Put rendered page into temporary PDF.
    background_pdf = fitz.open()

    background_page = (
        background_pdf.new_page(
            width=page_width,
            height=page_height
        )
    )

    background_page.insert_image(
        background_page.rect,
        filename=original_image_path
    )

    # Extract static PDF text.
    page_data = (
        original_page.get_text(
            "dict"
        )
    )

    # Cover original static text with white.
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

                if not text.strip():

                    continue

                bbox = span.get(
                    "bbox"
                )

                if not bbox:

                    continue

                rect = fitz.Rect(
                    bbox
                )

                rect.x0 -= 0.5
                rect.x1 += 0.5

                rect.y0 -= 0.5
                rect.y1 += 0.5

                background_page.draw_rect(
                    rect,
                    color=None,
                    fill=(
                        1,
                        1,
                        1
                    ),
                    overlay=True
                )

    # Render cleaned visual background.
    background_pixmap = (
        background_page.get_pixmap(
            matrix=fitz.Matrix(
                2,
                2
            ),
            alpha=False
        )
    )

    background_image_path = (
        os.path.join(
            temp_dir,
            f"form_background_{page_number}.png"
        )
    )

    background_pixmap.save(
        background_image_path
    )

    background_pdf.close()

    return (
        background_image_path,
        page_width,
        page_height,
        page_data
    )


# =========================================
# INTERACTIVE FORM → EDITABLE DOCX
# =========================================

def convert_interactive_form_to_docx(
    input_path,
    output_path,
    temp_dir
):

    print(
        "Starting dedicated interactive "
        "form conversion...",
        flush=True
    )

    pdf = fitz.open(
        input_path
    )

    word = Document()

    # =====================================
    # IMPORTANT FIX:
    #
    # A new python-docx Document can contain
    # ZERO paragraphs.
    #
    # We create one explicitly instead of
    # trying word.paragraphs[0].
    # =====================================

    first_paragraph = (
        word.add_paragraph()
    )

    print(
        "Word document initialized.",
        flush=True
    )

    for page_number in range(
        pdf.page_count
    ):

        print(
            f"Building Word page "
            f"{page_number + 1} "
            f"of {pdf.page_count}...",
            flush=True
        )

        (
            background_path,
            page_width,
            page_height,
            page_data
        ) = create_form_background(
            pdf,
            page_number,
            temp_dir
        )

        print(
            "Form background created.",
            flush=True
        )

        # =================================
        # PAGE SECTION
        # =================================

        if page_number == 0:

            section = (
                word.sections[0]
            )

            paragraph = (
                first_paragraph
            )

        else:

            section = word.add_section(
                WD_SECTION.NEW_PAGE
            )

            paragraph = (
                word.add_paragraph()
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

        paragraph.paragraph_format.space_before = (
            Pt(0)
        )

        paragraph.paragraph_format.space_after = (
            Pt(0)
        )

        paragraph.paragraph_format.line_spacing = (
            Pt(1)
        )

        # =================================
        # FORM BACKGROUND
        # =================================

        background_run = (
            paragraph.add_run()
        )

        background_run.add_picture(
            background_path,
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

        print(
            "Form background inserted into Word.",
            flush=True
        )

        # =================================
        # EDITABLE STATIC PDF TEXT
        # =================================

        text_span_count = 0

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

                    if not text.strip():

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
                        in lower_font
                    )

                    italic = (
                        "italic"
                        in lower_font
                        or
                        "oblique"
                        in lower_font
                    )

                    color = (
                        pdf_color_to_hex(
                            span.get(
                                "color",
                                0
                            )
                        )
                    )

                    add_editable_textbox(
                        paragraph=paragraph,
                        text=text,
                        x=x0,
                        y=max(
                            0,
                            y0 - 0.7
                        ),
                        width=width,
                        height=height,
                        font_name=font_name,
                        font_size=font_size,
                        color=color,
                        bold=bold,
                        italic=italic
                    )

                    text_span_count += 1

        print(
            "Editable static text spans added:",
            text_span_count,
            flush=True
        )

        # =================================
        # FORM FIELD VALUES
        # =================================

        source_page = (
            pdf.load_page(
                page_number
            )
        )

        widgets = (
            source_page.widgets()
        )

        field_count = 0

        if widgets:

            for widget in widgets:

                field_count += 1

                field_type = (
                    widget.field_type_string
                    or
                    ""
                ).lower()

                field_value = (
                    widget.field_value
                    or
                    ""
                )

                field_rect = (
                    widget.rect
                )

                if (
                    "text"
                    in field_type
                    and
                    str(
                        field_value
                    ).strip()
                ):

                    add_editable_textbox(
                        paragraph=paragraph,
                        text=str(
                            field_value
                        ),
                        x=field_rect.x0 + 2,
                        y=field_rect.y0 + 1,
                        width=max(
                            10,
                            field_rect.width - 4
                        ),
                        height=max(
                            10,
                            field_rect.height - 2
                        ),
                        font_name="Arial",
                        font_size=max(
                            8,
                            min(
                                11,
                                field_rect.height *
                                0.60
                            )
                        ),
                        color="000000"
                    )

        print(
            "PDF form fields processed:",
            field_count,
            flush=True
        )

    pdf.close()

    print(
        "Saving interactive-form DOCX...",
        flush=True
    )

    word.save(
        output_path
    )

    print(
        "Interactive form DOCX created:",
        output_path,
        flush=True
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

    uploaded_file = request.files[
        "file"
    ]

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
            +
            ".docx"
        )

        output_path = os.path.join(
            temp_dir,
            output_name
        )

        try:

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

            is_interactive_form = (
                pdf_has_interactive_form(
                    input_path
                )
            )

            if is_interactive_form:

                print(
                    "Interactive PDF detected.",
                    flush=True
                )

                print(
                    "Using dedicated fixed-layout "
                    "form converter.",
                    flush=True
                )

                convert_interactive_form_to_docx(
                    input_path,
                    output_path,
                    temp_dir
                )

            else:

                print(
                    "Normal PDF detected.",
                    flush=True
                )

                print(
                    "Using ExactDoc.",
                    flush=True
                )

                run_exactdoc(
                    input_path,
                    output_path
                )

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
                "PDF converter exit code:",
                error.returncode,
                flush=True
            )

            print(
                "PDF converter stdout:",
                error.stdout,
                flush=True
            )

            print(
                "PDF converter stderr:",
                error.stderr,
                flush=True
            )

            return jsonify({
                "error":
                "The PDF could not be converted to Word."
            }), 500

        except Exception as error:

            print(
                "PDF to Word error:",
                repr(error),
                flush=True
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
