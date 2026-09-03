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
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import fitz
import os
import re
import subprocess
import tempfile


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

app = Flask(__name__)

MAX_FILE_SIZE = 25 * 1024 * 1024

app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE


# ============================================================
# STATIC WEBSITE ROUTES
# ============================================================

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


# ============================================================
# FILE SIZE ERROR
# ============================================================

@app.errorhandler(413)
def file_too_large(error):

    return jsonify({
        "error":
        "File is too large. Maximum size is 25 MB."
    }), 413


# ============================================================
# WORD → PDF
# ============================================================

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
                    "LibreOffice:",
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


# ============================================================
# EXACTDOC
# NORMAL PDF → WORD
# ============================================================

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


# ============================================================
# INTERACTIVE FORM DETECTION
# ============================================================

def pdf_has_interactive_form(
    input_path
):

    pdf = None


    try:

        pdf = fitz.open(
            input_path
        )


        widget_count = 0


        for page in pdf:

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

        if pdf is not None:

            pdf.close()


# ============================================================
# WORD XML HELPERS
# ============================================================

def set_cell_border(
    cell,
    top=None,
    bottom=None,
    left=None,
    right=None
):

    tc = cell._tc

    tcPr = tc.get_or_add_tcPr()


    tcBorders = tcPr.first_child_found_in(
        "w:tcBorders"
    )


    if tcBorders is None:

        tcBorders = OxmlElement(
            "w:tcBorders"
        )

        tcPr.append(
            tcBorders
        )


    border_map = {
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right
    }


    for edge_name, edge in border_map.items():

        if edge is None:

            continue


        tag = "w:" + edge_name

        element = tcBorders.find(
            qn(tag)
        )


        if element is None:

            element = OxmlElement(
                tag
            )

            tcBorders.append(
                element
            )


        element.set(
            qn("w:val"),
            edge.get(
                "val",
                "single"
            )
        )

        element.set(
            qn("w:sz"),
            str(
                edge.get(
                    "sz",
                    6
                )
            )
        )

        element.set(
            qn("w:color"),
            edge.get(
                "color",
                "000000"
            )
        )


def remove_all_cell_borders(
    cell
):

    no_border = {
        "val": "nil",
        "sz": 0,
        "color": "FFFFFF"
    }


    set_cell_border(
        cell,
        top=no_border,
        bottom=no_border,
        left=no_border,
        right=no_border
    )


def set_cell_margins(
    cell,
    top=0,
    start=40,
    bottom=0,
    end=40
):

    tc = cell._tc

    tcPr = tc.get_or_add_tcPr()


    tcMar = tcPr.first_child_found_in(
        "w:tcMar"
    )


    if tcMar is None:

        tcMar = OxmlElement(
            "w:tcMar"
        )

        tcPr.append(
            tcMar
        )


    values = {
        "top": top,
        "start": start,
        "bottom": bottom,
        "end": end
    }


    for margin_name, value in values.items():

        node = tcMar.find(
            qn(
                "w:" + margin_name
            )
        )


        if node is None:

            node = OxmlElement(
                "w:" + margin_name
            )

            tcMar.append(
                node
            )


        node.set(
            qn("w:w"),
            str(value)
        )

        node.set(
            qn("w:type"),
            "dxa"
        )


def set_table_fixed_layout(
    table
):

    tblPr = table._tbl.tblPr


    tblLayout = tblPr.first_child_found_in(
        "w:tblLayout"
    )


    if tblLayout is None:

        tblLayout = OxmlElement(
            "w:tblLayout"
        )

        tblPr.append(
            tblLayout
        )


    tblLayout.set(
        qn("w:type"),
        "fixed"
    )


def set_repeat_table_header(
    row
):

    trPr = row._tr.get_or_add_trPr()

    tblHeader = OxmlElement(
        "w:tblHeader"
    )

    tblHeader.set(
        qn("w:val"),
        "true"
    )

    trPr.append(
        tblHeader
    )


# ============================================================
# FONT NAME CLEANER
# ============================================================

def clean_font_name(
    font_name
):

    if not font_name:

        return "Arial"


    font_name = re.sub(
        r"^[A-Z]{6}\+",
        "",
        str(font_name)
    )


    name_lower = font_name.lower()


    if "helvetica" in name_lower:

        return "Arial"


    if "times" in name_lower:

        return "Times New Roman"


    if "courier" in name_lower:

        return "Courier New"


    return font_name


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_page_text_items(
    page
):

    data = page.get_text(
        "dict"
    )


    items = []


    for block in data.get(
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

            line_spans = line.get(
                "spans",
                []
            )


            for span in line_spans:

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


                font_name = clean_font_name(
                    span.get(
                        "font",
                        "Arial"
                    )
                )


                raw_font = str(
                    span.get(
                        "font",
                        ""
                    )
                ).lower()


                items.append({
                    "text":
                        text,

                    "x0":
                        float(
                            bbox[0]
                        ),

                    "y0":
                        float(
                            bbox[1]
                        ),

                    "x1":
                        float(
                            bbox[2]
                        ),

                    "y1":
                        float(
                            bbox[3]
                        ),

                    "font_size":
                        float(
                            span.get(
                                "size",
                                9
                            )
                        ),

                    "font_name":
                        font_name,

                    "bold":
                        (
                            "bold"
                            in raw_font
                            or
                            "black"
                            in raw_font
                            or
                            "heavy"
                            in raw_font
                        ),

                    "italic":
                        (
                            "italic"
                            in raw_font
                            or
                            "oblique"
                            in raw_font
                        )
                })


    return items


# ============================================================
# GROUP TEXT INTO VISUAL ROWS
# ============================================================

def group_items_into_rows(
    items,
    tolerance=4.0
):

    items = sorted(
        items,
        key=lambda item: (
            item["y0"],
            item["x0"]
        )
    )


    rows = []


    for item in items:

        found_row = None


        for row in rows:

            if abs(
                row["y"] -
                item["y0"]
            ) <= tolerance:

                found_row = row

                break


        if found_row is None:

            found_row = {
                "y":
                    item["y0"],

                "items":
                    []
            }

            rows.append(
                found_row
            )


        found_row[
            "items"
        ].append(
            item
        )


    for row in rows:

        row["items"] = sorted(
            row["items"],
            key=lambda item:
                item["x0"]
        )


    return sorted(
        rows,
        key=lambda row:
            row["y"]
    )


# ============================================================
# PDF FORM WIDGET EXTRACTION
# ============================================================

def extract_widget_items(
    page
):

    result = []


    widgets = page.widgets()


    if not widgets:

        return result


    for widget in widgets:

        rect = widget.rect


        field_type = (
            widget.field_type_string
            or
            ""
        ).lower()


        value = (
            widget.field_value
            or
            ""
        )


        item = {
            "x0":
                float(
                    rect.x0
                ),

            "y0":
                float(
                    rect.y0
                ),

            "x1":
                float(
                    rect.x1
                ),

            "y1":
                float(
                    rect.y1
                ),

            "type":
                field_type,

            "value":
                str(
                    value
                ),

            "name":
                str(
                    widget.field_name
                    or
                    ""
                )
        }


        result.append(
            item
        )


    return result


# ============================================================
# DETERMINE COLUMN
#
# We use 3 main page regions for form-style PDFs.
#
# This works much better for structured court forms such
# as LA-350 than trying to make floating Word text boxes.
# ============================================================

def determine_column(
    x_center,
    page_width
):

    first_boundary = (
        page_width
        * 0.34
    )


    second_boundary = (
        page_width
        * 0.67
    )


    if x_center < first_boundary:

        return 0


    if x_center < second_boundary:

        return 1


    return 2


# ============================================================
# ADD PDF TEXT INTO A NATIVE WORD CELL
# ============================================================

def add_text_item_to_cell(
    cell,
    item,
    add_space=True
):

    paragraph = cell.paragraphs[
        0
    ]


    paragraph.paragraph_format.space_before = Pt(
        0
    )

    paragraph.paragraph_format.space_after = Pt(
        0
    )


    if (
        add_space
        and
        paragraph.text
    ):

        paragraph.add_run(
            " "
        )


    run = paragraph.add_run(
        item["text"]
    )


    run.font.name = item[
        "font_name"
    ]


    run.font.size = Pt(
        max(
            6,
            min(
                item[
                    "font_size"
                ],
                14
            )
        )
    )


    run.bold = item[
        "bold"
    ]


    run.italic = item[
        "italic"
    ]


# ============================================================
# ADD CHECKBOX / FORM FIELD INTO CELL
# ============================================================

def add_widget_to_cell(
    cell,
    widget
):

    paragraph = cell.paragraphs[
        0
    ]


    if paragraph.text:

        paragraph.add_run(
            " "
        )


    field_type = widget[
        "type"
    ]


    value = widget[
        "value"
    ].strip()


    # --------------------------------------
    # Checkbox
    # --------------------------------------

    if "check" in field_type:

        checked_values = {
            "yes",
            "on",
            "1",
            "true",
            "checked"
        }


        checked = (
            value.lower()
            in checked_values
        )


        symbol = (
            "☒"
            if checked
            else
            "☐"
        )


        run = paragraph.add_run(
            symbol
        )

        run.font.name = "Arial"

        run.font.size = Pt(
            9
        )

        return


    # --------------------------------------
    # Radio
    # --------------------------------------

    if "radio" in field_type:

        selected = bool(
            value
            and
            value.lower()
            not in {
                "off",
                "0",
                "false"
            }
        )


        symbol = (
            "●"
            if selected
            else
            "○"
        )


        run = paragraph.add_run(
            symbol
        )

        run.font.name = "Arial"

        run.font.size = Pt(
            8
        )

        return


    # --------------------------------------
    # Text / Combo / List
    # --------------------------------------

    if (
        "text" in field_type
        or
        "combo" in field_type
        or
        "list" in field_type
    ):

        if value:

            run = paragraph.add_run(
                value
            )

            run.font.name = "Arial"

            run.font.size = Pt(
                9
            )


# ============================================================
# BUILD PAGE AS NATIVE WORD TABLE
# ============================================================

def add_native_form_page(
    document,
    page,
    page_number
):

    page_width = float(
        page.rect.width
    )


    page_height = float(
        page.rect.height
    )


    print(
        f"Native form page {page_number + 1}: "
        f"{page_width:.1f} x {page_height:.1f}",
        flush=True
    )


    text_items = extract_page_text_items(
        page
    )


    widgets = extract_widget_items(
        page
    )


    print(
        "Text items extracted:",
        len(text_items),
        flush=True
    )


    print(
        "Widget items extracted:",
        len(widgets),
        flush=True
    )


    rows = group_items_into_rows(
        text_items,
        tolerance=4.0
    )


    # Add widget positions as row anchors too.
    for widget in widgets:

        target = None


        for row in rows:

            if abs(
                row["y"] -
                widget["y0"]
            ) <= 5:

                target = row

                break


        if target is None:

            rows.append({
                "y":
                    widget[
                        "y0"
                    ],

                "items":
                    []
            })


    rows = sorted(
        rows,
        key=lambda row:
            row["y"]
    )


    # Limit absurdly dense PDFs.
    # The LA-350 form is comfortably below this.
    if len(rows) > 120:

        rows = rows[
            :120
        ]


    if not rows:

        paragraph = document.add_paragraph(
            "No editable text could be extracted from this form."
        )

        return


    # --------------------------------------------------------
    # Three-column fixed native Word table.
    # --------------------------------------------------------

    table = document.add_table(
        rows=len(rows),
        cols=3
    )


    table.alignment = (
        WD_TABLE_ALIGNMENT.CENTER
    )


    table.autofit = False


    set_table_fixed_layout(
        table
    )


    # Use available width slightly smaller than full page.
    # Court forms typically use US Letter with margins.
    usable_width = max(
        400,
        page_width - 52
    )


    column_widths = [
        usable_width * 0.34,
        usable_width * 0.33,
        usable_width * 0.33
    ]


    for row in table.rows:

        for column_index, cell in enumerate(
            row.cells
        ):

            cell.width = Pt(
                column_widths[
                    column_index
                ]
            )


            cell.vertical_alignment = (
                WD_CELL_VERTICAL_ALIGNMENT.TOP
            )


            set_cell_margins(
                cell,
                top=0,
                bottom=0,
                start=30,
                end=30
            )


            # Start borderless.
            remove_all_cell_borders(
                cell
            )


    # --------------------------------------------------------
    # Build each visual row.
    # --------------------------------------------------------

    previous_y = rows[
        0
    ][
        "y"
    ]


    for row_index, row_data in enumerate(
        rows
    ):

        word_row = table.rows[
            row_index
        ]


        current_y = row_data[
            "y"
        ]


        gap = max(
            7,
            min(
                28,
                current_y -
                previous_y
            )
        )


        # Word row height
        word_row.height = Pt(
            gap
        )


        previous_y = current_y


        # Native PDF text
        for item in row_data[
            "items"
        ]:

            center_x = (
                item["x0"]
                +
                item["x1"]
            ) / 2


            column = determine_column(
                center_x,
                page_width
            )


            add_text_item_to_cell(
                word_row.cells[
                    column
                ],
                item
            )


        # Widgets belonging to same visual row
        for widget in widgets:

            if abs(
                widget["y0"] -
                current_y
            ) > 5:

                continue


            center_x = (
                widget["x0"]
                +
                widget["x1"]
            ) / 2


            column = determine_column(
                center_x,
                page_width
            )


            add_widget_to_cell(
                word_row.cells[
                    column
                ],
                widget
            )


    # --------------------------------------------------------
    # Add borders around rows that look like major form bands.
    #
    # This uses real Word table borders.
    # --------------------------------------------------------

    for row_index, row_data in enumerate(
        rows
    ):

        text = " ".join(
            item["text"]
            for item in row_data[
                "items"
            ]
        ).lower()


        major_sections = (
            "services",
            "languages",
            "language assistance",
            "service area",
            "date",
            "signature",
            "name"
        )


        if any(
            phrase in text
            for phrase in major_sections
        ):

            for cell in table.rows[
                row_index
            ].cells:

                border = {
                    "val":
                        "single",

                    "sz":
                        5,

                    "color":
                        "000000"
                }


                set_cell_border(
                    cell,
                    top=border,
                    bottom=border,
                    left=border,
                    right=border
                )


    # --------------------------------------------------------
    # Small spacing after form table.
    # --------------------------------------------------------

    paragraph = document.add_paragraph()

    paragraph.paragraph_format.space_before = Pt(
        0
    )

    paragraph.paragraph_format.space_after = Pt(
        0
    )


# ============================================================
# INTERACTIVE FORM → NATIVE WORD DOCX
# ============================================================

def convert_interactive_form_to_docx(
    input_path,
    output_path
):

    print(
        "Starting native Word form reconstruction...",
        flush=True
    )


    pdf = fitz.open(
        input_path
    )


    document = Document()


    # --------------------------------------------------------
    # PAGE SETUP
    # --------------------------------------------------------

    if pdf.page_count > 0:

        first_page = pdf[
            0
        ]


        section = document.sections[
            0
        ]


        section.page_width = Pt(
            first_page.rect.width
        )


        section.page_height = Pt(
            first_page.rect.height
        )


        # Tight margins help us retain the form on one page.
        section.top_margin = Pt(
            14
        )

        section.bottom_margin = Pt(
            14
        )

        section.left_margin = Pt(
            18
        )

        section.right_margin = Pt(
            18
        )


        section.header_distance = Pt(
            0
        )

        section.footer_distance = Pt(
            0
        )


    # --------------------------------------------------------
    # REMOVE THE DEFAULT EMPTY PARAGRAPH WHEN POSSIBLE
    # --------------------------------------------------------

    if document.paragraphs:

        paragraph = document.paragraphs[
            0
        ]

        paragraph.paragraph_format.space_before = Pt(
            0
        )

        paragraph.paragraph_format.space_after = Pt(
            0
        )

        paragraph.paragraph_format.line_spacing = Pt(
            1
        )


    # --------------------------------------------------------
    # BUILD PAGES
    # --------------------------------------------------------

    for page_number in range(
        pdf.page_count
    ):

        if page_number > 0:

            document.add_page_break()


        page = pdf[
            page_number
        ]


        add_native_form_page(
            document,
            page,
            page_number
        )


    pdf.close()


    print(
        "Saving native Word form...",
        flush=True
    )


    document.save(
        output_path
    )


    print(
        "Native form DOCX created:",
        output_path,
        flush=True
    )


# ============================================================
# PDF → WORD ROUTE
# ============================================================

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
            + ".docx"
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


            interactive = (
                pdf_has_interactive_form(
                    input_path
                )
            )


            if interactive:

                print(
                    "Interactive PDF detected.",
                    flush=True
                )


                print(
                    "Using native Word table/form reconstruction.",
                    flush=True
                )


                convert_interactive_form_to_docx(
                    input_path,
                    output_path
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


        # ----------------------------------------------------
        # VERIFY DOCX
        # ----------------------------------------------------

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


# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )


    app.run(
        host="0.0.0.0,
        port=port
    )
