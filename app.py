from flask import Flask, request, send_file, send_from_directory, jsonify
from werkzeug.utils import secure_filename

from docx import Document
from docx.shared import Pt, Inches, RGBColor
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


# ============================================================
# STATIC WEBSITE ROUTES
# ============================================================

@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/privacy.html")
def privacy():
    return send_from_directory(BASE_DIR, "privacy.html")


@app.route("/terms.html")
def terms():
    return send_from_directory(BASE_DIR, "terms.html")


@app.route("/about.html")
def about():
    return send_from_directory(BASE_DIR, "about.html")


@app.route("/contact.html")
def contact():
    return send_from_directory(BASE_DIR, "contact.html")


@app.route("/style.css")
def styles():
    return send_from_directory(BASE_DIR, "style.css")


@app.route("/script.js")
def scripts():
    return send_from_directory(BASE_DIR, "script.js")


@app.route("/convertdocgoose-logo.png")
def logo():
    return send_from_directory(BASE_DIR, "convertdocgoose-logo.png")


@app.route("/favicon.png")
def favicon():
    return send_from_directory(BASE_DIR, "favicon.png")


# ============================================================
# ERRORS
# ============================================================

@app.errorhandler(413)
def file_too_large(error):
    return jsonify({
        "error": "File is too large. Maximum size is 25 MB."
    }), 413


# ============================================================
# WORD -> PDF
# ============================================================

@app.route("/convert/word-to-pdf", methods=["POST"])
def word_to_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    uploaded_file = request.files["file"]

    if uploaded_file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    filename = secure_filename(uploaded_file.filename)

    if not filename.lower().endswith((".doc", ".docx")):
        return jsonify({"error": "Please upload a DOC or DOCX file."}), 400

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(temp_dir, filename)
        uploaded_file.save(input_path)

        try:
            result = subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    temp_dir,
                    input_path,
                ],
                check=True,
                timeout=120,
                capture_output=True,
                text=True,
            )

            if result.stdout:
                print("LibreOffice:", result.stdout, flush=True)

            if result.stderr:
                print("LibreOffice warnings:", result.stderr, flush=True)

        except subprocess.TimeoutExpired:
            return jsonify({"error": "The conversion took too long."}), 504

        except subprocess.CalledProcessError as error:
            print("LibreOffice error:", error.stderr, flush=True)
            return jsonify({
                "error": "The Word document could not be converted."
            }), 500

        except Exception as error:
            print("Word to PDF error:", repr(error), flush=True)
            return jsonify({
                "error": "The Word document could not be converted."
            }), 500

        output_name = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(temp_dir, output_name)

        if not os.path.exists(output_path):
            return jsonify({"error": "The PDF could not be created."}), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/pdf",
        )


# ============================================================
# EXACTDOC - NORMAL DIGITAL PDF -> WORD
# ============================================================

def run_exactdoc(input_path, output_path):
    print("Running ExactDoc on:", input_path, flush=True)

    result = subprocess.run(
        ["exactdoc", input_path, "-o", output_path],
        check=True,
        timeout=300,
        capture_output=True,
        text=True,
    )

    if result.stdout:
        print("ExactDoc output:", result.stdout, flush=True)

    if result.stderr:
        print("ExactDoc warnings:", result.stderr, flush=True)

    return result


# ============================================================
# PDF / FORM DETECTION
# ============================================================

def pdf_widget_count(input_path):
    pdf = None

    try:
        pdf = fitz.open(input_path)
        count = 0

        for page in pdf:
            widgets = page.widgets()
            if widgets:
                count += len(list(widgets))

        print("Interactive widget count:", count, flush=True)
        return count

    except Exception as error:
        print("Form detection error:", repr(error), flush=True)
        return 0

    finally:
        if pdf is not None:
            pdf.close()


def is_la350_form(input_path):
    pdf = None

    try:
        pdf = fitz.open(input_path)

        if pdf.page_count != 1:
            return False

        text = pdf[0].get_text("text") or ""
        normalized = re.sub(r"\s+", " ", text).lower()

        return (
            "la-350" in normalized
            and "notice of available language" in normalized
            and "assistance" in normalized
            and "service provider" in normalized
        )

    except Exception as error:
        print("LA-350 detection error:", repr(error), flush=True)
        return False

    finally:
        if pdf is not None:
            pdf.close()


# ============================================================
# WORD XML HELPERS
# ============================================================

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")

    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    borders = {
        "top": top,
        "bottom": bottom,
        "left": left,
        "right": right,
    }

    for edge_name, edge in borders.items():
        if edge is None:
            continue

        tag = "w:" + edge_name
        element = tcBorders.find(qn(tag))

        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)

        element.set(qn("w:val"), edge.get("val", "single"))
        element.set(qn("w:sz"), str(edge.get("sz", 6)))
        element.set(qn("w:color"), edge.get("color", "000000"))


def set_cell_margins(cell, top=20, start=45, bottom=20, end=45):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")

    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)

    values = {
        "top": top,
        "start": start,
        "bottom": bottom,
        "end": end,
    }

    for margin_name, value in values.items():
        node = tcMar.find(qn("w:" + margin_name))

        if node is None:
            node = OxmlElement("w:" + margin_name)
            tcMar.append(node)

        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_fixed_layout(table):
    tblPr = table._tbl.tblPr
    tblLayout = tblPr.first_child_found_in("w:tblLayout")

    if tblLayout is None:
        tblLayout = OxmlElement("w:tblLayout")
        tblPr.append(tblLayout)

    tblLayout.set(qn("w:type"), "fixed")


def set_cell_width(cell, width_inches):
    cell.width = Inches(width_inches)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.first_child_found_in("w:tcW")

    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)

    tcW.set(qn("w:w"), str(int(width_inches * 1440)))
    tcW.set(qn("w:type"), "dxa")


def set_row_height(row, points, exact=False):
    trPr = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(points * 20)))
    trHeight.set(qn("w:hRule"), "exact" if exact else "atLeast")
    trPr.append(trHeight)


def keep_table_row_together(row):
    trPr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    trPr.append(cant_split)


def set_paragraph_keep_with_next(paragraph, value=True):
    pPr = paragraph._p.get_or_add_pPr()
    existing = pPr.find(qn("w:keepNext"))

    if value:
        if existing is None:
            pPr.append(OxmlElement("w:keepNext"))
    elif existing is not None:
        pPr.remove(existing)


def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))

    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)

    shd.set(qn("w:fill"), fill)


def set_table_borders(table, size=6):
    border = {"val": "single", "sz": size, "color": "000000"}

    for row in table.rows:
        for cell in row.cells:
            set_cell_border(
                cell,
                top=border,
                bottom=border,
                left=border,
                right=border,
            )


def remove_cell_borders(cell):
    none = {"val": "nil", "sz": 0, "color": "FFFFFF"}
    set_cell_border(
        cell,
        top=none,
        bottom=none,
        left=none,
        right=none,
    )


def set_run_font(run, size=8, bold=False, italic=False, name="Arial"):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic

    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.rFonts

    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)


def format_paragraph(
    paragraph,
    size=8,
    bold=False,
    italic=False,
    align=None,
    before=0,
    after=0,
    line=1.0,
):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line

    if align is not None:
        paragraph.alignment = align

    for run in paragraph.runs:
        set_run_font(
            run,
            size=size,
            bold=bold,
            italic=italic,
        )


def add_run(
    paragraph,
    text,
    size=8,
    bold=False,
    italic=False,
):
    run = paragraph.add_run(text)

    set_run_font(
        run,
        size=size,
        bold=bold,
        italic=italic,
    )

    return run


def clear_cell(cell):
    cell.text = ""

    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0

    return paragraph


def set_normal_document_defaults(document):
    styles = document.styles
    normal = styles["Normal"]

    normal.font.name = "Arial"
    normal.font.size = Pt(8)

    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1.0


# ============================================================
# LA-350 WIDGET EXTRACTION
# ============================================================

def widget_short_name(full_name):
    if not full_name:
        return ""

    last = full_name.split(".")[-1]
    return re.sub(r"\[\d+\]$", "", last)


def checkbox_is_checked(value):
    if value is None:
        return False

    text = str(value).strip().lower()

    return text not in {
        "",
        "off",
        "0",
        "false",
        "none",
        "null",
        "no",
    }


def extract_la350_values(input_path):
    values = {
        "date": "",
        "court": "",
        "typed_name": "",
        "provider": "",
        "address": "",
        "telephone": "",
        "contact_name": "",
        "email": "",
        "web": "",
        "calendar_year": "",
        "service_specify": "",
        "language_specify": "",
        "assistance_specify": "",
        "service_area": "",
        "narrative": False,
        "services": [False] * 10,
        "languages": [False] * 12,
        "assistance": [False] * 5,
    }

    pdf = fitz.open(input_path)

    try:
        page = pdf[0]
        widgets = list(page.widgets() or [])

        for widget in widgets:
            full_name = widget.field_name or ""
            short = widget_short_name(full_name)
            value = widget.field_value or ""

            text_map = {
                "DateTimeField1": "date",
                "CourtInfo_ft": "court",
                "YourName_ft": "typed_name",
                "TextField3": "provider",
                "TextField4": "address",
                "TextField5": "telephone",
                "TextField7": "contact_name",
                "TextField8": "email",
                "TextField20": "web",
                "TextField21": "service_specify",
                "TextField22": "language_specify",
                "TextField23": "assistance_specify",
                "TextField24": "service_area",
                "TextField25": "calendar_year",
            }

            if short in text_map:
                values[text_map[short]] = str(value).strip()
                continue

            if short == "CheckBox29":
                values["narrative"] = checkbox_is_checked(value)
                continue

            if "Table1" in full_name and short.startswith("CheckBox"):
                number_match = re.search(r"CheckBox(\d+)", short)

                if number_match:
                    number = int(number_match.group(1))

                    service_number_to_index = {
                        1: 0,
                        2: 1,
                        3: 2,
                        4: 3,
                        5: 4,
                        6: 5,
                        7: 6,
                        8: 7,
                        9: 8,
                        28: 9,
                    }

                    index = service_number_to_index.get(number)

                    if index is not None:
                        values["services"][index] = checkbox_is_checked(value)

                continue

            if "Table2" in full_name and short.startswith("CheckBox"):
                number_match = re.search(r"CheckBox(\d+)", short)

                if number_match:
                    number = int(number_match.group(1))

                    if 11 <= number <= 22:
                        values["languages"][number - 11] = checkbox_is_checked(value)

                continue

            if "Table3" in full_name and short.startswith("CheckBox"):
                number_match = re.search(r"CheckBox(\d+)", short)

                if number_match:
                    number = int(number_match.group(1))

                    if 23 <= number <= 27:
                        values["assistance"][number - 23] = checkbox_is_checked(value)

                continue

        return values

    finally:
        pdf.close()


# ============================================================
# SMALL WORD FORM HELPERS
# ============================================================

def checkbox_symbol(checked):
    return "☒" if checked else "☐"


def add_checkbox_line(
    cell,
    label,
    checked=False,
    size=7.2,
    bold=False,
):
    paragraph = clear_cell(cell)

    add_run(
        paragraph,
        checkbox_symbol(checked) + " ",
        size=size,
    )

    add_run(
        paragraph,
        label,
        size=size,
        bold=bold,
    )

    return paragraph


def add_form_line(
    paragraph,
    label,
    value="",
    width_chars=30,
    size=7.6,
    label_bold=False,
):
    add_run(
        paragraph,
        label,
        size=size,
        bold=label_bold,
    )

    if value:
        add_run(
            paragraph,
            " " + value,
            size=size,
        )
    else:
        add_run(
            paragraph,
            " " + "_" * width_chars,
            size=size,
        )


def add_number_circle(cell, number):
    paragraph = clear_cell(cell)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = add_run(
        paragraph,
        str(number),
        size=8,
        bold=True,
    )

    set_cell_border(
        cell,
        top={"val": "single", "sz": 8, "color": "000000"},
        bottom={"val": "single", "sz": 8, "color": "000000"},
        left={"val": "single", "sz": 8, "color": "000000"},
        right={"val": "single", "sz": 8, "color": "000000"},
    )

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    set_cell_margins(
        cell,
        top=0,
        bottom=0,
        start=0,
        end=0,
    )

    return run


def make_section_marker_row(
    document,
    number,
    text,
    right_text=None,
    right_checked=False,
):
    table = document.add_table(
        rows=1,
        cols=2,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    set_table_fixed_layout(table)

    set_cell_width(
        table.cell(0, 0),
        0.28,
    )

    set_cell_width(
        table.cell(0, 1),
        7.12,
    )

    remove_cell_borders(table.cell(0, 0))
    remove_cell_borders(table.cell(0, 1))

    p0 = clear_cell(table.cell(0, 0))

    add_run(
        p0,
        f"{number} ",
        size=8,
        bold=True,
    )

    p1 = clear_cell(table.cell(0, 1))

    add_run(
        p1,
        text,
        size=7.6,
    )

    if right_text:
        add_run(
            p1,
            "    " + checkbox_symbol(right_checked) + " ",
            size=7.6,
        )

        add_run(
            p1,
            right_text,
            size=7.6,
        )

    set_row_height(
        table.rows[0],
        14,
        exact=False,
    )

    keep_table_row_together(table.rows[0])

    return table


# ============================================================
# LA-350 TEMPLATE RECONSTRUCTION
# ============================================================

def add_la350_header(document):
    table = document.add_table(
        rows=1,
        cols=3,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    set_table_fixed_layout(table)

    widths = [
        1.45,
        3.32,
        2.63,
    ]

    for index, width in enumerate(widths):
        set_cell_width(
            table.cell(0, index),
            width,
        )

        set_cell_margins(
            table.cell(0, index),
            top=25,
            bottom=25,
            start=55,
            end=55,
        )

    left = table.cell(0, 0)

    set_cell_shading(
        left,
        "000000",
    )

    p = clear_cell(left)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_run(
        p,
        "LA-350",
        size=15,
        bold=True,
    )

    for run in p.runs:
        run.font.color.rgb = RGBColor(
            255,
            255,
            255,
        )

    middle = table.cell(0, 1)

    p = clear_cell(middle)

    add_run(
        p,
        "Notice of Available Language\nAssistance—Service Provider",
        size=12.5,
        bold=True,
    )

    right = table.cell(0, 2)

    p = clear_cell(right)

    add_run(
        p,
        "Clerk stamps date here when form is received.",
        size=6.4,
        italic=True,
    )

    set_cell_border(
        right,
        top={"val": "single", "sz": 8, "color": "000000"},
        bottom={"val": "single", "sz": 8, "color": "000000"},
        left={"val": "single", "sz": 8, "color": "000000"},
        right={"val": "single", "sz": 8, "color": "000000"},
    )

    remove_cell_borders(left)
    remove_cell_borders(middle)

    set_row_height(
        table.rows[0],
        48,
        exact=True,
    )

    keep_table_row_together(table.rows[0])

    return table


def add_la350_intro(document, values):
    table = document.add_table(
        rows=2,
        cols=2,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    set_table_fixed_layout(table)

    set_cell_width(table.cell(0, 0), 4.80)
    set_cell_width(table.cell(0, 1), 2.60)
    set_cell_width(table.cell(1, 0), 4.80)
    set_cell_width(table.cell(1, 1), 2.60)

    for row in table.rows:
        keep_table_row_together(row)

        for cell in row.cells:
            remove_cell_borders(cell)

            set_cell_margins(
                cell,
                top=15,
                bottom=10,
                start=25,
                end=35,
            )

    cell = table.cell(0, 0)

    p = clear_cell(cell)

    add_run(
        p,
        "Use this form to:",
        size=7.5,
        bold=True,
    )

    p2 = cell.add_paragraph()

    format_paragraph(
        p2,
        size=7.2,
        before=0,
        after=0,
        line=0.95,
    )

    add_run(
        p2,
        "• Tell the court that you are a service provider, program, or professional offering language assistance with services that may be ordered by a court; and",
        size=7.2,
    )

    p3 = cell.add_paragraph()

    format_paragraph(
        p3,
        size=7.2,
        before=0,
        after=0,
        line=0.95,
    )

    add_run(
        p3,
        "• Provide information about the services you provide, the languages and types of language assistance available, and your service area.",
        size=7.2,
    )

    cell = table.cell(0, 1)

    p = clear_cell(cell)

    add_run(
        p,
        "Fill in court name and address:",
        size=6.4,
        italic=True,
    )

    p = cell.add_paragraph()

    format_paragraph(
        p,
        before=1,
        after=0,
    )

    add_run(
        p,
        "Superior Court of California, County of",
        size=7.1,
        bold=True,
    )

    court_value = values["court"] or ""

    p = cell.add_paragraph()

    format_paragraph(
        p,
        before=2,
        after=0,
    )

    add_run(
        p,
        court_value if court_value else " ",
        size=7.1,
    )

    set_cell_border(
        cell,
        top={"val": "single", "sz": 6, "color": "000000"},
        bottom={"val": "single", "sz": 6, "color": "000000"},
        left={"val": "single", "sz": 6, "color": "000000"},
        right={"val": "single", "sz": 6, "color": "000000"},
    )

    cell = table.cell(1, 0)

    p = clear_cell(cell)

    add_run(
        p,
        "1  ",
        size=8,
        bold=True,
    )

    add_run(
        p,
        "This form should be filed with the court by January 31 of each year to indicate services that will be provided during the calendar year. You may also submit this form to let the court know your services have changed.",
        size=7.0,
    )

    p = cell.add_paragraph()

    format_paragraph(
        p,
        before=1,
        after=0,
        line=0.95,
    )

    add_run(
        p,
        "    The information in this form describes services available during calendar year: ",
        size=7.0,
    )

    add_run(
        p,
        values["calendar_year"]
        if values["calendar_year"]
        else "____________",
        size=7.0,
    )

    table.cell(1, 1).text = ""

    set_row_height(
        table.rows[0],
        87,
        exact=False,
    )

    set_row_height(
        table.rows[1],
        52,
        exact=False,
    )

    return table


def add_la350_provider_section(document, values):
    table = document.add_table(
        rows=3,
        cols=2,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    set_table_fixed_layout(table)

    for row in table.rows:
        keep_table_row_together(row)

    set_cell_width(
        table.cell(0, 0),
        0.32,
    )

    set_cell_width(
        table.cell(0, 1),
        7.08,
    )

    table.cell(1, 0).merge(
        table.cell(1, 1)
    )

    table.cell(2, 0).merge(
        table.cell(2, 1)
    )

    for row in table.rows:
        for cell in row.cells:
            remove_cell_borders(cell)

            set_cell_margins(
                cell,
                top=0,
                bottom=0,
                start=15,
                end=15,
            )

    p = clear_cell(table.cell(0, 0))

    add_run(
        p,
        "2",
        size=8,
        bold=True,
    )

    p = clear_cell(table.cell(0, 1))

    add_form_line(
        p,
        "Name of service provider:",
        values["provider"],
        width_chars=45,
        size=7.4,
    )

    p = clear_cell(table.cell(1, 0))

    add_form_line(
        p,
        "Address:",
        values["address"],
        width_chars=77,
        size=7.4,
    )

    p = clear_cell(table.cell(2, 0))

    add_form_line(
        p,
        "Telephone:",
        values["telephone"],
        width_chars=18,
        size=7.4,
    )

    add_run(
        p,
        "        ",
        size=7.4,
    )

    add_form_line(
        p,
        "Web address:",
        values["web"],
        width_chars=30,
        size=7.4,
    )

    p2 = table.cell(2, 0).add_paragraph()

    format_paragraph(
        p2,
        before=1,
        after=0,
    )

    add_form_line(
        p2,
        "Contact name:",
        values["contact_name"],
        width_chars=24,
        size=7.4,
    )

    add_run(
        p2,
        "      ",
        size=7.4,
    )

    add_form_line(
        p2,
        "E-mail:",
        values["email"],
        width_chars=29,
        size=7.4,
    )

    set_row_height(
        table.rows[0],
        18,
        exact=False,
    )

    set_row_height(
        table.rows[1],
        16,
        exact=False,
    )

    set_row_height(
        table.rows[2],
        31,
        exact=False,
    )

    return table


def add_option_table(
    parent_cell,
    title,
    labels,
    checked_values,
    specify_value="",
    include_service_area=False,
    service_area_value="",
):
    rows_needed = 1 + len(labels) + 1

    if include_service_area:
        rows_needed += 1

    table = parent_cell.add_table(
        rows=rows_needed,
        cols=1,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    set_table_fixed_layout(table)
    set_table_borders(table, size=5)

    header = table.cell(0, 0)

    p = clear_cell(header)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_run(
        p,
        title,
        size=8.2,
        bold=True,
    )

    p2 = header.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

    format_paragraph(
        p2,
        before=0,
        after=0,
    )

    add_run(
        p2,
        "(select all that apply)",
        size=6.8,
        italic=True,
    )

    set_cell_margins(
        header,
        top=20,
        bottom=20,
        start=35,
        end=35,
    )

    set_row_height(
        table.rows[0],
        28,
        exact=True,
    )

    for index, label in enumerate(labels):
        cell = table.cell(index + 1, 0)

        add_checkbox_line(
            cell,
            label,
            checked_values[index],
            size=7.0,
        )

        set_cell_margins(
            cell,
            top=12,
            bottom=12,
            start=30,
            end=30,
        )

        height = 17

        if len(label) > 30:
            height = 26
        elif len(label) > 23:
            height = 21

        set_row_height(
            table.rows[index + 1],
            height,
            exact=False,
        )

        keep_table_row_together(
            table.rows[index + 1]
        )

    specify_row_index = 1 + len(labels)

    cell = table.cell(
        specify_row_index,
        0,
    )

    p = clear_cell(cell)

    add_form_line(
        p,
        "Specify:",
        specify_value,
        width_chars=17,
        size=7.0,
    )

    set_cell_margins(
        cell,
        top=12,
        bottom=10,
        start=30,
        end=30,
    )

    set_row_height(
        table.rows[specify_row_index],
        21,
        exact=False,
    )

    if include_service_area:
        service_row_index = specify_row_index + 1

        cell = table.cell(
            service_row_index,
            0,
        )

        p = clear_cell(cell)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        add_run(
            p,
            "Service Area",
            size=8.2,
            bold=True,
        )

        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

        format_paragraph(
            p2,
            before=0,
            after=0,
        )

        add_run(
            p2,
            "(county or region)",
            size=6.8,
        )

        p3 = cell.add_paragraph()

        format_paragraph(
            p3,
            before=3,
            after=0,
        )

        add_run(
            p3,
            service_area_value
            if service_area_value
            else "\n\n________________________",
            size=7.0,
        )

        set_cell_margins(
            cell,
            top=20,
            bottom=20,
            start=35,
            end=35,
        )

        set_row_height(
            table.rows[service_row_index],
            90,
            exact=False,
        )

    return table


def add_la350_services_section(document, values):
    heading = document.add_table(
        rows=1,
        cols=2,
    )

    heading.alignment = WD_TABLE_ALIGNMENT.CENTER
    heading.autofit = False

    set_table_fixed_layout(heading)

    set_cell_width(
        heading.cell(0, 0),
        3.05,
    )

    set_cell_width(
        heading.cell(0, 1),
        4.35,
    )

    for cell in heading.rows[0].cells:
        remove_cell_borders(cell)

        set_cell_margins(
            cell,
            top=0,
            bottom=0,
            start=10,
            end=10,
        )

    p = clear_cell(
        heading.cell(0, 0)
    )

    add_run(
        p,
        "3  Information about the services provided:",
        size=7.5,
    )

    p = clear_cell(
        heading.cell(0, 1)
    )

    add_run(
        p,
        checkbox_symbol(values["narrative"]) + " ",
        size=7.4,
    )

    add_run(
        p,
        "Check here to attach a narrative description of the services offered.",
        size=7.2,
    )

    set_row_height(
        heading.rows[0],
        20,
        exact=False,
    )

    keep_table_row_together(
        heading.rows[0]
    )

    outer = document.add_table(
        rows=1,
        cols=3,
    )

    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    outer.autofit = False

    set_table_fixed_layout(outer)

    widths = [
        2.58,
        1.95,
        2.87,
    ]

    for index, width in enumerate(widths):
        cell = outer.cell(0, index)

        set_cell_width(
            cell,
            width,
        )

        remove_cell_borders(cell)

        set_cell_margins(
            cell,
            top=0,
            bottom=0,
            start=0,
            end=0,
        )

        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

        clear_cell(cell)

    service_labels = [
        "Mediation",
        "Child custody recommending counseling",
        "Professional supervised child visitation",
        "Parenting education classes",
        "Anger management classes",
        "Mental health counseling",
        "Batterer intervention–MEN",
        "Batterer intervention–WOMEN",
        "Alcohol/substance abuse treatment",
        "Other",
    ]

    language_labels = [
        "Any language",
        "American Sign Language",
        "Spanish",
        "Mandarin",
        "Cantonese",
        "Farsi",
        "Korean",
        "Punjabi",
        "Russian",
        "Tagalog",
        "Vietnamese",
        "Other",
    ]

    assistance_labels = [
        "Program offered directly in language",
        "In-person interpreter",
        "Telephone interpreter",
        "Translated materials",
        "Other",
    ]

    add_option_table(
        outer.cell(0, 0),
        "Services",
        service_labels,
        values["services"],
        specify_value=values["service_specify"],
    )

    add_option_table(
        outer.cell(0, 1),
        "Languages Available",
        language_labels,
        values["languages"],
        specify_value=values["language_specify"],
    )

    add_option_table(
        outer.cell(0, 2),
        "Types of Language\nAssistance",
        assistance_labels,
        values["assistance"],
        specify_value=values["assistance_specify"],
        include_service_area=True,
        service_area_value=values["service_area"],
    )

    keep_table_row_together(
        outer.rows[0]
    )

    return outer


def add_la350_signature_and_footer(document, values):
    p = document.add_paragraph()

    format_paragraph(
        p,
        before=1,
        after=0,
    )

    add_form_line(
        p,
        "Date:",
        values["date"],
        width_chars=21,
        size=7.2,
    )

    table = document.add_table(
        rows=2,
        cols=2,
    )

    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    set_table_fixed_layout(table)

    set_cell_width(
        table.cell(0, 0),
        3.70,
    )

    set_cell_width(
        table.cell(0, 1),
        3.70,
    )

    set_cell_width(
        table.cell(1, 0),
        3.70,
    )

    set_cell_width(
        table.cell(1, 1),
        3.70,
    )

    for row in table.rows:
        keep_table_row_together(row)

        for cell in row.cells:
            remove_cell_borders(cell)

            set_cell_margins(
                cell,
                top=0,
                bottom=0,
                start=0,
                end=0,
            )

    p = clear_cell(
        table.cell(0, 0)
    )

    add_run(
        p,
        values["typed_name"]
        if values["typed_name"]
        else "________________________________________",
        size=7.2,
    )

    p = clear_cell(
        table.cell(0, 1)
    )

    add_run(
        p,
        "________________________________________",
        size=7.2,
    )

    p = clear_cell(
        table.cell(1, 0)
    )

    add_run(
        p,
        "Type or print your name",
        size=6.8,
        italic=True,
    )

    p = clear_cell(
        table.cell(1, 1)
    )

    add_run(
        p,
        "Sign your name",
        size=6.8,
        italic=True,
    )

    set_row_height(
        table.rows[0],
        12,
        exact=False,
    )

    set_row_height(
        table.rows[1],
        12,
        exact=False,
    )

    footer = document.add_table(
        rows=1,
        cols=3,
    )

    footer.alignment = WD_TABLE_ALIGNMENT.CENTER
    footer.autofit = False

    set_table_fixed_layout(footer)

    widths = [
        2.25,
        3.55,
        1.60,
    ]

    for index, width in enumerate(widths):
        set_cell_width(
            footer.cell(0, index),
            width,
        )

        remove_cell_borders(
            footer.cell(0, index)
        )

        set_cell_margins(
            footer.cell(0, index),
            top=15,
            bottom=0,
            start=0,
            end=0,
        )

    p = clear_cell(
        footer.cell(0, 0)
    )

    add_run(
        p,
        "Judicial Council of California, www.courts.ca.gov\n"
        "New September 1, 2019, Optional Form\n"
        "Cal. Rules of Court, rule 1.300",
        size=5.2,
    )

    p = clear_cell(
        footer.cell(0, 1)
    )

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_run(
        p,
        "Notice of Available Language\n"
        "Assistance—Service Provider",
        size=9.5,
        bold=True,
    )

    p = clear_cell(
        footer.cell(0, 2)
    )

    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    add_run(
        p,
        "LA-350, Page 1 of 1",
        size=6.0,
        bold=True,
    )

    set_row_height(
        footer.rows[0],
        36,
        exact=False,
    )

    keep_table_row_together(
        footer.rows[0]
    )


# ============================================================
# LA-350 -> DOCX
# ============================================================

def convert_la350_to_docx(input_path, output_path):
    print(
        "LA-350 detected.",
        flush=True,
    )

    print(
        "Using dedicated native Word LA-350 reconstruction.",
        flush=True,
    )

    values = extract_la350_values(
        input_path
    )

    document = Document()

    set_normal_document_defaults(
        document
    )

    section = document.sections[0]

    section.page_width = Inches(8.5)
    section.page_height = Inches(11)

    section.top_margin = Inches(0.10)
    section.bottom_margin = Inches(0.10)
    section.left_margin = Inches(0.15)
    section.right_margin = Inches(0.15)

    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)

    if document.paragraphs:
        first = document.paragraphs[0]

        first.paragraph_format.space_before = Pt(0)
        first.paragraph_format.space_after = Pt(0)
        first.paragraph_format.line_spacing = 0.1

        run = first.add_run("")

        set_run_font(
            run,
            size=1,
        )

    add_la350_header(
        document
    )

    add_la350_intro(
        document,
        values,
    )

    add_la350_provider_section(
        document,
        values,
    )

    add_la350_services_section(
        document,
        values,
    )

    add_la350_signature_and_footer(
        document,
        values,
    )

    document.save(
        output_path
    )

    print(
        "LA-350 native DOCX created:",
        output_path,
        flush=True,
    )


# ============================================================
# UNKNOWN INTERACTIVE FORM FALLBACK
# ============================================================

def convert_interactive_form_image_fallback(
    input_path,
    output_path,
):
    print(
        "Unknown interactive form.",
        flush=True,
    )

    print(
        "Using safe image-preserving Word fallback.",
        flush=True,
    )

    pdf = fitz.open(
        input_path
    )

    document = Document()

    set_normal_document_defaults(
        document
    )

    try:
        for page_index in range(
            pdf.page_count
        ):
            page = pdf[
                page_index
            ]

            if page_index == 0:
                section = document.sections[0]
            else:
                document.add_page_break()
                section = document.sections[-1]

            section.page_width = Pt(
                page.rect.width
            )

            section.page_height = Pt(
                page.rect.height
            )

            section.top_margin = Pt(6)
            section.bottom_margin = Pt(6)
            section.left_margin = Pt(6)
            section.right_margin = Pt(6)

            section.header_distance = Pt(0)
            section.footer_distance = Pt(0)

            pix = page.get_pixmap(
                matrix=fitz.Matrix(2, 2),
                alpha=False,
                annots=True,
            )

            image_path = os.path.join(
                os.path.dirname(output_path),
                f"interactive_page_{page_index + 1}.png",
            )

            pix.save(
                image_path
            )

            paragraph = document.add_paragraph()

            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)

            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            available_width = max(
                page.rect.width - 12,
                72,
            )

            run = paragraph.add_run()

            run.add_picture(
                image_path,
                width=Pt(available_width),
            )

        document.save(
            output_path
        )

    finally:
        pdf.close()

    print(
        "Interactive-form fallback DOCX created:",
        output_path,
        flush=True,
    )


# ============================================================
# PDF -> WORD
# ============================================================

@app.route(
    "/convert/pdf-to-word",
    methods=["POST"],
)
def pdf_to_word():
    if "file" not in request.files:
        return jsonify({
            "error": "No file uploaded."
        }), 400

    uploaded_file = request.files[
        "file"
    ]

    if uploaded_file.filename == "":
        return jsonify({
            "error": "No file selected."
        }), 400

    filename = secure_filename(
        uploaded_file.filename
    )

    if not filename.lower().endswith(
        ".pdf"
    ):
        return jsonify({
            "error": "Please upload a PDF file."
        }), 400

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = os.path.join(
            temp_dir,
            filename,
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
            output_name,
        )

        try:
            print(
                "=================================",
                flush=True,
            )

            print(
                "PDF TO WORD START:",
                filename,
                flush=True,
            )

            print(
                "=================================",
                flush=True,
            )

            widget_count = pdf_widget_count(
                input_path
            )

            if widget_count > 0:
                if is_la350_form(
                    input_path
                ):
                    convert_la350_to_docx(
                        input_path,
                        output_path,
                    )

                else:
                    convert_interactive_form_image_fallback(
                        input_path,
                        output_path,
                    )

            else:
                print(
                    "Normal PDF detected.",
                    flush=True,
                )

                print(
                    "Using ExactDoc.",
                    flush=True,
                )

                run_exactdoc(
                    input_path,
                    output_path,
                )

        except subprocess.TimeoutExpired:
            print(
                "PDF conversion timed out.",
                flush=True,
            )

            return jsonify({
                "error": "The PDF conversion took too long."
            }), 504

        except subprocess.CalledProcessError as error:
            print(
                "PDF converter exit code:",
                error.returncode,
                flush=True,
            )

            print(
                "PDF converter stdout:",
                error.stdout,
                flush=True,
            )

            print(
                "PDF converter stderr:",
                error.stderr,
                flush=True,
            )

            return jsonify({
                "error": "The PDF could not be converted to Word."
            }), 500

        except Exception as error:
            print(
                "PDF to Word error:",
                repr(error),
                flush=True,
            )

            return jsonify({
                "error": "The PDF could not be converted to Word."
            }), 500

        if not os.path.exists(
            output_path
        ):
            return jsonify({
                "error": "The Word document could not be created."
            }), 500

        if os.path.getsize(
            output_path
        ) == 0:
            return jsonify({
                "error": "The Word document was created but was empty."
            }), 500

        print(
            "=================================",
            flush=True,
        )

        print(
            "PDF TO WORD SUCCESS:",
            output_name,
            flush=True,
        )

        print(
            "=================================",
            flush=True,
        )

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )


# ============================================================
# START SERVER
# ============================================================

if __name__ == "__main__":
    port = int(
        os.environ.get(
            "PORT",
            10000,
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
    )
