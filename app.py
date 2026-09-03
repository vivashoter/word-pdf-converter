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
