"""
==============================================================================
NGLESHIE AMANFROM POLYCLINIC LABORATORY — REPORT GENERATOR
==============================================================================
Reads a Mindray BS240 (or similar analyser) PDF and produces a patient report
laid out exactly like the laboratory's Word template:

    • Ghana Health Service logo + centred laboratory name on every page
    • PATH. NO / SPECIMEN DATE / PATIENT NAME / SEX / AGE / SPECIMEN block
    • Red section titles:  RENAL FUNCTION TEST · LIVER FUNCTION TEST ·
                           LIPID PROFILE · GLYCATED HAEMOGLOBIN · URIC ACID
    • Borderless 5-column tables: INVESTIGATION | VALUE | UNIT | FLAG |
                                  REFERENCE INTERVAL
    • "Date printed …" / "Electronically Signed by QC manager" footer
    • "- -END OF REPORT- -" closing line

HOW TO RUN IT
-------------
  Double-click the file  →  the drag-and-drop window opens.

  From a terminal / PowerShell:
    python lab_report.py --demo              (sample renal/liver/lipid report)
    python lab_report.py --demo-hba1c        (sample HbA1c report)
    python lab_report.py "patient.pdf"
    python lab_report.py "patient.pdf" --path-no CHE901/26 --sex FEMALE
                          --age 46 --specimen SERUM
                          --specimen-date 27/04/2026 --output report.pdf

WHAT YOU MAY WANT TO EDIT
-------------------------
  SECTION 2  — laboratory name, signatory line, page geometry, SHOW_FLAGS
  SECTION 6  — reference intervals and units (TEST_LIBRARY)
  SECTION 5  — analyser short codes (TEST_CODE_MAP)
==============================================================================
"""

# ==============================================================================
# SECTION 1 — IMPORTS
# "Importing" means borrowing tools that other people already wrote.
# Think of it like getting equipment out of a toolbox before starting a job.
# ==============================================================================

import argparse     # Reads what you type after "python lab_report.py ..."
import base64       # Decodes the laboratory logo that is embedded in this file
import io           # Lets us treat the decoded logo as if it were a file
import os           # File paths, folders, "does this file exist?"
import re           # Text pattern matching (used all over the parser)
import sys          # Command-line arguments and clean exits
import threading    # Runs PDF generation in the background so the GUI responds

from datetime import datetime
# datetime = the current date and time, used for the "Date printed" footer.

# ── GUI imports (tkinter + optional drag-and-drop) ────────────────────────────
# Wrapped in try/except so the script still runs in CLI mode on machines that
# have no tkinter (e.g. a bare Linux server). Windows and macOS always have it.
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    TK_AVAILABLE = True
except ImportError:
    TK_AVAILABLE = False
    tk = None

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkinterDnD = None

# ── ReportLab: the library that actually draws the PDF ────────────────────────
from reportlab.lib.pagesizes import letter      # 8.5in × 11in paper
from reportlab.lib import colors                # Named and hex colours
from reportlab.lib.units import inch            # So we can write 0.4*inch
from reportlab.lib.styles import ParagraphStyle # Reusable text styles
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.lib.utils import ImageReader     # Turns the logo bytes into art
from reportlab.platypus import (
    Paragraph,        # A block of text with a style applied to it
    Spacer,           # An invisible vertical gap
    Table,            # A grid of cells
    TableStyle,       # Rules controlling how a Table looks
    KeepTogether,     # Keeps a group of elements on the same page
    BaseDocTemplate,  # A document whose every page we can customise
    Frame,            # The rectangular content box on each page
    PageTemplate,     # "What does a page look like, and what runs on each one"
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ==============================================================================
# SECTION 2 — LABORATORY CONFIGURATION
# ✏️  These are the only lines most people ever need to change.
# ==============================================================================

LAB_NAME  = "NGLESHIE AMANFROM POLYCLINIC LABORATORY"
SIGNATORY = "Electronically Signed by QC manager"

# The two-digit year suffix used when building a path number like "CHE901/26".
# Leave as None to use the current year automatically.
PATH_NO_YEAR = None

# Set to True to have the script work out High / Low itself and print it in the
# FLAG column in red. Left False so the report matches the Word template, where
# the FLAG column stays blank unless the analyser itself printed a flag.
SHOW_FLAGS = False

# Print "- -END OF REPORT- -" after the last section.
SHOW_END_OF_REPORT = True

# ── Page geometry (inches) ────────────────────────────────────────────────────
# These numbers were measured off the Word template so the PDF lines up with it.
PAGE_LEFT    = 0.40    # Left margin — the results table starts here
PAGE_RIGHT   = 0.40    # Right margin
PAGE_TOP     = 2.70    # Big top margin: the header block is drawn inside it
PAGE_BOTTOM  = 1.20    # Bottom margin: the footer line is drawn inside it

HDR_TEXT_X   = 1.00    # Patient lines start 1 inch in, as in the template
HDR_COL2_X   = 5.00    # SPECIMEN DATE / SEX start here
HDR_RULE_END = 7.27    # The horizontal rule stops here

LOGO_X, LOGO_TOP, LOGO_W, LOGO_H = 0.08, 1.08, 0.97, 0.92

RED = colors.HexColor("#FF0000")   # The section-title red used in the template


# ==============================================================================
# SECTION 2B — EMBEDDED GHANA HEALTH SERVICE LOGO
#
# The logo is stored inside this file as base64 text (a way of writing binary
# data using ordinary letters) so the script is a single self-contained file —
# there is no separate image to keep next to it or lose.
#
# ✏️  To use a different logo: save it as a PNG, then either
#       (a) put the file next to this script and name it "lab_logo.png"
#           — the script prefers that file whenever it exists — or
#       (b) replace the text below with the base64 of your own image.
# ==============================================================================

LOGO_FILENAME = "lab_logo.png"   # Optional override sitting beside the script

GHS_LOGO_B64 = """
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
AAAASUVORK5CYII=
"""


def get_logo_image():
    """
    Returns something ReportLab can draw as the laboratory logo, or None.

    Order of preference:
      1. A file called lab_logo.png sitting next to this script
      2. The Ghana Health Service logo embedded above
    """
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        override = os.path.join(here, LOGO_FILENAME)
        if os.path.isfile(override):
            return ImageReader(override)
    except Exception:
        pass

    try:
        raw = base64.b64decode("".join(GHS_LOGO_B64.split()))
        return ImageReader(io.BytesIO(raw))
    except Exception:
        return None


# ==============================================================================
# SECTION 3 — FONTS
#
# The Word template uses Cambria (header), Algerian (red section titles),
# Berlin Sans FB (column headings) and Times New Roman (the results).
# Those first three only exist on Windows, so the script looks for them and
# quietly falls back to the fonts built into ReportLab when they are absent.
# The report is laid out identically either way — only the lettering differs.
# ==============================================================================

FONT_DIRS = [
    r"C:\Windows\Fonts",
    os.path.expanduser(r"~\AppData\Local\Microsoft\Windows\Fonts"),
    "/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts"),
    "/usr/share/fonts/truetype/msttcorefonts", "/usr/share/fonts",
]

# Filled in by _register_fonts(); every drawing routine reads these names.
F_HEADER      = "Times-Bold"    # Laboratory name + patient block
F_TITLE       = "Helvetica-Bold"  # Red section titles
F_COLHDR      = "Helvetica-Bold"  # INVESTIGATION / VALUE / UNIT / …
F_BODY        = "Times-Roman"   # Test names, units
F_BODY_BOLD   = "Times-Bold"    # Values and reference intervals
F_BODY_ITALIC = "Times-Italic"  # HbA1c test name
F_BODY_BI     = "Times-BoldItalic"  # eGFR explanatory note

# True once a real TrueType body font has been found. ReportLab's built-in
# fonts cannot draw µ or ≥, so when this stays False those characters are
# swapped for plain-text equivalents rather than silently disappearing.
HAVE_TTF_BODY = False


def _find_font_file(*filenames):
    """
    Returns the full path of the first font file that exists, else None.

    Windows keeps everything in one folder; Linux buries fonts a few folders
    deep, so each location is searched three levels down.
    """
    wanted = {name.lower(): name for name in filenames}

    for directory in FONT_DIRS:
        if not os.path.isdir(directory):
            continue

        # The plain case first — the file sitting right there
        for name in filenames:
            candidate = os.path.join(directory, name)
            if os.path.isfile(candidate):
                return candidate

        # Then a shallow walk, in case it is tucked into sub-folders
        try:
            base_depth = directory.rstrip(os.sep).count(os.sep)
            for root, dirs, files in os.walk(directory):
                if root.count(os.sep) - base_depth >= 3:
                    dirs[:] = []          # deep enough — stop descending
                    continue
                for file_name in files:
                    if file_name.lower() in wanted:
                        return os.path.join(root, file_name)
        except Exception:
            pass

    return None


def _try_register(internal_name, *filenames):
    """Registers a TrueType font if it can be found. Returns True on success."""
    path = _find_font_file(*filenames)
    if not path:
        return False
    try:
        pdfmetrics.registerFont(TTFont(internal_name, path))
        return True
    except Exception:
        return False


# The serif families to try for the results, best match first. Times New Roman
# is what the template uses; Liberation Serif has identical letter widths and
# ships with most Linux systems; DejaVu Serif is the last resort.
BODY_FONT_CANDIDATES = [
    # (regular, bold, italic, bold-italic)
    ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    ("Times New Roman.ttf", "Times New Roman Bold.ttf",
     "Times New Roman Italic.ttf", "Times New Roman Bold Italic.ttf"),
    ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf",
     "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"),
    ("DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf",
     "DejaVuSerif-Italic.ttf", "DejaVuSerif-BoldItalic.ttf"),
]


def _register_fonts():
    """Looks for the template's fonts and points the F_* names at them."""
    global F_HEADER, F_TITLE, F_COLHDR
    global F_BODY, F_BODY_BOLD, F_BODY_ITALIC, F_BODY_BI, HAVE_TTF_BODY

    # ── The results themselves — a serif face ─────────────────────────────────
    for regular, bold, italic, bold_italic in BODY_FONT_CANDIDATES:
        if not (_try_register("LabSerif", regular)
                and _try_register("LabSerif-Bold", bold)
                and _try_register("LabSerif-Italic", italic)
                and _try_register("LabSerif-BoldItalic", bold_italic)):
            continue
        try:
            pdfmetrics.registerFontFamily(
                "LabSerif", normal="LabSerif", bold="LabSerif-Bold",
                italic="LabSerif-Italic", boldItalic="LabSerif-BoldItalic")
        except Exception:
            continue
        F_BODY, F_BODY_BOLD = "LabSerif", "LabSerif-Bold"
        F_BODY_ITALIC, F_BODY_BI = "LabSerif-Italic", "LabSerif-BoldItalic"
        F_HEADER = "LabSerif-Bold"   # A sensible default until Cambria is found
        HAVE_TTF_BODY = True
        break

    # ── Cambria Bold — the laboratory name and the patient block ──────────────
    if _try_register("Cambria-Bold", "cambriab.ttf", "Cambria Bold.ttf"):
        F_HEADER = "Cambria-Bold"

    # ── Algerian — the red section titles ─────────────────────────────────────
    if _try_register("Algerian", "ALGER.TTF", "alger.ttf", "Algerian.ttf"):
        F_TITLE = "Algerian"

    # ── Berlin Sans FB Demi Bold — the column headings ────────────────────────
    if _try_register("BerlinSansFB-Bold", "BRLNSDB.TTF", "brlnsdb.ttf"):
        F_COLHDR = "BerlinSansFB-Bold"


_register_fonts()


# Characters the built-in ReportLab fonts have no glyph for. When no TrueType
# serif could be found, these are swapped for plain equivalents so a unit or a
# reference interval never comes out with a character missing.
PLAIN_SUBSTITUTES = {
    "µ": "u", "μ": "u",          # micro sign, Greek mu
    "≥": "&gt;=", "≤": "&lt;=",  # written as entities: these go inside markup
    "–": "-", "—": "-",          # en dash, em dash
    "≈": "~", "×": "x", "·": ".", "²": "2",
}


def safe_text(text, markup=True):
    """
    Returns text the current fonts can actually draw.

    markup=True  — the text is going into a Paragraph, where "<" and "&" have
                   a special meaning, so replacements are written as entities.
    markup=False — the text is being drawn straight onto the page.
    """
    text = "" if text is None else str(text)
    if HAVE_TTF_BODY:
        return text
    for character, replacement in PLAIN_SUBSTITUTES.items():
        if not markup:
            replacement = replacement.replace("&gt;", ">").replace("&lt;", "<")
        text = text.replace(character, replacement)
    return text


# ==============================================================================
# SECTION 4 — SECTIONS OF THE REPORT
#
# Each entry describes one red-titled block of the report:
#   "title"     — exactly what is printed in red, centred
#   "ref_label" — the wording of the last column heading
#   "show_flag" — whether the FLAG column appears at all
# ==============================================================================

RENAL = "RENAL FUNCTION TEST"
LIVER = "LIVER FUNCTION TEST"
LIPID = "LIPID PROFILE"
HBA1C = "GLYCATED HAEMOGLOBIN"
URIC  = "URIC ACID"
OTHER = "OTHER INVESTIGATIONS"

# "indent" and "cols" are inches, and only needed where a section's columns sit
# somewhere other than the usual five positions — the glycated haemoglobin page
# of the template is laid out differently from the rest, so it says so here.
SECTIONS = {
    RENAL: {"ref_label": "REFERENCE INTERVAL", "show_flag": True,
            "cols": [1.329, 1.480, 1.500, 1.397, 1.997]},
    LIVER: {"ref_label": "REFERENCE INTERVAL", "show_flag": True,
            "cols": [1.682, 1.333, 1.310, 1.142, 2.234]},
    LIPID: {"ref_label": "REFERENCE  VALUE",   "show_flag": True,
            "cols": [2.123, 1.353, 0.922, 0.868, 2.433]},
    HBA1C: {"ref_label": "REFERENCE VALUE",    "show_flag": False,
            "indent": 0.60, "cols": [2.95, 0.85, 0.90, 2.40]},
    URIC:  {"ref_label": "REFERENCE VALUE",    "show_flag": True},
    OTHER: {"ref_label": "REFERENCE VALUE",    "show_flag": True},
}

# Top-to-bottom order of the sections in the finished report
SECTION_ORDER = [RENAL, LIVER, LIPID, HBA1C, URIC, OTHER]


# ==============================================================================
# SECTION 5 — TEST LIBRARY
#
# One row per test the laboratory reports. Everything the report needs to know
# about a test lives here, so this is the single place to edit when a reference
# interval changes or a new test is introduced.
#
#   name  — printed in the INVESTIGATION column, and used to order the rows
#   panel — which red-titled section it belongs to
#   unit  — printed in the UNIT column
#   ref   — one or more lines printed in the reference column
#   low / high — optional numeric limits, used only when SHOW_FLAGS is True
#   source_ref — when True, whatever the analyser printed for this test is used
#                instead of the "ref" line above, and "ref" becomes the fallback
#                for when the analyser printed nothing
#
# ✏️  Reference intervals below are the laboratory's own. Creatinine and uric
#     acid are marked source_ref, so each report shows the interval the
#     analyser itself gave for that patient. Every other test — alkaline
#     phosphatase included — uses the fixed interval written here.
# ==============================================================================

EGFR_STAGES = [
    "Normal Kidney Function: 90+",
    "Mildly reduced Kidney Function: 60 - 89",
    "Moderately reduced Kidney Function: 30 – 59",
    "Severely reduced Kidney Function: 15 - 29",
    "End stage Kidney failure:&lt;15",
]

HBA1C_INTERPRETATION = [
    "NON – DIABETIC:   3.4% – 6.0%",
    "PREDIABETES: 5.7% - 6.4%",
    "DIABETIC:",
    "GOOD CONTROL: 6.1 – 8.0 %",
    "FAIR CONTROL:  8.1 – 10.0 %",
    "POOR CONTROL:   ABOVE 10 %",
]

EGFR_NOTE = (
    "Estimated Glomerular Filtration Rate (GFR) was calculated using the "
    "CKD-EPI (2021) creatinine equation, which does not include a race "
    "variable. eGFR is an estimate and should be interpreted alongside "
    "clinical findings, urine albumin-to-creatinine ratio (uACR), and patient "
    "history. Values may be affected by extremes of muscle mass, malnutrition, "
    "amputations, or acute illness. A single eGFR result is insufficient for "
    "CKD staging; confirmation requires two measurements at least 90 days "
    "apart. For clinical decisions, physician correlation is advised."
)

TEST_LIBRARY = [
    # ── Renal Function ────────────────────────────────────────────────────────
    {"name": "CREATININE", "panel": RENAL, "unit": "µmol/L",
     "ref": ["53.0 – 140.0"], "source_ref": True,
     "low": 53.0, "high": 140.0},

    {"name": "UREA", "panel": RENAL, "unit": "mmol/L",
     "ref": ["1.7 – 8.3"], "low": 1.7, "high": 8.3},

    {"name": "SODIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["135 - 150"], "low": 135, "high": 150},

    {"name": "POTASSIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["3.5 - 5.5"], "low": 3.5, "high": 5.5},

    {"name": "CHLORIDE", "panel": RENAL, "unit": "mmol/L",
     "ref": ["98.0 - 110.0"], "low": 98.0, "high": 110.0},

    {"name": "TOTAL CO2", "panel": RENAL, "unit": "mmol/L",
     "ref": ["22.0 - 29.0"], "low": 22.0, "high": 29.0},

    {"name": "CALCIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["2.15 - 2.55"], "low": 2.15, "high": 2.55},

    {"name": "PHOSPHATE", "panel": RENAL, "unit": "mmol/L",
     "ref": ["0.81 - 1.45"], "low": 0.81, "high": 1.45},

    {"name": "MAGNESIUM", "panel": RENAL, "unit": "mmol/L",
     "ref": ["0.75 - 1.00"], "low": 0.75, "high": 1.00},

    {"name": "eGFR (CKD-EPI)", "panel": RENAL, "unit": "ml/min/1.73m<super>2</super>",
     "ref": EGFR_STAGES, "low": 90, "high": None, "note": EGFR_NOTE},

    # ── Liver Function ────────────────────────────────────────────────────────
    {"name": "TOTAL PROTEIN", "panel": LIVER, "unit": "g/L",
     "ref": ["60.0 – 87.0"], "low": 60.0, "high": 87.0},

    {"name": "ALBUMIN", "panel": LIVER, "unit": "g/L",
     "ref": ["35.0 – 50.0"], "low": 35.0, "high": 50.0},

    {"name": "GLOBULIN", "panel": LIVER, "unit": "g/L",
     "ref": ["25.0 – 35.0"], "low": 25.0, "high": 35.0},

    {"name": "A/G RATIO", "panel": LIVER, "unit": "",
     "ref": ["1.0 – 2.2"], "low": 1.0, "high": 2.2},

    {"name": "BILIRUBIN (TOTAL)", "panel": LIVER, "unit": "µmol/L",
     "ref": ["1.0 – 22.0"], "low": 1.0, "high": 22.0},

    {"name": "BILIRUBIN (DIRECT)", "panel": LIVER, "unit": "µmol/L",
     "ref": ["1.0 – 9.0"], "low": 1.0, "high": 9.0},

    {"name": "BILIRUBIN (INDIRECT)", "panel": LIVER, "unit": "µmol/L",
     "ref": ["0.0 – 17.0"], "low": 0.0, "high": 17.0},

    {"name": "GAMMA GT", "panel": LIVER, "unit": "U/L",
     "ref": ["5.0 – 50.0"], "low": 5.0, "high": 50.0},

    {"name": "ALT (SGPT)", "panel": LIVER, "unit": "U/L",
     "ref": ["0.0 – 45.0"], "low": 0.0, "high": 45.0},

    {"name": "AST (SGOT)", "panel": LIVER, "unit": "U/L",
     "ref": ["0.0 – 40.0"], "low": 0.0, "high": 40.0},

    {"name": "AST/ALT RATIO", "panel": LIVER, "unit": "",
     "ref": ["0.8 – 1.4"], "low": 0.8, "high": 1.4},

    {"name": "ALKALINE PHOSPHATASE", "panel": LIVER, "unit": "U/L",
     "ref": ["40 – 150"], "low": 40.0, "high": 150.0},

    # ── Lipid Profile ─────────────────────────────────────────────────────────
    {"name": "TOTAL CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["2.80 – 5.20"], "low": 2.80, "high": 5.20},

    {"name": "HDL-CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["≥ 0.90"], "low": 0.90, "high": None},

    {"name": "LDL-CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["&lt;4.00"], "low": None, "high": 4.00},

    {"name": "TRIGLYCERIDES", "panel": LIPID, "unit": "mmol/L",
     "ref": ["0.50 - 1.80"], "low": 0.50, "high": 1.80},

    {"name": "VLDL-CHOLESTEROL", "panel": LIPID, "unit": "mmol/L",
     "ref": ["0.10 - 1.00"], "low": 0.10, "high": 1.00},

    {"name": "CORONARY RISK", "panel": LIPID, "unit": "",
     "ref": ["&lt;4.0"], "low": None, "high": 4.0},

    {"name": "LDL/HDL RATIO", "panel": LIPID, "unit": "",
     "ref": ["&lt;3.5"], "low": None, "high": 3.5},

    # ── Glycated Haemoglobin ──────────────────────────────────────────────────
    {"name": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)", "panel": HBA1C,
     "unit": "%", "ref": HBA1C_INTERPRETATION,
     "italic": True, "low": 3.4, "high": 6.0},

    # ── Uric Acid ─────────────────────────────────────────────────────────────
    {"name": "URIC ACID", "panel": URIC, "unit": "µmol/L",
     "ref": ["155 – 428"], "source_ref": True,
     "low": 155, "high": 428},
]

# Quick lookups built once from the library above
TEST_INFO  = {t["name"]: t for t in TEST_LIBRARY}
TEST_ORDER = {t["name"]: i for i, t in enumerate(TEST_LIBRARY)}


# ==============================================================================
# SECTION 6 — ANALYSER SHORT CODES
#
# The BS240 prints terse codes ("T-bil", "CREA", "γ-GT"). This dictionary turns
# each one into the name used in the report. Everything is lower-case here so
# matching ignores capitalisation.
#
# "__SKIP__" removes the row from the report entirely — used for the HbA1c
# sub-fractions (IFCC, eAG) that the laboratory does not report.
#
# ✏️  Add a line here whenever the analyser produces an abbreviation the report
#     does not recognise. The name on the right must exist in TEST_LIBRARY.
# ==============================================================================

TEST_CODE_MAP = {
    # ── Renal ─────────────────────────────────────────────────────────────────
    "crea": "CREATININE", "creat": "CREATININE", "creatinine": "CREATININE",
    "urea": "UREA", "bun": "UREA", "blood urea nitrogen (bun)": "UREA",
    "na+": "SODIUM", "na": "SODIUM", "sodium": "SODIUM",
    "k+": "POTASSIUM", "k": "POTASSIUM", "potassium": "POTASSIUM",
    "cl-": "CHLORIDE", "cl": "CHLORIDE", "chloride": "CHLORIDE",
    "tco2": "TOTAL CO2", "co2": "TOTAL CO2",
    "carbon dioxide, total": "TOTAL CO2", "bicarbonate": "TOTAL CO2",
    "ca": "CALCIUM", "ca2+": "CALCIUM", "calcium": "CALCIUM",
    "phos": "PHOSPHATE", "phosphate": "PHOSPHATE",
    "mg": "MAGNESIUM", "magnesium": "MAGNESIUM",
    "egfr": "eGFR (CKD-EPI)", "estimated gfr": "eGFR (CKD-EPI)",
    "egfr (ckd-epi)": "eGFR (CKD-EPI)", "egfr (ckd-epi 2021)": "eGFR (CKD-EPI)",

    # ── Liver ─────────────────────────────────────────────────────────────────
    "tp": "TOTAL PROTEIN", "tp-o": "TOTAL PROTEIN",
    "total protein": "TOTAL PROTEIN", "protein, total": "TOTAL PROTEIN",
    "alb": "ALBUMIN", "albumin": "ALBUMIN",
    "glob": "GLOBULIN", "globulin": "GLOBULIN", "globulin, total": "GLOBULIN",
    "a/g": "A/G RATIO", "a/g ratio": "A/G RATIO",
    "t-bil": "BILIRUBIN (TOTAL)", "tbil": "BILIRUBIN (TOTAL)",
    "tb": "BILIRUBIN (TOTAL)", "total bilirubin": "BILIRUBIN (TOTAL)",
    "bilirubin, total": "BILIRUBIN (TOTAL)",
    "d-bil": "BILIRUBIN (DIRECT)", "dbil": "BILIRUBIN (DIRECT)",
    "direct bil": "BILIRUBIN (DIRECT)", "direct bilirubin": "BILIRUBIN (DIRECT)",
    "ibil": "BILIRUBIN (INDIRECT)", "ind bill": "BILIRUBIN (INDIRECT)",
    "ind bil": "BILIRUBIN (INDIRECT)", "ind-bil": "BILIRUBIN (INDIRECT)",
    "indirect bilirubin": "BILIRUBIN (INDIRECT)",
    # Depending on how the analyser's PDF encodes the Greek gamma, the code can
    # come out of the text extractor as γ-GT, y-GT or g-GT — all three are here.
    "γ-gt": "GAMMA GT", "y-gt": "GAMMA GT", "g-gt": "GAMMA GT",
    "ggt": "GAMMA GT", "ggtp": "GAMMA GT", "gamma gt": "GAMMA GT",
    "gamma-gt": "GAMMA GT", "gamma glutamyl transferase": "GAMMA GT",
    "alt": "ALT (SGPT)", "sgpt": "ALT (SGPT)", "alt (sgpt)": "ALT (SGPT)",
    "alanine aminotrans": "ALT (SGPT)",
    "ast": "AST (SGOT)", "sgot": "AST (SGOT)", "ast (sgot)": "AST (SGOT)",
    "aspartate aminotrans": "AST (SGOT)",
    "ast/alt": "AST/ALT RATIO",
    "alp": "ALKALINE PHOSPHATASE",
    "alkaline phos.": "ALKALINE PHOSPHATASE",
    "alkaline phosphatase": "ALKALINE PHOSPHATASE",

    # ── Lipids ────────────────────────────────────────────────────────────────
    "tc": "TOTAL CHOLESTEROL", "chol": "TOTAL CHOLESTEROL",
    "total chol": "TOTAL CHOLESTEROL", "total chol.": "TOTAL CHOLESTEROL",
    "cholesterol": "TOTAL CHOLESTEROL",
    "cholesterol, total": "TOTAL CHOLESTEROL",
    "hdl": "HDL-CHOLESTEROL", "hdl-c": "HDL-CHOLESTEROL",
    "hdl-chol.": "HDL-CHOLESTEROL", "hdl cholesterol": "HDL-CHOLESTEROL",
    "ldl": "LDL-CHOLESTEROL", "ldl-c": "LDL-CHOLESTEROL",
    "ldl cholesterol": "LDL-CHOLESTEROL",
    "ldl chol calc (nih)": "LDL-CHOLESTEROL",
    "tg": "TRIGLYCERIDES", "trig": "TRIGLYCERIDES",
    "triglyceride": "TRIGLYCERIDES", "triglycerides": "TRIGLYCERIDES",
    "vldl": "VLDL-CHOLESTEROL", "vldl-c": "VLDL-CHOLESTEROL",
    "vldl cholesterol cal": "VLDL-CHOLESTEROL",
    "coronary r": "CORONARY RISK", "coronary risk": "CORONARY RISK",
    "chol/hdl": "CORONARY RISK",
    "ldl/hdl": "LDL/HDL RATIO",

    # ── Glycated haemoglobin — only the NGSP percentage is reported ───────────
    "a1c(ngsp)": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "a1c (ngsp)": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "ngsp": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "hba1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "hb a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "hemoglobin a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "haemoglobin a1c": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "glycated": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "glycated haemoglobin (hba1c)": "GLYCATED HAEMOGLOBIN (HB<sub>A1C</sub> – NGSP)",
    "a1c(ifcc)": "__SKIP__", "a1c (ifcc)": "__SKIP__", "ifcc": "__SKIP__",
    "eag": "__SKIP__", "estimated average glucose": "__SKIP__",

    # ── Uric acid ─────────────────────────────────────────────────────────────
    "ua": "URIC ACID", "urate": "URIC ACID", "uric acid": "URIC ACID",
}


def resolve_test_name(raw_name):
    """
    Turns an analyser code into the name used in the report.

    Returns (display_name, panel). If the code is unknown the original name is
    returned unchanged and the panel is "OTHER INVESTIGATIONS", so nothing the
    analyser measured is ever silently lost.
    """
    key = (raw_name or "").strip().lower()

    # Some parsers hand us "CODE — FULL NAME"; the code alone is what we want.
    if " — " in key:
        key = key.split(" — ")[0].strip()

    # Trailing punctuation and stray spaces are common in analyser output
    key = key.strip(" .:")

    name = TEST_CODE_MAP.get(key)
    if name is None:
        # Maybe the parser already gave us a proper report name
        for candidate in TEST_INFO:
            if candidate.lower() == key:
                name = candidate
                break

    if name is None:
        return raw_name.strip(), OTHER
    if name == "__SKIP__":
        return "__SKIP__", ""
    return name, TEST_INFO[name]["panel"]


# ==============================================================================
# SECTION 7 — TIDYING UP THE PARSED RESULTS
#
# apply_test_library() is the bridge between the parser (which reads whatever
# the analyser printed) and the report (which wants the laboratory's own names,
# units and reference intervals). It runs once, before anything is drawn.
# ==============================================================================

def _to_number(text):
    """Turns '4.01' or '<0.20' into a number, or None if it isn't numeric."""
    if text is None:
        return None
    cleaned = re.sub(r"[^\d.\-]", "", str(text))
    if cleaned in ("", "-", ".", "-."):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def compute_flag(name, value):
    """
    Works out "High" / "Low" / "" for one result.
    Only consulted when SHOW_FLAGS is True.
    """
    info = TEST_INFO.get(name)
    if not info:
        return ""
    number = _to_number(value)
    if number is None:
        return ""
    low, high = info.get("low"), info.get("high")
    if high is not None and number > high:
        return "High"
    if low is not None and number < low:
        return "Low"
    return ""


def tidy_reference(text):
    """
    Smartens a reference interval as printed by the analyser so it matches the
    look of the rest of the report: "53.0-120.0" becomes "53.0 – 120.0", and
    the characters that upset the PDF markup are escaped.
    """
    text = str(text or "").strip()
    if not text:
        return ""

    # A hyphen sitting between two numbers is a range dash, not a minus sign
    text = re.sub(r"(?<=[\d.])\s*[-–—]\s*(?=[\d.])", " – ", text)

    # "<" and "&" would be read as markup inside a Paragraph
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return text


def apply_test_library(results, patient=None):
    """
    Takes the raw list of results from the parser and returns a clean list
    ready for printing:

      • analyser codes replaced with the laboratory's own test names
      • __SKIP__ rows (HbA1c IFCC, eAG) dropped
      • units and reference intervals filled in from TEST_LIBRARY
      • tests marked source_ref keep the interval the analyser printed
      • each row tagged with the section it belongs to
      • a flag computed when SHOW_FLAGS is on

    The analyser's own unit and reference interval are kept whenever
    TEST_LIBRARY has nothing to say about that test.
    """
    cleaned = []
    seen = set()

    for raw in results:
        name, panel = resolve_test_name(raw.get("test", ""))
        if name == "__SKIP__":
            continue

        # Guard against an analyser printing the same test twice
        if name in seen:
            continue
        seen.add(name)

        row = dict(raw)
        row["test"] = name
        row["panel"] = panel

        source_ref = tidy_reference(raw.get("ref", ""))

        info = TEST_INFO.get(name)
        if info:
            row["unit"] = info["unit"]

            if info.get("source_ref") and source_ref:
                # Creatinine, ALP, uric acid: print what the analyser gave
                row["ref_lines"] = [source_ref]
            else:
                row["ref_lines"] = list(info["ref"])

            if info.get("note"):
                row["note"] = info["note"]
            if info.get("italic"):
                row["italic"] = True
        else:
            # Unknown test: show whatever the analyser printed
            row["ref_lines"] = [source_ref] if source_ref else []

        if SHOW_FLAGS and not row.get("flag"):
            row["flag"] = compute_flag(name, row.get("result", ""))

        cleaned.append(row)

    return cleaned


def group_results_by_section(results):
    """
    Sorts the cleaned results into their sections and puts each section's rows
    into the order used by the laboratory's template.

    Returns a dictionary: {"RENAL FUNCTION Test": [row, row, …], …}
    """
    grouped = {section: [] for section in SECTION_ORDER}

    for row in results:
        panel = row.get("panel", OTHER)
        if panel not in grouped:
            panel = OTHER
        grouped[panel].append(row)

    for section in grouped:
        grouped[section].sort(
            key=lambda r: TEST_ORDER.get(r["test"], 9999)
        )

    return grouped


# ==============================================================================
# SECTION 8 — TEXT STYLES
#
# Like the Styles panel in Word: define each kind of text once, then apply it
# by name everywhere it is needed.
# ==============================================================================

def styles():
    return dict(
        # Red centred section title, e.g. "RENAL FUNCTION Test"
        section_title=ParagraphStyle(
            "SectionTitle", fontName=F_TITLE, fontSize=18, leading=24,
            textColor=RED, alignment=TA_CENTER, spaceAfter=0),

        # Column headings: INVESTIGATION | VALUE | UNIT | FLAG | REFERENCE …
        col_hdr=ParagraphStyle(
            "ColHdr", fontName=F_COLHDR, fontSize=11.5, leading=14,
            textColor=colors.black, alignment=TA_CENTER),

        # Test name — left aligned, regular weight
        test_name=ParagraphStyle(
            "TestName", fontName=F_BODY, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_LEFT),

        # Test name for HbA1c — italic, as in the template
        test_name_italic=ParagraphStyle(
            "TestNameItalic", fontName=F_BODY_ITALIC, fontSize=10, leading=13,
            textColor=colors.black, alignment=TA_LEFT),

        # The measured value — centred and bold
        value=ParagraphStyle(
            "Value", fontName=F_BODY_BOLD, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_CENTER),

        # Unit — centred, regular, slightly smaller
        unit=ParagraphStyle(
            "Unit", fontName=F_BODY, fontSize=10, leading=13,
            textColor=colors.black, alignment=TA_CENTER),

        # FLAG column — bold red when there is something to say
        flag=ParagraphStyle(
            "Flag", fontName=F_BODY_BOLD, fontSize=11, leading=14,
            textColor=RED, alignment=TA_CENTER),

        # Reference interval — bold, left aligned inside its column.
        # spaceAfter puts a gap between the Female and Male lines, and between
        # the kidney-function stages, exactly as the Word template does.
        ref=ParagraphStyle(
            "Ref", fontName=F_BODY_BOLD, fontSize=10, leading=14,
            spaceAfter=8, textColor=colors.black, alignment=TA_LEFT),

        # The eGFR explanatory paragraph — bold italic, justified
        note=ParagraphStyle(
            "Note", fontName=F_BODY_BI, fontSize=9.5, leading=15,
            textColor=colors.black, alignment=TA_JUSTIFY),

        # "- -END OF REPORT- -"
        end_of_report=ParagraphStyle(
            "EndOfReport", fontName=F_BODY_BI, fontSize=11, leading=14,
            textColor=colors.black, alignment=TA_CENTER),
    )


# ==============================================================================
# SECTION 9 — THE PAGE ITSELF
#
# LabReportDoc is a document that knows how to draw the laboratory header and
# the footer on every single page. Everything above the rule and below the
# results is painted here; the results themselves flow in between.
# ==============================================================================

class LabReportDoc(BaseDocTemplate):

    def __init__(self, filename, patient, **kwargs):
        super().__init__(filename, **kwargs)
        self.patient = patient or {}
        self.logo = get_logo_image()

        frame = Frame(
            self.leftMargin, self.bottomMargin,
            self.width, self.height, id="body",
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        )
        self.addPageTemplates([
            PageTemplate(id="All", frames=frame, onPage=self._on_page)
        ])

    # ── Drawn automatically at the start of every page ────────────────────────
    def _on_page(self, canvas, doc):
        canvas.saveState()
        self._draw_header(canvas)
        self._draw_footer(canvas)
        canvas.restoreState()

    def _draw_header(self, canvas):
        p = self.patient
        page_w, page_h = self.pagesize

        # ── Logo, top left ────────────────────────────────────────────────────
        if self.logo is not None:
            try:
                canvas.drawImage(
                    self.logo,
                    LOGO_X * inch,
                    page_h - (LOGO_TOP + LOGO_H) * inch,
                    width=LOGO_W * inch, height=LOGO_H * inch,
                    mask="auto", preserveAspectRatio=True, anchor="sw",
                )
            except Exception:
                pass   # A missing logo must never stop a report printing

        # ── Laboratory name, centred ──────────────────────────────────────────
        canvas.setFont(F_HEADER, 14)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(page_w / 2.0, page_h - 0.98 * inch,
                                 safe_text(LAB_NAME, markup=False))

        # ── Patient block ─────────────────────────────────────────────────────
        canvas.setFont(F_HEADER, 11)
        x1 = HDR_TEXT_X * inch
        x2 = HDR_COL2_X * inch

        def line(y_from_top, left_text, right_text=None):
            y = page_h - y_from_top * inch
            canvas.drawString(x1, y, safe_text(left_text, markup=False))
            if right_text:
                canvas.drawString(x2, y, safe_text(right_text, markup=False))

        line(1.34, f"PATH. NO: {p.get('path_no', '')}",
                   f"SPECIMEN DATE: {p.get('specimen_date', '')}")
        line(1.66, f"PATIENT NAME: {p.get('name', '')}",
                   f"SEX: {p.get('sex', '')}")
        line(2.08, f"AGE: {p.get('age', '')} YEARS")
        line(2.31, f"SPECIMEN: {p.get('specimen', '')}")

        # ── The rule that closes the header ───────────────────────────────────
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(1.0)
        y_rule = page_h - 2.41 * inch
        canvas.line(x1, y_rule, HDR_RULE_END * inch, y_rule)

    def _draw_footer(self, canvas):
        # 0.92 inch up from the bottom edge — measured off the Word template
        y = 0.92 * inch

        printed = self.patient.get("date_printed") or \
            datetime.today().strftime("%d-%b-%y")

        canvas.setFillColor(colors.black)
        canvas.setFont(F_HEADER, 11)
        canvas.drawString(HDR_TEXT_X * inch, y, "Date printed ")
        label_w = canvas.stringWidth("Date printed ", F_HEADER, 11)
        canvas.setFont(F_BODY, 11)
        canvas.drawString(HDR_TEXT_X * inch + label_w, y, printed)

        canvas.setFont(F_HEADER, 11)
        canvas.drawString(4.55 * inch, y, safe_text(SIGNATORY, markup=False))


# ==============================================================================
# SECTION 10 — BUILDING ONE SECTION
#
# Returns the list of drawing instructions for a single red-titled block:
# the title, the column headings, and one row per result.
# ==============================================================================

# Column widths as fractions of the usable page width, measured off the
# Word template so the columns land in the same place.
COL_FRACTIONS_5 = [0.1726, 0.1921, 0.1948, 0.1813, 0.2592]  # with FLAG
COL_FRACTIONS_4 = [0.1726, 0.1921, 0.1948, 0.4405]          # without FLAG


def _ref_paragraphs(lines, s):
    """Turns one or more reference lines into a stack of paragraphs."""
    return [Paragraph(safe_text(line), s["ref"])
            for line in lines if str(line).strip()]


def build_section(section_name, rows, W, s):
    """
    Builds one section of the report.

    section_name — e.g. "RENAL FUNCTION Test"
    rows         — the results belonging to it, already in template order
    W            — usable page width in points
    s            — the styles dictionary
    """
    config = SECTIONS.get(section_name, SECTIONS[OTHER])
    show_flag = config["show_flag"]

    # ── Where the columns sit ─────────────────────────────────────────────────
    # Most sections use the five positions measured off the template. A section
    # may override them with its own widths, and ask for a left indent — the
    # indent is simply an extra empty column at the front.
    indent = config.get("indent", 0) * inch
    if config.get("cols"):
        widths = [c * inch for c in config["cols"]]
        overflow = (indent + sum(widths)) - W
        if overflow > 0:                    # Never run off the edge of the page
            widths[-1] -= overflow
    else:
        fractions = COL_FRACTIONS_5 if show_flag else COL_FRACTIONS_4
        widths = [(W - indent) * f for f in fractions]

    if indent:
        widths = [indent] + widths

    def with_indent(cells):
        """Adds the empty spacer cell when the section is indented."""
        return ([""] + cells) if indent else cells

    block = []

    # ── The red centred title ─────────────────────────────────────────────────
    block.append(Paragraph(safe_text(section_name), s["section_title"]))
    block.append(Spacer(1, 20))

    # ── Column headings ───────────────────────────────────────────────────────
    heads = ["INVESTIGATION", "VALUE", "UNIT"]
    if show_flag:
        heads.append("FLAG")
    heads.append(config["ref_label"])

    hdr_table = Table(
        [with_indent([Paragraph(h, s["col_hdr"]) for h in heads])],
        colWidths=widths,
    )
    hdr_table.setStyle(TableStyle([
        ("VALIGN",        (0, 0), (-1, -1), "BOTTOM"),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
    ]))
    block.append(hdr_table)
    block.append(Spacer(1, 14))

    # ── One table row per result ──────────────────────────────────────────────
    for row in rows:
        name_style = s["test_name_italic"] if row.get("italic") else s["test_name"]

        cells = [
            Paragraph(safe_text(row["test"]), name_style),
            Paragraph(safe_text(row.get("result", "")), s["value"]),
            Paragraph(safe_text(row.get("unit", "")), s["unit"]),
        ]
        if show_flag:
            cells.append(Paragraph(safe_text(row.get("flag", "") or ""), s["flag"]))
        cells.append(_ref_paragraphs(row.get("ref_lines", []), s))

        row_table = Table([with_indent(cells)], colWidths=widths)
        row_table.setStyle(TableStyle([
            ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",    (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING",   (0, 0), (-1, -1), 2),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
        ]))
        block.append(row_table)

        # ── The eGFR explanatory paragraph, if this row carries one ───────────
        # It sits under the first columns while the reference column above it
        # continues with the CKD staging list — exactly as in the template.
        if row.get("note"):
            note_width = sum(widths[:-1])
            note_table = Table(
                [[Paragraph(safe_text(row["note"]), s["note"])]],
                colWidths=[note_width], hAlign="LEFT",
            )
            note_table.setStyle(TableStyle([
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING",    (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING",   (0, 0), (-1, -1), 12),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
            ]))
            block.append(note_table)

    return block


# ==============================================================================
# SECTION 11 — BUILDING THE WHOLE REPORT
# ==============================================================================

def generate_report(patient, results, output_path):
    """
    Assembles and saves the finished PDF.

    patient     — dict with path_no, specimen_date, name, sex, age, specimen
    results     — the flat list of results from the parser
    output_path — where to save the PDF
    """
    doc = LabReportDoc(
        output_path, patient,
        pagesize=letter,
        leftMargin=PAGE_LEFT * inch,
        rightMargin=PAGE_RIGHT * inch,
        topMargin=PAGE_TOP * inch,
        bottomMargin=PAGE_BOTTOM * inch,
        title=f"Laboratory Report — {patient.get('name', '')}",
        author=LAB_NAME,
    )

    s = styles()
    W = doc.width
    story = []

    cleaned = apply_test_library(results, patient)
    grouped = group_results_by_section(cleaned)

    # Which sections actually have something in them
    present = [name for name in SECTION_ORDER if grouped.get(name)]

    for position, section_name in enumerate(present):
        block = build_section(section_name, grouped[section_name], W, s)

        # The closing line belongs with the last section, not adrift on a page
        # of its own — so it goes inside that section's block.
        if SHOW_END_OF_REPORT and position == len(present) - 1:
            block.append(Spacer(1, 26))
            block.append(Paragraph("- -END OF REPORT- -", s["end_of_report"]))
        else:
            block.append(Spacer(1, 26))

        # KeepTogether moves the whole section to the next page rather than
        # splitting it, unless it is too tall to fit on a page at all — in
        # which case ReportLab splits it as it must.
        story.append(KeepTogether(block))

    doc.build(story)
    print(f"Report saved: {output_path}")
    return output_path


# ==============================================================================
# SECTION 12 — DEMO DATA
# Fake but realistic data so you can see a finished report without an analyser.
# Reproduces the two Word templates the report was modelled on.
# ==============================================================================

DEMO_PATIENT = {
    "path_no":       "CHE901/26",
    "specimen_date": "27/04/2026",
    "name":          "ADWOA SAMPLE BOATENG",
    "sex":           "FEMALE",
    "age":           "46",
    "specimen":      "SERUM",
}

DEMO_RESULTS = [
    {"test": "CREATININE",           "result": "66.0"},
    {"test": "UREA",                 "result": "2.8"},
    {"test": "SODIUM",               "result": "141.4"},
    {"test": "POTASSIUM",            "result": "4.7"},
    {"test": "CHLORIDE",             "result": "103.1"},
    {"test": "eGFR (CKD-EPI)",       "result": "100"},
    {"test": "TOTAL PROTEIN",        "result": "74.6"},
    {"test": "ALBUMIN",              "result": "47.1"},
    {"test": "GLOBULIN",             "result": "27.5"},
    {"test": "BILIRUBIN (TOTAL)",    "result": "15.6"},
    {"test": "BILIRUBIN (DIRECT)",   "result": "4.9"},
    {"test": "BILIRUBIN (INDIRECT)", "result": "10.7"},
    {"test": "GAMMA GT",             "result": "14.8"},
    {"test": "ALT (SGPT)",           "result": "11.4"},
    {"test": "AST (SGOT)",           "result": "15.3"},
    {"test": "ALKALINE PHOSPHATASE", "result": "52.3"},
    {"test": "TOTAL CHOLESTEROL",    "result": "4.01"},
    {"test": "HDL-CHOLESTEROL",      "result": "1.62"},
    {"test": "LDL-CHOLESTEROL",      "result": "2.12"},
    {"test": "TRIGLYCERIDES",        "result": "0.59"},
    {"test": "CORONARY RISK",        "result": "2.5"},
]

DEMO_HBA1C_PATIENT = {
    "path_no":       "CHE902/26",
    "specimen_date": "15/07/2026",
    "name":          "AKUA EXAMPLE",
    "sex":           "FEMALE",
    "age":           "32",
    "specimen":      "WHOLE BLOOD",
}

DEMO_HBA1C_RESULTS = [
    {"test": "A1c(NGSP)", "result": "5.64"},
]

# ==============================================================================
# HELPER — build_template_patient()
#
# The parser reads whatever the analyser printed at the top of its PDF, using
# its own field names (Patient, Sample ID, Gender, Test Date …). The report
# header needs six specific fields instead:
#
#     PATH. NO · SPECIMEN DATE · PATIENT NAME · SEX · AGE · SPECIMEN
#
# This function does that translation, and tidies the values on the way:
# names and specimen types are put in capitals, the age is reduced to a plain
# number, and the date is rewritten in day/month/year order.
# ==============================================================================

# The Mindray BS240 prints dates American-style (month/day/year).
# Set this to False if your analyser already prints day/month/year.
ANALYSER_DATES_ARE_US = True


def _reformat_date(text):
    """
    Rewrites a date as DD/MM/YYYY, which is what the template shows.
    Anything it cannot make sense of is passed straight through unchanged.
    """
    if not text:
        return ""

    raw = str(text).strip()

    # Keep only the date part — analysers often append a time
    match = re.search(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{2,4})", raw)
    if not match:
        return raw

    a, b, c = match.group(1), match.group(2), match.group(3)

    # ISO order: 2026-04-27
    if len(a) == 4:
        year, month, day = a, b, c
    else:
        year = c
        if ANALYSER_DATES_ARE_US:
            month, day = a, b
        else:
            day, month = a, b
        # A "month" above 12 can only be a day — the analyser swapped them
        try:
            if int(month) > 12 >= int(day):
                day, month = month, day
        except ValueError:
            pass

    if len(year) == 2:
        year = "20" + year

    try:
        return f"{int(day):02d}/{int(month):02d}/{year}"
    except ValueError:
        return raw


def _year_suffix(date_text):
    """Returns the two-digit year used to build a path number, e.g. '26'."""
    if PATH_NO_YEAR:
        return str(PATH_NO_YEAR)[-2:]
    match = re.search(r"(\d{4})\s*$", str(date_text or "").strip())
    if match:
        return match.group(1)[-2:]
    return datetime.today().strftime("%y")


def build_template_patient(parsed):
    """
    Turns the parser's patient dictionary into the six fields the report header
    prints. The original keys are kept as well, so nothing is lost.
    """
    patient = dict(parsed or {})

    # ── SPECIMEN DATE ─────────────────────────────────────────────────────────
    # Prefer the collection date; fall back to whatever date we do have.
    raw_date = (parsed.get("collected") or parsed.get("received")
                or parsed.get("reported") or "")
    if str(raw_date).strip().upper() in ("", "N/A"):
        raw_date = ""
    specimen_date = _reformat_date(raw_date)

    # ── PATH. NO ──────────────────────────────────────────────────────────────
    # The laboratory's path number looks like "CHE901/26". The analyser usually
    # prints the "CHE901" part as the Patient ID or the Sample ID.
    path_no = ""
    for key in ("req", "id"):
        candidate = str(parsed.get(key, "")).strip()
        if candidate and candidate.upper() != "N/A":
            path_no = candidate
            break
    if path_no and "/" not in path_no:
        path_no = f"{path_no}/{_year_suffix(specimen_date)}"

    # ── AGE — digits only ─────────────────────────────────────────────────────
    age = re.sub(r"[^\d]", "", str(parsed.get("age", "")))

    # ── SEX ───────────────────────────────────────────────────────────────────
    sex = str(parsed.get("gender", "")).strip().upper()
    if sex in ("N/A", "UNKNOWN"):
        sex = ""

    # ── SPECIMEN TYPE ─────────────────────────────────────────────────────────
    specimen = str(parsed.get("sample_type", "")).strip().upper()
    if specimen in ("N/A", ""):
        specimen = "SERUM"          # The laboratory's usual specimen

    # ── PATIENT NAME ──────────────────────────────────────────────────────────
    name = str(parsed.get("name", "")).strip().upper()
    if name in ("UNKNOWN", "N/A"):
        name = ""

    patient.update({
        "path_no":       path_no,
        "specimen_date": specimen_date,
        "name":          name,
        "sex":           sex,
        "age":           age,
        "specimen":      specimen,
    })
    return patient


# ==============================================================================
# SECTION 13 — READING THE ANALYSER'S PDF
#
# Everything from here to SECTION 14 has one job: opening the PDF the analyser
# produced and turning it into a plain list of results. None of it knows
# anything about how the finished report looks.
# ==============================================================================

# ==============================================================================
# SECTION 13A — extract_text_from_pdf() FUNCTION
#
# Opens a PDF exported by the Mindray BS240 analyser and extracts all text
# AND positional word data so we can reconstruct the column layout.
#
# WHY POSITIONAL EXTRACTION?
#   The BS240 produces PDFs in (at least) two different layouts:
#
#   Layout A — Old-style (Weija Gbawe format):
#     Uses real table borders → pdfplumber's extract_tables() works perfectly.
#     Columns: Test | Full Name | Concentration | Unit | Results | Remark | Reference
#
#   Layout B — New-style (Yaa Specimen / CHE900 format):
#     NO table borders — just text positioned at fixed x-coordinates.
#     extract_tables() returns 0 tables so the old approach gets nothing.
#     The columns are identified purely by their horizontal (x) position:
#       x ≈   7  → Test name (e.g. "T-bil", "ALT", "CORONARY R")
#       x ≈ 116  → Result value (e.g. "43.7", "0.8")
#       x ≈ 228  → Unit (e.g. "U/L", "mmol/L", "μmol/L")
#       x ≈ 303  → Flag (just "H" or "L" — only present if abnormal)
#       x ≈ 398  → Reference range (e.g. "0.0-41.0", "44.0-106.0")
#
# This function now returns a dict with THREE keys:
#   "header_text"   → full plain text string (used for patient field regex)
#   "table_rows"    → rows from extract_tables() — populated in Layout A only
#   "word_rows"     → rows built from positional words — populated for Layout B
#                     Each row: [test_name, result, unit, flag, ref_range]
# ==============================================================================

def extract_text_from_pdf(pdf_path):
    """
    Extracts header text, table rows (if any), and positional word rows
    from a Mindray BS240 PDF export.

    Returns:
      {
        "header_text": "Patient: YAA SPECIMEN ...",
        "table_rows":  [...],   # non-empty only when real table borders exist
        "word_rows":   [        # always populated from positional grouping
            ["ALT",  "43.7",  "U/L",    "H", "0.0-41.0"],
            ["AST",  "34.1",  "U/L",    "",  "0.0-40.0"],
            ["CREA", "123.1", "μmol/L", "H", "44.0-106.0"],
            ...
        ]
      }

    Requires pdfplumber — install with:  pip install pdfplumber
    """

    try:
        import pdfplumber
    except ImportError:
        print("❌ pdfplumber is not installed.")
        print("   Run:  pip install pdfplumber")
        raise

    header_text = ""   # Plain text for patient field regex
    table_rows  = []   # Rows from extract_tables() — only for bordered-table PDFs
    word_rows   = []   # Rows reconstructed from x-position grouping

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            # ── Plain text: used for patient header fields ─────────────────────
            plain = page.extract_text()
            if plain:
                header_text += plain + "\n"

            # ── Table extraction: works when table borders exist ───────────────
            # This succeeds for Layout A (old Weija Gbawe style).
            # For Layout B it returns an empty list — we handle that below.
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or all(c is None or str(c).strip() == "" for c in row):
                        continue
                    cleaned = [str(c).strip() if c is not None else "" for c in row]
                    table_rows.append(cleaned)

            # ── Positional word grouping: works for both layouts ───────────────
            # extract_words() returns every word with its x,y coordinates.
            # We group words that share the same vertical position (same row)
            # then assign each word to a column based on its x coordinate.
            #
            # COLUMN BOUNDARIES (from inspecting the actual PDF word positions):
            #   x < 100     → Column 0: Test name
            #   100 ≤ x < 200 → Column 1: Result value
            #   200 ≤ x < 280 → Column 2: Unit
            #   280 ≤ x < 360 → Column 3: Flag (H or L)
            #   360 ≤ x       → Column 4: Reference range
            #
            # These boundaries work for both the old and new BS240 PDF formats
            # because the result columns are always in the same relative positions.
            #
            words = page.extract_words()

            # ── Group the words into rows ─────────────────────────────────────
            # Words that look level on the page rarely have identical
            # coordinates: a letter with a descender (the γ of "γ-GT") or a
            # slightly smaller unit sits a point or two off its neighbours.
            #
            # So rather than rounding each word to a grid — which splits a row
            # whenever two words fall either side of a boundary — we sort the
            # words down the page and start a new row only when the gap from
            # the current one is bigger than ROW_TOLERANCE.
            #
            # Each word is measured at its vertical middle rather than its top,
            # which ignores differences in letter height altogether.
            ROW_TOLERANCE = 4.0     # points

            def middle(w):
                return (w["top"] + w["bottom"]) / 2.0

            rows = []                    # list of (middle, [words])
            for word in sorted(words, key=middle):
                centre = middle(word)
                if rows and abs(centre - rows[-1][0]) <= ROW_TOLERANCE:
                    rows[-1][1].append(word)
                else:
                    rows.append((centre, [word]))

            # Process each row in top-to-bottom order
            for _centre, words_in_row in rows:
                row_words = sorted(words_in_row, key=lambda w: w["x0"])

                # Assign each word to a column slot based on its x position
                # slots[0] = test name words, slots[1] = result, etc.
                slots = ["", "", "", "", ""]   # 5 slots for 5 columns

                for word in row_words:
                    x = word["x0"]
                    text = word["text"]

                    if x < 100:
                        # Column 0: Test name — words can be multi-word so we append
                        slots[0] = (slots[0] + " " + text).strip()
                    elif x < 200:
                        # Column 1: Result value
                        slots[1] = (slots[1] + " " + text).strip()
                    elif x < 280:
                        # Column 2: Unit (e.g. "U/L", "mmol/L", "μmol/L")
                        slots[2] = (slots[2] + " " + text).strip()
                    elif x < 360:
                        # Column 3: Flag — only "H" or "L" will appear here
                        slots[3] = (slots[3] + " " + text).strip()
                    else:
                        # Column 4: Reference range (e.g. "0.0-41.0")
                        slots[4] = (slots[4] + " " + text).strip()

                # Only keep rows that have something in both col 0 (test name)
                # AND col 1 (result value), and where col 1 contains a digit.
                # This filters out the page header, column label row, footer etc.
                if slots[0] and slots[1] and re.search(r"\d", slots[1]):
                    word_rows.append(slots)

    return {
        "header_text": header_text,
        "table_rows":  table_rows,
        "word_rows":   word_rows,
    }


# ==============================================================================
# SECTION 13B — parse_pdf_data() FUNCTION
#
# Takes the extracted data from extract_text_from_pdf() and turns it into
# two clean Python structures:
#   1. patient  — a dictionary with patient details (name, age, gender etc.)
#   2. results  — a list of test result dictionaries
#
# PARSER PRIORITY (chosen automatically):
#   1. parse_bs240_table()    — when bordered table rows were found (Layout A)
#   2. parse_bs240_words()    — when positional word rows were found (Layout B)
#   3. parse_generic_fallback() — last resort for unknown formats
# ==============================================================================

def parse_pdf_data(extracted):
    """
    Master parser. Accepts the dict returned by extract_text_from_pdf().
    Returns: (patient_dict, list_of_result_dicts)

    patient dict keys:
      name, age, gender, dob, id, req, physician,
      collected, received, reported, status, fasting,
      sample_type, blood_type, bed_no, priority, diagnosis

    result dict keys:
      test, full_name, result, unit, flag, remark, ref
    """


    # ── Unpack ────────────────────────────────────────────────────────────────
    if isinstance(extracted, dict):
        header_text = extracted.get("header_text", "")
        table_rows  = extracted.get("table_rows",  [])
        word_rows   = extracted.get("word_rows",   [])
    else:
        # Backwards compatibility: plain string passed in
        header_text = extracted
        table_rows  = []
        word_rows   = []

    # ── Helper: find a regex match in header_text ─────────────────────────────
    def find(pattern, default="N/A"):
        m = re.search(pattern, header_text, re.IGNORECASE)
        return m.group(1).strip() if m else default

    # ==========================================================================
    # PART A — PATIENT INFORMATION
    #
    # This parser handles BOTH BS240 header formats:
    #
    # FORMAT 1 (Old / Weija Gbawe style):
    #   Patient: Enoch Sipitey che 320   Gender: Male   Age: 65Year   Sample ID: 26
    #   Sample: Serum    MRN.:    Zone:    Bed No.:
    #   Sent from:   Priority: No
    #   Test Date:  12/18/2023 3:36:41 PM
    #
    # FORMAT 2 (New / Yaa Specimen style):
    #   Patient: YAA SPECIMEN          Sample ID: CHE900
    #   Patient ID: CHE900               Sample Type: Serum
    #   Date of Birth:                   Bar Code: CHE900
    #   Age: 68 Year                     Collection Date: 06/23/2025
    #   Gender: Female                   Collection Time:
    #   Physician:                       Department:
    #   Diagnosis:                       Comment:
    #   Ordering Date/Time: 06/23/2025   Print Date/Time: 06/23/2025 04:18:58 PM
    # ==========================================================================

    patient = {

        # Patient name — stops at 2+ spaces OR known next-column label
        # Handles both "Patient:" and "Patient ID:" (we want the first one)
        "name": find(
            r"Patient\s*:\s*([A-Za-z][A-Za-z0-9 '\-\.]+?)(?:\s{2,}|Sample\s*ID|Patient\s*ID|\n)",
            "Unknown"
        ),

        # Age — digits followed by optional "Year"/"Yr"
        "age": find(r"Age\s*:\s*(\d+\s*(?:Year|Yr|Y)?)", "N/A"),

        # Gender
        "gender": find(r"Gender\s*:\s*(Male|Female|Other|\w+)", "N/A"),

        # Date of Birth — Format 2 has "Date of Birth:" field
        # Matches dates like 01/15/1980 or 1980-01-15
        "dob": find(
            r"Date\s*of\s*Birth\s*:\s*([\d]{1,2}[\/\-][\d]{1,2}[\/\-][\d]{2,4}|[\d]{4}[\/\-][\d]{2}[\/\-][\d]{2})",
            "N/A"
        ),

        # Sample ID — both formats use "Sample ID:"
        "id": find(r"Sample\s*ID\s*:\s*(\S+)", "N/A"),

        # Patient ID — Format 2 has a separate "Patient ID:" field
        # We use this as req/requisition number if MRN is absent
        "req": find(r"Patient\s*ID\s*:\s*(\S+)",
               find(r"MRN\.?\s*:\s*((?!Zone|Bed|Sample|Sender|\s)\S+)", "N/A")),

        # Sample type — Format 1: "Sample: Serum", Format 2: "Sample Type: Serum"
        "sample_type": find(r"Sample\s*(?:Type)?\s*:\s*(\w+)", "N/A"),

        # Physician — Format 2 has "Physician:" field
        "physician": find(r"Physician\s*:\s*([A-Za-z][A-Za-z ,\.]+?)(?:\s{2,}|\n|Department)", "N/A"),

        # Diagnosis — Format 2 has "Diagnosis:" field
        "diagnosis": find(r"Diagnosis\s*:\s*([^\n]+?)(?:\s{2,}|Comment\s*:|\s*$)", "N/A"),

        # Blood type — only if explicitly filled in
        "blood_type": find(r"Blood\s*(?:Type|Group)\s*:\s*([ABO][B]?[+\-]?(?:pos|neg|itive|ative)?)", "N/A"),

        # Bed number (old format only)
        "bed_no": find(r"Bed\s*No\.?\s*:\s*((?!Sender|Sent|\s)\S+)", "N/A"),

        # Priority (old format only)
        "priority": find(r"Priority\s*:\s*(\S+)", "N/A"),

        # Collection date —
        #   Format 1: "Test Date: 12/18/2023 3:36:41 PM"
        #   Format 2: "Collection Date: 06/23/2025" + "Ordering Date/Time: 06/23/2025"
        "collected": find(
            r"(?:Collection\s*Date|Ordering\s*Date(?:/Time)?|Test\s*Date)\s*[:/]\s*([\d/]+(?:\s+[\d:]+\s*(?:AM|PM)?)?)",
            "N/A"
        ),

        # Received date —
        #   Format 1: "Send Date: 12/18/2023 3:36:41 PM"
        #   Format 2: "Test Date/Time: 06/23/2025"
        "received": find(
            r"(?:Test\s*Date(?:/Time)?|Send\s*Date)\s*[:/]\s*([\d/]+(?:\s+[\d:]+\s*(?:AM|PM)?)?)",
            "N/A"
        ),

        # Reported / Print date
        "reported": find(
            r"Print\s*Date(?:/Time)?\s*[:/]\s*([\d/]+\s+[\d:]+\s*(?:AM|PM)?)",
            "N/A"
        ),

        # Fasting status — not present in either BS240 format
        "fasting": "N/A",

        # Always Final for a completed BS240 printout
        "status": "Final",
    }

    # ── Clean up age (remove "Year"/"Yr" suffix) ──────────────────────────────
    age_raw   = patient["age"]
    age_clean = re.sub(r"[^\d]", "", age_raw)
    patient["age"] = age_clean if age_clean else age_raw

    # ── Clean up name (strip trailing spaces) ─────────────────────────────────
    patient["name"] = patient["name"].strip()

    # ── Clean up physician — "N/A" if only whitespace was captured ───────────
    if patient["physician"].strip() in ("", "N/A"):
        patient["physician"] = "N/A"

    # ==========================================================================
    # PART B — CHOOSE RESULTS PARSER
    #
    # Priority:
    #   1. Bordered-table rows  (Layout A — old Weija Gbawe style)
    #   2. Positional word rows (Layout B — new Yaa Specimen / CHE900 style)
    #   3. Generic text fallback (unknown formats)
    # ==========================================================================

    # Count usable table rows (must have a number in the Concentration column)
    data_row_count = sum(
        1 for row in table_rows
        if len(row) >= 4 and re.search(r"\d", row[2] if len(row) > 2 else "")
    )

    # Count usable word rows (must have a number in the Result slot)
    word_row_count = sum(
        1 for row in word_rows
        if len(row) >= 2 and re.search(r"\d", row[1])
    )

    if data_row_count >= 1:
        print(f"   Using BS240 table parser  ({data_row_count} bordered-table rows)")
        results = parse_bs240_table(table_rows)

    elif word_row_count >= 1:
        print(f"   Using BS240 word parser   ({word_row_count} positional word rows)")
        results = parse_bs240_words(word_rows)

    else:
        print("   No structured data found — using generic text fallback parser")
        results = parse_generic_fallback(header_text)

    # ==========================================================================
    # PART C — A NOTE ON NAMES AND REFERENCE INTERVALS
    #
    # The parser deliberately leaves the analyser's own short codes alone
    # ("T-bil", "γ-GT", "CREA"). They are translated into the laboratory's
    # test names — and given the laboratory's units and reference intervals —
    # further along, by apply_test_library() in SECTION 7.
    #
    # Keeping the two jobs apart means the reference intervals live in exactly
    # one place (TEST_LIBRARY), which is where you go to change them.
    # ==========================================================================

    # ==========================================================================
    # PART D — CALCULATE eGFR (CKD-EPI 2021 formula)
    #
    # eGFR (estimated Glomerular Filtration Rate) measures how well the kidneys
    # are filtering waste from the blood. It is NOT printed on the BS240 report
    # because the analyser only measures Creatinine — eGFR must be calculated.
    #
    # FORMULA: CKD-EPI 2021 (race-free version, the current clinical standard)
    #
    #   eGFR = 142 × min(Scr/κ, 1)^α × max(Scr/κ, 1)^(-1.200)
    #          × 0.9938^Age × (1.012 if Female)
    #
    #   Where:
    #     Scr = serum creatinine in mg/dL
    #     κ   = 0.7 (Female) or 0.9 (Male)
    #     α   = -0.241 (Female) or -0.302 (Male)
    #     Age = patient age in years
    #
    # UNIT CONVERSION:
    #   The BS240 reports creatinine in µmol/L.
    #   To convert to mg/dL: divide by 88.42
    #
    # REFERENCE RANGES (KDIGO staging):
    #   >=90         Normal or high (G1)
    #   60-89        Mildly decreased (G2)
    #   45-59        Mildly to moderately decreased (G3a)
    #   30-44        Moderately to severely decreased (G3b)
    #   15-29        Severely decreased (G4)
    #   <15          Kidney failure (G5)
    #   We report simply: >=60 as the lower acceptable bound.
    # ==========================================================================

    results = calculate_egfr(results, patient)

    # ==========================================================================
    # PART E — TRANSLATE THE HEADER FIELDS
    #
    # The analyser labels its header fields its own way (Patient, Sample ID,
    # Gender …). build_template_patient() turns them into the six lines the
    # report prints: PATH. NO, SPECIMEN DATE, PATIENT NAME, SEX, AGE, SPECIMEN.
    # ==========================================================================

    patient = build_template_patient(patient)

    return patient, results


# ==============================================================================
# HELPER — calculate_egfr()
#
# eGFR (estimated Glomerular Filtration Rate) says how well the kidneys are
# filtering waste out of the blood. The BS240 does not print it, because the
# analyser only measures creatinine — the rest is arithmetic, done here.
#
# FORMULA — CKD-EPI 2021, the race-free version that is the current standard:
#
#   eGFR = 142 × min(Scr/κ, 1)^α × max(Scr/κ, 1)^(-1.200)
#          × 0.9938^Age × (1.012 if female)
#
#   Scr = serum creatinine in mg/dL   (the BS240 reports µmol/L: divide by 88.42)
#   κ   = 0.7 female, 0.9 male
#   α   = -0.241 female, -0.302 male
#
# The result is inserted into the list directly after creatinine, so it prints
# at the foot of the renal section as the template shows.
# ==============================================================================

CREATININE_CODES = {
    "crea", "creat", "creatinine", "cre",
}


def calculate_egfr(results, patient):
    """
    Adds an eGFR row to the results, worked out from the creatinine result.

    Returns the list unchanged if creatinine is missing, if the age is unknown,
    or if the analyser already reported an eGFR of its own.
    """
    results = list(results)

    # ── Do we already have one? ───────────────────────────────────────────────
    for r in results:
        name, _ = resolve_test_name(r.get("test", ""))
        if name == "eGFR (CKD-EPI)":
            return results

    # ── Find creatinine ───────────────────────────────────────────────────────
    crea_index = None
    crea_value = None
    crea_unit = ""
    for i, r in enumerate(results):
        code = str(r.get("test", "")).strip().lower()
        if " — " in code:
            code = code.split(" — ")[0].strip()
        if code.strip(" .:") in CREATININE_CODES:
            crea_index = i
            crea_value = r.get("result", "")
            crea_unit = str(r.get("unit", ""))
            break

    if crea_index is None:
        return results

    scr = _to_number(crea_value)
    if scr is None or scr <= 0:
        return results

    # ── Age ───────────────────────────────────────────────────────────────────
    age_text = re.sub(r"[^\d]", "", str(patient.get("age", "")))
    if not age_text:
        return results
    age = int(age_text)
    if age <= 0:
        return results

    # ── Convert creatinine to mg/dL if it came in µmol/L ──────────────────────
    unit_lower = crea_unit.lower()
    looks_like_micromol = ("mol" in unit_lower) or (not unit_lower and scr > 30)
    if looks_like_micromol:
        scr = scr / 88.42

    # ── Sex-specific constants ────────────────────────────────────────────────
    sex = str(patient.get("gender") or patient.get("sex") or "").strip().lower()
    is_female = sex.startswith("f")

    if is_female:
        kappa, alpha, sex_factor = 0.7, -0.241, 1.012
    else:
        kappa, alpha, sex_factor = 0.9, -0.302, 1.0

    ratio = scr / kappa
    egfr = (142
            * (min(ratio, 1.0) ** alpha)
            * (max(ratio, 1.0) ** -1.200)
            * (0.9938 ** age)
            * sex_factor)

    # The laboratory reports eGFR as a whole number
    egfr_value = int(round(egfr))

    results.insert(crea_index + 1, {
        "test":   "eGFR (CKD-EPI)",
        "result": str(egfr_value),
        "unit":   "ml/min/1.73m<super>2</super>",
        "flag":   "",
        "ref":    "",
    })

    return results


# ==============================================================================
# HELPER — parse_bs240_table()
#
# Used for Layout A PDFs that have real table borders.
# Columns arrive already separated by pdfplumber's extract_tables():
#   [0] Test code  [1] Full Name  [2] Concentration  [3] Unit
#   [4] Results (Normal/high/low)  [5] Remark (RFH/RFL)  [6] Reference
# ==============================================================================

def parse_bs240_table(table_rows):
    """
    Parses bordered-table rows (Layout A — old Weija Gbawe format).
    Returns a list of result dicts.
    """
    results = []

    HEADER_WORDS = {
        "test", "full name", "concentration", "unit", "results",
        "remark", "reference", "result", "fullname", "chemistry",
    }

    for row in table_rows:
        if len(row) < 4:
            continue

        test_code  = row[0] if len(row) > 0 else ""
        full_name  = row[1] if len(row) > 1 else ""
        conc       = row[2] if len(row) > 2 else ""
        unit       = row[3] if len(row) > 3 else ""
        result_txt = row[4] if len(row) > 4 else ""
        remark     = row[5] if len(row) > 5 else ""
        ref        = row[6] if len(row) > 6 else ""

        if test_code.lower().strip() in HEADER_WORDS:
            continue

        conc_clean = conc.strip()
        if not re.search(r"\d", conc_clean):
            continue

        result_lower = result_txt.strip().lower()
        if result_lower == "high" or remark.upper() in ("RFH",):
            flag = "High"
        elif result_lower == "low" or remark.upper() in ("RFL",):
            flag = "Low"
        else:
            flag = ""

        if full_name and full_name.strip() and full_name.strip() != test_code.strip():
            display_name = f"{test_code.strip()} — {full_name.strip()}"
        else:
            display_name = test_code.strip()

        results.append({
            "test":      display_name,
            "full_name": full_name.strip(),
            "result":    conc_clean,
            "unit":      unit.strip(),
            "flag":      flag,
            "remark":    remark.strip(),
            "ref":       ref.strip().replace("–", "-").replace("—", "-"),
        })

    return results


# ==============================================================================
# HELPER — parse_bs240_words()
#
# Used for Layout B PDFs — no table borders, columns identified by x-position.
# Each row in word_rows is already a 5-slot list built in extract_text_from_pdf:
#   [0] test_name  [1] result  [2] unit  [3] flag (H/L/"")  [4] ref_range
#
# This format comes from the "Yaa Specimen / CHE900" style PDF where:
#   - Flag column contains just "H" or "L" (not "Normal"/"high"/"low")
#   - Some rows have no unit and no ref range (e.g. AST/ALT ratio, CORONARY R)
#   - Test names can contain special chars: γ-GT, K+, Cl-, Na+, T-bil
# ==============================================================================

def parse_bs240_words(word_rows):
    """
    Parses positional word rows (Layout B — new CHE900 / Yaa Specimen format).
    Returns a list of result dicts.
    """
    results = []

    # Words that identify header rows or footer rows — skip these
    SKIP_NAMES = {
        "chemistry", "result", "unit", "flag", "ref", "range",
        "ordering", "operator", "reviewer", "page", "the",
        "print", "test", "collection", "patient", "gender",
        "age", "physician", "diagnosis", "sample", "date",
        "ordering date/time:", "print date/time:", "test date/time:",
        "send date/time:", "the result", "the results",
    }

    # Also skip any row where the test name contains "date" or "time" or "page"
    # These are footer/header lines, not test results
    DATE_TIME_SKIP = re.compile(r"(?:date|time|page|operator|reviewer|result\s+are)", re.IGNORECASE)

    for row in word_rows:
        if len(row) < 2:
            continue

        test_name  = row[0].strip()
        result_val = row[1].strip()
        unit       = row[2].strip() if len(row) > 2 else ""
        flag_raw   = row[3].strip() if len(row) > 3 else ""
        ref        = row[4].strip() if len(row) > 4 else ""

        # Skip header/footer rows — by exact name match
        if test_name.lower() in SKIP_NAMES:
            continue
        # Skip any row whose test name looks like a date/time label or footer line
        if DATE_TIME_SKIP.search(test_name):
            continue

        # Skip if result has no digit
        if not re.search(r"\d", result_val):
            continue

        # Normalise flag:
        #   Layout B uses "H" or "L" in the Flag column
        #   Layout A uses "Normal"/"high"/"low" in the Results column
        flag_upper = flag_raw.upper()
        if flag_upper in ("H", "HIGH", "RFH"):
            flag = "High"
        elif flag_upper in ("L", "LOW", "RFL"):
            flag = "Low"
        else:
            flag = ""

        results.append({
            "test":      test_name,
            "full_name": "",       # Layout B doesn't have a separate Full Name column
            "result":    result_val,
            "unit":      unit,
            "flag":      flag,
            "remark":    flag_raw,  # Store original H/L for reference
            "ref":       ref.replace("–", "-"),
        })

    return results


# ==============================================================================
# HELPER — parse_generic_fallback()
#
# Last-resort parser used when neither table extraction nor positional word
# grouping produces results. Scans every line of plain text with regex.
# Less accurate but works for unknown analyser formats and simple text PDFs.
# ==============================================================================

def parse_generic_fallback(text):
    """
    Regex line-scanner fallback. Returns result dicts in the same format
    as parse_bs240_table() and parse_bs240_words().
    """
    results = []

    SKIP_WORDS = {
        "test", "full name", "concentration", "unit", "results", "chemistry",
        "remark", "reference", "patient", "sample", "gender", "flag", "ref",
        "sender", "sent", "diagnosis", "test date", "send date", "print date",
        "ordering", "operator", "reviewer", "collection", "physician",
    }

    # Pattern: TEST_NAME   NUMBER   UNIT   (H|L|flag)   (ref_range)
    pattern = re.compile(
        r"^([A-Za-zγµ][A-Za-z0-9 /\-\(\)\.\+γµ]+?)\s+"
        r"([<>]?\d+\.?\d*)\s*"
        r"([µμa-zA-Z/%\.]+(?:/[a-zA-Z\.]+)?)?\s*"
        r"(Normal|high|low|H\b|L\b|RFH|RFL|\*)?\s*"
        r"([<>=]{0,2}[\d\.\-–]+(?:\s*[-–]\s*[\d\.]+)?)?",
        re.MULTILINE | re.IGNORECASE
    )

    for m in pattern.finditer(text):
        test_name  = m.group(1).strip()
        result_val = m.group(2).strip()
        unit       = m.group(3).strip() if m.group(3) else ""
        flag_word  = m.group(4).strip() if m.group(4) else ""
        ref        = m.group(5).strip() if m.group(5) else ""

        if test_name.lower() in SKIP_WORDS or len(test_name) < 2:
            continue

        flag_upper = flag_word.upper()
        if flag_upper in ("HIGH", "H", "RFH"):
            flag = "High"
        elif flag_upper in ("LOW", "L", "RFL"):
            flag = "Low"
        else:
            flag = ""

        results.append({
            "test":      test_name,
            "full_name": "",
            "result":    result_val,
            "unit":      unit,
            "flag":      flag,
            "remark":    flag_word,
            "ref":       ref.replace("–", "-"),
        })

    return results


# ==============================================================================
# SECTION 14 — main()  —  running the script from a terminal
#
# Double-clicking the file opens the window instead; this is here for batch
# work, for testing, and for anyone who prefers the command line.
# ==============================================================================

def process_pdf(input_pdf, overrides=None):
    """
    Reads one analyser PDF and returns (patient, results) ready for printing.
    `overrides` is an optional dict of header fields typed by the user, which
    always win over whatever was parsed out of the PDF.
    """
    extracted = extract_text_from_pdf(input_pdf)
    parsed, results = parse_pdf_data(extracted)
    patient = build_template_patient(parsed)

    for key, value in (overrides or {}).items():
        if str(value).strip():
            patient[key] = str(value).strip()

    return patient, results


def default_output_name(patient, input_pdf=None):
    """Builds a sensible file name, e.g. 'CHE901_26_ADWOA_SAMPLE_BOATENG.pdf'."""
    path_no = re.sub(r"[^A-Za-z0-9]+", "_", patient.get("path_no", "")).strip("_")
    name = re.sub(r"[^A-Za-z0-9]+", "_", patient.get("name", "")).strip("_")
    stem = "_".join(part for part in (path_no, name) if part)
    if not stem:
        stem = os.path.splitext(os.path.basename(input_pdf or "report"))[0]
    return f"{stem}.pdf"


def main():
    parser = argparse.ArgumentParser(
        description="Ngleshie Amanfrom Polyclinic Laboratory — report generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Sample renal / liver / lipid report:
    python lab_report.py --demo

  Sample glycated haemoglobin report:
    python lab_report.py --demo-hba1c

  A real analyser PDF:
    python lab_report.py "CHE901.pdf"

  Correcting the header by hand:
    python lab_report.py "CHE901.pdf" --path-no CHE901/26 --sex FEMALE
        --age 46 --specimen SERUM --specimen-date 27/04/2026
        --output "CHE901 ADWOA SAMPLE BOATENG.pdf"
        """,
    )

    parser.add_argument("input_pdf", nargs="?",
                        help="The PDF exported from the analyser")

    # ── Header fields — each one overrides what was read from the PDF ─────────
    parser.add_argument("--path-no",       default=None, help="e.g. CHE901/26")
    parser.add_argument("--specimen-date", default=None, help="e.g. 27/04/2026")
    parser.add_argument("--name",          default=None, help="Patient name")
    parser.add_argument("--sex",           default=None, help="MALE or FEMALE")
    parser.add_argument("--age",           default=None, help="Age in years")
    parser.add_argument("--specimen",      default=None,
                        help="e.g. SERUM or WHOLE BLOOD")

    parser.add_argument("--lab-name",  default=None, help="Laboratory name in the header")
    parser.add_argument("--signatory", default=None, help="Footer signature line")
    parser.add_argument("--output",    default=None, help="Output PDF file name")
    parser.add_argument("--flags", action="store_true",
                        help="Work out High/Low and print it in the FLAG column")

    parser.add_argument("--demo",       action="store_true",
                        help="Sample renal / liver / lipid report")
    parser.add_argument("--demo-hba1c", action="store_true",
                        help="Sample glycated haemoglobin report")

    args = parser.parse_args()

    global LAB_NAME, SIGNATORY, SHOW_FLAGS
    if args.lab_name:
        LAB_NAME = args.lab_name
    if args.signatory:
        SIGNATORY = args.signatory
    if args.flags:
        SHOW_FLAGS = True

    # ── Demo modes ────────────────────────────────────────────────────────────
    if args.demo:
        out = args.output or "demo_report.pdf"
        generate_report(DEMO_PATIENT, DEMO_RESULTS, out)
        return

    if args.demo_hba1c:
        out = args.output or "demo_hba1c_report.pdf"
        generate_report(DEMO_HBA1C_PATIENT, DEMO_HBA1C_RESULTS, out)
        return

    # ── A real analyser PDF ───────────────────────────────────────────────────
    if args.input_pdf:
        overrides = {
            "path_no":       args.path_no,
            "specimen_date": args.specimen_date,
            "name":          (args.name or "").upper() or None,
            "sex":           (args.sex or "").upper() or None,
            "age":           args.age,
            "specimen":      (args.specimen or "").upper() or None,
        }
        overrides = {k: v for k, v in overrides.items() if v}

        print(f"Reading: {args.input_pdf}")

        try:
            patient, results = process_pdf(args.input_pdf, overrides)
        except FileNotFoundError:
            print(f"File not found: {args.input_pdf}")
            print("   Check the path. On Windows, put quotes around paths with spaces.")
            return
        except Exception as exc:
            print(f"Could not read the PDF: {exc}")
            return

        if not results:
            print("No test results could be read from this PDF.")
            print("   The layout may not match what the parser expects.")
            print("   Try --demo first to confirm the rest of the script works.")
            return

        print(f"Found {len(results)} results for {patient.get('name') or 'this patient'}")

        out = args.output or default_output_name(patient, args.input_pdf)
        generate_report(patient, results, out)
        return

    parser.print_help()


# ==============================================================================
# SECTION 15 — THE WINDOW (drag-and-drop application)
#
# Double-clicking the script opens this. It takes one or more analyser PDFs,
# shows you the patient details it read from each, and writes one report per
# file.
#
#   pip install reportlab pdfplumber tkinterdnd2
#         (tkinterdnd2 is optional — without it, use the Add PDF(s) button)
# ==============================================================================

# ── COLOUR THEME — every GUI colour in one place ──────────────────────────────
THEME = {
    "bg":           "#1e1e2e",
    "card":         "#2a2a3e",
    "border":       "#3a3a5c",
    "accent":       "#7c6af7",
    "accent_hover": "#9d8fff",
    "accent_dark":  "#5a4fd4",
    "text":         "#e0e0f0",
    "text_dim":     "#8888aa",
    "text_bright":  "#ffffff",
    "success":      "#4caf84",
    "error":        "#e05555",
    "warning":      "#e0a040",
    "drop_idle":    "#252538",
    "drop_hover":   "#2a3a2a",
    "drop_border":  "#7c6af7",
    "input_bg":     "#16162a",
    "input_border": "#44446a",
}


# ==============================================================================
# COLOUR THEME
# All GUI colours in one place — edit here to restyle the whole app at once.
# ==============================================================================

THEME = {
    "bg":           "#1e1e2e",
    "card":         "#2a2a3e",
    "border":       "#3a3a5c",
    "accent":       "#7c6af7",
    "accent_hover": "#9d8fff",
    "accent_dark":  "#5a4fd4",
    "text":         "#e0e0f0",
    "text_dim":     "#8888aa",
    "text_bright":  "#ffffff",
    "success":      "#4caf84",
    "error":        "#e05555",
    "warning":      "#e0a040",
    "drop_idle":    "#252538",
    "drop_hover":   "#2a3a2a",
    "drop_border":  "#7c6af7",
    "input_bg":     "#16162a",
    "input_border": "#44446a",
}


# ==============================================================================
# APP CLASS
# The main window. All GUI logic lives here.
# ==============================================================================

def _build_gui_app():
    """
    Creates and returns the App class.
    Defined inside a function so the class body is only evaluated when
    tkinter is confirmed available — avoids errors on headless servers.
    """

    BaseClass = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

    class App(BaseClass):

        def __init__(self):
            super().__init__()
            self.title("Laboratory Report Generator")
            self.geometry("780x880")
            self.minsize(680, 720)
            self.configure(bg=THEME["bg"])
            self.resizable(True, True)

            # ── State variables ───────────────────────────────────────────────
            self.input_pdfs     = []   # PDF paths to process (multiple supported)
            self.lab_name_var   = tk.StringVar(value=LAB_NAME)
            self.signatory_var  = tk.StringVar(value=SIGNATORY)
            self.output_dir_var = tk.StringVar()
            self.filename_var   = tk.StringVar(value="auto")
            self.review_var     = tk.BooleanVar(value=True)
            self.flags_var      = tk.BooleanVar(value=SHOW_FLAGS)

            self._build_ui()
            self._load_settings()

            # Centre window on screen
            self.update_idletasks()
            x = (self.winfo_screenwidth()  - self.winfo_width())  // 2
            y = (self.winfo_screenheight() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")

        # ──────────────────────────────────────────────────────────────────────
        # UI CONSTRUCTION
        # ──────────────────────────────────────────────────────────────────────

        def _build_ui(self):
            # Outer scrollable canvas so the window can be made smaller
            self.canvas = tk.Canvas(self, bg=THEME["bg"], highlightthickness=0)
            scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
            self.canvas.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side="right", fill="y")
            self.canvas.pack(side="left", fill="both", expand=True)

            self.inner = tk.Frame(self.canvas, bg=THEME["bg"])
            self.canvas_window = self.canvas.create_window(
                (0, 0), window=self.inner, anchor="nw"
            )
            self.inner.bind("<Configure>",   self._on_frame_configure)
            self.canvas.bind("<Configure>",  self._on_canvas_configure)
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

            self._build_header()
            self._build_drop_zone()
            self._build_lab_details()
            self._build_output_section()
            self._build_generate_btn()
            self._build_status_bar()

        def _build_header(self):
            hdr = tk.Frame(self.inner, bg=THEME["accent_dark"], pady=14)
            hdr.pack(fill="x")

            tk.Label(
                hdr, text="🔬  Lab Report Generator",
                font=("Segoe UI", 18, "bold"),
                bg=THEME["accent_dark"], fg=THEME["text_bright"]
            ).pack()

            tk.Label(
                hdr,
                text="Renal (+ eGFR)  ·  Liver  ·  Lipid Profile  ·  HbA1c  ·  Uric Acid",
                font=("Segoe UI", 9),
                bg=THEME["accent_dark"], fg=THEME["text_dim"]
            ).pack()

        def _build_drop_zone(self):
            self._section_label("1.  Input PDFs  (add one or more patient files)")

            # ── Drop target area ──────────────────────────────────────────────
            self.drop_frame = tk.Frame(
                self.inner, bg=THEME["drop_idle"],
                highlightbackground=THEME["drop_border"],
                highlightthickness=2, pady=16, padx=20
            )
            self.drop_frame.pack(fill="x", padx=16, pady=(0, 6))

            self.drop_icon = tk.Label(
                self.drop_frame, text="📄",
                font=("Segoe UI", 26),
                bg=THEME["drop_idle"], fg=THEME["text_dim"]
            )
            self.drop_icon.pack()

            self.drop_label = tk.Label(
                self.drop_frame,
                text="Drag & drop PDFs here  (multiple files supported)" if DND_AVAILABLE
                     else "Use the buttons below to add patient PDF files",
                font=("Segoe UI", 10),
                bg=THEME["drop_idle"], fg=THEME["text_dim"]
            )
            self.drop_label.pack(pady=(4, 6))

            btn_row = tk.Frame(self.drop_frame, bg=THEME["drop_idle"])
            btn_row.pack()
            self._make_button(btn_row, "➕  Add PDF(s)…",
                              self._browse_input, style="secondary").pack(side="left", padx=4)
            self._make_button(btn_row, "🗑  Remove Selected",
                              self._remove_selected, style="secondary").pack(side="left", padx=4)
            self._make_button(btn_row, "✖  Clear All",
                              self._clear_all_pdfs, style="secondary").pack(side="left", padx=4)

            # Register drag-and-drop if available
            if DND_AVAILABLE:
                for widget in [self.drop_frame, self.drop_icon, self.drop_label]:
                    widget.drop_target_register(DND_FILES)
                    widget.dnd_bind("<<Drop>>",       self._on_drop)
                    widget.dnd_bind("<<DragEnter>>",  self._on_drag_enter)
                    widget.dnd_bind("<<DragLeave>>",  self._on_drag_leave)

            # ── File list box ─────────────────────────────────────────────────
            list_frame = tk.Frame(self.inner, bg=THEME["bg"])
            list_frame.pack(fill="x", padx=16, pady=(0, 8))

            self.file_listbox = tk.Listbox(
                list_frame,
                font=("Segoe UI", 9),
                bg=THEME["input_bg"], fg=THEME["text"],
                selectbackground=THEME["accent"], selectforeground=THEME["text_bright"],
                relief="flat", highlightthickness=1,
                highlightbackground=THEME["input_border"],
                height=5, activestyle="none",
                selectmode=tk.EXTENDED   # Allow Ctrl+click / Shift+click multi-select
            )
            self.file_listbox.pack(side="left", fill="x", expand=True)

            sb = ttk.Scrollbar(list_frame, orient="vertical",
                               command=self.file_listbox.yview)
            sb.pack(side="right", fill="y")
            self.file_listbox.configure(yscrollcommand=sb.set)

            # Counter label below the list
            self.file_count_label = tk.Label(
                self.inner, text="No files added yet",
                font=("Segoe UI", 8), bg=THEME["bg"], fg=THEME["text_dim"], anchor="w", padx=16
            )
            self.file_count_label.pack(fill="x")

        def _build_lab_details(self):
            self._section_label("2.  Laboratory Details")
            card = self._card()

            self._labeled_entry(card, "Lab Name *", self.lab_name_var,
                                "Printed centred at the top of every page")
            self._labeled_entry(card, "Signed by", self.signatory_var,
                                "Printed at the bottom right of every page")

            # ── Two options that change how the report is produced ────────────
            tk.Checkbutton(
                card,
                text="Check patient details before each report",
                variable=self.review_var,
                font=("Segoe UI", 9),
                bg=THEME["card"], fg=THEME["text"],
                activebackground=THEME["card"], activeforeground=THEME["text"],
                selectcolor=THEME["input_bg"], anchor="w",
            ).pack(fill="x", pady=(8, 0))

            tk.Checkbutton(
                card,
                text="Work out High / Low myself and print it in the FLAG column",
                variable=self.flags_var,
                font=("Segoe UI", 9),
                bg=THEME["card"], fg=THEME["text"],
                activebackground=THEME["card"], activeforeground=THEME["text"],
                selectcolor=THEME["input_bg"], anchor="w",
            ).pack(fill="x")

        def _build_output_section(self):
            self._section_label("3.  Output Settings")
            card = self._card()

            self.output_dir_var = tk.StringVar()
            self.filename_var   = tk.StringVar(value="auto")

            # Output folder
            row = tk.Frame(card, bg=THEME["card"])
            row.pack(fill="x", pady=(0, 8))
            tk.Label(
                row, text="Save to folder",
                font=("Segoe UI", 9), bg=THEME["card"], fg=THEME["text_dim"], width=14, anchor="w"
            ).pack(side="left")
            tk.Entry(
                row, textvariable=self.output_dir_var,
                font=("Segoe UI", 10), bg=THEME["input_bg"],
                fg=THEME["text"], insertbackground=THEME["text"],
                relief="flat", highlightthickness=1,
                highlightbackground=THEME["input_border"],
                highlightcolor=THEME["accent"]
            ).pack(side="left", fill="x", expand=True, padx=(6, 6))
            self._make_button(row, "Browse…", self._browse_output_dir,
                              style="secondary", width=8).pack(side="left")

            self._labeled_entry(card, "Filename", self.filename_var,
                                "Leave as 'auto' to name from patient name")

        def _build_generate_btn(self):
            btn_frame = tk.Frame(self.inner, bg=THEME["bg"], pady=10)
            btn_frame.pack(fill="x", padx=16)

            self.generate_btn = self._make_button(
                btn_frame, "⚡  Generate Report", self._generate,
                style="primary", height=2
            )
            self.generate_btn.pack(fill="x")

            self.progress = ttk.Progressbar(
                btn_frame, mode="indeterminate", length=200
            )
            self.progress.pack(fill="x", pady=(6, 0))

        def _build_status_bar(self):
            self.status_var = tk.StringVar(value="Ready — drag a PDF or click Browse to start")
            self.status_bar = tk.Label(
                self.inner, textvariable=self.status_var,
                font=("Segoe UI", 9), bg=THEME["border"],
                fg=THEME["text_dim"], anchor="w", padx=10, pady=5
            )
            self.status_bar.pack(fill="x", side="bottom")

        # ──────────────────────────────────────────────────────────────────────
        # WIDGET HELPERS
        # ──────────────────────────────────────────────────────────────────────

        def _section_label(self, text):
            tk.Label(
                self.inner, text=text,
                font=("Segoe UI", 10, "bold"),
                bg=THEME["bg"], fg=THEME["accent"],
                anchor="w", padx=16
            ).pack(fill="x", pady=(10, 2))

        def _card(self):
            frame = tk.Frame(
                self.inner, bg=THEME["card"],
                highlightbackground=THEME["border"],
                highlightthickness=1, padx=14, pady=10
            )
            frame.pack(fill="x", padx=16, pady=(0, 6))
            return frame

        def _labeled_entry(self, parent, label_text, variable, placeholder=""):
            row = tk.Frame(parent, bg=THEME["card"])
            row.pack(fill="x", pady=4)

            tk.Label(
                row, text=label_text,
                font=("Segoe UI", 9), bg=THEME["card"],
                fg=THEME["text_dim"], width=14, anchor="w"
            ).pack(side="left")

            entry = tk.Entry(
                row, textvariable=variable,
                font=("Segoe UI", 10), bg=THEME["input_bg"],
                fg=THEME["text"], insertbackground=THEME["text"],
                relief="flat", highlightthickness=1,
                highlightbackground=THEME["input_border"],
                highlightcolor=THEME["accent"]
            )
            entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

            if placeholder:
                self._add_placeholder(entry, variable, placeholder)

            # Auto-save settings whenever the user leaves this field
            entry.bind("<FocusOut>", lambda e: self._save_settings())

            return entry

        def _add_placeholder(self, entry, variable, placeholder):
            if not variable.get():
                entry.insert(0, placeholder)
                entry.config(fg=THEME["text_dim"])

            def on_focus_in(event):
                if entry.get() == placeholder:
                    entry.delete(0, "end")
                    entry.config(fg=THEME["text"])

            def on_focus_out(event):
                if not entry.get():
                    entry.insert(0, placeholder)
                    entry.config(fg=THEME["text_dim"])

            def on_key(event):
                if entry.cget("fg") == THEME["text_dim"]:
                    entry.delete(0, "end")
                    entry.config(fg=THEME["text"])

            entry.bind("<FocusIn>",  on_focus_in)
            entry.bind("<FocusOut>", on_focus_out)
            entry.bind("<Key>",      on_key)

        def _make_button(self, parent, text, command,
                         style="primary", width=None, height=1):
            colours = {
                "primary":   (THEME["accent"],      THEME["text_bright"], THEME["accent_hover"]),
                "secondary": (THEME["card"],         THEME["text"],        THEME["border"]),
                "success":   (THEME["success"],      THEME["text_bright"], "#3da870"),
            }
            bg, fg, hover = colours.get(style, colours["primary"])

            btn = tk.Button(
                parent, text=text, command=command,
                font=("Segoe UI", 10, "bold" if style == "primary" else "normal"),
                bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
                relief="flat", cursor="hand2",
                padx=14, pady=6 * height,
                width=width or 0
            )

            btn.bind("<Enter>", lambda e: btn.config(bg=hover))
            btn.bind("<Leave>", lambda e: btn.config(bg=bg))
            return btn

        # ──────────────────────────────────────────────────────────────────────
        # EVENT HANDLERS
        # ──────────────────────────────────────────────────────────────────────

        def _on_drop(self, event):
            # tkinterdnd2 delivers multiple files as space-separated paths
            # wrapped in {} braces. Parse them all.
            raw = event.data.strip()
            # Split on "} {" to handle multiple files from drag
            paths = [p.strip().strip("{}") for p in raw.replace("} {", "\n").split("\n") if p.strip()]
            added = 0
            for path in paths:
                if self._add_pdf(path):
                    added += 1
            if added:
                self._set_status(f"✅  Added {added} file(s)", "success")
            self._on_drag_leave(event)

        def _on_drag_enter(self, event):
            self.drop_frame.config(bg=THEME["drop_hover"],
                                   highlightbackground=THEME["success"])
            for w in [self.drop_icon, self.drop_label]:
                w.config(bg=THEME["drop_hover"])

        def _on_drag_leave(self, event):
            self.drop_frame.config(bg=THEME["drop_idle"],
                                   highlightbackground=THEME["drop_border"])
            for w in [self.drop_icon, self.drop_label]:
                w.config(bg=THEME["drop_idle"])

        def _browse_input(self):
            paths = filedialog.askopenfilenames(
                title="Select Patient PDF(s)",
                filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
            )
            added = 0
            for path in paths:
                if self._add_pdf(path):
                    added += 1
            if added:
                self._set_status(f"✅  Added {added} file(s)", "success")

        def _add_pdf(self, path):
            """Add a single PDF path to the list. Returns True if added."""
            if not path.lower().endswith(".pdf"):
                self._set_status(f"⚠️  Skipped (not a PDF): {os.path.basename(path)}", "warning")
                return False
            if not os.path.isfile(path):
                self._set_status(f"⚠️  File not found: {path}", "warning")
                return False
            if path in self.input_pdfs:
                return False   # Already in list — silently skip duplicate
            self.input_pdfs.append(path)
            self.file_listbox.insert(tk.END, os.path.basename(path))
            # Auto-fill output folder from the first file added
            if len(self.input_pdfs) == 1 and not self.output_dir_var.get():
                self.output_dir_var.set(os.path.dirname(path))
            self._update_file_count()
            return True

        def _remove_selected(self):
            """Remove whichever rows are selected in the listbox."""
            selected = list(self.file_listbox.curselection())
            for idx in reversed(selected):   # reverse so indices stay valid
                self.file_listbox.delete(idx)
                del self.input_pdfs[idx]
            self._update_file_count()

        def _clear_all_pdfs(self):
            self.input_pdfs.clear()
            self.file_listbox.delete(0, tk.END)
            self._update_file_count()
            self._set_status("File list cleared.", "info")

        def _update_file_count(self):
            n = len(self.input_pdfs)
            if n == 0:
                self.file_count_label.config(text="No files added yet", fg=THEME["text_dim"])
            elif n == 1:
                self.file_count_label.config(text="1 file ready to process", fg=THEME["success"])
            else:
                self.file_count_label.config(text=f"{n} files ready — one report per patient", fg=THEME["success"])

        def _browse_output_dir(self):
            folder = filedialog.askdirectory(title="Select Output Folder")
            if folder:
                self.output_dir_var.set(folder)

        def _get_clean_value(self, var, placeholder=""):
            val = var.get().strip()
            if val == placeholder:
                return ""
            return val

        # ──────────────────────────────────────────────────────────────────────
        # REPORT GENERATION
        # ──────────────────────────────────────────────────────────────────────

        # Generating a report happens in three steps, so that the window never
        # freezes and so the patient details can be checked in between:
        #
        #   1. _generate()        reads the PDFs in the background
        #   2. _reading_done()    shows the check-the-details window per patient
        #   3. _run_generation()  writes the PDFs in the background
        #
        # Steps 1 and 3 run off the main thread; step 2 must be on it, because
        # only the main thread is allowed to open windows.

        def _generate(self):
            # ── Check we have something to work with ──────────────────────────
            if not self.input_pdfs:
                self._set_status("Please add at least one PDF first.", "warning")
                return

            lab_name = self._get_clean_value(
                self.lab_name_var, "Printed centred at the top of every page")
            if not lab_name:
                self._set_status("The laboratory name is required.", "warning")
                messagebox.showwarning(
                    "Missing Field",
                    "Please enter the laboratory name before generating.")
                return

            signatory = self._get_clean_value(
                self.signatory_var, "Printed at the bottom right of every page")

            # Apply the settings that affect every report in this run
            global LAB_NAME, SIGNATORY, SHOW_FLAGS
            LAB_NAME   = lab_name
            SIGNATORY  = signatory or SIGNATORY
            SHOW_FLAGS = bool(self.flags_var.get())

            self.generate_btn.config(state="disabled", text="Reading…")
            self.progress.start(12)
            self._set_status(f"Reading {len(self.input_pdfs)} file(s)…", "info")

            threading.Thread(
                target=self._read_files,
                args=(list(self.input_pdfs),),
                daemon=True,
            ).start()

        # ── Step 1: read and parse every PDF (background thread) ──────────────
        def _read_files(self, pdf_paths):
            jobs, errors = [], []

            for i, input_pdf in enumerate(pdf_paths, 1):
                try:
                    self.after(0, lambda i=i, n=len(pdf_paths):
                               self._set_status(f"Reading file {i} of {n}…", "info"))

                    patient, results = process_pdf(input_pdf)
                    if not results:
                        raise ValueError("no test results could be read")
                    jobs.append({"pdf": input_pdf, "patient": patient,
                                 "results": results})

                except Exception as exc:
                    errors.append(f"{os.path.basename(input_pdf)}: {exc}")

            self.after(0, lambda: self._reading_done(jobs, errors))

        # ── Step 2: let the user check each patient's details (main thread) ───
        def _reading_done(self, jobs, errors):
            if not jobs:
                self.progress.stop()
                self.generate_btn.config(state="normal", text="⚡  Generate Report")
                self._set_status("Nothing could be read from those files.", "error")
                if errors:
                    messagebox.showerror("Nothing to generate", "\n".join(errors))
                return

            if self.review_var.get():
                self.progress.stop()
                kept = []
                for job in jobs:
                    edited = self._ask_patient_details(job["patient"], job["pdf"])
                    if edited is None:          # the user pressed Skip
                        continue
                    job["patient"] = edited
                    kept.append(job)
                jobs = kept

                if not jobs:
                    self.generate_btn.config(state="normal",
                                             text="⚡  Generate Report")
                    self._set_status("Cancelled — nothing was generated.", "info")
                    return
                self.progress.start(12)

            out_dir = (self.output_dir_var.get().strip()
                       or os.path.dirname(jobs[0]["pdf"]))

            filename_raw = self.filename_var.get().strip()
            custom_name = "" if filename_raw in (
                "", "auto", "Leave as 'auto' to name from patient name"
            ) else filename_raw
            if custom_name and not custom_name.lower().endswith(".pdf"):
                custom_name += ".pdf"

            self.generate_btn.config(state="disabled", text="Generating…")
            self._set_status(f"Writing {len(jobs)} report(s)…", "info")

            threading.Thread(
                target=self._run_generation,
                args=(jobs, out_dir, custom_name, errors),
                daemon=True,
            ).start()

        # ── The check-the-details window ──────────────────────────────────────
        def _ask_patient_details(self, patient, pdf_path):
            """
            Shows the six header fields, filled in from the PDF, so they can be
            corrected before the report is written.

            Returns the corrected dictionary, or None if the user pressed Skip.
            """
            dialog = tk.Toplevel(self)
            dialog.title("Patient Details")
            dialog.configure(bg=THEME["bg"])
            dialog.transient(self)
            dialog.resizable(False, False)

            tk.Label(
                dialog, text=os.path.basename(pdf_path),
                font=("Segoe UI", 10, "bold"),
                bg=THEME["bg"], fg=THEME["text_bright"],
                anchor="w", padx=16,
            ).pack(fill="x", pady=(12, 0))

            tk.Label(
                dialog,
                text="These six lines print at the top of the report. "
                     "Correct anything the analyser got wrong.",
                font=("Segoe UI", 8), wraplength=430, justify="left",
                bg=THEME["bg"], fg=THEME["text_dim"],
                anchor="w", padx=16,
            ).pack(fill="x", pady=(2, 8))

            body = tk.Frame(dialog, bg=THEME["card"], padx=14, pady=12)
            body.pack(fill="both", padx=16, pady=(0, 10))

            fields = [
                ("path_no",       "PATH. NO"),
                ("specimen_date", "SPECIMEN DATE"),
                ("name",          "PATIENT NAME"),
                ("sex",           "SEX"),
                ("age",           "AGE (years)"),
                ("specimen",      "SPECIMEN"),
            ]
            variables = {}

            for key, label in fields:
                row = tk.Frame(body, bg=THEME["card"])
                row.pack(fill="x", pady=3)
                tk.Label(
                    row, text=label, width=16, anchor="w",
                    font=("Segoe UI", 9),
                    bg=THEME["card"], fg=THEME["text_dim"],
                ).pack(side="left")

                var = tk.StringVar(value=str(patient.get(key, "")))
                variables[key] = var
                tk.Entry(
                    row, textvariable=var, font=("Segoe UI", 10), width=30,
                    bg=THEME["input_bg"], fg=THEME["text"],
                    insertbackground=THEME["text"], relief="flat",
                    highlightthickness=1,
                    highlightbackground=THEME["input_border"],
                    highlightcolor=THEME["accent"],
                ).pack(side="left", fill="x", expand=True)

            outcome = {"ok": False}

            def confirm():
                outcome["ok"] = True
                dialog.destroy()

            def skip():
                dialog.destroy()

            buttons = tk.Frame(dialog, bg=THEME["bg"], pady=8)
            buttons.pack(fill="x", padx=16, pady=(0, 12))
            self._make_button(buttons, "Skip this file", skip,
                              style="secondary").pack(side="left")
            self._make_button(buttons, "Looks right — generate", confirm,
                              style="primary").pack(side="right")

            dialog.bind("<Return>", lambda _e: confirm())
            dialog.bind("<Escape>", lambda _e: skip())

            # Centre the dialog over the main window
            dialog.update_idletasks()
            x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_width()) // 2
            y = self.winfo_rooty() + 120
            dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")

            dialog.grab_set()
            self.wait_window(dialog)

            if not outcome["ok"]:
                return None

            edited = dict(patient)
            for key, var in variables.items():
                value = var.get().strip()
                edited[key] = value.upper() if key in (
                    "name", "sex", "specimen") else value
            return edited

        # ── Step 3: write the PDFs (background thread) ────────────────────────
        def _run_generation(self, jobs, out_dir, custom_name, errors):
            errors  = list(errors)
            outputs = []

            for i, job in enumerate(jobs, 1):
                try:
                    self.after(0, lambda i=i, n=len(jobs):
                               self._set_status(f"Writing report {i} of {n}…", "info"))

                    patient = job["patient"]

                    if custom_name and len(jobs) == 1:
                        out_path = os.path.join(out_dir, custom_name)
                    else:
                        out_path = os.path.join(
                            out_dir, default_output_name(patient, job["pdf"]))

                    generate_report(patient, job["results"], out_path)
                    outputs.append(out_path)

                except Exception as exc:
                    errors.append(f"{os.path.basename(job['pdf'])}: {exc}")

            self.after(0, lambda: self._generation_done(outputs, errors))

        def _generation_done(self, outputs, errors):
            """Called on the GUI thread when all files are done."""
            self.progress.stop()
            self.generate_btn.config(state="normal", text="⚡  Generate Report")
            self._save_settings()

            if not outputs:
                # Nothing was written
                self._set_status(
                    f"None of the {len(errors)} file(s) could be turned into a "
                    "report." if errors else "No reports were generated.",
                    "error")
                if errors:
                    messagebox.showerror("Nothing generated", "\n".join(errors))
                return

            if errors:
                # Partial success
                messagebox.showwarning(
                    "Partial Success",
                    f"{len(outputs)} report(s) saved.\n\nFailed:\n" + "\n".join(errors)
                )

            # Success
            n = len(outputs)
            self._set_status(f"✅  {n} report(s) saved to: {os.path.dirname(outputs[0])}", "success")

            if n == 1:
                msg = f"Report saved:\n{outputs[0]}\n\nOpen it now?"
                if messagebox.askyesno("Report Ready", msg):
                    self._open_file(outputs[0])
            else:
                msg = f"{n} reports saved to:\n{os.path.dirname(outputs[0])}\n\nOpen folder?"
                if messagebox.askyesno("Reports Ready", msg):
                    self._open_file(os.path.dirname(outputs[0]))

        def _open_file(self, path):
            """Opens a file or folder using the OS default viewer/explorer."""
            import subprocess
            try:
                if sys.platform == "win32":
                    os.startfile(path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
            except Exception:
                pass

        def _set_status(self, message, level="info"):
            colours = {
                "info":    THEME["text_dim"],
                "success": THEME["success"],
                "warning": THEME["warning"],
                "error":   THEME["error"],
            }
            self.status_var.set(message)
            self.status_bar.config(fg=colours.get(level, THEME["text_dim"]))

        def _on_frame_configure(self, event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        def _on_canvas_configure(self, event):
            self.canvas.itemconfig(self.canvas_window, width=event.width)

        def _on_mousewheel(self, event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # ──────────────────────────────────────────────────────────────────────
        # SETTINGS PERSISTENCE
        # ──────────────────────────────────────────────────────────────────────

        def _settings_path(self):
            return os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "lab_report_settings.txt"
            )

        def _save_settings(self):
            try:
                with open(self._settings_path(), "w", encoding="utf-8") as f:
                    f.write(f"lab_name={self._get_clean_value(self.lab_name_var)}\n")
                    f.write(f"signatory={self._get_clean_value(self.signatory_var)}\n")
                    f.write(f"output_dir={self.output_dir_var.get()}\n")
                    f.write(f"review={int(bool(self.review_var.get()))}\n")
                    f.write(f"flags={int(bool(self.flags_var.get()))}\n")
            except Exception:
                pass

        def _load_settings(self):
            try:
                path = self._settings_path()
                if not os.path.exists(path):
                    return
                settings = {}
                with open(path, encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if "=" in line:
                            key, _, value = line.partition("=")
                            settings[key] = value
                if settings.get("lab_name"):
                    self.lab_name_var.set(settings["lab_name"])
                if settings.get("signatory"):
                    self.signatory_var.set(settings["signatory"])
                if settings.get("output_dir"):
                    self.output_dir_var.set(settings["output_dir"])
                if settings.get("review") in ("0", "1"):
                    self.review_var.set(settings["review"] == "1")
                if settings.get("flags") in ("0", "1"):
                    self.flags_var.set(settings["flags"] == "1")
            except Exception:
                pass

    return App


# ==============================================================================
# SECTION 16 — WHAT HAPPENS WHEN YOU RUN THE FILE
#
#   Double-click the file            →  the window opens
#   python lab_report.py --demo      →  sample report, no analyser needed
#   python lab_report.py "file.pdf"  →  a real report, straight to a PDF
# ==============================================================================

if __name__ == "__main__":

    if len(sys.argv) > 1:
        # Something was typed after the file name — run in command-line mode
        main()

    else:
        # Nothing typed — open the window
        if not TK_AVAILABLE:
            print("tkinter is not available on this computer, so the window")
            print("cannot open. The command line still works, for example:")
            print("    python lab_report.py --demo")
            sys.exit(0)

        # Check the two libraries the script cannot do without
        missing = []
        try:
            import reportlab           # noqa: F401  (imported only as a test)
        except ImportError:
            missing.append("reportlab")
        try:
            import pdfplumber          # noqa: F401
        except ImportError:
            missing.append("pdfplumber")

        if missing:
            message = ("Please install the missing package(s) first:\n\n"
                       "    pip install " + " ".join(missing) + "\n\n"
                       "Then run the script again.")
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Missing Packages", message)
                root.destroy()
            except Exception:
                print(message)
            sys.exit(1)

        if not DND_AVAILABLE:
            print("Drag-and-drop is off. To switch it on: pip install tkinterdnd2")

        AppClass = _build_gui_app()
        app = AppClass()
        app.mainloop()
