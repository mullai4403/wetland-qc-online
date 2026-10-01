"""
config.py - ALL tunable settings for the Wetland QC Tool live here.

You can edit this file directly, or (recommended for the .exe) create a
`user_config.json` next to config.py / the .exe that overrides any value, e.g.

    {"LAYOUTS": {"landscape": {"header": [0, 0, 1, 0.15], "legend": [0.6, 0.5, 1, 1]}},
     "WORKERS": 3}
"""
import os
import sys

APP_TITLE = "Wetland QC Online - 2019 / 2021 / 2024"
OUTPUT_EXCEL_NAME = "Wetland_QC_Result.xlsx"
REVIEW_DIR_NAME = "QC_Review"
DEBUG_LOG_NAME = "QC_OCR_Debug.txt"
IMAGE_EXTS = (".jpg", ".jpeg", ".png")

# Parallel map workers (each has its own OCR engine). 2 is safe; raise if you
# have a strong CPU and lots of RAM.
WORKERS = 2

# --------------------------------------------------------------------------
# CROP REGIONS  (x0, y0, x1, y1) as FRACTIONS of image width / height.
#   (0, 0, 1, 1) = whole image.   (0.5, 0, 1, 0.2) = top-right quarter strip.
#
# header : wetland ID, wetland name, taluk/tahasil
# legend : the block containing the 2019 / 2021 / 2024 area lines
#
# !! These are STARTING GUESSES. Use the "Calibrate on one map" button in the
# !! GUI to see the boxes drawn on your map and the OCR text they produce,
# !! then adjust the numbers below.
# --------------------------------------------------------------------------
LAYOUTS = {
    "landscape": {
        # Tuned from the supplied Maharashtra wetland map samples.
        # Top-right grey strip containing Wetland Name / Wetland ID / Taluk.
        "header": (0.615, 0.085, 0.992, 0.155),

        # Right-side Wetland Boundary legend with 2019 / 2021 / 2024 values.
        "legend": (0.835, 0.245, 0.992, 0.405),
    },
    "portrait": {
        "header": (0.00, 0.00, 1.00, 0.15),
        "legend": (0.00, 0.70, 1.00, 1.00),
    },
}
# "auto" = pick landscape/portrait by image shape, or force a layout name.
ACTIVE_LAYOUT = "auto"

# --------------------------------------------------------------------------
# OCR / preprocessing
# --------------------------------------------------------------------------
OCR_TARGET_WIDTH = 1800                  # each crop is scaled towards this width
# Preprocessing variants tried in order. Variant 1 runs for every map; the rest
# only run when a field is not yet confidently resolved.
VARIANTS = ["base", "clahe", "otsu", "sharp"]
MAX_VARIANTS = 3

# --------------------------------------------------------------------------
# Decision thresholds (OCR confidence is 0..1)
# --------------------------------------------------------------------------
HIGH_CONF = 0.90         # "very sure" reading
MATCH_MIN_CONF = 0.50    # minimum confidence to accept an exact match as YES
MIN_NO_CONF = 0.60       # minimum confidence for two agreeing readings to say NO
FUZZY_RATIO = 0.90       # text that differs but is this similar AND low-confidence -> MANUAL CHECK

# If the Wetland ID cannot be read inside the map:
#   True  -> Wetland ID check = MANUAL CHECK (strict, recommended)
#   False -> Wetland ID check = YES based on the filename (use if maps do not print the ID)
ID_REQUIRED_IN_MAP = True

# If Excel has more decimals than the map prints (Excel 12.5833, map 12.58),
# round Excel to the map's precision before comparing. 12.50 vs 12.58 stays a MISMATCH.
ROUND_EXCEL_TO_MAP_PRECISION = True

# --------------------------------------------------------------------------
# Labels searched inside the MAP (regex, case-insensitive)
# --------------------------------------------------------------------------
_T = r"(?:taluk|taluka|tahasil|tahsil|tehsil|tehasil)"
NAME_LABELS = [r"wetland\s*name", r"name\s*of\s*(?:the\s*)?wetland"]
TALUK_LABELS = [rf"\b{_T}(?:\s*[/,|]\s*{_T})*"]
# Another label that marks the end of a value on the same line (colon required)
STOP_LABEL = rf"\b(?:{_T}|district|village|state|wetland\s*(?:id|code|name))\s*[:\-\u2013=]"
# Wetland ID token pattern (2 letters + digits; O/I/l tolerated as OCR noise)
ID_TOKEN_PATTERN = r"(?<![A-Za-z0-9])[A-Za-z]{2}[0-9OoIl]{4,9}(?![A-Za-z0-9])"
# A legend line is only considered an area line if it contains one of these
AREA_KEYWORDS = r"area|\bha\b|hectare|\(ha\)"

# --------------------------------------------------------------------------
# Excel / CSV header detection (matched after lower-casing and removing
# every non-alphanumeric character).  Exact aliases first, then LOOSE regex.
# --------------------------------------------------------------------------
COLUMN_ALIASES = {
    "id": ["Wetland ID", "Wetland_ID", "Wetland Code", "Wetland_Code", "Code", "ID"],
    "name": ["Wetland Name", "Wetland_Name", "Name"],
    "taluk": ["Taluk", "Tahasil", "Tahsil", "Tehsil", "Taluka"],
    "a2021": ["2021 Area", "SAC 2021", "SAC Area", "Area 2021", "Wetland Area 2021"],
    "a2024": ["2024 Area", "NCSCM 2024", "Ground Truthed Area", "Ground Truth Area",
              "GT Area", "Area 2024"],
    "a2019": ["2019 Area", "Area 2019", "Wetland Area 2019", "NCSCM 2019"],
}
COLUMN_LOOSE = {
    "id": r"wetland.*(id|code)",
    "name": r"wetland.*name|nameofwetland",
    "taluk": r"taluk|taluka|tahasil|tahsil|tehsil|tehasil",
    "a2021": r"(2021.*(area|sac|ha))|((area|sac).*2021)",
    "a2024": r"(2024.*(area|ncscm|gt|ground))|((area|ncscm|ground).*2024)|groundtruth",
    "a2019": r"(2019.*(area|ncscm|ha))|((area|ncscm).*2019)",
}
HEADER_SCAN_ROWS = 50     # how many top rows of each sheet to scan for the header


# --------------------------------------------------------------------------
# Optional JSON override (works inside the PyInstaller .exe too)
# --------------------------------------------------------------------------
def _load_overrides():
    import json
    base = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "user_config.json")
    if not os.path.exists(path):
        return
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        g = globals()
        for k, v in data.items():
            if k in g and not k.startswith("_"):
                g[k] = v
    except Exception as e:  # never crash on a bad override file
        print("user_config.json ignored:", e)


_load_overrides()
