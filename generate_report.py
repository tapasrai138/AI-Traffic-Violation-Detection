"""
generate_report.py
Generates a fully-formatted Word document explaining the
AI Traffic Violation Detection project — code, workflow, and architecture.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── PAGE MARGINS ──────────────────────────────────────────────
section = doc.sections[0]
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(3.0)
section.right_margin  = Cm(2.5)

# ── HELPER FUNCTIONS ──────────────────────────────────────────
def set_font(run, name="Calibri", size=11, bold=False, italic=False, color=None):
    run.font.name  = name
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)

def heading(text, level=1):
    para = doc.add_heading(text, level=level)
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in para.runs:
        run.font.name  = "Calibri"
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D) if level == 1 else \
                             RGBColor(0x2E, 0x74, 0xB5) if level == 2 else \
                             RGBColor(0x5B, 0x9B, 0xD5)
    return para

def body(text, bold=False, italic=False, color=None):
    para = doc.add_paragraph()
    run  = para.add_run(text)
    set_font(run, bold=bold, italic=italic, color=color)
    para.paragraph_format.space_after = Pt(4)
    return para

def bullet(text, bold_prefix=None):
    para = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r1 = para.add_run(bold_prefix)
        set_font(r1, bold=True)
    r2 = para.add_run(text)
    set_font(r2)
    para.paragraph_format.space_after = Pt(2)

def code_block(lines):
    """Add a shaded code block paragraph for each line."""
    for line in lines:
        para = doc.add_paragraph()
        para.paragraph_format.left_indent = Cm(0.5)
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(0)
        run = para.add_run(line if line else " ")
        run.font.name  = "Courier New"
        run.font.size  = Pt(9)
        run.font.color.rgb = RGBColor(0x1E, 0x1E, 0x1E)
        # light grey shading
        pPr = para._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  "F2F2F2")
        pPr.append(shd)
    doc.add_paragraph()  # spacing after block

def table_2col(headers, rows, col_widths=(4, 10)):
    tbl = doc.add_table(rows=1+len(rows), cols=2)
    tbl.style = "Table Grid"
    # Header row
    hdr = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        run = hdr[i].paragraphs[0].runs[0]
        set_font(run, bold=True, color=(0x1F, 0x49, 0x7D))
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Data rows
    for ri, row in enumerate(rows):
        cells = tbl.rows[ri+1].cells
        for ci, val in enumerate(row):
            cells[ci].text = val
            set_font(cells[ci].paragraphs[0].runs[0])
    # Column widths
    for row in tbl.rows:
        row.cells[0].width = Inches(col_widths[0])
        row.cells[1].width = Inches(col_widths[1])
    doc.add_paragraph()

def divider():
    para = doc.add_paragraph("─" * 90)
    for run in para.runs:
        run.font.size  = Pt(7)
        run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after  = Pt(4)

# ══════════════════════════════════════════════════════════════
#  TITLE PAGE
# ══════════════════════════════════════════════════════════════
para = doc.add_paragraph()
para.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = para.add_run("\nAI-Powered Traffic Violation Detection System")
r.font.name  = "Calibri"
r.font.size  = Pt(22)
r.font.bold  = True
r.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

para2 = doc.add_paragraph()
para2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = para2.add_run("Complete Project Report — Code Explanation, Workflow & Architecture")
set_font(r2, size=13, italic=True, color=(0x44, 0x72, 0xC4))

doc.add_paragraph()

para3 = doc.add_paragraph()
para3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = para3.add_run("Tapas Rai  |  24BRS1366  |  DA2 Submission")
set_font(r3, size=11, color=(0x60, 0x60, 0x60))

doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  TABLE OF CONTENTS (manual)
# ══════════════════════════════════════════════════════════════
heading("Table of Contents", 1)
toc = [
    "1.  Project Overview",
    "2.  Tech Stack",
    "3.  Project Workflow (Step-by-Step)",
    "4.  System Architecture",
    "5.  File: database_setup.py",
    "6.  File: night_enhancement.py",
    "       6.1  is_low_light()",
    "       6.2  enhance_frame()",
    "       6.3  apply_gamma()",
    "       6.4  draw_night_mode_badge()",
    "7.  File: main.py",
    "       7.1  Imports & Configuration",
    "       7.2  select_video_source()",
    "       7.3  log_to_db()",
    "       7.4  Model Loading",
    "       7.5  person_overlap_with_zone()",
    "       7.6  analyze_bike_crop()",
    "       7.7  draw_bike_panel()",
    "       7.8  draw_stats_hud()",
    "       7.9  Main Detection Loop",
    "8.  Violation Rules",
    "9.  Night Mode — USP Explanation",
    "10. Database Schema",
    "11. One Frame — Complete Flow Diagram",
]
for line in toc:
    para = doc.add_paragraph(line)
    set_font(para.runs[0], size=11)
    para.paragraph_format.space_after = Pt(2)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  SECTION 1: PROJECT OVERVIEW
# ══════════════════════════════════════════════════════════════
heading("1. Project Overview")
body(
    "This project is an AI-powered system that detects motorcycle traffic violations "
    "from live CCTV footage or recorded video in real time. It automatically identifies "
    "the following two violations:"
)
bullet("No Helmet — ", "any rider on a motorcycle who is not wearing a helmet.")
bullet("Triple Riding — ", "more than 2 people sitting on a single motorcycle.")

body(
    "\nWhen a violation is detected, the system automatically:"
)
bullet("Saves a screenshot of the frame as evidence.")
bullet("Runs OCR (Optical Character Recognition) to read the number plate.")
bullet("Logs the violation, plate number, and confidence score into a MySQL database.")

body(
    "\nThe key USP (Unique Selling Point) added to this project is 24/7 low-light and "
    "night-time detection capability. Most traffic violations happen at night, and standard "
    "detection systems fail in the dark. This system uses software-based image enhancement "
    "(CLAHE + Gamma Correction) to brighten dark frames before sending them to the AI — "
    "without requiring any infrared cameras or special hardware."
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 2: TECH STACK
# ══════════════════════════════════════════════════════════════
heading("2. Tech Stack")
table_2col(
    ["Library / Tool", "Purpose"],
    [
        ("Python 3.13", "Main programming language"),
        ("OpenCV (cv2)", "Video capture, image processing, drawing on frames"),
        ("YOLOv8 (ultralytics)", "Real-time object detection and tracking"),
        ("helmet.pt (custom YOLO)", "Custom-trained model: detects helmets, no-helmets, license plates"),
        ("EasyOCR", "Reads license plate text from cropped images"),
        ("PyTorch", "Deep learning backend for YOLO and EasyOCR"),
        ("MySQL + mysql-connector", "Stores every violation record in a database"),
        ("Tkinter", "Built-in Python GUI — used for the file browser dialog"),
        ("NumPy", "Array math for image processing"),
        ("night_enhancement.py", "Custom module — CLAHE + gamma brightness enhancement"),
    ]
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 3: PROJECT WORKFLOW
# ══════════════════════════════════════════════════════════════
heading("3. Project Workflow (Step-by-Step)")
body(
    "The following is the complete pipeline for every frame processed by the system. "
    "Each step happens in order, roughly 25–38 times per second."
)

steps = [
    ("STEP 1 — Startup",
     "User runs 'python main.py'. The program asks: Use webcam or browse for a video file? "
     "All three AI models are loaded into memory: YOLOv8s (traffic), helmet.pt (helmet/plate), EasyOCR."),
    ("STEP 2 — Frame Capture",
     "OpenCV reads one frame from the camera or video file using cap.read(). "
     "A frame is just a single image — like one photo in a flipbook."),
    ("STEP 3 — Night Mode Check",
     "The frame is converted to grayscale and its average brightness is calculated. "
     "If brightness < 80 (out of 255), it is considered a low-light frame."),
    ("STEP 4 — Low-Light Enhancement (if needed)",
     "Dark frames are enhanced using CLAHE on the LAB color space + Gamma Correction. "
     "This brightens the image so YOLO can see people and bikes that would otherwise be invisible. "
     "Bright frames skip this step entirely."),
    ("STEP 5 — YOLOv8 Detection & Tracking",
     "The (possibly enhanced) frame is passed to YOLOv8. It detects all motorcycles (class 3) "
     "and persons (class 0) in the frame and assigns a persistent tracking ID to each motorcycle. "
     "The same bike keeps the same ID across frames."),
    ("STEP 6 — Rider-Bike Association",
     "For each detected motorcycle, the system builds an expanded 'search zone' around it "
     "(120% upward, 15% wider). Any person whose bounding box overlaps this zone by >= 20% "
     "is counted as a rider on that bike."),
    ("STEP 7 — Helmet Model Analysis",
     "The region around each bike is cropped and passed to helmet.pt. This model detects "
     "individual heads as 'withHelmet' (class 2) or 'withoutHelmet' (class 3), and also "
     "detects license plates (class 0). The head count gives a second rider count estimate."),
    ("STEP 8 — Violation Rule Check",
     "Two rules are applied: (1) If rider count > 2 → Triple Riding violation. "
     "(2) If any head detected as 'withoutHelmet' → No Helmet violation."),
    ("STEP 9 — Evidence & OCR (only on violations)",
     "When a violation is confirmed: a screenshot is saved to the evidence/ folder. "
     "EasyOCR reads the license plate text from the detected plate region. "
     "Both the violation and plate number are saved to MySQL."),
    ("STEP 10 — HUD Overlay",
     "Per-bike panels (Bike ID, rider count, helmet status, plate, violation alert) and "
     "a global stats panel (FPS, bike count, total violations, night mode status, time) "
     "are drawn on the frame."),
    ("STEP 11 — Display",
     "The annotated frame is displayed in a window using cv2.imshow(). "
     "The loop repeats from Step 2. Press Q to quit."),
]

for i, (title, desc) in enumerate(steps):
    para = doc.add_paragraph()
    r1   = para.add_run(f"  {title}:  ")
    set_font(r1, bold=True, color=(0x1F, 0x49, 0x7D))
    r2   = para.add_run(desc)
    set_font(r2)
    para.paragraph_format.space_after = Pt(6)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 4: ARCHITECTURE
# ══════════════════════════════════════════════════════════════
heading("4. System Architecture")
body(
    "The system is organized into 3 Python files, each with a distinct responsibility:"
)
table_2col(
    ["File", "Responsibility"],
    [
        ("database_setup.py", "One-time setup: creates the MySQL database and violations table."),
        ("night_enhancement.py", "Standalone module: all brightness detection and image enhancement logic."),
        ("main.py", "The main pipeline: video loop, YOLO detection, rider association, violation rules, OCR, HUD drawing."),
    ]
)

body("Data flow between components:")
code_block([
    "Camera / Video File",
    "       |",
    "       v",
    " [night_enhancement.py]",
    "   is_low_light(frame) --> if dark --> enhance_frame(frame)",
    "       |",
    "       v",
    " [main.py — YOLOv8s]",
    "   Detect: motorcycles (class 3) + persons (class 0)",
    "   Track:  persistent ID per motorcycle",
    "       |",
    "       v",
    " [main.py — Rider Association]",
    "   Expand bike zone -> overlap check -> rider_count (YOLOv8)",
    "       |",
    "       v",
    " [main.py — helmet.pt]",
    "   Crop bike region -> detect heads -> rider_count + no_helmet flag + plate boxes",
    "       |",
    "       v",
    " [main.py — Violation Rules]",
    "   rider_count > 2  --> Triple Riding",
    "   no_helmet = True --> No Helmet",
    "       |",
    "       v",
    " [main.py — On Violation Only]",
    "   cv2.imwrite()  -> save evidence screenshot",
    "   EasyOCR        -> read license plate",
    "   log_to_db()    -> save to MySQL",
    "       |",
    "       v",
    " [main.py — HUD]",
    "   draw_bike_panel() + draw_stats_hud() + draw_night_mode_badge()",
    "       |",
    "       v",
    " cv2.imshow() --> display to screen",
    "       |",
    "       ^--- loop back",
])

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 5: database_setup.py
# ══════════════════════════════════════════════════════════════
heading("5. File: database_setup.py")
body(
    "This file is run once before starting the project. It connects to MySQL and creates the "
    "database and table needed to store violations. You run it with: python database_setup.py"
)

heading("Full Code", 2)
code_block([
    "import mysql.connector",
    "",
    "def create_db():",
    "    try:",
    "        db = mysql.connector.connect(",
    "            host='localhost',",
    "            user='root',",
    "            password='@ritesh'",
    "        )",
    "        cursor = db.cursor()",
    "",
    "        cursor.execute('CREATE DATABASE IF NOT EXISTS traffic_ai')",
    "        cursor.execute('USE traffic_ai')",
    "",
    "        cursor.execute('''",
    "        CREATE TABLE IF NOT EXISTS violations (",
    "            id              INT AUTO_INCREMENT PRIMARY KEY,",
    "            violation_type  VARCHAR(50),",
    "            plate_number    VARCHAR(20),",
    "            confidence      FLOAT,",
    "            timestamp       DATETIME DEFAULT CURRENT_TIMESTAMP,",
    "            location        VARCHAR(100) DEFAULT 'Main Gate Camera 1',",
    "            evidence_image  VARCHAR(255) DEFAULT NULL",
    "        )''')",
    "",
    "        print('Database and Table Created Successfully!')",
    "        db.close()",
    "    except Exception as e:",
    "        print(f'Error: {e}')",
    "",
    "if __name__ == '__main__':",
    "    create_db()",
])

heading("Line-by-Line Explanation", 2)
body("import mysql.connector")
body(
    "Imports the MySQL connector library — a bridge between Python and the MySQL server. "
    "Without this, Python cannot communicate with MySQL.",
    italic=True
)

body("mysql.connector.connect(host='localhost', user='root', password='@ritesh')")
body(
    "Opens a connection to the MySQL database server. 'localhost' means it's running on your "
    "own computer. This is like logging into MySQL from within Python code. The connection "
    "object 'db' is used for all further operations.",
    italic=True
)

body("cursor = db.cursor()")
body(
    "Creates a cursor object. Think of it as a 'pen' that can write SQL commands to the database. "
    "All SQL queries are sent through this cursor.",
    italic=True
)

body("cursor.execute('CREATE DATABASE IF NOT EXISTS traffic_ai')")
body(
    "'IF NOT EXISTS' means the command won't fail if the database already exists. "
    "Safe to run multiple times.",
    italic=True
)

body("violation_type VARCHAR(50)")
body("Stores what type of violation occurred — 'No Helmet', 'Triple Riding', or both.", italic=True)

body("timestamp DATETIME DEFAULT CURRENT_TIMESTAMP")
body(
    "Automatically fills in the current date and time when a row is inserted. "
    "You never need to set this manually.",
    italic=True
)

body("evidence_image VARCHAR(255) DEFAULT NULL")
body(
    "Stores the file path to the saved screenshot of the violation frame. "
    "NULL means no screenshot was saved (fallback).",
    italic=True
)

body("if __name__ == '__main__': create_db()")
body(
    "This is a standard Python pattern. It means: only call create_db() when this file is "
    "run directly. If another file imports this file, create_db() will NOT automatically run.",
    italic=True
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 6: night_enhancement.py
# ══════════════════════════════════════════════════════════════
heading("6. File: night_enhancement.py")
body(
    "This is a custom module created specifically for this project's USP. It adds night-time "
    "detection capability using classical image processing — no retraining of any AI model is "
    "needed. The module has 4 functions, each with a single clear responsibility."
)
body(
    "It uses only OpenCV and NumPy, which are already installed for the main project. "
    "No additional libraries required."
)

heading("6.1  is_low_light(frame, threshold=80)", 2)
code_block([
    "def is_low_light(frame, threshold=80):",
    "    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)",
    "    brightness = float(gray.mean())",
    "    return brightness < threshold",
])
body("Purpose: Determines whether a frame needs brightness enhancement.")
body(
    "cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)  —  Converts the 3-channel colour image "
    "(Blue, Green, Red) into a 1-channel grayscale image. Each pixel becomes a single "
    "number from 0 (pure black) to 255 (pure white).",
    italic=True
)
body(
    "gray.mean()  —  Calculates the average brightness of all pixels in the image. "
    "A typical night CCTV frame might average around 30–60 out of 255.",
    italic=True
)
body(
    "return brightness < threshold  —  Returns True (dark, needs enhancement) if the "
    "average pixel value is below 80. This threshold was chosen empirically — it works "
    "well for most CCTV setups. Can be tuned by changing LOW_LIGHT_THRESHOLD at the top of the file.",
    italic=True
)

heading("6.2  enhance_frame(frame)", 2)
code_block([
    "def enhance_frame(frame):",
    "    # Step 1: Convert BGR -> LAB color space",
    "    lab = cv2.cvtColor(frame, cv2.COLOR_BGR2LAB)",
    "    l_channel, a_channel, b_channel = cv2.split(lab)",
    "",
    "    # Step 2: Apply CLAHE to only the L (Lightness) channel",
    "    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))",
    "    l_enhanced = clahe.apply(l_channel)",
    "",
    "    # Step 3: Merge back and convert to BGR",
    "    lab_enhanced = cv2.merge((l_enhanced, a_channel, b_channel))",
    "    enhanced = cv2.cvtColor(lab_enhanced, cv2.COLOR_LAB2BGR)",
    "",
    "    # Step 4: Apply gamma correction for extra brightness",
    "    enhanced = apply_gamma(enhanced, gamma=1.4)",
    "    return enhanced",
])
body("Purpose: Brightens a dark frame using a 4-step pipeline. This is the core of night mode.")

body(
    "STEP 1 — BGR to LAB Conversion:"
)
body(
    "LAB is a colour space where L = Lightness, A = green-to-red axis, B = blue-to-yellow axis. "
    "By converting to LAB, we can edit ONLY the brightness (L channel) without changing any colours. "
    "If we simply boosted brightness in BGR space, colours would shift and distort.",
    italic=True
)
body("STEP 2 — CLAHE (Contrast Limited Adaptive Histogram Equalization):")
body(
    "Normal histogram equalization spreads pixel values across 0–255 uniformly, but this "
    "causes bright areas (like headlights) to blow out completely. CLAHE divides the image "
    "into small 8x8 tiles and equalizes each tile independently. The 'Contrast Limited' part "
    "caps the maximum amplification (clipLimit=2.0) to prevent noise from being amplified. "
    "Result: dark areas become visible while bright areas remain controlled.",
    italic=True
)
body("STEP 3 — Merge and convert back:")
body(
    "The enhanced L channel is recombined with the unchanged A and B channels, then converted "
    "back to BGR for OpenCV to display. Colours are perfectly preserved.",
    italic=True
)
body("STEP 4 — Gamma Correction (gamma=1.4):")
body(
    "An additional overall brightness boost. The formula is: output = (input/255)^(1/gamma) * 255. "
    "With gamma=1.4, mid-tones get lifted significantly while very bright pixels (near 255) "
    "are barely changed — preventing overexposure.",
    italic=True
)

heading("6.3  apply_gamma(frame, gamma=1.4)", 2)
code_block([
    "def apply_gamma(frame, gamma=1.4):",
    "    inv_gamma = 1.0 / gamma",
    "    table = np.array(",
    "        [((i / 255.0) ** inv_gamma) * 255 for i in range(256)],",
    "        dtype=np.uint8",
    "    )",
    "    return cv2.LUT(frame, table)",
])
body("Purpose: Applies a brightness curve to every pixel efficiently using a lookup table.")
body(
    "inv_gamma = 1.0 / 1.4 = 0.714  —  The exponent used in the gamma formula.",
    italic=True
)
body(
    "table  —  A list of 256 values (one for each possible pixel brightness 0–255). "
    "For each input value i, the output is (i/255)^0.714 * 255. This precomputes all possible "
    "outputs so we never do math per-pixel in a slow loop.",
    italic=True
)
body(
    "cv2.LUT(frame, table)  —  'Lookup Table'. Replaces every pixel in the frame with its "
    "corresponding value from the table. Extremely fast — processes the entire frame in one C++ call.",
    italic=True
)

heading("6.4  draw_night_mode_badge(frame, is_active)", 2)
code_block([
    "def draw_night_mode_badge(frame, is_active):",
    "    h, w = frame.shape[:2]",
    "    label = 'NIGHT MODE: ON' if is_active else 'NIGHT MODE: OFF'",
    "    color = (0, 200, 255) if is_active else (0, 200, 0)",
    "    cv2.rectangle(frame, (10, h-45), (240, h-10), (20,20,20), -1)",
    "    cv2.rectangle(frame, (10, h-45), (240, h-10), color, 1)",
    "    cv2.putText(frame, label, (18, h-18), cv2.FONT_HERSHEY_SIMPLEX, 0.55, color, 1)",
])
body("Purpose: Draws a badge in the bottom-left corner showing Night Mode status.")
body(
    "frame.shape[:2]  —  Returns (height, width). The badge is positioned h-45 from the top, "
    "which is 45 pixels from the bottom regardless of frame resolution.",
    italic=True
)
body(
    "cv2.rectangle(..., -1)  —  The -1 as the last argument means 'fill the rectangle solid'. "
    "Positive numbers draw only the border with that thickness.",
    italic=True
)
body(
    "OpenCV uses BGR colour order (not RGB). So (0, 200, 255) = Blue=0, Green=200, Red=255 "
    "which appears orange. (0, 200, 0) = pure green.",
    italic=True
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 7: main.py
# ══════════════════════════════════════════════════════════════
heading("7. File: main.py")
body(
    "This is the main brain of the entire system. It ties together all the components: "
    "video capture, night enhancement, YOLOv8 detection, rider counting, helmet checking, "
    "violation logging, OCR, evidence saving, and on-screen display."
)

heading("7.1  Imports & Configuration", 2)
code_block([
    "import cv2                    # Video/image processing",
    "import os                     # File and folder operations",
    "import time                   # For FPS calculation",
    "import easyocr                # OCR engine for license plates",
    "import mysql.connector        # MySQL database connection",
    "from ultralytics import YOLO  # YOLOv8 object detection",
    "import torch                  # Deep learning + GPU check",
    "from datetime import datetime # Timestamps for filenames",
    "import tkinter as tk          # GUI library (file dialog only)",
    "from tkinter import filedialog",
    "from night_enhancement import is_low_light, enhance_frame, draw_night_mode_badge",
])
body(
    "Each import statement loads a library — a collection of pre-written code. "
    "The last line imports 3 specific functions from our own night_enhancement.py module."
)

code_block([
    "TRAFFIC_MODEL_PATH = 'yolov8s.pt'",
    "if not os.path.exists(TRAFFIC_MODEL_PATH):",
    "    TRAFFIC_MODEL_PATH = 'yolov8n.pt'",
    "",
    "MOTO_CONF   = 0.30   # Min confidence to accept a motorcycle detection",
    "PERSON_CONF = 0.25   # Min confidence to accept a person detection",
    "HELMET_CONF = 0.30   # Min confidence for helmet model detections",
    "",
    "CLS_LICENSE_PLATE  = 0   # helmet.pt class: license plate",
    "CLS_WITH_HELMET    = 2   # helmet.pt class: rider WITH helmet",
    "CLS_WITHOUT_HELMET = 3   # helmet.pt class: rider WITHOUT helmet",
    "",
    "DEBUG = False   # Toggle with D key while running",
])
body(
    "yolov8s.pt (Small) is preferred over yolov8n.pt (Nano) — it detects more accurately. "
    "The fallback ensures the program works even if yolov8s hasn't been downloaded yet."
)
body(
    "Confidence thresholds control how certain YOLO must be before accepting a detection. "
    "Lower values catch more objects but may include false positives. "
    "PERSON_CONF = 0.25 is deliberately low to catch riders at difficult angles."
)
body(
    "The class IDs (0, 2, 3) for the helmet model were determined by loading the model "
    "and printing model.names — this is the ground truth, not an assumption."
)

heading("7.2  select_video_source()", 2)
code_block([
    "def select_video_source():",
    "    choice = input('Enter choice (1 or 2): ').strip()",
    "    if choice == '2':",
    "        root = tk.Tk()",
    "        root.withdraw()               # Hide the blank tkinter window",
    "        root.attributes('-topmost', True)  # Show dialog on top",
    "        path = filedialog.askopenfilename(",
    "            title='Select Video File',",
    "            filetypes=[('Video Files', '*.mp4 *.avi *.mov *.mkv')]",
    "        )",
    "        root.destroy()",
    "        if path: return path",
    "        return 0   # No file chosen -> use webcam",
    "    return 0       # User chose webcam (or invalid input)",
])
body("Purpose: Asks the user at startup whether to use webcam or a video file.")
body(
    "tkinter is Python's built-in GUI library. We use it only for the file browser dialog. "
    "root.withdraw() hides the main tkinter window so only the file picker appears. "
    "askopenfilename() opens a standard Windows file browser and returns the selected path.",
    italic=True
)
body(
    "Returning 0 is OpenCV's convention for 'use the first connected webcam'. "
    "Any string path opens that video file instead.",
    italic=True
)

heading("7.3  log_to_db()", 2)
code_block([
    "def log_to_db(violation_type, plate_text, bike_id, conf):",
    "    try:",
    "        db = mysql.connector.connect(",
    "            host='localhost', user='root',",
    "            password=DB_PASSWORD, database='traffic_ai'",
    "        )",
    "        cur = db.cursor()",
    "        cur.execute(",
    "            'INSERT INTO violations (violation_type, plate_number, location, confidence)'",
    "            ' VALUES (%s, %s, %s, %s)',",
    "            (violation_type, plate_text, f'Bike_ID_{bike_id}', float(conf))",
    "        )",
    "        db.commit()",
    "        db.close()",
    "    except Exception as e:",
    "        print(f'[DB ERROR] {e}')",
])
body("Purpose: Saves one violation record to MySQL. Called only when a violation is confirmed.")
body(
    "A new connection is opened for each violation. This is simple and safe — no need to "
    "maintain a persistent connection for an infrequent operation.",
    italic=True
)
body(
    "The %s placeholders are filled safely by the tuple in the second argument. "
    "This prevents SQL injection attacks — user data is never directly inserted into the SQL string.",
    italic=True
)
body(
    "db.commit() is REQUIRED. Without it, the INSERT is never actually saved to disk. "
    "It's like pressing Save after typing in a document.",
    italic=True
)
body(
    "The try/except ensures that if MySQL is not running, the program does not crash — "
    "it just prints the error and continues processing the video.",
    italic=True
)

heading("7.4  Model Loading", 2)
code_block([
    "traffic_model = YOLO('yolov8s.pt')",
    "traffic_model.to(device)   # Move to GPU if available, else CPU",
    "",
    "try:",
    "    helmet_model = YOLO('helmet.pt')",
    "    helmet_model.to(device)",
    "    HAS_HELMET = True",
    "except:",
    "    HAS_HELMET = False   # Run without helmet checks if file missing",
    "",
    "ocr = easyocr.Reader(['en'], gpu=(device == 'cuda'))",
])
body(
    "YOLO('yolov8s.pt') loads the neural network weights from disk into memory. "
    "First run auto-downloads the file from Ultralytics servers if not present locally."
)
body(
    ".to(device) moves the model's parameters to the selected hardware. "
    "On GPU this is essential for speed; on CPU it's a no-op. "
    "device = 'cuda' if torch.cuda.is_available() else 'cpu'.",
    italic=True
)
body(
    "EasyOCR is initialized with ['en'] for English. gpu=True only if CUDA is available — "
    "OCR on GPU is 5–10x faster than CPU. With this project running on CPU, OCR takes ~140ms "
    "per invocation but is only called on confirmed violations, keeping it off the critical path.",
    italic=True
)

heading("7.5  person_overlap_with_zone()", 2)
code_block([
    "def person_overlap_with_zone(person_box, zone_box):",
    "    ax1,ay1,ax2,ay2 = person_box",
    "    bx1,by1,bx2,by2 = zone_box",
    "    ix1 = max(ax1, bx1);  iy1 = max(ay1, by1)   # Top-left of intersection",
    "    ix2 = min(ax2, bx2);  iy2 = min(ay2, by2)   # Bottom-right of intersection",
    "    inter = max(0, ix2-ix1) * max(0, iy2-iy1)   # Area of overlap",
    "    area_person = max(1, (ax2-ax1) * (ay2-ay1)) # Area of person box",
    "    return inter / area_person",
])
body(
    "Purpose: Calculates what fraction of a person bounding box lies inside the bike's search zone. "
    "If this is >= 0.20 (20%), that person is counted as a rider on this bike."
)
body(
    "A bounding box is defined by its top-left corner (x1, y1) and bottom-right corner (x2, y2). "
    "The intersection of two boxes is found by taking the maximum of the left edges (max of x1s) "
    "and the minimum of the right edges (min of x2s). max(0, ...) handles the case where the "
    "boxes don't overlap at all — they would produce a negative area without this.",
    italic=True
)
body(
    "Why 20%? Because a rider sitting on a motorcycle will have their lower body overlapping "
    "the bike zone significantly, but their head and upper body may extend well above it. "
    "20% is enough to confirm association without being too strict.",
    italic=True
)

heading("7.6  analyze_bike_crop()", 2)
code_block([
    "def analyze_bike_crop(crop, offset_x, offset_y):",
    "    rider_count, no_helmet, plate_boxes = 0, False, []",
    "    if not HAS_HELMET or crop is None or crop.size == 0:",
    "        return rider_count, no_helmet, plate_boxes",
    "",
    "    results = helmet_model(crop, verbose=False)",
    "    for r in results:",
    "        for box in r.boxes:",
    "            cls  = int(box.cls[0])",
    "            conf = float(box.conf[0])",
    "            if conf < HELMET_CONF: continue",
    "",
    "            if cls == CLS_WITH_HELMET:       rider_count += 1",
    "            elif cls == CLS_WITHOUT_HELMET:  rider_count += 1; no_helmet = True",
    "            elif cls == CLS_LICENSE_PLATE:",
    "                x1,y1,x2,y2 = map(int, box.xyxy[0])",
    "                plate_boxes.append((x1+offset_x, y1+offset_y,",
    "                                    x2+offset_x, y2+offset_y))",
    "    return rider_count, no_helmet, plate_boxes",
])
body(
    "Purpose: Runs the helmet model on a cropped region around one bike. "
    "Returns rider count, whether any rider has no helmet, and plate box locations."
)
body(
    "crop is a small sub-image cut from the full frame. It contains only the bike and "
    "the space above it where riders sit. Running the model on a small crop is faster "
    "than running it on the entire frame.",
    italic=True
)
body(
    "offset_x, offset_y are the coordinates where the crop starts in the full frame. "
    "The helmet model returns box coordinates relative to the crop (starting from 0,0). "
    "Adding the offset converts them back to full-frame coordinates for drawing.",
    italic=True
)
body(
    "The rider count from this function counts HEADS detected by the helmet model. "
    "Each head is either 'withHelmet' (class 2) or 'withoutHelmet' (class 3). "
    "Both increment rider_count. Only 'withoutHelmet' sets no_helmet = True.",
    italic=True
)

heading("7.7  draw_bike_panel()", 2)
code_block([
    "def draw_bike_panel(frame, bx1,by1,bx2,by2, bid, rider_count, violations, plate_text=''):",
    "    is_viol   = len(violations) > 0",
    "    box_color = (0,0,220) if is_viol else (0,200,80)  # Red if violation, Green if safe",
    "    cv2.rectangle(frame, (bx1,by1), (bx2,by2), box_color, 3)",
    "",
    "    # Semi-transparent background panel",
    "    ov = frame.copy()",
    "    cv2.rectangle(ov, (bx1,py1), (px2,py2), (15,15,15), -1)",
    "    cv2.addWeighted(ov, 0.75, frame, 0.25, 0, frame)",
    "",
    "    # Draw text lines",
    "    cv2.putText(frame, f'Bike #{bid}', (bx1+6, y), FONT, 0.47, (180,220,255), 1)",
    "    cv2.putText(frame, f'Riders: {rider_count}', (bx1+6, y+22), FONT, 0.47, rc, 1)",
    "    ...",
])
body("Purpose: Draws an info panel above each detected motorcycle showing its status.")
body(
    "The semi-transparent background is achieved by: (1) copying the frame, (2) drawing a "
    "solid dark rectangle on the copy, (3) blending the copy with the original using "
    "cv2.addWeighted(). A weight of 0.75 means 75% from the copy (dark rectangle) and "
    "25% from the original frame beneath — giving a see-through dark panel.",
    italic=True
)
body(
    "cv2.LINE_AA enables anti-aliasing — sub-pixel rendering that makes text edges smooth "
    "instead of jagged. Important for readability at small font sizes.",
    italic=True
)

heading("7.8  draw_stats_hud()", 2)
body(
    "Draws a live statistics panel in the top-right corner of every frame. Shows:"
)
bullet("FPS — frames per second (smoothed using exponential moving average)")
bullet("Bikes in frame — current number of tracked motorcycles")
bullet("Violations — total violations detected this session (turns red when > 0)")
bullet("Night Mode — ON (orange) or OFF (green)")
bullet("Model — which YOLO model is running (yolov8s or yolov8n)")
bullet("Time — current system time in HH:MM:SS")
body(
    "Uses the same semi-transparent overlay technique as draw_bike_panel."
)

heading("7.9  Main Detection Loop", 2)
code_block([
    "cap = cv2.VideoCapture(VIDEO_SOURCE)   # Open camera or video file",
    "cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)",
    "cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)",
    "",
    "while True:",
    "    ret, frame = cap.read()   # Read next frame",
    "    if not ret: break         # End of video or disconnected",
    "",
    "    # FPS (exponential moving average — smooth display)",
    "    fps = 0.9*fps + 0.1*(1.0 / (now - prev_t + 1e-6))",
    "",
    "    # Night mode",
    "    if is_low_light(frame):",
    "        frame = enhance_frame(frame)",
    "",
    "    # YOLO detection + tracking",
    "    res = traffic_model.track(frame, persist=True, classes=[0,3])",
    "",
    "    # Expand search zone per bike (120% upward, 15% sideways)",
    "    zone_y1 = max(by1 - int(bh * 1.20), 0)",
    "",
    "    # Count riders (YOLOv8 persons overlap zone >= 20%)",
    "    riders_yolo = sum(1 for p in persons if person_overlap_with_zone(p, zone) >= 0.20)",
    "",
    "    # Helmet model on bike crop",
    "    rider_count_helmet, no_helmet, plate_boxes = analyze_bike_crop(crop, cx1, cy1)",
    "",
    "    # Final count = max of both methods",
    "    rider_count = max(rider_count_helmet, riders_yolo)",
    "",
    "    # Violation rules",
    "    if rider_count > 2: violations.append('Triple Riding')",
    "    if no_helmet:       violations.append('No Helmet')",
    "",
    "    # On violation: save screenshot + OCR plate + log to DB",
    "    if violations:",
    "        cv2.imwrite(ev_path, frame)",
    "        for (_,txt,prob) in ocr.readtext(plate_crop):",
    "            if prob > 0.30 and len(clean) > 3:",
    "                log_to_db(vtext, clean, bid, prob)",
    "",
    "    key = cv2.waitKey(1) & 0xFF",
    "    if key == ord('q'): break",
    "    elif key == ord('d'): DEBUG = not DEBUG",
])

body(
    "persist=True in traffic_model.track() tells the tracker to maintain the same ID "
    "for each motorcycle across consecutive frames. This is critical — without it, a bike "
    "would get a new random ID every frame, making it impossible to track violations over time "
    "or avoid logging the same violation multiple times."
)
body(
    "classes=[0, 3] restricts YOLO to only detect class 0 (person) and class 3 (motorcycle). "
    "This ignores all other 78 COCO classes (cars, trucks, dogs, etc.) — making detection "
    "faster and reducing false positives."
)
body(
    "The expanded search zone (zone_y1 = by1 - 1.20 * bh) solves a key problem: "
    "YOLO draws the motorcycle bounding box around the vehicle body, but riders sit ON TOP "
    "of the bike. Their heads and upper bodies are above the bike's bounding box. "
    "Extending the search zone 120% upward captures the full rider."
)
body(
    "rider_count = max(rider_count_helmet, riders_yolo) — Two independent methods count riders. "
    "Taking the maximum prevents undercounting in cases where one method misses a rider."
)
body(
    "cv2.waitKey(1) waits 1 millisecond for a keypress. This is required for cv2.imshow() "
    "to actually render the frame on screen. Without it, the window would freeze. "
    "The & 0xFF mask extracts the last 8 bits for cross-platform key code compatibility."
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 8: VIOLATION RULES
# ══════════════════════════════════════════════════════════════
heading("8. Violation Rules")
table_2col(
    ["Rule", "Explanation"],
    [
        ("Triple Riding", "More than 2 riders are detected on a single motorcycle. "
                          "rider_count > 2 triggers this rule."),
        ("No Helmet",     "The helmet model detects at least one rider with class 3 "
                          "(withoutHelmet) in the bike crop. Even one unhelmeted rider "
                          "out of three triggers this violation."),
        ("Both Together", "A single motorcycle can trigger both violations simultaneously. "
                          "The violation list would be ['Triple Riding', 'No Helmet'] "
                          "and logged as 'Triple Riding & No Helmet'."),
    ]
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 9: NIGHT MODE USP
# ══════════════════════════════════════════════════════════════
heading("9. Night Mode — USP Explanation")
body(
    "The key innovation of this project is adding reliable low-light detection as a USP "
    "(Unique Selling Point) without modifying any of the existing AI models. Here is why "
    "each technology choice was made:"
)
table_2col(
    ["Technology", "Why It Was Chosen"],
    [
        ("LAB colour space",
         "Separates brightness (L) from colour (A, B). We can boost brightness without "
         "causing colour distortion — which would happen if we just multiplied BGR values."),
        ("CLAHE",
         "Smarter than basic histogram equalization. Divides image into tiles and equalizes "
         "each independently. Bright headlights don't blow out the rest of the frame."),
        ("Gamma Correction",
         "A fast, mathematically smooth brightness curve applied via lookup table. "
         "Lifts shadows without clipping highlights."),
        ("Conditional Application",
         "Enhancement only runs when is_low_light() returns True. Daytime frames are "
         "completely unaffected — no added latency, no colour distortion on bright footage."),
        ("No Retraining Needed",
         "The enhancement happens BEFORE YOLO sees the frame. YOLO is trained on normal "
         "images, and the enhanced dark frames now look similar to normal images — "
         "so detection works without any model changes."),
    ]
)
body("Honest Limitation: This is software enhancement of visible-light footage. "
     "In near-total darkness where the camera captures essentially no signal, "
     "enhancement cannot recover detail. True infrared hardware would be needed for that.")

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 10: DATABASE SCHEMA
# ══════════════════════════════════════════════════════════════
heading("10. Database Schema")
body("Database name: traffic_ai    Table name: violations")
table_2col(
    ["Column", "Type / Description"],
    [
        ("id",             "INT, AUTO_INCREMENT, PRIMARY KEY — unique ID for each violation row"),
        ("violation_type", "VARCHAR(50) — 'No Helmet', 'Triple Riding', or 'No Helmet & Triple Riding'"),
        ("plate_number",   "VARCHAR(20) — license plate text read by EasyOCR"),
        ("confidence",     "FLOAT — OCR confidence score (0.0 to 1.0)"),
        ("timestamp",      "DATETIME — automatically set to current date/time on INSERT"),
        ("location",       "VARCHAR(100) — camera identifier, default 'Main Gate Camera 1'"),
        ("evidence_image", "VARCHAR(255) — file path to the saved violation screenshot"),
    ]
)

divider()

# ══════════════════════════════════════════════════════════════
#  SECTION 11: COMPLETE FRAME FLOW
# ══════════════════════════════════════════════════════════════
heading("11. Complete Flow — One Frame, Start to Finish")
steps_flow = [
    ("1",  "cap.read()",                   "Read one raw frame from webcam or video file"),
    ("2",  "is_low_light(frame)",          "Convert to grayscale, compute average brightness"),
    ("3",  "enhance_frame(frame)",         "If dark: BGR→LAB, CLAHE on L-channel, gamma correction"),
    ("4",  "traffic_model.track()",        "YOLO detects all motorcycles+persons, assigns tracking IDs"),
    ("5",  "Build search zone",            "Expand bike box: 120% up, 15% wide, 20% down"),
    ("6",  "person_overlap_with_zone()",   "Count persons overlapping zone >= 20% as riders (YOLOv8 count)"),
    ("7",  "analyze_bike_crop()",          "Run helmet.pt on crop → rider heads + no_helmet + plate boxes"),
    ("8",  "rider_count = max(6, 7)",      "Take higher of the two rider counts"),
    ("9",  "Violation check",              "rider_count > 2 → Triple Riding | no_helmet → No Helmet"),
    ("10", "cv2.imwrite()",                "Save annotated frame as JPEG evidence"),
    ("11", "ocr.readtext()",               "EasyOCR reads license plate text (on violation only)"),
    ("12", "log_to_db()",                  "INSERT violation record into MySQL"),
    ("13", "draw_bike_panel()",            "Draw per-bike info overlay on frame"),
    ("14", "draw_stats_hud()",             "Draw top-right stats panel (FPS, count, time, night mode)"),
    ("15", "draw_night_mode_badge()",      "Draw bottom-left NIGHT MODE badge"),
    ("16", "cv2.imshow()",                 "Render the annotated frame in the display window"),
    ("17", "cv2.waitKey(1)",               "Wait 1ms, check for Q (quit) or D (debug toggle)"),
    ("18", "Loop back to step 1",          "Repeat for every frame until video ends or Q is pressed"),
]
table_2col(
    ["Step", "What Happens"],
    [(f"{s} — {code}", desc) for s, code, desc in steps_flow],
    col_widths=(3, 11)
)

# ── SAVE ──────────────────────────────────────────────────────
out_path = r"C:\Users\raita\Downloads\AI_Traffic_Violation_Code_Explanation.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
