import cv2
import os
import time
import easyocr
import mysql.connector
from ultralytics import YOLO
import torch
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from night_enhancement import is_low_light, enhance_frame, draw_night_mode_badge

# ── CONFIGURATION ──────────────────────────────────────────────
# Try yolov8s (small) first — much better accuracy than nano (n)
# Falls back to yolov8n if yolov8s is not available
TRAFFIC_MODEL_PATH = "yolov8s.pt"
import os as _os
if not _os.path.exists(TRAFFIC_MODEL_PATH):
    TRAFFIC_MODEL_PATH = "yolov8n.pt"

MOTO_CONF   = 0.30   # Motorcycle detection threshold  (lower = detect more)
PERSON_CONF = 0.25   # Person detection threshold      (lower = catch more riders)
HELMET_CONF = 0.30   # Helmet/plate detection threshold

DB_PASSWORD  = "@ritesh"
EVIDENCE_DIR = "evidence"
os.makedirs(EVIDENCE_DIR, exist_ok=True)

# Helmet model class IDs  (verified by inspecting helmet.pt)
CLS_LICENSE_PLATE  = 0
CLS_MOTORCYCLE_H   = 1   # Not used
CLS_WITH_HELMET    = 2   # Rider wearing helmet -> safe
CLS_WITHOUT_HELMET = 3   # Rider NOT wearing helmet -> violation

# Debug mode: press D while running to toggle debug overlay
DEBUG = False

# ── VIDEO SOURCE SELECTOR ──────────────────────────────────────
def select_video_source():
    print("\n" + "="*52)
    print("    AI Traffic Violation Detection System")
    print("="*52)
    print("  [1] Use Webcam")
    print("  [2] Browse for a Video File")
    print("-"*52)
    choice = input("  Enter choice (1 or 2): ").strip()
    if choice == "2":
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select Video File",
            filetypes=[("Video Files","*.mp4 *.avi *.mov *.mkv *.wmv *.flv"),("All Files","*.*")]
        )
        root.destroy()
        if path:
            print(f"  Selected: {os.path.basename(path)}\n"); return path
        print("  No file chosen. Using webcam.\n"); return 0
    print("  Using Webcam.\n"); return 0

VIDEO_SOURCE = select_video_source()

# ── HARDWARE ───────────────────────────────────────────────────
device  = "cuda" if torch.cuda.is_available() else "cpu"
hw_name = torch.cuda.get_device_name(0) if device == "cuda" else "CPU (no GPU)"
print(f"[SYSTEM] Hardware      : {hw_name}")
print(f"[SYSTEM] Traffic Model : {TRAFFIC_MODEL_PATH}")

# ── DATABASE ───────────────────────────────────────────────────
def log_to_db(violation_type, plate_text, bike_id, conf):
    try:
        db = mysql.connector.connect(
            host="localhost", user="root", password=DB_PASSWORD, database="traffic_ai")
        cur = db.cursor()
        cur.execute(
            "INSERT INTO violations (violation_type,plate_number,location,confidence) VALUES (%s,%s,%s,%s)",
            (violation_type, plate_text, f"Bike_ID_{bike_id}", float(conf))
        )
        db.commit(); db.close()
        print(f"      [DB] Saved -> {violation_type} | Plate: {plate_text}")
    except Exception as e:
        print(f"      [DB ERROR] {e}")

# ── LOAD MODELS ────────────────────────────────────────────────
print(f"[SYSTEM] Loading traffic model ({TRAFFIC_MODEL_PATH}) ...")
traffic_model = YOLO(TRAFFIC_MODEL_PATH); traffic_model.to(device)

print("[SYSTEM] Loading helmet detection model ...")
try:
    helmet_model = YOLO("helmet.pt"); helmet_model.to(device)
    HAS_HELMET = True
    print("[SYSTEM] Helmet model OK  (2=withHelmet, 3=withoutHelmet, 0=plate)")
except Exception as e:
    print(f"[SYSTEM] helmet.pt missing ({e})")
    HAS_HELMET = False

print("[SYSTEM] Loading EasyOCR ...")
ocr = easyocr.Reader(["en"], gpu=(device=="cuda"))
print("[SYSTEM] All models ready. Press Q=quit, D=debug toggle.\n" + "="*52 + "\n")


# ── HELPER: Intersection-over-person (IoP) ─────────────────────
def person_overlap_with_zone(person_box, zone_box):
    """
    Returns what fraction of the PERSON box overlaps with zone_box.
    A value > 0.20 means the person is significantly inside the zone.
    """
    ax1,ay1,ax2,ay2 = person_box
    bx1,by1,bx2,by2 = zone_box
    ix1 = max(ax1, bx1); iy1 = max(ay1, by1)
    ix2 = min(ax2, bx2); iy2 = min(ay2, by2)
    inter = max(0, ix2-ix1) * max(0, iy2-iy1)
    area_person = max(1, (ax2-ax1) * (ay2-ay1))
    return inter / area_person


# ── HELPER: Run helmet model on bike crop ──────────────────────
def analyze_bike_crop(crop, offset_x, offset_y):
    """
    Run helmet.pt on a bike crop region.
    Returns:
        rider_count (int)  - heads detected (withHelmet + withoutHelmet)
        no_helmet   (bool) - True if any withoutHelmet detected
        plate_boxes (list) - plate bounding boxes in full-frame coordinates
    """
    rider_count, no_helmet, plate_boxes = 0, False, []
    if not HAS_HELMET or crop is None or crop.size == 0:
        return rider_count, no_helmet, plate_boxes

    results = helmet_model(crop, verbose=False)
    for r in results:
        for box in r.boxes:
            cls  = int(box.cls[0])
            conf = float(box.conf[0])
            if conf < HELMET_CONF:
                continue
            x1,y1,x2,y2 = map(int, box.xyxy[0])
            if cls == CLS_WITH_HELMET:
                rider_count += 1
            elif cls == CLS_WITHOUT_HELMET:
                rider_count += 1
                no_helmet = True
            elif cls == CLS_LICENSE_PLATE:
                # Convert crop coords to full-frame coords
                plate_boxes.append((x1+offset_x, y1+offset_y,
                                    x2+offset_x, y2+offset_y))
    return rider_count, no_helmet, plate_boxes


# ── HUD FONT ───────────────────────────────────────────────────
FONT = cv2.FONT_HERSHEY_SIMPLEX

def draw_bike_panel(frame, bx1,by1,bx2,by2, bid, rider_count, violations, plate_text=""):
    is_viol   = len(violations) > 0
    box_color = (0,0,220) if is_viol else (0,200,80)
    accent    = (0,50,255) if is_viol else (0,160,60)
    cv2.rectangle(frame,(bx1,by1),(bx2,by2), box_color, 3 if is_viol else 2)

    lines = [f"Bike #{bid}"]
    rc    = (0,80,255) if rider_count > 2 else (80,255,80)
    lines.append(f"Riders : {rider_count}")
    lines.append("Helmet : NO HELMET!" if "No Helmet" in violations else "Helmet : OK")
    if is_viol:
        lines.append("!! " + " | ".join(violations))
    if plate_text:
        lines.append(f"Plate  : {plate_text}")

    ph   = len(lines)*22 + 10
    py1  = max(by1 - ph - 4, 0)
    py2  = max(by1 - 4, py1 + ph)
    px2  = min(bx1 + 250, frame.shape[1]-1)

    ov = frame.copy()
    cv2.rectangle(ov,(bx1,py1),(px2,py2),(15,15,15),-1)
    cv2.addWeighted(ov,0.75,frame,0.25,0,frame)
    cv2.rectangle(frame,(bx1,py1),(px2,py2),accent,1)

    y = py1 + 17
    color_map = {
        0:(180,220,255),
        1:rc,
        2:(0,80,255) if "No Helmet" in violations else (80,255,80),
        3:(0,60,255),
        4:(0,230,230)
    }
    for i,line in enumerate(lines):
        cv2.putText(frame, line,(bx1+6,y), FONT,0.47, color_map.get(i,(200,200,200)),1,cv2.LINE_AA)
        y += 22


def draw_stats_hud(frame, fps, n_bikes, n_viols, night_on, model_name):
    h,w   = frame.shape[:2]
    pw    = 250
    x1,y1 = w-pw-8, 8
    x2,y2 = w-8, 210
    ov = frame.copy()
    cv2.rectangle(ov,(x1,y1),(x2,y2),(10,10,10),-1)
    cv2.addWeighted(ov,0.70,frame,0.30,0,frame)
    cv2.rectangle(frame,(x1,y1),(x2,y2),(55,55,55),1)
    y = y1+18
    cv2.putText(frame,"  LIVE STATS",(x1+8,y),FONT,0.52,(180,220,255),1,cv2.LINE_AA)
    cv2.line(frame,(x1+5,y+5),(x2-5,y+5),(50,50,50),1); y+=26
    rows = [
        (f"FPS           :  {fps:5.1f}",          (200,200,200)),
        (f"Bikes in frame:  {n_bikes}",            (200,200,200)),
        (f"Violations    :  {n_viols}",            (0,80,255) if n_viols>0 else (200,200,200)),
        (f"Night Mode    :  {'ON' if night_on else 'OFF'}", (0,200,255) if night_on else (100,200,100)),
        (f"Model         :  {model_name}",         (180,180,180)),
        (f"Time          :  {datetime.now().strftime('%H:%M:%S')}", (200,200,200)),
    ]
    for txt,col in rows:
        cv2.putText(frame,txt,(x1+8,y),FONT,0.43,col,1,cv2.LINE_AA); y+=23


# ── MAIN LOOP ──────────────────────────────────────────────────
cap = cv2.VideoCapture(VIDEO_SOURCE)
cap.set(cv2.CAP_PROP_FRAME_WIDTH,  1280)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
print("[SYSTEM] Running. Press Q=quit  D=debug\n")

session_viols = 0
prev_t = time.time()
fps    = 0.0
model_name = TRAFFIC_MODEL_PATH.replace(".pt","")

while True:
    ret, frame = cap.read()
    if not ret: break

    # Smooth FPS
    now = time.time()
    fps = 0.9*fps + 0.1*(1.0/(now - prev_t + 1e-6))
    prev_t = now

    # Night mode: enhance if dark
    night_on = is_low_light(frame)
    if night_on:
        frame = enhance_frame(frame)

    H, W = frame.shape[:2]

    # ── STEP 1: Run YOLOv8 - detect ALL motorcycles and persons ─
    res = traffic_model.track(frame, persist=True, verbose=False, classes=[0, 3])
    bikes, persons = [], []
    for r in res:
        for box in r.boxes:
            x1,y1,x2,y2 = map(int, box.xyxy[0])
            tid  = int(box.id[0]) if box.id is not None else -1
            cls  = int(box.cls[0])
            conf = float(box.conf[0])

            if cls == 3 and conf > MOTO_CONF:
                bikes.append({"box":[x1,y1,x2,y2], "id":tid, "conf":conf})
                if DEBUG:
                    cv2.putText(frame, f"bike {conf:.2f}", (x1,y1-5),
                                FONT, 0.4, (0,255,255), 1)
            elif cls == 0 and conf > PERSON_CONF:
                persons.append({"box":[x1,y1,x2,y2], "conf":conf})
                if DEBUG:
                    px1,py1,px2,py2 = x1,y1,x2,y2
                    cv2.rectangle(frame,(px1,py1),(px2,py2),(255,255,0),1)
                    cv2.putText(frame, f"p {conf:.2f}", (px1,py1-4),
                                FONT, 0.35, (255,255,0), 1)

    # ── STEP 2: Per-bike analysis ──────────────────────────────
    for bike in bikes:
        bx1,by1,bx2,by2 = bike["box"]
        bid = bike["id"]
        bw  = bx2 - bx1
        bh  = by2 - by1

        # Build an expanded "rider search zone":
        # Extend upward by 120% of bike height (riders sit on top)
        # Extend sideways by 15% (riders lean sideways)
        # Extend downward by 20% (feet below chassis)
        zone_x1 = max(bx1 - int(bw * 0.15), 0)
        zone_y1 = max(by1 - int(bh * 1.20), 0)
        zone_x2 = min(bx2 + int(bw * 0.15), W)
        zone_y2 = min(by2 + int(bh * 0.20), H)
        zone    = [zone_x1, zone_y1, zone_x2, zone_y2]

        if DEBUG:
            # Show the rider search zone in blue
            cv2.rectangle(frame,(zone_x1,zone_y1),(zone_x2,zone_y2),(255,100,0),1)
            cv2.putText(frame,"ZONE",(zone_x1+2,zone_y1+12),FONT,0.35,(255,100,0),1)

        # Count riders from YOLOv8: person overlaps zone by >= 20%
        riders_yolo = 0
        for p in persons:
            pb = p["box"]
            if person_overlap_with_zone(pb, zone) >= 0.20:
                riders_yolo += 1
                px1,py1_,px2,py2_ = pb
                cv2.rectangle(frame,(px1,py1_),(px2,py2_),(0,200,80),2)  # Green = rider

        # ── STEP 3: Helmet model crop ──────────────────────────
        # Crop from the zone (which already includes above-bike space)
        # Add extra top padding to make sure heads are included
        crop_x1 = zone_x1
        crop_y1 = max(zone_y1 - int(bh * 0.3), 0)  # extra buffer at top
        crop_x2 = zone_x2
        crop_y2 = zone_y2
        crop = frame[crop_y1:crop_y2, crop_x1:crop_x2]

        if DEBUG and crop.size > 0:
            # Show crop outline in magenta
            cv2.rectangle(frame,(crop_x1,crop_y1),(crop_x2,crop_y2),(255,0,255),1)

        rider_count_helmet, no_helmet, plate_boxes = analyze_bike_crop(
            crop, crop_x1, crop_y1)

        # Final rider count = max of both methods
        rider_count = max(rider_count_helmet, riders_yolo)

        if DEBUG:
            cv2.putText(frame,
                f"yolo:{riders_yolo} helm:{rider_count_helmet} final:{rider_count}",
                (bx1, by2+15), FONT, 0.38, (255,255,0), 1)

        # ── STEP 4: Violation rules ────────────────────────────
        violations = []
        if rider_count > 2:
            violations.append("Triple Riding")
        if no_helmet:
            violations.append("No Helmet")

        # ── STEP 5: On violation — save evidence + OCR plate ──
        plate_text = ""
        if violations:
            session_viols += 1
            vtext = " & ".join(violations)
            print(f"  [VIOLATION] Bike#{bid} -> {vtext} | Riders:{rider_count}")

            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            ev_path = os.path.join(EVIDENCE_DIR, f"{ts}_Bike{bid}.jpg")
            cv2.imwrite(ev_path, frame)

            # Try OCR on helmet-model detected plate boxes first
            if plate_boxes:
                for (plx1,ply1,plx2,ply2) in plate_boxes:
                    cv2.rectangle(frame,(plx1,ply1),(plx2,ply2),(0,255,255),2)
                    plate_crop = frame[ply1:ply2, plx1:plx2]
                    if plate_crop.size > 0:
                        try:
                            for (_,txt,prob) in ocr.readtext(plate_crop):
                                if prob > 0.30:
                                    clean = "".join(c for c in txt if c.isalnum())
                                    if len(clean) > 3:
                                        plate_text = clean
                                        print(f"      [OCR-plate] {clean} ({prob:.2f})")
                                        log_to_db(vtext, clean, bid, prob)
                                        break
                        except: pass
                    if plate_text: break

            if not plate_text:
                # Fallback: scan bottom 40% of bike bounding box
                proi = frame[by1 + int(bh*0.60):by2, bx1:bx2]
                try:
                    for (_,txt,prob) in ocr.readtext(proi):
                        if prob > 0.30:
                            clean = "".join(c for c in txt if c.isalnum())
                            if len(clean) > 3:
                                plate_text = clean
                                print(f"      [OCR-fallback] {clean} ({prob:.2f})")
                                log_to_db(vtext, clean, bid, prob)
                                break
                except: pass
        else:
            print(f"  [OK]  Bike#{bid} -> Safe | Riders:{rider_count}")

        # Draw bike info panel
        draw_bike_panel(frame, bx1,by1,bx2,by2, bid, rider_count, violations, plate_text)

    # ── Frame HUD ──────────────────────────────────────────────
    draw_stats_hud(frame, fps, len(bikes), session_viols, night_on, model_name)
    draw_night_mode_badge(frame, night_on)

    if DEBUG:
        cv2.putText(frame,"[DEBUG ON - press D to hide]",(10,30),FONT,0.55,(0,255,255),1)

    cv2.imshow("AI Traffic Violation Detection", frame)

    key = cv2.waitKey(1) & 0xFF
    if key == ord("q"):
        break
    elif key == ord("d") or key == ord("D"):
        DEBUG = not DEBUG
        print(f"[DEBUG] {'ON' if DEBUG else 'OFF'}")

cap.release()
cv2.destroyAllWindows()
print(f"\n[SESSION DONE]  Violations: {session_viols}  |  Evidence: {EVIDENCE_DIR}/")