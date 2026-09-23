# 🚦 AI Traffic Violation Detection System

An AI-powered system that detects motorcycle traffic violations from CCTV/video footage in real time.

## 🎯 Features

| Feature | Description |
|---|---|
| 🏍️ Motorcycle Tracking | Persistent ID tracking across frames using YOLOv8 |
| 👥 Rider Counting | Geometry-based association of persons to bikes |
| 🪖 Helmet Detection | Custom YOLO model detects `withHelmet` / `withoutHelmet` |
| ⚠️ Triple Riding | Flags when more than 2 riders are on one bike |
| 🔢 License Plate OCR | EasyOCR reads plate numbers on violation |
| 🌙 Night Mode | Auto low-light enhancement using CLAHE (no hardware needed) |
| 📸 Evidence Saving | Violation screenshots saved automatically to `evidence/` |
| 💾 MySQL Logging | All violations logged to database with plate, type, confidence |

## 🛠️ Tech Stack

- Python 3.10+
- YOLOv8 (`ultralytics`) — traffic detection & tracking
- Custom `helmet.pt` — helmet/plate detection
- OpenCV — video processing & HUD rendering
- EasyOCR — license plate text recognition
- PyTorch (CPU/CUDA)
- MySQL — violation logging
- Tkinter — video file browser dialog

## 🚀 Setup

### 1. Install dependencies
```bash
pip install ultralytics opencv-python easyocr mysql-connector-python torch torchvision
```

### 2. Download YOLO models
```bash
# Auto-downloaded on first run, or manually:
python -c "from ultralytics import YOLO; YOLO('yolov8s.pt')"
```

> **Note:** `helmet.pt` must be placed manually in the project root.

### 3. Set up MySQL
```bash
python database_setup.py
```

### 4. Run
```bash
python main.py
```
Choose webcam or browse for a video file at the startup prompt.

## 🌙 Night Mode

Automatically activates when average frame brightness < 80/255.  
Uses **CLAHE on LAB color space** — no IR hardware or additional training required.

Press **D** while running to toggle debug overlay (shows detection zones, raw counts).

## 📁 Project Structure

```
├── main.py                  # Main detection pipeline
├── night_enhancement.py     # Low-light enhancement module
├── database_setup.py        # MySQL schema setup
├── helmet.pt                # Custom helmet detection model
├── requirements.txt         # Dependencies
└── evidence/                # Auto-created — violation screenshots
```

## ⚠️ Notes

- Helmet model classes: `0=licensePlate`, `2=withHelmet`, `3=withoutHelmet`
- `yolov8s.pt` and `yolov8n.pt` are not committed (large files) — auto-downloaded on first run
- Change `DB_PASSWORD` in `main.py` to match your MySQL setup

## 📋 Violation Rules

- **No Helmet** — any rider detected without a helmet
- **Triple Riding** — more than 2 riders on a single motorcycle
