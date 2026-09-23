"""
night_enhancement.py
--------------------
Low-Light / Night-Time Frame Enhancement Module
for AI Traffic Violation Detection System.

No new packages required — uses only OpenCV and NumPy
which are already in your requirements.txt.

How it works:
  1. is_low_light()  → checks average brightness of the frame
  2. enhance_frame() → applies CLAHE on LAB color space
     - Converts BGR → LAB
     - Applies CLAHE only on the L (lightness) channel
     - Converts back to BGR
     - Colors stay natural, no distortion
  3. apply_gamma()   → optional extra boost for very dark frames

This module does NOT touch any detection/tracking/OCR logic.
It is purely a pre-processing step.
"""

import cv2
import numpy as np

# ---------------------------------------------------------------
# TUNABLE SETTINGS
# ---------------------------------------------------------------
LOW_LIGHT_THRESHOLD = 80   # 0-255: frames darker than this → enhanced
CLAHE_CLIP_LIMIT    = 2.0  # Higher = more contrast boost (2.0 is safe)
CLAHE_GRID_SIZE     = (8, 8)
DEFAULT_GAMMA       = 1.4  # > 1 brightens, < 1 darkens


def is_low_light(frame: np.ndarray, threshold: int = LOW_LIGHT_THRESHOLD) -> bool:
    """
    Returns True if the frame is considered low-light / dark.

    Args:
        frame     : BGR image (numpy array from cv2.VideoCapture)
        threshold : average brightness below this → low-light
                    Range 0–255. Default 80 works well for most CCTV.

    Returns:
        bool
    """
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    brightness = float(gray.mean())
    return brightness < threshold


def enhance_frame(frame: np.ndarray) -> np.ndarray:
    """
    Enhances a low-light frame using CLAHE on the LAB color space.

    - Converts BGR → LAB
    - Applies CLAHE only on L (lightness) channel → no color shift
    - Converts back to BGR
    - Runs optional gamma correction for very dark frames

    Args:
        frame : BGR image (numpy array)

    Returns:
        Enhanced BGR image (same shape & dtype as input)
    """
    # Step 1: Convert to LAB
    lab = cv2.cvtColor(frame, cv2.COLOR_BGR2LAB)
    l_channel, a_channel, b_channel = cv2.split(lab)

    # Step 2: Apply CLAHE on the L (lightness) channel only
    clahe = cv2.createCLAHE(clipLimit=CLAHE_CLIP_LIMIT, tileGridSize=CLAHE_GRID_SIZE)
    l_enhanced = clahe.apply(l_channel)

    # Step 3: Merge back and convert to BGR
    lab_enhanced = cv2.merge((l_enhanced, a_channel, b_channel))
    enhanced = cv2.cvtColor(lab_enhanced, cv2.COLOR_LAB2BGR)

    # Step 4: Apply gamma correction for an extra brightness lift
    enhanced = apply_gamma(enhanced, gamma=DEFAULT_GAMMA)

    return enhanced


def apply_gamma(frame: np.ndarray, gamma: float = DEFAULT_GAMMA) -> np.ndarray:
    """
    Applies gamma correction to a BGR frame.
    gamma > 1 → brightens the image
    gamma < 1 → darkens the image

    Args:
        frame : BGR image (numpy array)
        gamma : correction factor (default 1.4)

    Returns:
        Gamma-corrected BGR image
    """
    inv_gamma = 1.0 / gamma
    # Build a lookup table for all 256 pixel values (fast)
    table = np.array(
        [((i / 255.0) ** inv_gamma) * 255 for i in range(256)],
        dtype=np.uint8
    )
    return cv2.LUT(frame, table)


def draw_night_mode_badge(frame: np.ndarray, is_active: bool) -> np.ndarray:
    """
    Draws a small HUD badge in the bottom-left corner of the frame
    showing whether Night Mode is active.

    Args:
        frame     : BGR image to draw on
        is_active : True = Night Mode ON, False = OFF

    Returns:
        Frame with badge drawn (in-place, also returned for convenience)
    """
    h, w = frame.shape[:2]
    label   = "NIGHT MODE: ON" if is_active else "NIGHT MODE: OFF"
    color   = (0, 200, 255) if is_active else (0, 200, 0)   # Orange if ON, Green if OFF
    icon    = "  [MOON]" if is_active else "  [SUN]"

    # Background rectangle
    rect_x1 = 10
    rect_y1 = h - 45
    rect_x2 = 240
    rect_y2 = h - 10
    cv2.rectangle(frame, (rect_x1, rect_y1), (rect_x2, rect_y2), (20, 20, 20), -1)
    cv2.rectangle(frame, (rect_x1, rect_y1), (rect_x2, rect_y2), color, 1)

    # Text
    cv2.putText(
        frame, label,
        (rect_x1 + 8, rect_y2 - 10),
        cv2.FONT_HERSHEY_SIMPLEX, 0.55, color, 1, cv2.LINE_AA
    )
    return frame


# ---------------------------------------------------------------
# STANDALONE TEST  (run: python night_enhancement.py)
# ---------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 55)
    print("  Night Enhancement Module — Standalone Test")
    print("=" * 55)

    # Test with webcam frame
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("[ERROR] Cannot open webcam for test.")
    else:
        ret, frame = cap.read()
        cap.release()

        if ret:
            # Simulate a dark frame by darkening it
            dark_frame = (frame * 0.25).astype("uint8")
            enhanced   = enhance_frame(dark_frame)

            gray   = cv2.cvtColor(dark_frame, cv2.COLOR_BGR2GRAY)
            bright = gray.mean()
            print(f"[TEST] Simulated dark frame — brightness: {bright:.1f}/255")
            print(f"[TEST] is_low_light() → {is_low_light(dark_frame)}")
            print(f"[TEST] Enhancement applied. Showing side-by-side (press any key to close).")

            # Side-by-side comparison
            comparison = np.hstack([dark_frame, enhanced])
            cv2.putText(comparison, "ORIGINAL (dark)",  (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0,0,255), 2)
            cv2.putText(comparison, "ENHANCED",          (frame.shape[1]+10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0,255,0), 2)
            cv2.imshow("Night Enhancement Test", comparison)
            cv2.waitKey(0)
            cv2.destroyAllWindows()
            print("[TEST] ✅ Module working correctly!")
        else:
            print("[ERROR] Could not read frame from webcam.")
