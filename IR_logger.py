import mss
import cv2
import numpy as np
import pytesseract
from datetime import datetime
import re
import time
import pandas as pd
import os

# Path to Tesseract OCR executable (adjust if needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Screen region to capture the ROI stats panel
region = {
    "top": 1070,
    "left": 30,
    "width": 1800,
    "height": 250
}

EXCEL_FILE = "thermal_roi_log.xlsx"

# Create the Excel file if it doesn't exist
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "Timestamp",
            "Belt 1-2 Mean", "Belt 1-2 Min", "Belt 1-2 Max",
            "Belt 3-4 Mean", "Belt 3-4 Min", "Belt 3-4 Max",
            "Belt 5-6 Mean", "Belt 5-6 Min", "Belt 5-6 Max",
            "Belt 7 Mean", "Belt 7 Min", "Belt 7 Max"
        ])
        df.to_excel(EXCEL_FILE, index=False)


# Capture the defined screen region
def capture_region():
    with mss.mss() as sct:
        return np.array(sct.grab(region))[:, :, :3]

# Enhance image for OCR
def preprocess_image(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    zoom = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    inverted = cv2.bitwise_not(zoom)
    _, thresh = cv2.threshold(inverted, 160, 255, cv2.THRESH_BINARY)
    return thresh

# Extract only the correct mean, min, and max values per polygon
def extract_all_roi_values(img):
    prepped = preprocess_image(img)
    text = pytesseract.image_to_string(prepped, config="--psm 6")

    print("\nOCR Output:\n", text)

    blocks = re.split(r'Polygon\s+\d', text, flags=re.IGNORECASE)
    results = []

    for block in blocks:
        # Extract standalone decimal numbers NOT inside (x, y) coordinate pairs
        numbers = re.findall(r"(?<!\()\b\d+\.\d+\b(?![\d\s,]*\))", block)
        floats = [float(n) for n in numbers]

        if len(floats) >= 4:
            mean = floats[0]
            min_ = floats[2]
            max_ = floats[3]  # ← CORRECT max
            results.append((mean, min_, max_))
        else:
            print(f"Skipping block (not enough data): {block.strip()}")

    # Flatten list for easy mapping to belts
    flat = [val for triple in results for val in triple]
    return flat



    # Flatten the list for belt assignment
    flat = [val for triple in results for val in triple]
    return flat

# Write one row of values to Excel
def log_roi_to_excel(timestamp, roi_data):
    row = {
        "Timestamp": timestamp,
        "Belt 1-2 Mean": roi_data["belt2"]["mean"],
        "Belt 1-2 Min": roi_data["belt2"]["min"],
        "Belt 1-2 Max": roi_data["belt2"]["max"],
        "Belt 3-4 Mean": roi_data["belt1"]["mean"],
        "Belt 3-4 Min": roi_data["belt1"]["min"],
        "Belt 3-4 Max": roi_data["belt1"]["max"],
        "Belt 5-6 Mean": roi_data["belt3"]["mean"],
        "Belt 5-6 Min": roi_data["belt3"]["min"],
        "Belt 5-6 Max": roi_data["belt3"]["max"],
        "Belt 7 Mean": roi_data["belt4"]["mean"],
        "Belt 7 Min": roi_data["belt4"]["min"],
        "Belt 7 Max": roi_data["belt4"]["max"],
    }
    df = pd.read_excel(EXCEL_FILE)
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

# Main loop
def run_logger(interval_seconds=10):
    init_excel()
    while True:
        img = capture_region()
        temps = extract_all_roi_values(img)
        timestamp = datetime.now().isoformat()

        if len(temps) >= 12:
            roi_data = {
                "belt1": {"mean": temps[3], "min": temps[4], "max": temps[5]},   # FLIR 48898 Poly 2
                "belt2": {"mean": temps[0], "min": temps[1], "max": temps[2]},   # FLIR 48898 Poly 1
                "belt3": {"mean": temps[6], "min": temps[7], "max": temps[8]},   # FLIR 48007 Poly 1
                "belt4": {"mean": temps[9], "min": temps[10], "max": temps[11]}, # FLIR 48007 Poly 2
            }
            log_roi_to_excel(timestamp, roi_data)
            print(f"[{timestamp}] Logged data to Excel.")
        else:
            print(f"[{timestamp}] OCR failed or returned incomplete data.")

        time.sleep(interval_seconds)

# Start script
if __name__ == "__main__":
    run_logger(interval_seconds=10)
