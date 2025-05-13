# IR ROI Logger

## Requirements

- Python (https://www.python.org/downloads/)  
- Tesseract OCR (https://github.com/UB-Mannheim/tesseract/wiki)  
- VS Code  

---

## Setup

1. **Install Python**  
   - Download and install from https://www.python.org/downloads/  
   - Check the box “Add Python to PATH” during install

2. **Install Tesseract OCR**  
   - Download and install from: https://github.com/UB-Mannheim/tesseract/wiki  
   - In `IR_logger.py`, update this line with your install path:
     ```python
     pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
     ```

3. **Download the Project**  
   - From GitHub, click **Code > Download ZIP**  
   - Unzip it to your desktop

4. **Open in VS Code**  
   - Open VS Code → File → Open Folder → Select the unzipped project folder  
   - Open the `IR_logger.py` file

5. **Install Python Packages**  
   - Go to Terminal → New Terminal  
   - Run:
     ```bash
     pip install -r requirements.txt
     ```

6. **Run the Script**  
   - Ensure the FLIR stats panel is visible on your screen  
   - Click the green “Run” button in VS Code, or run in terminal:
     ```bash
     python IR_logger.py
     ```

---

## Output

- Creates and updates `thermal_roi_log.xlsx` with timestamped data every 10 seconds

---

## Troubleshooting

- **No data?** Make sure the FLIR window is visible and Tesseract is installed  
- **Excel not updating?** Confirm OCR is reading valid numbers from the screen region

---
