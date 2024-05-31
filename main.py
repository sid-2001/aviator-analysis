import easyocr
from PIL import ImageGrab
import openpyxl
import time
from datetime import datetime
import re

# Define the region to capture (left, top, right, bottom)
region = (75, 556, 365, 686)  # Example region; adjust as needed

# Create a new Excel workbook and select the active worksheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Add headers to the sheet
sheet['A1'] = 'Timestamp'
sheet['B1'] = 'Extracted Text'

# Initialize EasyOCR reader
reader = easyocr.Reader(['en'], gpu=False)

# Function to capture the region and perform OCR
def capture_and_save(region, sheet, prev_value):
    # Capture the specified region of the screen
    screenshot = ImageGrab.grab(bbox=region)
    screenshot.save('screenshot.png')

    # Perform OCR on the captured image
    results = reader.readtext('screenshot.png', detail=0)

    # Combine results into a single string
    text = ' '.join(results)

    # Extract numeric value from the text
    numeric_text = re.sub(r'[^0-9.-]', '', text)
    try:
        numeric_value = abs(float(numeric_text))
    except ValueError:
        numeric_value = None

    # If the numeric value is different from the previous one and is valid
    if numeric_value is not None and numeric_value != prev_value:
        # Get the current timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Find the next empty row
        next_row = sheet.max_row + 1

        # Write the timestamp and extracted text to the sheet
        sheet[f'A{next_row}'] = timestamp
        sheet[f'B{next_row}'] = numeric_value

        # Return the new numeric value
        return numeric_value

    # Return the previous numeric value if no change
    return prev_value

# Initialize previous value
prev_value = None

try:
    while True:
        # Capture and save data
        prev_value = capture_and_save(region, sheet, prev_value)

        # Save the workbook to a file
        workbook.save('extracted_text_with_timestamps.xlsx')

        # Wait for 10 seconds
        time.sleep(4)
except KeyboardInterrupt:
    # Save the workbook one last time if the script is interrupted
    workbook.save('extracted_text_with_timestamps.xlsx')
    print('Script terminated and data saved to extracted_text_with_timestamps.xlsx')
