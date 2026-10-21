import pyautogui
import time
import subprocess
from datetime import datetime
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

today = datetime.now()

date_string = today.strftime("%Y-%m-%d")
time_string = today.strftime("%Y-%m-%d %H:%M:%S")

REPORT_DIR = BASE_DIR / "reports"
REPORT_DIR.mkdir(exist_ok=True)

excel_file = REPORT_DIR / f"daily_report_{date_string}.xlsx"
screenshot_file = REPORT_DIR / f"daily_report_{date_string}.png"

# ============================================================
# HELPER FUNCTION
# ============================================================

def wait(seconds):
    """Pause so the application has time to respond."""
    time.sleep(seconds)


# ============================================================
# START
# ============================================================

print("=" * 60)
print("DAILY STATUS REPORT BOT")
print("=" * 60)

print()
print("Starting bot in 3 seconds...")
wait(3)


# ============================================================
# STEP 1 - OPEN GOOGLE CHROME
# ============================================================

print("Opening Google Chrome...")

pyautogui.hotkey(
    "command",
    "space"
)

wait(1)

pyautogui.write(
    "Google Chrome",
    interval=0.05
)

pyautogui.press("enter")

wait(4)

print("Chrome opened.")


# ============================================================
# STEP 2 - OPEN CHENNAI WEATHER
# ============================================================

print("Opening Chennai weather...")

pyautogui.hotkey(
    "command",
    "l"
)

pyautogui.write(
    "https://www.google.com/search?q=Chennai+weather",
    interval=0.03
)

pyautogui.press("enter")

# Give Google time to load
wait(6)

print("Weather page loaded.")


# ============================================================
# STEP 3 - COPY WEATHER INFORMATION
# ============================================================

print("Selecting weather information...")

# Click somewhere on the webpage first
pyautogui.click(
    x=700,
    y=450
)

wait(1)

# Select all visible webpage text
pyautogui.hotkey(
    "command",
    "a"
)

wait(1)

# Copy the selected text
pyautogui.hotkey(
    "command",
    "c"
)

wait(2)

# Read clipboard using macOS pbpaste
try:

    weather_data = subprocess.check_output(
        ["pbpaste"]
    ).decode(
        "utf-8",
        errors="ignore"
    ).strip()

except Exception:

    weather_data = ""


# If clipboard is empty, use a fallback
if not weather_data:

    weather_data = (
        "Chennai weather information "
        "could not be copied automatically."
    )


# Keep the report cell reasonably short
weather_lines = [
    line.strip()
    for line in weather_data.splitlines()
    if line.strip()
]

if weather_lines:

    # Take first few useful lines
    weather_summary = " | ".join(
        weather_lines[:8]
    )

else:

    weather_summary = weather_data


print()
print("Weather information copied:")
print(weather_summary)


# ============================================================
# STEP 4 - CREATE COMMENT
# ============================================================

comment = "Good for checking today's outdoor conditions."


# ============================================================
# STEP 5 - OPEN NUMBERS
# ============================================================

print()
print("Opening Numbers...")

pyautogui.hotkey(
    "command",
    "space"
)

wait(1)

pyautogui.write(
    "Numbers",
    interval=0.05
)

pyautogui.press(
    "enter"
)

wait(5)

print("Numbers opened.")


# ============================================================
# STEP 6 - CREATE NEW SPREADSHEET
# ============================================================

print("Creating new spreadsheet...")

# Command + N creates a new spreadsheet
pyautogui.hotkey(
    "command",
    "n"
)

wait(4)


# ============================================================
# STEP 7 - ENTER REPORT DATA
# ============================================================

print("Entering report data...")

# Headers
pyautogui.write(
    "Date & Time",
    interval=0.03
)

pyautogui.press("tab")

pyautogui.write(
    "Weather",
    interval=0.03
)

pyautogui.press("tab")

pyautogui.write(
    "Comment",
    interval=0.03
)

pyautogui.press("enter")


# Date and time
pyautogui.write(
    time_string,
    interval=0.03
)

pyautogui.press("tab")


# Weather information
pyautogui.write(
    weather_summary[:500],
    interval=0.001
)

pyautogui.press("tab")


# Comment
pyautogui.write(
    comment,
    interval=0.03
)

wait(3)

print("Report data entered.")


# ============================================================
# STEP 8 - SAVE EXCEL FILE
# ============================================================

# ============================================================
# STEP 8 - SAVE EXCEL FILE
# ============================================================

print()
print("Saving Excel file...")

# Use Command + S
pyautogui.hotkey(
    "command",
    "s"
)

wait(3)

# Type the desired filename
pyautogui.hotkey(
    "command",
    "a"
)

pyautogui.write(
    str(excel_file),
    interval=0.01
)

wait(1)

pyautogui.press(
    "enter"
)

wait(4)

# Numbers may ask whether to use Numbers format
# or Excel format.
#
# Press Enter to accept the currently selected option.
pyautogui.press(
    "enter"
)

wait(4)

# Verify that the file really exists
if excel_file.exists():

    print(
        f"Excel file verified successfully:"
    )

    print(
        excel_file
    )

else:

    print(
        "WARNING: Excel file was not found."
    )

    print(
        "Please check the Numbers save dialog."
    )

# ============================================================
# STEP 9 - TAKE SCREENSHOT
# ============================================================

# ============================================================
# STEP 9 - TAKE SCREENSHOT
# ============================================================

print()
print("Taking screenshot of final report...")

# Make sure Numbers is active
pyautogui.click(
    x=700,
    y=400
)

wait(2)

try:

    screenshot = pyautogui.screenshot()

    screenshot.save(
        str(screenshot_file)
    )

    wait(2)

except Exception as error:

    print(
        f"Screenshot error: {error}"
    )


if screenshot_file.exists():

    print(
        "Screenshot verified successfully:"
    )

    print(
        screenshot_file
    )

else:

    print(
        "WARNING: Screenshot was not created."
    )
# ============================================================
# FINISHED
# ============================================================

print()
print("=" * 60)
print("DAILY REPORT COMPLETED")
print("=" * 60)

print()
print("Excel:")
print(excel_file)

print()
print("Screenshot:")
print(screenshot_file)

print()
print("Bot finished successfully.")
