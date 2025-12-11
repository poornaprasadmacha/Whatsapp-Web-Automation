"""
send_whatsapp_selenium.py

Usage:
    python send_whatsapp_selenium.py

Requirements:
    pip install selenium webdriver-manager pandas openpyxl

Notes:
 - The first time you run this, Chrome will open and you must scan the WhatsApp Web QR code.
 - The script uses a local Chrome profile folder (./wa_profile) to preserve login between runs.
 - Template file should use '{name}' placeholder for personalization.
"""

import re
import time
import csv
import sys
from pathlib import Path
from urllib.parse import quote

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ---------- Configuration ----------
EXCEL_PATH = "Whatsapp List_Main.xlsx"
MESSAGE_PATH = "WHATSDRAFT.txt"
LOG_CSV = "whatsapp_send_log.csv"
CHROME_PROFILE_DIR = "./wa_profile"   # persists login; keep this folder safe
PAGE_LOAD_TIMEOUT = 60                # seconds to wait for page load
SEND_WAIT_TIMEOUT = 30                # seconds to wait for send button / input
DELAY_BETWEEN_MESSAGES = 3            # seconds between messages (rate-limit)
# -----------------------------------


def sanitize_contact(raw):
    """Remove any non-digit chars. WhatsApp expects country code without leading +."""
    if pd.isna(raw):
        return ""
    s = str(raw)
    s = re.sub(r"\D", "", s)  # remove non-digits
    # Remove leading zeros if any (user should include country code)
    s = s.lstrip("0")
    return s


def load_template(path):
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Message template not found: {path}")
    return p.read_text(encoding="utf-8")


def prepare_driver(profile_dir: str):
    """Set up Chrome driver with a persistent profile (so you stay logged into WhatsApp)."""
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={profile_dir}")
    # optional: run headless (NOT recommended because WhatsApp Web often blocks or shows different UI)
    # options.add_argument("--headless=new")
    # recommended for reliability on some systems:
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    return driver


def wait_for_send_ready(driver, timeout=SEND_WAIT_TIMEOUT):
    """
    Wait for either:
      - the send button to be clickable (data-testid or aria-label), OR
      - the editable text box to be present (so we can press ENTER).
    Returns a tuple (send_button_element_or_None, input_box_element_or_None)
    """
    wait = WebDriverWait(driver, timeout)
    send_btn = None
    input_box = None

    # Try several ways to find the send button and input box
    try:
        # common send button data-testid used in WhatsApp web
        send_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='compose-btn-send']")))
        return send_btn, None
    except Exception:
        send_btn = None

    # fallback to aria-label='Send'
    try:
        send_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Send']")))
        return send_btn, None
    except Exception:
        send_btn = None

    # fallback: find the contenteditable input box (typical pattern)
    try:
        # the text box for message often is a div with contenteditable="true" and a data-tab attribute
        input_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and @data-tab]")))
        return None, input_box
    except Exception:
        input_box = None

    # nothing found
    return None, None


def send_whatsapp_selenium(excel_path, message_path):
    # Read data
    if not Path(excel_path).exists():
        print(f"Excel file not found: {excel_path}")
        sys.exit(1)

    df = pd.read_excel(excel_path, dtype={"Contact": str})
    template = load_template(message_path)

    # Initialize driver
    driver = prepare_driver(CHROME_PROFILE_DIR)
    wait = WebDriverWait(driver, PAGE_LOAD_TIMEOUT)

    # Prepare log CSV
    log_file = open(LOG_CSV, mode="w", newline="", encoding="utf-8")
    csv_writer = csv.writer(log_file)
    csv_writer.writerow(["row", "name", "contact", "sanitized_contact", "status", "details"])

    try:
        counter = 0
        total = len(df)
        for idx, row in df.iterrows():
            raw_name = row.get("Name", "")
            raw_contact = row.get("Contact", "")

            name = "" if pd.isna(raw_name) else str(raw_name)
            sanitized = sanitize_contact(raw_contact)

            if not sanitized:
                print(f"[{idx}] Skipping empty or invalid contact: {raw_contact}")
                csv_writer.writerow([idx, name, raw_contact, sanitized, "skipped", "empty contact"])
                continue

            # Prepare message. Use {name} placeholder in template.
            try:
                message = template.format(name=name)
            except Exception as e:
                # If template formatting fails, fallback to simple replacement
                print(f"Template format error for row {idx}: {e}. Using simple replace.")
                message = template.replace("{name}", name)

            # Create url and open
            url = f"https://web.whatsapp.com/send?phone={sanitized}&text={quote(message)}"
            try:
                driver.get(url)
            except Exception as e:
                print(f"[{idx}] Error loading URL for {sanitized}: {e}")
                csv_writer.writerow([idx, name, raw_contact, sanitized, "error", f"get_error: {e}"])
                continue

            # Wait for the UI to be ready (send button or input box)
            send_btn, input_box = wait_for_send_ready(driver, timeout=SEND_WAIT_TIMEOUT)

            # If we have a send button element, click it
            try:
                if send_btn is not None:
                    send_btn.click()
                elif input_box is not None:
                    # send ENTER in the input box (works if message prefilled)
                    # Many times the text is already in the box (from the URL param), so Enter will send.
                    input_box.send_keys(Keys.ENTER)
                else:
                    # Nothing found; try one more fallback: press Enter on body
                    print(f"[{idx}] Send controls not found for {sanitized}. Trying fallback Enter on body.")
                    try:
                        body = driver.find_element(By.TAG_NAME, "body")
                        body.send_keys(Keys.ENTER)
                    except Exception as e:
                        print(f"[{idx}] Fallback also failed: {e}")
                        csv_writer.writerow([idx, name, raw_contact, sanitized, "error", "no_send_control"])
                        continue

                # Wait a small bit to ensure message is sent
                time.sleep(1.5)
                counter += 1
                print(f"[{counter}/{total}] Sent to {name} ({sanitized})")
                csv_writer.writerow([idx, name, raw_contact, sanitized, "sent", "ok"])

            except Exception as e:
                print(f"[{idx}] Failed to send to {sanitized}: {e}")
                csv_writer.writerow([idx, name, raw_contact, sanitized, "failed", str(e)])
                continue

            # Rate limiting / politeness
            time.sleep(DELAY_BETWEEN_MESSAGES)

    finally:
        log_file.close()
        driver.quit()
        print("Finished. Log is saved to", LOG_CSV)


if __name__ == "__main__":
    send_whatsapp_selenium(EXCEL_PATH, MESSAGE_PATH)
