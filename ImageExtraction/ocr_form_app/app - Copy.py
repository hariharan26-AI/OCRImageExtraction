"""
OCR -> Review -> Auto-Fill Internal Form
=========================================
Windows-friendly Flask app for your internal data-entry automation.

Flow:
  1. Upload a scanned form image.
  2. OCR.space extracts text WITH word bounding boxes (isOverlayRequired=True).
  3. Words are grouped by position (label row -> value row) instead of
     relying on sequential token order, so the mapping doesn't break when
     OCR misses/adds a word.
  4. You see the extracted JSON on a review page (with a zoomable image)
     and can fix any field.
  5. Only after you click "Confirm & Fill Form" does Selenium log in to your
     internal system and fill the real form.
  6. Once the form fill succeeds, the uploaded image is deleted from disk
     (static/uploads) and from the session.

SETUP (Windows):
    python -m venv venv
    venv\\Scripts\\activate
    pip install -r requirements.txt
    copy .env.example .env      # then edit .env with your real values
    python app.py

Chrome + chromedriver: webdriver-manager auto-downloads the right
chromedriver for your installed Chrome version, so you don't need to
manually manage driver binaries on Windows. The driver *service* is now
cached at module level (see get_driver_service) so we don't re-check /
re-download it on every single form submission -- that repeated network
check was one of the two big causes of slowness.
"""

import os
import re
import json
import atexit
import shutil
import threading
import subprocess
import requests
from flask import Flask, request, render_template, session, redirect, url_for
from dotenv import load_dotenv
from rapidfuzz import fuzz
from datetime import datetime
from num2words import num2words

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

load_dotenv()  # reads .env file

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-change-me")

UPLOAD_FOLDER = os.path.join("static", "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SOURCE_IMAGE_FOLDER = r"C:\Users\Kishore\Downloads\DataImages"

BANNED_CHARS = ["'", "\\", "£", "|"]


def fix_punctuation_spacing(text: str) -> str:
    """OCR reports punctuation as its own separate word with a space on
    both sides -- e.g. "Mr" "." becomes "Mr ." when joined, and "Floor"
    "," "South" becomes "Floor , South". This removes the space BEFORE
    ,.;: so titles keep their dot attached ("Mr.") and addresses read
    naturally ("Floor, South") instead of "Mr ." / "Floor , South"."""
    if not text:
        return text
    text = re.sub(r'\s+([,.;:])', r'\1', text)
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

FIELD_ORDER = [
    "File No", "Form No", "Title", "First Name", "Last Name", "Email", "Father name", "DOB", "Gender",
    "Profession", "Mailing street", "Mailing city", "Mailing postal code",
    "Mailing country", "Service provider", "File no", "Reference number",
    "Sim no", "Type of network", "Cell model number", "IMMEI-1", "IMMEI-2",
    "Type of plan", "Credit card type", "Contact value", "Date of issue",
    "Date of renewal", "Installments", "Amount in words", "Remarks",
]

FUZZY_MATCH_THRESHOLD = 60


def clean_value(value: str) -> str:
    if value is None:
        return ""
    for ch in BANNED_CHARS:
        value = value.replace(ch, "")
    value = fix_punctuation_spacing(value)
    return value.strip()


def extract_file_and_form_no(filename):
    filename = os.path.splitext(filename)[0]

    match = re.match(r"(\d+)\s+(\d+)", filename)

    if match:
        return match.group(1), match.group(2)

    return "", ""


def extract_numeric_value(value):
    if not value:
        return 0

    match = re.search(r"[\d.]+", str(value))

    if match:
        return float(match.group())

    return 0


def calculate_installments(contact_value, issue_date, renewal_date):
    try:
        issue_year = datetime.strptime(
            issue_date,
            "%d/%m/%Y"
        ).year

        renewal_year = datetime.strptime(
            renewal_date,
            "%d/%m/%Y"
        ).year

        months = (renewal_year - issue_year) * 12

        if months <= 0:
            return 0

        amount = (contact_value / months) + 10.33

        return round(amount, 2)

    except Exception:
        return 0


def amount_to_words(value):
    try:
        text = num2words(value)

        text = text.replace("-", " ")

        return text.title()

    except Exception:
        return ""


# ---------------------------------------------------------------------------
# OCR with bounding boxes via OCR.space overlay
# ---------------------------------------------------------------------------

def ocr_with_boxes(image_path: str) -> list:
    """Calls OCR.space with isOverlayRequired=True and returns a flat list
    of words with their pixel positions."""
    api_key = os.environ.get("OCR_API_KEY")
    if not api_key:
        raise RuntimeError("OCR_API_KEY missing from .env")

    url = "https://api.ocr.space/parse/image"
    with open(image_path, "rb") as f:
        response = requests.post(
            url,
            files={"file": f},
            data={
                "apikey": api_key,
                "language": "eng",
                # IMPORTANT: OCR.space expects the literal strings "true"/"2",
                # not Python booleans/ints. requests.post would otherwise send
                # "True" (capital T) which OCR.space silently ignores, so you
                # get a 200 OK with NO overlay/bounding-box data at all -> every
                # field ends up blank with no error shown. This was the bug.
                "isOverlayRequired": "true",
                "OCREngine": "2",
            },
        )
    result = response.json()

    # Full raw response so you can see exactly what OCR.space returned.
    print("---- OCR.space raw response ----")
    print(json.dumps(result, indent=2)[:2000])
    print("---------------------------------")

    if "ParsedResults" not in result or not result["ParsedResults"]:
        raise RuntimeError(f"OCR API Error: {result.get('ErrorMessage', 'Unknown error')}")

    parsed = result["ParsedResults"][0]
    if parsed.get("ErrorMessage") or parsed.get("IsErroredOnProcessing"):
        raise RuntimeError(f"OCR processing error: {parsed.get('ErrorMessage')}")

    overlay = parsed.get("TextOverlay", {})
    words = []
    for line in overlay.get("Lines", []):
        for w in line.get("Words", []):
            words.append({
                "text": w["WordText"],
                "left": w["Left"],
                "top": w["Top"],
                "width": w["Width"],
                "height": w["Height"],
            })

    print(f"OCR extracted {len(words)} words with bounding boxes")
    if not words:
        raise RuntimeError(
            "OCR returned success but no word-level overlay data. "
            "Check that OCR_API_KEY in .env is a real, valid key "
            "(not the placeholder) and that the image has readable text."
        )

    return words


def group_into_lines(words: list, y_tolerance: int = 10) -> list:
    lines = []
    for w in sorted(words, key=lambda w: w["top"]):
        placed = False
        for line in lines:
            if abs(line[0]["top"] - w["top"]) <= y_tolerance:
                line.append(w)
                placed = True
                break
        if not placed:
            lines.append([w])
    for line in lines:
        line.sort(key=lambda w: w["left"])
    lines.sort(key=lambda line: line[0]["top"])

    # Debug: show exactly how words were grouped into visual rows, so we can
    # tell whether y_tolerance is merging rows that should be separate (or
    # splitting rows that should be together).
    print(f"---- grouped into {len(lines)} lines (y_tolerance={y_tolerance}) ----")
    for idx, line in enumerate(lines):
        tops = [w["top"] for w in line]
        text_preview = " | ".join(w["text"] for w in line)
        print(f"line {idx}: top_range=[{min(tops)}-{max(tops)}]  words: {text_preview}")
    print("--------------------------------------------------------")

    return lines


def cluster_by_gap(words: list, gap_threshold: int = 35) -> list:
    if not words:
        return []
    clusters = [[words[0]]]
    for w in words[1:]:
        prev = clusters[-1][-1]
        gap = w["left"] - (prev["left"] + prev["width"])
        if gap > gap_threshold:
            clusters.append([w])
        else:
            clusters[-1].append(w)
    return clusters


def add_dash_spacing(value: str) -> str:
    """IMEI-style values arrive from OCR in wildly inconsistent shapes --
    sometimes glued as one token ('795244-96-733620-8'), sometimes with
    stray/uneven spacing. The internal form expects a single space on
    both sides of every dash ('795244 - 96 - 733620 - 8'), so this
    normalizes ANY dash spacing to that shape, regardless of how OCR
    reported it. Safe to call even on an empty string."""
    if not value:
        return value
    value = re.sub(r'\s*-\s*', ' - ', value)
    value = re.sub(r'\s{2,}', ' ', value)
    return value.strip()


DASH_SPACED_FIELDS = ("File no", "IMMEI-1", "IMMEI-2")


def normalize_dash_fields(result: dict) -> dict:
    """Applies add_dash_spacing to every field where the internal form
    expects spaced dashes, regardless of which extraction path produced
    the value (grid-form labels or the unlabeled continuous row)."""
    for field in DASH_SPACED_FIELDS:
        if field in result:
            result[field] = add_dash_spacing(result[field])
    return result


def map_labels_to_schema(raw_pairs: dict) -> dict:
    result = {field: "" for field in FIELD_ORDER}
    for label, value in raw_pairs.items():
        best_field, best_score = None, 0
        for field in FIELD_ORDER:
            score = fuzz.ratio(label.lower(), field.lower())
            if score > best_score:
                best_field, best_score = field, score
        if best_score >= FUZZY_MATCH_THRESHOLD:
            result[best_field] = value
    return normalize_dash_fields(result)


def extract_grid_form(image_path: str) -> dict:
    words = ocr_with_boxes(image_path)
    lines = group_into_lines(words)

    raw_pairs = {}
    i = 0
    while i < len(lines) - 1:
        label_clusters = cluster_by_gap(lines[i])
        value_clusters = cluster_by_gap(lines[i + 1])

        for label_cluster in label_clusters:
            label_text = " ".join(w["text"] for w in label_cluster)
            label_x = label_cluster[0]["left"]

            if value_clusters:
                nearest = min(value_clusters, key=lambda vc: abs(vc[0]["left"] - label_x))
                value_text = " ".join(w["text"] for w in nearest)
            else:
                value_text = ""

            raw_pairs[label_text] = clean_value(value_text)
        i += 2

    return map_labels_to_schema(raw_pairs)


# ---------------------------------------------------------------------------
# MODE 2: UNLABELED CONTINUOUS ROW (no labels at all, e.g. new11.JPG style —
# all 25 field values run together across just 2-3 OCR lines). There is no
# position-based label to match against here, so this uses a different
# technique: merge OCR's punctuation-split tokens back into whole values
# (dates, emails, IMEI-style codes), then anchor on a handful of values we
# CAN recognize with confidence (title words, email, date pattern, Male/
# Female, CDMA/GSM, card brand names). Everything between two confident
# anchors gets assigned to the field(s) that live in that gap.
#
# IMPORTANT / HONEST LIMITATION: fields with no delimiter between them
# (e.g. "Profession, Mailing street, Mailing city, Mailing postal code" all
# just space-separated free text with zero anchors between them) genuinely
# CANNOT be split with certainty by any automated method — there's no
# information in the text to say where one ends and the next begins. For
# that stretch, this code keeps the raw combined text together (in
# "Mailing street") and leaves the sibling fields blank rather than
# guessing wrong silently. Fix those by hand on the review screen, using
# the side-by-side image.
# ---------------------------------------------------------------------------

TITLE_VOCAB = {"mr", "mrs", "ms", "miss", "dr", "prof"}
GENDER_VOCAB = {"male", "female"}
NETWORK_KEYWORDS = ["gsm", "cdma"]
CARD_KEYWORDS = ["visa", "master", "amex", "discover", "maestro"]
COUNTRY_VOCAB = {"uk", "england", "scotland", "wales", "ireland", "usa", "us", "canada", "australia"}
PROVIDER_VOCAB = {"vodafone", "hutchison", "whampoa", "o2", "orange", "ee",
                   "three", "t-mobile", "tmobile", "at&t", "verizon", "sprint"}

DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
CONNECTORS = {"/", "@", "-", "#", "+"}


def merge_connector_runs(tokens: list) -> list:
    """OCR splits values on punctuation even when it's one logical value:
    '04' '/' '08' '/' '1980' -> '04/08/1980'
    'robert.marsh' '@' 'powellpadilla.co.uk' -> 'robert.marsh@powellpadilla.co.uk'
    'Delta' '-' '299771866' -> 'Delta-299771866'
    'HuWh85cz' '#' '8' '#' '063779525' -> 'HuWh85cz#8#063779525'
    This glues a token back together with what follows as long as they keep
    alternating content/connector/content, and stops the moment two content
    tokens are simply space-separated (a real field boundary)."""
    merged = []
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        if tok == ".":
            # standalone period is an abbreviation dot (e.g. "Mr" "." ->
            # should read "Mr."). Previously this was dropped entirely,
            # which silently deleted the dot from titles. Attach it to the
            # token already collected instead of discarding it.
            if merged:
                merged[-1] += "."
            i += 1
            continue
        if tok in CONNECTORS:
            merged.append(tok)
            i += 1
            continue
        buffer = tok
        j = i + 1
        while j + 1 < n and tokens[j] in CONNECTORS and tokens[j + 1] not in CONNECTORS and tokens[j + 1] != ".":
            buffer += tokens[j] + tokens[j + 1]
            j += 2
        merged.append(buffer)
        i = j
    return merged


def looks_like_date(tok: str) -> bool:
    return bool(DATE_RE.match(tok))


def extract_unlabeled_row(lines: list) -> dict:
    tokens = []
    for line in lines:
        tokens.extend(w["text"] for w in line)
    tokens = merge_connector_runs(tokens)

    result = {field: "" for field in FIELD_ORDER}
    n = len(tokens)
    idx = 0

    # Title
    if idx < n and tokens[idx].lower().rstrip('.') in TITLE_VOCAB:
        result["Title"] = tokens[idx]
        idx += 1

    # Name segment ends at the email anchor
    email_idx = next((i for i in range(idx, n) if "@" in tokens[i]), None)
    if email_idx is not None:
        name_tokens = tokens[idx:email_idx]
        if len(name_tokens) >= 2:
            result["First Name"] = name_tokens[0]
            result["Last Name"] = " ".join(name_tokens[1:])
        elif len(name_tokens) == 1:
            result["First Name"] = name_tokens[0]
        result["Email"] = tokens[email_idx]
        idx = email_idx + 1

    # Father name segment ends at the first date (DOB)
    dob_idx = next((i for i in range(idx, n) if looks_like_date(tokens[i])), None)
    if dob_idx is not None:
        result["Father name"] = " ".join(tokens[idx:dob_idx])
        result["DOB"] = tokens[dob_idx]
        idx = dob_idx + 1

    # Gender anchor
    gender_idx = next((i for i in range(idx, n) if tokens[i].lower() in GENDER_VOCAB), None)
    if gender_idx is not None:
        result["Gender"] = tokens[gender_idx]
        idx = gender_idx + 1

    # Network type anchor (bounds the big ambiguous middle chunk)
    network_idx = next((i for i in range(idx, n) if any(k in tokens[i].lower() for k in NETWORK_KEYWORDS)), None)
    search_end = network_idx if network_idx is not None else n

    provider_idx = next((i for i in range(idx, search_end) if tokens[i].lower() in PROVIDER_VOCAB), None)
    country_idx = next((i for i in range(idx, provider_idx if provider_idx is not None else search_end)
                         if tokens[i].lower() in COUNTRY_VOCAB), None)

    # File no pattern: "<code>-<number>" already glued by merge_connector_runs
    file_no_idx = next((i for i in range((provider_idx + 1) if provider_idx is not None else idx, search_end)
                         if re.match(r'^[A-Za-z]+-\d+$', tokens[i])), None)

    # Everything before the country/provider/file-no anchor = Profession +
    # Mailing street + Mailing city + Mailing postal code, with NO way to
    # split them reliably (see limitation note above) — keep combined.
    ambiguous_stop = next((x for x in [country_idx, provider_idx, file_no_idx, network_idx] if x is not None), n)
    ambiguous_tokens = tokens[idx:ambiguous_stop]
    if ambiguous_tokens:
        result["Mailing street"] = " ".join(ambiguous_tokens)
        idx = ambiguous_stop

    if country_idx is not None:
        result["Mailing country"] = tokens[country_idx]
        idx = max(idx, country_idx + 1)
    if provider_idx is not None:
        result["Service provider"] = tokens[provider_idx]
        idx = max(idx, provider_idx + 1)
    if file_no_idx is not None:
        result["File no"] = add_dash_spacing(tokens[file_no_idx])
        idx = max(idx, file_no_idx + 1)

    # Whatever's left before the network anchor = Reference number + Sim no
    if network_idx is not None:
        remaining = tokens[idx:network_idx]
        if len(remaining) >= 2:
            result["Reference number"] = remaining[0]
            result["Sim no"] = " ".join(remaining[1:])
        elif len(remaining) == 1:
            result["Reference number"] = remaining[0]
        result["Type of network"] = tokens[network_idx]
        idx = network_idx + 1

    # IMEI anchors: digit groups joined by dashes, e.g. "616776-26-952876-4"
    imei_re = re.compile(r'^\d+(-\d+){2,}$')
    imei_indices = [i for i in range(idx, n) if imei_re.match(tokens[i])]

    card_idx = next((i for i in range(idx, n) if any(k in tokens[i].lower() for k in CARD_KEYWORDS)), None)
    cell_end = imei_indices[0] if imei_indices else (card_idx if card_idx is not None else n)
    result["Cell model number"] = " ".join(tokens[idx:cell_end])
    if imei_indices:
        result["IMMEI-1"] = tokens[imei_indices[0]]
        idx = imei_indices[0] + 1
        if len(imei_indices) > 1:
            result["IMMEI-2"] = tokens[imei_indices[1]]
            idx = imei_indices[1] + 1

    # Last two dates = Date of issue / Date of renewal
    date_indices = [i for i in range(idx, n) if looks_like_date(tokens[i])]
    if len(date_indices) >= 2:
        result["Date of issue"] = tokens[date_indices[-2]]
        result["Date of renewal"] = tokens[date_indices[-1]]
        dates_start = date_indices[-2]
    else:
        dates_start = n

    if card_idx is not None:
        result["Type of plan"] = " ".join(tokens[idx:card_idx])
        # take the card keyword plus the word right after it (e.g. "Master Card Silver")
        card_end = min(card_idx + 3, dates_start)
        result["Credit card type"] = " ".join(tokens[card_idx:card_end])
        # whatever's between the card type and the final two dates = contact value
        result["Contact value"] = " ".join(tokens[card_end:dates_start])

    for field in FIELD_ORDER:
        result[field] = clean_value(result[field])

    return normalize_dash_fields(result)


def detect_extraction_mode(lines: list) -> str:
    """Grid forms produce many short lines (one per row of labels/values).
    Unlabeled continuous rows produce very few, very long lines (all fields
    packed into 2-4 lines because of page-width wrapping). Use average words
    per line as the signal."""
    if not lines:
        return "grid"
    avg_words_per_line = sum(len(line) for line in lines) / len(lines)
    if len(lines) <= 5 and avg_words_per_line >= 8:
        return "unlabeled"
    return "grid"


def extract_fields(image_path: str) -> dict:
    words = ocr_with_boxes(image_path)
    lines = group_into_lines(words)
    mode = detect_extraction_mode(lines)
    print(f"Detected extraction mode: {mode}")

    if mode == "unlabeled":
        return extract_unlabeled_row(lines)

    raw_pairs = {}
    i = 0
    while i < len(lines) - 1:
        label_clusters = cluster_by_gap(lines[i])
        value_clusters = cluster_by_gap(lines[i + 1])
        for label_cluster in label_clusters:
            label_text = " ".join(w["text"] for w in label_cluster)
            label_x = label_cluster[0]["left"]
            if value_clusters:
                nearest = min(value_clusters, key=lambda vc: abs(vc[0]["left"] - label_x))
                value_text = " ".join(w["text"] for w in nearest)
            else:
                value_text = ""
            raw_pairs[label_text] = clean_value(value_text)
        i += 2
    return map_labels_to_schema(raw_pairs)


# ---------------------------------------------------------------------------
# Selenium form filling
# ---------------------------------------------------------------------------

# Cache the chromedriver Service at module level. ChromeDriverManager().install()
# makes a network round-trip to check/download the right chromedriver build
# EVERY time it's called. Doing that on every single form submission was one
# of the two big causes of slowness -- now it only happens once per app run.
_driver_service = None


def get_driver_service():
    global _driver_service
    if _driver_service is None:
        _driver_service = Service(ChromeDriverManager().install())
    return _driver_service


def set_input_value(driver, element, value):
    """Fills one input the way a real keyboard would, but done entirely in
    JavaScript in a single round trip -- fast, and never depends on the
    Windows OS keyboard layout.

    ROOT CAUSE of the special-character mistakes (!, $, @, #, %, ^, &, *):
    the target site likely has its own JS listening for 'keydown'/'keypress'
    on these inputs (input masks, validation, formatting). Just setting
    el.value directly (no keyboard-like events at all) skips that JS
    entirely, so on some fields the site's own script then "corrects" or
    strips the character because it never saw it arrive normally.

    THE FIX: type character-by-character *inside the browser*, dispatching
    real keydown/keypress/input/keyup events for every character (with the
    correct key + charCode/keyCode), while also setting el.value directly
    after each keystroke so the final value is guaranteed correct even if
    the site's script tries to block/alter a keypress. This is still ONE
    execute_script() call per field (the loop runs in JS, not in Python),
    so it stays fast -- only the truly slow part (Selenium's own per-char
    send_keys round trips) is gone.

    Returns the value actually left in the field after typing, so the
    caller can detect and log a mismatch if the site's own JS still
    altered something.
    """
    actual_value = driver.execute_script(
        """
        const el = arguments[0];
        const val = arguments[1];
        const proto = el.tagName === 'TEXTAREA'
            ? window.HTMLTextAreaElement.prototype
            : window.HTMLInputElement.prototype;
        const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;

        el.focus();
        setter.call(el, '');

        for (const ch of val) {
            const code = ch.codePointAt(0);
            const opts = {
                key: ch, bubbles: true, cancelable: true,
                charCode: code, keyCode: code, which: code
            };
            el.dispatchEvent(new KeyboardEvent('keydown', opts));
            el.dispatchEvent(new KeyboardEvent('keypress', opts));
            // Set the value ourselves regardless of what the keypress
            // handler did -- guarantees the character always lands.
            setter.call(el, el.value + ch);
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new KeyboardEvent('keyup', opts));
        }

        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur();
        return el.value;
        """,
        element, value
    )
    if actual_value != value:
        print(f"WARNING: field value differs after fill. wanted={value!r} got={actual_value!r}")
    return actual_value


# ---------------------------------------------------------------------------
# Persistent browser session
# ---------------------------------------------------------------------------
# ROOT CAUSE of "next record fails unless I close the previous browser
# first": every call to fill_form() used to open a BRAND NEW Chrome window
# and log in again, but the old window/chromedriver.exe from the previous
# record was never closed. On Windows, a running chromedriver.exe/Chrome
# process keeps files and ports locked, so starting another one on top of
# it is slow and sometimes fails outright -- and it only "works" once you
# manually close the earlier window and free those locks.
#
# THE FIX: keep ONE browser open and logged in, and reuse it for every
# record. This also removes the repeated login step, which was a second
# real source of the slowness. The browser is only replaced if it was
# closed (e.g. you closed the window yourself) or a network hiccup killed
# the session.

_driver = None
_driver_lock = threading.Lock()


def is_driver_alive(driver) -> bool:
    try:
        _ = driver.title  # any call that requires a live browser window
        return True
    except Exception:
        return False


def login_to_system(driver, username, password, target_login_url):
    driver.get(target_login_url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "username")))

    set_input_value(driver, driver.find_element(By.NAME, "username"), username)
    set_input_value(driver, driver.find_element(By.NAME, "password"), password)

    dropdown = Select(driver.find_element(By.TAG_NAME, "select"))
    dropdown.select_by_visible_text(os.environ.get("SERVER_OPTION", "SERVER 3"))

    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Sign In')]"))
    ).click()


def get_logged_in_driver():
    """Returns the single shared, already-logged-in browser, creating and
    logging into it only the first time (or after it was closed/crashed)."""
    global _driver

    username = os.environ.get("LOGIN_USERNAME")
    password = os.environ.get("LOGIN_PASSWORD")
    target_login_url = os.environ.get("TARGET_URL")

    with _driver_lock:
        if _driver is not None and is_driver_alive(_driver):
            return _driver  # reuse: still open, still logged in -- fast path

        service = get_driver_service()
        _driver = webdriver.Chrome(service=service)
        login_to_system(_driver, username, password, target_login_url)
        return _driver


@atexit.register
def _close_driver_on_exit():
    """Make sure Chrome doesn't get left running as an orphan process when
    the Flask app itself stops."""
    global _driver
    if _driver is not None:
        try:
            _driver.quit()
        except Exception:
            pass
        _driver = None


def _do_fill(driver, form_url, data_dict):
    """The actual field-filling work, split out so it can be retried
    against a freshly-recreated driver if the first attempt hits a dead
    browser session."""
    driver.get(form_url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "input")))

    inputs = driver.find_elements(By.TAG_NAME, "input")

    initial_value = "".join([
        data_dict.get("Title", "")[:1],
        data_dict.get("First Name", "")[:1],
        data_dict.get("Last Name", "")[:1],
    ])

    field_values = [
        data_dict["File No"], data_dict["Form No"],
        data_dict["Title"], data_dict["First Name"], data_dict["Last Name"], initial_value, data_dict["Email"],
        data_dict["Father name"], data_dict["DOB"], data_dict["Gender"], data_dict["Profession"],
        data_dict["Mailing street"], data_dict["Mailing city"], data_dict["Mailing postal code"],
        data_dict["Mailing country"], data_dict["Service provider"], data_dict["File no"],
        data_dict["Reference number"], data_dict["Sim no"], data_dict["Type of network"],
        data_dict["Cell model number"], data_dict["IMMEI-1"], data_dict["IMMEI-2"], data_dict["Type of plan"],
        data_dict["Credit card type"], data_dict["Contact value"], data_dict["Date of issue"],
        data_dict["Date of renewal"], data_dict["Installments"], data_dict["Amount in words"], data_dict["Remarks"],
    ]

    for i, value in enumerate(field_values):
        try:
            set_input_value(driver, inputs[i], value)
        except IndexError:
            print(f"Field index {i} does not exist on this form")
        except Exception as e:
            print(f"Could not fill field {i}: {e}")

    print("Form auto-filled successfully")


def fill_form(data_dict: dict):
    username = os.environ.get("LOGIN_USERNAME")
    password = os.environ.get("LOGIN_PASSWORD")
    target_login_url = os.environ.get("TARGET_URL")
    form_url = os.environ.get("FORM_URL")

    if not all([username, password, target_login_url, form_url]):
        raise RuntimeError("LOGIN_USERNAME / LOGIN_PASSWORD / TARGET_URL / FORM_URL missing from .env")

    driver = get_logged_in_driver()

    try:
        _do_fill(driver, form_url, data_dict)
    except WebDriverException as e:
        # Covers EVERY way the shared browser session can die between
        # records: you closed the window, Chrome crashed, the site
        # invalidated the session, "chrome not reachable", "no such
        # window", "invalid session id", etc. Selenium's real reason is
        # usually the first line of e.msg -- print it plainly so it's easy
        # to read instead of only the noisy chromedriver stacktrace.
        print(f"Browser session appears dead ({e.msg or e}). Restarting and retrying once...")
        global _driver
        with _driver_lock:
            _driver = None
        driver = get_logged_in_driver()
        _do_fill(driver, form_url, data_dict)
    # The browser is deliberately left open (not driver.quit()) so you can
    # visually verify and click Submit yourself -- the last safety
    # checkpoint before data goes live. It now stays open and logged in
    # for the NEXT record too, instead of piling up new windows.


def delete_uploaded_image(filename: str):
    """Removes the working copy from static/uploads once the form has been
    filled."""
    if not filename:
        return
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
            print(f"Deleted uploaded image: {filepath}")
    except Exception as e:
        print(f"Could not delete {filepath}: {e}")



def delete_source_image(filename: str):
    """Delete original image from fixed DataImages folder only after successful form fill."""
    if not filename:
        return

    source_path = os.path.join(SOURCE_IMAGE_FOLDER, filename)

    try:
        if os.path.exists(source_path):
            os.remove(source_path)
            print(f"Deleted original source image: {source_path}")
        else:
            print(f"Original image not found: {source_path}")
    except Exception as e:
        print(f"Could not delete original source image {source_path}: {e}")

def pick_file_dialog():
    """Opens a native Windows 'Open File' dialog (via PowerShell) and
    returns the full absolute path the user picked, or None if cancelled.

    WHY THIS EXISTS: a normal browser <input type="file"> NEVER reveals the
    real local path of the file you pick -- browsers hide it on purpose for
    security, sending only the file's bytes and its bare filename. That
    means the server has no way to find, let alone delete, the original
    file on your computer through a regular web upload -- this is a browser
    restriction, not something any server-side code can work around.

    Since this app is meant to run locally on your own Windows PC (you open
    it at localhost in your browser), we can sidestep the browser entirely
    for file selection: this spawns Windows' own native file-picker dialog
    directly on your machine and reads back the real path it returns. That
    real path is what lets delete_source_image() later remove the actual
    original file, not just our working copy.
    """
    ps_script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$f = New-Object System.Windows.Forms.OpenFileDialog;"
        "$f.Filter = 'Image files (*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff)|"
        "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff|All files (*.*)|*.*';"
        "$f.Title = 'Select scanned form image';"
        "if ($f.ShowDialog() -eq 'OK') { Write-Output $f.FileName }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True, text=True, timeout=120
        )
        path = result.stdout.strip()
        return path if path else None
    except Exception as e:
        print(f"Native file dialog failed: {e}")
        return None



def get_next_image_from_dataimages():
    """Return first image found in DataImages folder."""
    if not os.path.exists(SOURCE_IMAGE_FOLDER):
        return None

    exts = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")

    files = [
        f for f in os.listdir(SOURCE_IMAGE_FOLDER)
        if f.lower().endswith(exts)
    ]

    if not files:
        return None

    files.sort()
    return os.path.join(SOURCE_IMAGE_FOLDER, files[0])


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route("/browse-file", methods=["GET"])
def browse_file():
    """Triggers the native Windows file dialog and returns the chosen
    absolute path as JSON, so the landing page's 'Browse' button can show
    it and submit it -- instead of a normal <input type="file">."""
    path = pick_file_dialog()
    return {"path": path}


@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        source_path = get_next_image_from_dataimages()

        if not source_path:
            return render_template(
                "index.html",
                error="No images found in DataImages folder."
            ), 400

        filename = os.path.basename(source_path)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        # Copy (not move) into static/uploads -- this copy is what OCR reads
        # and what the review page displays. The ORIGINAL at source_path is
        # left untouched until the form fill actually succeeds.
        shutil.copy2(source_path, filepath)

        try:
            data_dict = extract_fields(filepath)
            file_no, form_no = extract_file_and_form_no(filename)

            data_dict["File No"] = file_no
            data_dict["Form No"] = form_no

            contact_value = extract_numeric_value(
                data_dict.get("Contact value", "")
            )

            installments = calculate_installments(
                contact_value,
                data_dict.get("Date of issue", ""),
                data_dict.get("Date of renewal", "")
            )

            data_dict["Installments"] = str(installments)

            data_dict["Amount in words"] = amount_to_words(
                installments
            )

            data_dict["Remarks"] = "Not Applicable"
        except Exception as e:
            # Show the image + the error together so you can see what was
            # uploaded and why extraction failed, instead of a bare 500 page.
            # Only the static/uploads working copy is removed here -- the
            # original file on your computer is left alone since nothing
            # succeeded yet.
            image_url = url_for("static", filename=f"uploads/{filename}")
            error_message = str(e)
            delete_uploaded_image(filename)
            return render_template("index.html", error=error_message, image_url=image_url), 500

        session["extracted_data"] = data_dict
        session["image_filename"] = filename
        return redirect(url_for("review"))

    return render_template("index.html")


@app.route("/review", methods=["GET", "POST"])
def review():
    if request.method == "POST":
        # Pull possibly-edited values back from the review form
        corrected = {field: request.form.get(field, "") for field in FIELD_ORDER}
        image_filename = session.get("image_filename")
        try:
            fill_form(corrected)
        except Exception as e:
            return f"Form fill failed: {e}", 500

        # Clean up in BOTH places now that the form has been filled:
        #   1. the working copy in static/uploads
        #   2. the ORIGINAL file you picked from your computer
        delete_uploaded_image(image_filename)
        delete_source_image(image_filename)
        session.pop("extracted_data", None)
        session.pop("image_filename", None)

        return "Reviewed data sent to the form. Please check the browser window and click Submit there."

    data_dict = session.get("extracted_data", {field: "" for field in FIELD_ORDER})
    image_filename = session.get("image_filename")
    image_url = url_for("static", filename=f"uploads/{image_filename}") if image_filename else None
    return render_template("review.html", fields=FIELD_ORDER, data=data_dict, image_url=image_url)


@app.route("/close-browser", methods=["POST"])
def close_browser():
    """Manual escape hatch: force-closes the shared browser session so the
    next record opens a completely fresh, freshly-logged-in window. You
    shouldn't need this in normal use since the browser now reuses itself
    automatically -- it's here in case the target site's session gets into
    a bad state and you want a clean restart without stopping the Flask app."""
    _close_driver_on_exit()
    return "Browser closed. The next form fill will open a fresh session."


if __name__ == "__main__":
    app.run(debug=True)
