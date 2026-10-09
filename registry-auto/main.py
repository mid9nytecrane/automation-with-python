import os
import sys

# FIX: must be set BEFORE playwright is imported/started. "0" tells Playwright to look for
# browsers inside its own package folder (playwright/driver/package/.local-browsers).
# That is the folder PyInstaller copies when Chromium was installed with the same setting
# at build time (see build_exe.bat). Without this the exe looks in %USERPROFILE%\AppData\...
# on the target PC, finds nothing, and crashes with "Executable doesn't exist".
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")

import re
import difflib
from pathlib import Path
from getpass import getpass
import pandas as pd
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# Folder that holds the .exe (frozen) or this .py file (development). The .env lives here.
base_path = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(base_path, ".env"))
# FIX: removed the second bare load_dotenv() - it just searched the current folder again.

# FIX: Path.home() instead of os.getlogin(): getlogin() can raise OSError (services, some
# RDP sessions) and hard-coding C:/Users/<name>/Desktop breaks when Desktop is redirected.
HOME = Path.home()
BASE_DIRECTORY = str(HOME / "Desktop")
if not os.path.isdir(BASE_DIRECTORY):  # OneDrive-redirected Desktop fallback
    onedrive_desktop = HOME / "OneDrive" / "Desktop"
    if onedrive_desktop.is_dir():
        BASE_DIRECTORY = str(onedrive_desktop)

BASE_URL = os.getenv("BASE_URL")
ERROR_SNAPSHOT_DIRECTORY = os.path.join(BASE_DIRECTORY, "ndc_registry_snapshot_errors")
CSV_ERROR_REPORT = os.path.join(BASE_DIRECTORY, "ndc_registry_csv_error_report")

HEADER_ROW = 5             # Excel row holding "First Name | Surname | ..." (rows 1-4 are the banner)
ACTION_TIMEOUT_MS = 8000   # fail fast instead of Playwright's 30 s default
MAX_CONSECUTIVE_ERRORS = 3 # stop a station if the script itself keeps failing (bad selector etc.)

# left = internal key, right = header text in your sheet
COLUMN_MAP = {
    "first_name": "First Name",
    "last_name": "Surname",
    "membership_id": "Party ID Number",
    "voter_id": "Voter ID Number",
    "phone": "Telephone Number",
}

ALL_FIELDS = ["first_name", "last_name", "middle_name", "gender", "dob", "phone",
              "occupation", "address", "ghana_card", "membership_id", "voter_id"]

REQUIRED = ["first_name", "last_name", "membership_id", "voter_id"]

GENDER_MAP = {"m": "Male", "male": "Male", "f": "Female", "female": "Female"}


def normalise(text):
    """Lowercase, unify apostrophes and collapse spaces: "Wura’s  Palace" -> "wura's palace"."""
    text = str(text).replace("’", "'").replace("‘", "'")
    return re.sub(r"\s+", " ", text).strip().lower()


def pause_and_exit(message, code=1):
    """FIX: in an .exe the console window closes instantly on exit, so errors were unreadable."""
    print(message)
    input("\nPress Enter to close...")
    sys.exit(code)


def selecting_excel_sheet():
    """Scans BASE_DIRECTORY for matching workbooks and returns the chosen filename (or None)."""
    with os.scandir(BASE_DIRECTORY) as entries:
        choose_file = [
            entry.name for entry in entries
            # FIX: the original was `is_file() and .xls  or  .xlsx and "2026 NDC Members"`.
            # `and` binds tighter than `or`, so EVERY .xls was listed (any name) and the
            # is_file() check only applied to .xls. Parentheses make all three rules apply together.
            if entry.is_file()
            and entry.name.lower().endswith((".xls", ".xlsx"))
            and "2026 ndc members" in entry.name.lower()
            and not entry.name.startswith("~$")  # Excel's temporary lock files
        ]

    if len(choose_file) == 0:
        print(f"No matching Excel file found in {BASE_DIRECTORY}.")
        typed = input("Type the filename manually (or press Enter to cancel): ").strip()
        return typed or None

    if len(choose_file) == 1:
        print(f"Using: {choose_file[0]}")
        return choose_file[0]

    print("More than one matching file found:")
    for i, file in enumerate(choose_file, start=1):
        print(f"  [{i}] {file}")

    while True:
        raw = input(f"Choose a file to work with (1-{len(choose_file)}, or 'q' to cancel): ").strip()
        if raw.lower() == "q":
            return None
        if not raw.isdigit():
            print("Please enter a number.")
            continue
        choice = int(raw)
        if 1 <= choice <= len(choose_file):
            selected = choose_file[choice - 1]
            print(f"Selected: {selected}")
            return selected
        print(f"Enter a number between 1 and {len(choose_file)}.")


def choose_sheet(sheet_names):
    """Keep asking until the typed name matches a sheet. Returns the exact sheet name, or None to quit."""
    lookup = {normalise(name): name for name in sheet_names}
    while True:
        typed = input("\nEnter polling station name (or 'q' to quit, 'l' to list all): ").strip()

        if normalise(typed) == "q":
            return None

        if normalise(typed) == "l":
            print("Available polling stations:")
            for name in sheet_names:
                print(f"  - {name}")
            continue

        if not typed:
            continue

        key = normalise(typed)

        if key in lookup:
            return lookup[key]

        partial = [name for k, name in lookup.items() if key in k]
        if len(partial) == 1:
            confirm = input(f"Did you mean '{partial[0]}'? (y/n): ").strip().lower()
            if confirm == "y":
                return partial[0]
            continue
        if len(partial) > 1:
            print("More than one sheet matches:")
            for name in partial:
                print(f"  - {name}")
            continue

        close = difflib.get_close_matches(key, lookup.keys(), n=3, cutoff=0.6)
        if close:
            print("No exact match. Similar names: " + ", ".join(lookup[c] for c in close))
        else:
            print("No matching sheet found. Type 'l' to see all polling stations.")


def ask_start_row():
    first_data_row = HEADER_ROW + 1
    typed = input(f"Start from Excel row (Enter = {first_data_row}): ").strip()
    return int(typed) if typed.isdigit() else first_data_row


def clean(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def prepare_row(raw):
    row = {key: "" for key in ALL_FIELDS}
    for key, header in COLUMN_MAP.items():
        row[key] = clean(raw.get(normalise(header), ""))

    row["phone"] = row["phone"].replace(" ", "")
    if row["phone"].isdigit() and len(row["phone"]) == 9:
        row["phone"] = "0" + row["phone"]

    if row["dob"]:
        parsed = pd.to_datetime(row["dob"], dayfirst=True, errors="coerce")
        row["dob"] = "" if pd.isna(parsed) else parsed.strftime("%Y-%m-%d")

    row["gender"] = GENDER_MAP.get(row["gender"].lower(), row["gender"])
    return row


def select_station(page, station):
    select = page.locator("select").last
    options = select.locator("option").evaluate_all(
        "els => els.map(e => ({value: e.value, text: e.textContent}))"
    )
    target = normalise(station)

    match = next((o for o in options if normalise(o["text"]) == target), None) \
        or next((o for o in options if target in normalise(o["text"])), None)

    if not match:
        raise ValueError(f"Polling station '{station}' not found in the dropdown")
    select.select_option(value=match["value"])


def fill_member(page, row, station):
    page.get_by_placeholder("e.g. Sheriff").fill(row["first_name"])
    page.get_by_placeholder("e.g. Sakara").fill(row["last_name"])
    page.get_by_placeholder("e.g. MXXX").fill(row["membership_id"])
    page.get_by_placeholder("e.g. XXXX").fill(row["voter_id"])

    if row["middle_name"]:
        page.get_by_placeholder("e.g. Shunkpa").fill(row["middle_name"])
    if row["phone"]:
        page.get_by_placeholder("e.g. 0244123456").fill(row["phone"])
    if row["occupation"]:
        page.get_by_placeholder("e.g. Farmer").fill(row["occupation"])
    if row["address"]:
        page.get_by_placeholder("Area / community / street").fill(row["address"])
    if row["ghana_card"]:
        page.get_by_placeholder("e.g. GHA-000000000-0").fill(row["ghana_card"])
    if row["gender"]:
        page.locator("select").first.select_option(label=row["gender"])
    if row["dob"]:
        page.locator('input[type="date"]').fill(row["dob"])

    select_station(page, station)


def enter_station(page, excel_path, register_url, station, start_row):
    """Read ONE sheet and enter every member on it, one form submission per row."""
    # FIX: excel_path is now a parameter instead of a module-level global that was only
    # assigned after top-level code ran; makes the file safe to import and easier to test.
    df = pd.read_excel(excel_path, sheet_name=station, header=HEADER_ROW - 1, dtype=str)
    df.columns = [normalise(c) for c in df.columns]

    missing = [h for h in COLUMN_MAP.values() if normalise(h) not in df.columns]
    if missing:
        print(f"Sheet '{station}' is missing columns {missing}. Found: {list(df.columns)}")
        return

    name_cols = [normalise(COLUMN_MAP["first_name"]), normalise(COLUMN_MAP["last_name"])]
    df = df.dropna(subset=name_cols, how="all")
    df = df[(df.index + HEADER_ROW + 1) >= start_row]

    print(f"\n=== {station}: {len(df)} members (from Excel row {start_row}) ===")
    results = []
    consecutive_errors = 0

    # FIX: create both output folders once, up front. Before, the validation-error screenshot
    # was written to the *current folder* (unpredictable/unwritable for an exe) as error_<id>.png.
    os.makedirs(ERROR_SNAPSHOT_DIRECTORY, exist_ok=True)
    os.makedirs(CSV_ERROR_REPORT, exist_ok=True)

    for idx, raw in df.iterrows():
        row = prepare_row(raw)
        excel_row = idx + HEADER_ROW + 1
        label = f"{row['first_name']} {row['last_name']} ({row['membership_id']})"

        if not all(row[k] for k in REQUIRED):
            print(f"SKIPPED (missing required data): Excel row {excel_row}")
            results.append((station, excel_row, label, "SKIPPED - missing data"))
            continue

        try:
            page.goto(register_url, wait_until="domcontentloaded")
            fill_member(page, row, station)

            with page.expect_navigation(wait_until="domcontentloaded", timeout=15000):
                page.get_by_role("button", name="Register Member").click()

            consecutive_errors = 0

            if page.url.rstrip("/") == register_url.rstrip("/"):
                shot = os.path.join(ERROR_SNAPSHOT_DIRECTORY, f"validation_{row['membership_id']}.png")
                page.screenshot(path=shot, full_page=True)
                results.append((station, excel_row, label, "FAILED - validation error"))
                print(f"FAILED (validation): {label} - see {shot}")
            else:
                results.append((station, excel_row, label, "OK"))
                print(f"OK: {label}")

        except Exception as e:
            consecutive_errors += 1
            # FIX: the screenshot itself can throw (browser closed/crashed) and would have
            # killed the whole run from inside the except block.
            try:
                page.screenshot(path=os.path.join(ERROR_SNAPSHOT_DIRECTORY, f"error_row_{excel_row}.png"), full_page=True)
            except Exception:
                pass
            results.append((station, excel_row, label, f"ERROR - {e}"))
            print(f"ERROR on {label}: {e}")

            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                print(f"\nStopping '{station}': {MAX_CONSECUTIVE_ERRORS} script errors in a row. "
                      f"Fix the issue, then resume from Excel row {excel_row - MAX_CONSECUTIVE_ERRORS + 1}.")
                break

    safe_name = re.sub(r"[^\w\-]+", "_", station)
    pd.DataFrame(results, columns=["Station", "Excel Row", "Member", "Status"]).to_csv(
        os.path.join(CSV_ERROR_REPORT, f"entry_report_{safe_name}.csv"), index=False
    )
    ok = sum(1 for r in results if r[3] == "OK")
    print(f"Done: {ok} saved, {len(results) - ok} need attention (see entry_report_{safe_name}.csv)")


def launch_browser(p):
    """
    FIX (new): try the Chromium bundled inside the exe first, then fall back to the Microsoft
    Edge / Google Chrome already installed on the PC. Windows 10/11 always ships with Edge,
    so the fallback means the tool still works even if the bundled browser is missing.
    """
    attempts = [{}, {"channel": "msedge"}, {"channel": "chrome"}]
    last_error = None
    for opts in attempts:
        try:
            browser = p.chromium.launch(headless=False, **opts)
            print(f"Browser started ({opts.get('channel', 'bundled Chromium')}).")
            return browser
        except Exception as e:
            last_error = e
    pause_and_exit(f"Could not start any browser (bundled Chromium, Edge, Chrome).\n{last_error}")


def main():
    # FIX: everything that used to run at import time (file picking, sys exit, etc.) now lives
    # here. Previously, if no file was chosen it printed "Exiting." but did NOT exit, then
    # crashed on os.path.join(BASE_DIRECTORY, None).
    if not BASE_URL:
        pause_and_exit(f"BASE_URL is not set. Put a .env file next to the program ({base_path}) containing:\nBASE_URL=https://your-registry-url/")

    filename = selecting_excel_sheet()
    if not filename:
        pause_and_exit("No file selected. Exiting.", code=0)

    excel_path = os.path.join(BASE_DIRECTORY, filename)
    if not os.path.isfile(excel_path):
        pause_and_exit(f"File not found: {excel_path}")
    print(f"Using workbook: {excel_path}")

    username = input("Enter username: ").strip()
    password = getpass("Registry password: ")

    # FIX: `assert` is stripped when Python runs with -O and gives an ugly traceback in an exe.
    if not password:
        pause_and_exit("Password is empty - check your .env has SHEFF/SHEFFPASS (or HARD/HARDPASS) set correctly.")

    sheet_names = pd.ExcelFile(excel_path).sheet_names
    station = choose_sheet(sheet_names)
    if station is None:
        print("Cancelled.")
        return
    start_row = ask_start_row()

    with sync_playwright() as p:
        browser = launch_browser(p)
        page = browser.new_page(viewport={"width": 1280, "height": 900})
        page.set_default_timeout(ACTION_TIMEOUT_MS)

        try:
            page.goto(BASE_URL, wait_until="domcontentloaded")
            page.locator('input[name="username"]').fill(username)
            page.locator("#password-input").fill(password)
            page.get_by_role("button", name="Sign In").click()
            page.wait_for_load_state("networkidle", timeout=15000)

            page.get_by_role("link", name="Register New Member").click()
            page.wait_for_url("**/members/register/")
            register_url = page.url

            while station:
                enter_station(page, excel_path, register_url, station, start_row)

                again = input("\nEnter another polling station? (y/n): ").strip().lower()
                if again == "y":
                    station = choose_sheet(sheet_names)
                    start_row = ask_start_row() if station else None
                else:
                    station = None

        finally:
            browser.close()

    input("\nAll done. Press Enter to close...")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:
        # FIX: keep the console open so the user can read unexpected errors on the target PC.
        import traceback
        traceback.print_exc()
        pause_and_exit(f"\nUnexpected error: {exc}")
