import os
import re
import difflib
from getpass import getpass
import pandas as pd
from playwright.sync_api import sync_playwright
import subprocess
from dotenv import load_dotenv



load_dotenv() # reads .env and loads its content into  the environment

#FILENAME = "2026 NDC Members Temp (Autosaved).xls"
PC_USRNAME = os.getlogin()
#BASE_URL = "http://127.0.0.1:8002/" # testing on local machine
BASE_URL = os.getenv("BASE_URL")
BASE_DIRECTORY = f"C:/Users/{PC_USRNAME}/Documents"
EXCEL_PATH = f"{BASE_DIRECTORY}" #"2026 NDC Members Temp.xls"  # .xls needs: pip install xlrd
USERNAME = os.getenv("REGISTRY_USER", "admin")

ERROR_SNAPSHOT_DIRECTORY = f"C:/Users/{PC_USRNAME}/Desktop/ndc_registry_snapshot_errors"
CSV_ERROR_REPORT = f"C:/Users/{PC_USRNAME}/Desktop/ndc_registry_csv_error_report"

HEADER_ROW = 5             # Excel row holding "First Name | Surname | ..." (rows 1-4 are the banner)
ACTION_TIMEOUT_MS = 8000   # fail fast instead of Playwright's 30 s default
MAX_CONSECUTIVE_ERRORS = 3 # stop a station if the script itself keeps failing (bad selector etc.)

# checking for filename
def selecting_excel_sheet():
    """
    Scans BASE_DIRECTORY for matching workbooks and returns the chosen filename.
    Returns None if the user cancels or nothing is found and nothing is typed.
    """
    # os.scandir gives directory entries, not just files, so filter to real files only.
    # entry.is_file() avoids a folder whose name happens to match the pattern being
    # picked up and later crashing pd.read_excel().
    with os.scandir(BASE_DIRECTORY) as entries:
        choose_file = [
            entry.name for entry in entries
            if entry.is_file()
            and entry.name.endswith(".xls")
            and "2026 NDC Members" in entry.name
        ]

    # --- No match at all ---
    if len(choose_file) == 0:
        print(f"No matching .xls file found in {BASE_DIRECTORY}.")
        typed = input("Type the filename manually (or press Enter to cancel): ").strip()
        return typed or None

    # --- Exactly one match: use it, no need to ask ---
    if len(choose_file) == 1:
        print(f"Using: {choose_file[0]}")
        return choose_file[0]

            

    # --- More than one match: show a numbered list and ask ---
    print("More than one matching file found:")
    for i, file in enumerate(choose_file, start=1):
        print(f"  [{i}] {file}")


    # Loop until a valid number is entered, instead of failing on the first bad input
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



# calling the excell select sheet function here
filename = selecting_excel_sheet()
if not filename:
    print("No file selected. Exiting.")
    

EXCEL_PATH = os.path.join(BASE_DIRECTORY, filename)
print(f"Using workbook: {EXCEL_PATH}")

# left = internal key, right = header text in your sheet
COLUMN_MAP = {
    "first_name": "First Name",
    "last_name": "Surname",
    "membership_id": "Party ID Number",
    "voter_id": "Voter ID Number",
    "phone": "Telephone Number",
}

# Every field the form code knows about; ones not in COLUMN_MAP stay "" and are skipped
ALL_FIELDS = ["first_name", "last_name", "middle_name", "gender", "dob", "phone",
              "occupation", "address", "ghana_card", "membership_id", "voter_id"]

# Fields marked * on the form. Rows missing any of these are skipped and reported
REQUIRED = ["first_name", "last_name", "membership_id", "voter_id"]

GENDER_MAP = {"m": "Male", "male": "Male", "f": "Female", "female": "Female"}


def normalise(text):
    """Lowercase, unify apostrophes and collapse spaces: "Wura’s  Palace" -> "wura's palace"."""
    text = str(text).replace("’", "'").replace("‘", "'")
    return re.sub(r"\s+", " ", text).strip().lower()


def choose_sheet(sheet_names):
    """Keep asking until the typed name matches a sheet. Returns the exact sheet name, or None to quit."""
    lookup = {normalise(name): name for name in sheet_names}  # normalised -> real sheet name
    #print("line 44: ", lookup)
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

        # 1) Exact match (case-insensitive)
        if key in lookup:
            return lookup[key]

        # 2) Partial match: typed text appears inside a sheet name
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

        # 3) Close spelling suggestions
        close = difflib.get_close_matches(key, lookup.keys(), n=3, cutoff=0.6)
        if close:
            print("No exact match. Similar names: " + ", ".join(lookup[c] for c in close))
        else:
            print("No matching sheet found. Type 'l' to see all polling stations.")


def ask_start_row():
    """Lets you resume part-way through a sheet (e.g. after an error) without re-entering earlier rows."""
    first_data_row = HEADER_ROW + 1
    typed = input(f"Start from Excel row (Enter = {first_data_row}): ").strip()
    return int(typed) if typed.isdigit() else first_data_row


def clean(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):  # Excel numbers sometimes arrive as "556450221.0"
        text = text[:-2]
    return text


def prepare_row(raw):
    row = {key: "" for key in ALL_FIELDS}  # start blank so missing columns are simply skipped
    for key, header in COLUMN_MAP.items():
        row[key] = clean(raw.get(normalise(header), ""))

    # Excel drops the leading 0 of phone numbers (556450221 -> 0556450221)
    row["phone"] = row["phone"].replace(" ", "")
    if row["phone"].isdigit() and len(row["phone"]) == 9:
        row["phone"] = "0" + row["phone"]

    if row["dob"]:
        parsed = pd.to_datetime(row["dob"], dayfirst=True, errors="coerce")
        row["dob"] = "" if pd.isna(parsed) else parsed.strftime("%Y-%m-%d")

    row["gender"] = GENDER_MAP.get(row["gender"].lower(), row["gender"])
    return row


def select_station(page, station):
    """Pick the dropdown option that matches the sheet name (ignoring case/spacing/apostrophe style)."""
    select = page.locator("select").last  # the polling station dropdown is the last <select> on the form
    options = select.locator("option").evaluate_all(
        "els => els.map(e => ({value: e.value, text: e.textContent}))"
    )
    target = normalise(station)

    # Try an exact match first, then "option text contains the sheet name"
    match = next((o for o in options if normalise(o["text"]) == target), None) \
        or next((o for o in options if target in normalise(o["text"])), None)
    
    if not match:
        raise ValueError(f"Polling station '{station}' not found in the dropdown")
    select.select_option(value=match["value"])


def fill_member(page, row, station):
    # Placeholders come straight from your screenshot. get_by_placeholder matches a
    # substring, so we don't need to count the X's exactly.
    page.get_by_placeholder("e.g. Sheriff").fill(row["first_name"])
    page.get_by_placeholder("e.g. Sakara").fill(row["last_name"])
    page.get_by_placeholder("e.g. MXXX").fill(row["membership_id"])   # Membership ID No.
    page.get_by_placeholder("e.g. XXXX").fill(row["voter_id"])        # Voter ID No.

    # Optional fields: only filled when the sheet has a value for them
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
        page.locator("select").first.select_option(label=row["gender"])  # first <select> = Gender
    if row["dob"]:
        page.locator('input[type="date"]').fill(row["dob"])              # needs yyyy-mm-dd

    select_station(page, station)


def enter_station(page, register_url, station, start_row):
    """Read ONE sheet and enter every member on it, one form submission per row."""
    # header=4 means "row 5 is the header", so the banner rows above it are ignored
    df = pd.read_excel(EXCEL_PATH, sheet_name=station, header=HEADER_ROW - 1, dtype=str)
    df.columns = [normalise(c) for c in df.columns]  # tolerate stray spaces/case in headers

    missing = [h for h in COLUMN_MAP.values() if normalise(h) not in df.columns]
    if missing:
        print(f"Sheet '{station}' is missing columns {missing}. Found: {list(df.columns)}")
        return

    # Drop pre-numbered blank rows: keep a row only if it has a first name or a surname
    name_cols = [normalise(COLUMN_MAP["first_name"]), normalise(COLUMN_MAP["last_name"])]
    df = df.dropna(subset=name_cols, how="all")

    # Resume support: pandas index 0 = Excel row HEADER_ROW + 1
    df = df[(df.index + HEADER_ROW + 1) >= start_row]

    print(f"\n=== {station}: {len(df)} members (from Excel row {start_row}) ===")
    results = []
    consecutive_errors = 0

    for idx, raw in df.iterrows():
        row = prepare_row(raw)
        excel_row = idx + HEADER_ROW + 1
        label = f"{row['first_name']} {row['last_name']} ({row['membership_id']})"

        if not all(row[k] for k in REQUIRED):
            print(f"SKIPPED (missing required data): Excel row {excel_row}")
            results.append((station, excel_row, label, "SKIPPED - missing data"))
            continue

        try:
            page.goto(register_url, wait_until="domcontentloaded")  # fresh blank form
            fill_member(page, row, station)

            # Click "Register Member" and wait for the resulting page load. A successful save
            # redirects; a validation error reloads the same URL. Both count as a navigation.
            with page.expect_navigation(wait_until="domcontentloaded", timeout=15000):
                page.get_by_role("button", name="Register Member").click()

            consecutive_errors = 0  # the script itself worked, whatever the server decided

            if page.url.rstrip("/") == register_url.rstrip("/"):
                # Still on the register page = Django showed a validation error (duplicate ID, bad format...)
                page.screenshot(path=f"error_{row['membership_id']}.png", full_page=True)
                results.append((station, excel_row, label, "FAILED - validation error"))
                print(f"FAILED (validation): {label} - see error_{row['membership_id']}.png")
            else:
                results.append((station, excel_row, label, "OK"))
                print(f"OK: {label}")

        except Exception as e:
            # checking if error snapshot directory is not created
            if not os.path.exists(ERROR_SNAPSHOT_DIRECTORY):
                os.makedirs(ERROR_SNAPSHOT_DIRECTORY) # create a directory

            consecutive_errors += 1
            page.screenshot(path=f"{ERROR_SNAPSHOT_DIRECTORY}/error_row_{excel_row}.png", full_page=True)
            results.append((station, excel_row, label, f"ERROR - {e}"))
            print(f"ERROR on {label}: {e}")

            # A broken selector would otherwise fail on every remaining row, so stop early
            if consecutive_errors >= MAX_CONSECUTIVE_ERRORS:
                print(f"\nStopping '{station}': {MAX_CONSECUTIVE_ERRORS} script errors in a row. "
                      f"Fix the issue, then resume from Excel row {excel_row - MAX_CONSECUTIVE_ERRORS + 1}.")
                break

    # Save a per-station report
    # checking if directory is not created then create one
    if not os.path.exists(CSV_ERROR_REPORT):
        os.makedirs(CSV_ERROR_REPORT)
    safe_name = re.sub(r"[^\w\-]+", "_", station)
    pd.DataFrame(results, columns=["Station", "Excel Row", "Member", "Status"]).to_csv(
        f"{CSV_ERROR_REPORT}/entry_report_{safe_name}.csv", index=False
    )
    ok = sum(1 for r in results if r[3] == "OK")
    print(f"Done: {ok} saved, {len(results) - ok} need attention (see entry_report_{safe_name}.csv)")


def main():

    USERNAME = input("Enter username: ").strip() #os.getenv("REGISTRY_USER") or getpass("Username: ")

    if USERNAME == os.getenv("SHEFF"):
        password = os.getenv("SHEFFPASS")

    elif USERNAME == os.getenv("HARD"):
        password = os.getenv("HARDPASS")
        
    else:
        password = getpass("Registry password: ")

    assert password, "Password is empty — check your .env file has SHEFF and SHEFFPASS set correctly"

    # Ask for the station BEFORE launching the browser, so a typo costs nothing
    sheet_names = pd.ExcelFile(EXCEL_PATH).sheet_names
    station = choose_sheet(sheet_names)
    if station is None:
        print("Cancelled.")
        return
    start_row = ask_start_row()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page(viewport={"width": 1280, "height": 900})
        page.set_default_timeout(ACTION_TIMEOUT_MS)

        try:
            # ---------- Login (once) ----------
            page.goto(BASE_URL, wait_until="domcontentloaded")
            page.locator('input[name="username"]').fill(USERNAME)
            page.locator("#password-input").fill(password)
            page.get_by_role("button", name="Sign In").click()
            page.wait_for_load_state("networkidle", timeout=15000)

            page.get_by_role("link", name="Register New Member").click()
            page.wait_for_url("**/members/register/")
            register_url = page.url

            # ---------- Process the chosen station, then offer another ----------
            while station:
                enter_station(page, register_url, station, start_row)

                again = input("\nEnter another polling station? (y/n): ").strip().lower()
                if again == "y":
                    station = choose_sheet(sheet_names)
                    start_row = ask_start_row() if station else None
                else:
                    station = None

        finally:
            browser.close()


if __name__ == "__main__":
    main()