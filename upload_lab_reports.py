"""
Upload Excel lab reports into the clinic system, one patient per file.

Folder layout (next to this script):
    reports/     put the Excel files here, named <patientID>-First_Last.xlsx
                 (a file is REMOVED from here as soon as its upload is confirmed)
    error/       every file that could not be uploaded is moved here, with the reason in results.csv
                 (move it back into reports/ after fixing it, to retry)
    screenshots/ one screenshot per failed attempt
    results.csv  one row per file per run (the log)

Usage:
    python upload_lab_reports.py --headed --limit 1     # first test, one file, watch it
    python upload_lab_reports.py --dry-run              # list what would be uploaded
    python upload_lab_reports.py                        # normal unattended run
"""
import argparse
import csv
import os
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import expect, sync_playwright

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
BASE_URL = os.getenv("CLINIC_BASE_URL", "https://indiaclinic.twinsdigital.com")
INBOX = Path("reports")
PROCESSED = Path("processed")      # only used with --keep-processed
FAILED = Path("error")             # files that could not be uploaded
SHOTS = Path("screenshots")
LOG_CSV = Path("results.csv")
AUTH_FILE = Path("auth.json")      # saved login session; treat like a password
RECYCLE_EVERY = 25                 # fresh browser context every N files (memory leaks)
TIMEOUT_MS = 30_000

# Where the Twin ID lives inside each Excel file.
# The sheets have a header "Twins Patient ID" (column A); the ID is typed by hand in the cell
# directly BELOW that header. The script finds the header (any column, so it still works if the
# column moves) and reads the cell under it.
# Option: if the ID is always in one fixed cell, set TWIN_ID_CELL = "A2" (overrides the header search).
TWIN_ID_LABEL = re.compile(r"twins?\s*(patient\s*)?_?id", re.I)
TWIN_ID_CELL = ""
SCAN_ROWS, SCAN_COLS = 10, 30

# After a successful upload the site shows a green "Saved successfully" bar and a pop-up
# "Saved Investigation(s): 16704201" with an Ok button (seen on the real site).
SUCCESS_TEXT = r"Saved Investigation"
SAVED_ID_RE = re.compile(r"Saved Investigation\(s\):\s*([\d,\s]+)")

# Label of the submit button at the bottom of the "Upload Lab Report" pop-up
# (you scroll down in the pop-up to reach it). If your button has another label, change this.
SUBMIT_NAME = re.compile(r"upload|submit|save", re.I)

# The pop-up shows a preview of the sheet. Before submitting, check that the preview contains
# the patient ID, so the file the site is about to save really is this patient's.
PREVIEW_CHECK = True

# Selectors, taken from the screenshots of the real login page and Investigations page.
# Anything not marked "seen" is still a best guess and may need adjusting after the first real run.
SEL = {
    "logged_in_marker": "text=Investigations",                  # left menu item; seen after login
    "login_user": "input[placeholder='username']",              # seen: login box with 'username' hint
    "login_pass": "input[type=password]",                       # seen
    "patient_box": "header input, .v-toolbar input",            # fallback; main lookup is by label "Patient"
    "patient_option": ".v-list-item, [role=option]",            # dropdown entries "5668101 - Ranjana Gupta (AGE: 60)"
    "investigations_menu": "text=Investigations",
    "upload_button": "button:has-text('Upload Lab Report')",    # seen: top right of Investigations
    "file_input": "input[type=file]",                           # seen: native "Choose File" in the pop-up
}

# Buttons are found by role + visible name, which also works when the text is wrapped in an inner
# element (the case on the real site). The Google button ("Sign in with ...") does not match
# because the whole name must be just "sign in".
LOGIN_BUTTON_NAME = re.compile(r"^\s*sign\s*in\s*$", re.I)   # seen: pink SIGN IN button
OK_BUTTON_NAME = re.compile(r"^\s*ok\s*$", re.I)             # seen: Ok in the pop-up after saving


# ----------------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------------
def log_row(pid, filename, status, detail=""):
    is_new = not LOG_CSV.exists()
    with LOG_CSV.open("a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if is_new:
            w.writerow(["timestamp", "patient_id", "file", "status", "detail"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), pid, filename, status, detail])


def move_file(path, dest_dir):
    dest_dir.mkdir(exist_ok=True)
    dest = dest_dir / path.name
    if dest.exists():
        dest = dest_dir / f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.move(str(path), str(dest))


def _digits(value):
    """Return the ID digits from a cell value (handles 15545, 15545.0, '15545')."""
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    m = re.search(r"\d+", str(value))
    return m.group(0) if m else None


def read_twin_id(path):
    """Read the Twin ID from inside the Excel file. Raises ValueError if not found."""
    wb = load_workbook(path, data_only=True)
    header_seen = False
    try:
        for ws in wb.worksheets:
            if TWIN_ID_CELL:
                found = _digits(ws[TWIN_ID_CELL].value)
                if found:
                    return found
                continue
            for row in ws.iter_rows(min_row=1, max_row=SCAN_ROWS, max_col=SCAN_COLS):
                for c in row:
                    if isinstance(c.value, str) and TWIN_ID_LABEL.search(c.value):
                        header_seen = True
                        # "Twins Patient ID: 15545" in one cell, otherwise the cell below the header
                        found = _digits(c.value.split(":", 1)[1]) if ":" in c.value else None
                        found = found or _digits(ws.cell(c.row + 1, c.column).value)
                        if found:
                            return found
    finally:
        wb.close()
    if header_seen:
        raise ValueError("'Twins Patient ID' header found but the cell below it is empty")
    raise ValueError("no 'Twins Patient ID' header found")


# File name format: "<patientID>-<first name>_<last name>.xlsx", e.g. 5668101-Ranjana_Gupta.xlsx
NAME_ID_RE = re.compile(r"^\s*(\d+)\s*-")


def resolve_patient_id(path):
    """Work out which patient a file belongs to.

    Returns (patient_id, sheet_has_id, source). The ID comes from the start of the file name
    (digits then a hyphen). If the sheet also holds an ID under "Twins Patient ID", the two must
    agree, otherwise the file is skipped: a wrong match would put a report on the wrong patient.
    A file whose name has no ID falls back to the ID inside the sheet.
    """
    m = NAME_ID_RE.match(path.name)
    name_id = m.group(1) if m else None
    sheet_id, sheet_err = None, ""
    try:
        sheet_id = read_twin_id(path)
    except Exception as e:  # noqa: BLE001
        sheet_err = str(e)

    if name_id and sheet_id:
        if name_id != sheet_id:
            raise ValueError(f"file name says patient {name_id} but the sheet says {sheet_id}")
        return name_id, True, "file name + sheet agree"
    if name_id:
        return name_id, False, "from file name (sheet has no ID)"
    if sheet_id:
        return sheet_id, True, "from sheet (file name has no ID)"
    raise ValueError(
        "no patient ID in the file name (expected '<id>-First_Last.xlsx') and none in the sheet"
        + (f" ({sheet_err})" if sheet_err else "")
    )


def login_if_needed(ctx, page, headed):
    page.goto(BASE_URL)
    try:
        page.locator(SEL["logged_in_marker"]).first.wait_for(timeout=10_000)
        return
    except PWTimeout:
        pass

    user, pw = os.getenv("CLINIC_USER"), os.getenv("CLINIC_PASS")
    if user and pw:
        page.fill(SEL["login_user"], user)
        page.fill(SEL["login_pass"], pw)
        try:
            page.get_by_role("button", name=LOGIN_BUTTON_NAME).first.click(timeout=8_000)
        except PWTimeout:
            page.locator(SEL["login_pass"]).press("Enter")  # fall back: submit the form with Enter
    elif headed:
        input("Log in in the browser window, then press Enter here... ")
    else:
        raise RuntimeError(
            "Not logged in. Set CLINIC_USER and CLINIC_PASS, or run once with --headed to log in by hand."
        )
    page.locator(SEL["logged_in_marker"]).first.wait_for(timeout=TIMEOUT_MS)
    ctx.storage_state(path=str(AUTH_FILE))


def new_session(browser, headed):
    kwargs = {"storage_state": str(AUTH_FILE)} if AUTH_FILE.exists() else {}
    kwargs["viewport"] = {"width": 1600, "height": 1000}  # wide enough that the left menu is not collapsed
    ctx = browser.new_context(**kwargs)
    ctx.set_default_timeout(TIMEOUT_MS)
    page = ctx.new_page()
    try:
        login_if_needed(ctx, page, headed)
    except Exception:
        try:  # leave a screenshot of the login screen for debugging (kept in the run-log download)
            SHOTS.mkdir(exist_ok=True)
            page.screenshot(path=str(SHOTS / f"login_failed_{datetime.now():%H%M%S}.png"))
        except Exception:  # noqa: BLE001
            pass
        raise
    return ctx, page


# ----------------------------------------------------------------------------
# THE ACTUAL UPLOAD FOR ONE PATIENT
# ----------------------------------------------------------------------------
CURRENT_STEP = ""
CLICK_TIMEOUT_MS = 10_000


def click_robust(locator):
    """Click normally; if Playwright's click checks time out, fire the click event directly.
    Vue/Vuetify widgets (dropdown items, menu entries) sometimes ignore or block a real mouse
    click (overlays, animations) even though their click handler works fine."""
    try:
        locator.click(timeout=CLICK_TIMEOUT_MS)
    except PWTimeout:
        print(f"    (normal click timed out at step '{CURRENT_STEP}', sending the click event directly)")
        locator.dispatch_event("click", timeout=CLICK_TIMEOUT_MS)


def step(name):
    """Remember which step we are on, so a failure can say exactly where it stopped."""
    global CURRENT_STEP
    CURRENT_STEP = name


class AfterSubmitError(Exception):
    """The submit button was clicked but no confirmation was seen. Do NOT retry: the report may
    already be saved, and a second attempt would create a duplicate. Check the clinic system."""


def _patient_box(page):
    box = page.get_by_label("Patient", exact=True)  # top box labelled "Patient"
    if box.count() == 0:
        box = page.locator(SEL["patient_box"]).first
    return box.first


def select_patient(page, pid):
    step("1 find and click the Patient box")
    box = _patient_box(page)
    try:
        box.click(timeout=CLICK_TIMEOUT_MS)
    except PWTimeout:
        box.focus()  # typing only needs focus
    step("2 type the patient ID")
    box.press("Control+A")
    box.type(pid, delay=80)
    step("3 press Enter to open the patient dashboard")
    page.wait_for_timeout(1500)  # let the server fill the suggestion list before Enter
    box.press("Enter")
    step("4 check the page moved to this patient's dashboard")

    # Safety check: never upload under the wrong patient. The page must leave the home screen
    # and show this patient's ID somewhere (header / patient banner).
    page.wait_for_url(lambda u: "/modules/home" not in u, timeout=TIMEOUT_MS)
    page.get_by_text(re.compile(rf"\b{pid}\b")).first.wait_for(state="attached", timeout=TIMEOUT_MS)


def upload_one(page, pid, path, sheet_has_id=True):
    """Returns (investigation IDs the clinic system reports, note), e.g. ('16704201', '')."""
    select_patient(page, pid)

    step("5 click Investigations in the left menu")
    click_robust(page.locator(SEL["investigations_menu"]).first)
    step("6 click the Upload Lab Report button")
    upload_btn = page.locator(SEL["upload_button"]).first
    upload_btn.wait_for(state="attached", timeout=TIMEOUT_MS)
    click_robust(upload_btn)

    # The "Upload Lab Report" pop-up has a native "Choose File" input. Attach the file to it.
    step("7 attach the Excel file in the pop-up")
    file_input = page.locator(SEL["file_input"]).first
    file_input.wait_for(state="attached")
    file_input.set_input_files(str(path.resolve()))

    # The pop-up previews the sheet; confirm the preview holds this patient's ID.
    step("8 wait for the sheet preview to show the patient ID")
    if PREVIEW_CHECK and sheet_has_id:  # when the sheet has no ID, the preview cannot show one
        page.locator("table").get_by_text(pid, exact=True).first.wait_for(timeout=15_000)

    # Submit button sits at the bottom of the pop-up (Playwright scrolls to it by itself).
    step("9 click the submit button at the bottom of the pop-up")
    submit = page.get_by_role("button", name=SUBMIT_NAME).last
    submit.wait_for(state="attached", timeout=TIMEOUT_MS)
    click_robust(submit)
    step("10 wait for the 'Saved Investigation' confirmation")

    # From here on a retry could create a duplicate, so failures raise AfterSubmitError.
    try:
        confirm = page.get_by_text(re.compile(SUCCESS_TEXT, re.I)).first
        confirm.wait_for(timeout=TIMEOUT_MS)
        text = confirm.inner_text()
    except PWTimeout as e:
        raise AfterSubmitError(
            "submitted but no 'Saved Investigation' confirmation appeared; check the clinic system "
            "before uploading this file again"
        ) from e

    m = SAVED_ID_RE.search(text)
    saved_ids = re.sub(r"\s+", "", m.group(1)).strip(",") if m else ""

    # Close the pop-up so the next file starts clean (failure here does not undo the save).
    try:
        page.get_by_role("button", name=OK_BUTTON_NAME).first.click(timeout=5_000)
    except Exception:  # noqa: BLE001
        try:
            page.get_by_role("button", name=OK_BUTTON_NAME).first.dispatch_event("click", timeout=3_000)
        except Exception:  # noqa: BLE001
            pass

    # Second confirmation: the new Investigation should now be a row in the list. This is only
    # a note (the report is already saved, so it must never trigger a retry).
    note = ""
    first_id = saved_ids.split(",")[0] if saved_ids else ""
    if first_id:
        try:
            page.get_by_text(first_id, exact=True).first.wait_for(timeout=8_000)
        except Exception:  # noqa: BLE001
            note = "WARNING: saved, but the new row was not seen in the Investigations list; check it"
    return saved_ids, note


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def move_in_drive(drive, file_id, subfolder):
    """Move the Drive copy to processed/error. Returns False if it could not be moved."""
    if not file_id:
        return True
    err = None
    for _ in range(3):
        try:
            drive.move(file_id, subfolder)
            return True
        except Exception as e:  # noqa: BLE001
            err = e
            time.sleep(2)
    print(f"WARNING: could not move file to '{subfolder}' in Drive: {err}")
    return False


def main():
    global INBOX, PROCESSED, FAILED, SHOTS, LOG_CSV, AUTH_FILE
    ap = argparse.ArgumentParser()
    ap.add_argument("--reports-dir", default="reports",
                    help="folder with the Excel files; error/, screenshots/, "
                         "results.csv and auth.json are created next to it")
    ap.add_argument("--headed", action="store_true", help="show the browser (use for testing)")
    ap.add_argument("--drive", action="store_true",
                    help="download Excel files from the Google Drive folder and move them to "
                         "processed/error there (used by GitHub Actions)")
    ap.add_argument("--dry-run", action="store_true", help="list files, do nothing")
    ap.add_argument("--keep-processed", action="store_true",
                    help="keep uploaded files in processed/ instead of deleting them")
    ap.add_argument("--limit", type=int, default=0, help="process at most N files")
    args = ap.parse_args()

    INBOX = Path(args.reports_dir)
    work = INBOX.resolve().parent
    PROCESSED, FAILED, SHOTS = work / "processed", work / "error", work / "screenshots"
    LOG_CSV, AUTH_FILE = work / "results.csv", work / "auth.json"

    for d in (INBOX, SHOTS):
        d.mkdir(exist_ok=True)
    if args.keep_processed:
        PROCESSED.mkdir(exist_ok=True)
    # error/ is created only when a file actually has to go there (move_file does it)

    drive, drive_ids = None, {}
    if args.drive:
        from drive_sync import DriveSync
        drive = DriveSync.from_env()
        drive_ids = drive.download_all(INBOX)
        print(f"Downloaded {len(drive_ids)} Excel file(s) from Drive")

    jobs = []
    for f in sorted(INBOX.glob("*.xlsx")):
        if f.name.startswith("~$"):  # Excel temp/lock file
            continue
        try:
            pid, sheet_has_id, source = resolve_patient_id(f)
        except Exception as e:  # noqa: BLE001
            log_row("", f.name, "SKIPPED", f"patient ID problem: {e}")
            print(f"SKIPPED {f.name}: {e}")
            if not args.dry_run:  # an error: set the file aside in error/ so it is not retried blindly
                if drive:
                    move_in_drive(drive, drive_ids.get(f.name), "error")
                move_file(f, FAILED)
            continue
        jobs.append((f, pid, sheet_has_id, source))
    if args.limit:
        jobs = jobs[: args.limit]

    if not jobs:
        print("Nothing to upload.")
        return 0
    if args.dry_run:
        for f, pid, _has, source in jobs:
            print(f"would upload {f.name} -> patient {pid} ({source})")
        return 0

    ok = failed = drive_problems = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed, slow_mo=300 if args.headed else 0)
        try:
            ctx, page = new_session(browser, args.headed)
            for i, (path, pid, sheet_has_id, _source) in enumerate(jobs, 1):
                status, detail = "FAIL", ""
                for attempt in (1, 2):
                    try:
                        saved, note = upload_one(page, pid, path, sheet_has_id)
                        status = "OK"
                        detail = "; ".join(x for x in (f"investigation {saved}" if saved else "", note) if x)
                        break
                    except Exception as e:  # noqa: BLE001
                        lines = str(e).strip().splitlines() or [repr(e)]
                        detail = f"attempt {attempt} at step '{CURRENT_STEP}': {lines[0][:200]}"
                        try:
                            where = page.url
                        except Exception:  # noqa: BLE001
                            where = "?"
                        print(f"  ! {pid} {detail} (page: {where})")
                        for extra in lines[1:8]:  # Playwright's call log: names the exact element it waited for
                            print("      " + extra[:220])
                        try:
                            page.screenshot(path=str(SHOTS / f"fail_{pid}_{datetime.now():%H%M%S}.png"))
                            page.goto(BASE_URL)
                        except Exception:  # noqa: BLE001
                            pass
                        if isinstance(e, AfterSubmitError):
                            break  # never retry after submitting: it could duplicate the report

                log_row(pid, path.name, status, detail)
                if status == "OK":
                    ok += 1
                else:
                    failed += 1
                subfolder = "processed" if status == "OK" else "error"
                if drive and not move_in_drive(drive, drive_ids.get(path.name), subfolder):
                    drive_problems += 1  # if this happens after an OK upload, the file may be re-uploaded next run
                if status == "OK":
                    if args.keep_processed:
                        move_file(path, PROCESSED)
                    else:
                        path.unlink()  # uploaded: remove it from reports/
                else:
                    move_file(path, FAILED)
                print(f"[{i}/{len(jobs)}] {pid} {status} {detail}")

                if i % RECYCLE_EVERY == 0 and i < len(jobs):
                    ctx.close()
                    ctx, page = new_session(browser, args.headed)
        finally:
            browser.close()

    print(f"Done: {ok} uploaded, {failed} failed. See {LOG_CSV}")
    return 1 if (failed or drive_problems) else 0


if __name__ == "__main__":
    sys.exit(main())
