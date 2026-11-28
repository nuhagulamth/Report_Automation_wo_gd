"""
Upload Excel lab reports into the clinic system, one patient per file.

Folder layout (next to this script):
    reports/     put the Excel files here, named <patientID>-First_Last.xlsx
                 (a file is REMOVED from here as soon as its upload is confirmed)
    error/       every file that could not be uploaded is moved here, with the reason in results.csv
                 (move it back into reports/ after fixing it, to retry)
    pdfs/        the PDF of each report (same name as its Excel file). After the Excel upload succeeds, the PDF is
                 attached to the patient's Attachments (category "Lab Report") and then removed from here.
                 If the Excel upload fails, or the attachment fails, the PDF is only removed from here (its copy in
                 Drive stays in "report pdf upload in process"); the reason is written to error/reasons.csv.
    pdfs_retry/  PDFs whose attachment failed earlier and are tried again (the Apps Script puts them here);
                 they are attached without an Excel upload, because their Excel was already uploaded
    status/      attached.csv: PDFs that were attached (the Apps Script then clears them from Drive)
                 attach_failed.csv: PDFs that could not be attached (the Apps Script decides: try again, or report)
    screenshots/ one screenshot per failed attempt (in the run-log download only)
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
PDF_DIR = Path("pdfs")             # PDFs waiting to be attached to the patient's Attachments
PDF_RETRY_DIR = Path("pdfs_retry") # PDFs whose attachment is tried again (Excel was uploaded earlier)
STATUS_DIR = Path("status")        # attached.csv: PDFs attached successfully
ATTACH_CATEGORY = "Lab Report"     # category chosen in the Attachments upload pop-up
LOG_CSV = Path("results.csv")
AUTH_FILE = Path("auth.json")      # saved login session; treat like a password
ATTACH_USER = os.getenv("ATTACH_USER")   # optional separate login for the Attachments step;
ATTACH_PASS = os.getenv("ATTACH_PASS")   # when empty, the normal CLINIC_USER login is used
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


def login_if_needed(ctx, page, headed, user=None, pw=None, auth_file=None):
    page.goto(BASE_URL)
    try:
        page.locator(SEL["logged_in_marker"]).first.wait_for(timeout=10_000)
        return
    except PWTimeout:
        pass

    user, pw = user or os.getenv("CLINIC_USER"), pw or os.getenv("CLINIC_PASS")
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
    ctx.storage_state(path=str(auth_file or AUTH_FILE))


def new_session(browser, headed, user=None, pw=None, auth_file=None):
    auth = auth_file or AUTH_FILE
    kwargs = {"storage_state": str(auth)} if Path(auth).exists() else {}
    kwargs["viewport"] = {"width": 1600, "height": 1000}  # wide enough that the left menu is not collapsed
    ctx = browser.new_context(**kwargs)
    ctx.set_default_timeout(TIMEOUT_MS)
    page = ctx.new_page()
    try:
        login_if_needed(ctx, page, headed, user, pw, auth_file)
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


class ReportRejectedError(Exception):
    """The site's own preview says the sheet has errors (Upload button stays disabled). Retrying
    cannot fix this: the file's values must be corrected or checked by a person."""


class AfterSubmitError(Exception):
    """The submit button was clicked but no confirmation was seen. Do NOT retry: the report may
    already be saved, and a second attempt would create a duplicate. Check the clinic system."""


def _patient_box(page):
    box = page.get_by_label("Patient", exact=True)  # top box labelled "Patient"
    if box.count() == 0:
        box = page.locator(SEL["patient_box"]).first
    return box.first


def select_patient(page, pid):
    """Type the ID, click the patient in the suggestion list, and confirm the box shows '<ID> - <Name>'.
    If the selection does not take, clear the box and try again (up to 3 times)."""
    shown = re.compile(rf"^\s*{pid}\s*-\s*\S")
    box = _patient_box(page)
    for attempt in (1, 2, 3):
        step(f"1 find and click the Patient box (try {attempt})")
        try:
            box.click(timeout=CLICK_TIMEOUT_MS)
        except PWTimeout:
            box.focus()  # typing only needs focus
        step(f"2 type the patient ID (try {attempt})")
        for _try in range(3):  # the first character is sometimes lost; never press Enter on a partial ID
            box.press("Control+A")
            box.press("Delete")
            box.type(pid, delay=80)
            if box.input_value().strip() == pid:
                break
        else:
            raise RuntimeError(f"could not type the full patient ID {pid} (box shows '{box.input_value()}')")

        # The site lists the matching patient under the box ("5967001 - Deepa Negi (AGE: 30)"). A person clicks
        # that entry; pressing Enter alone does not select it. So: wait for the entry, then click it.
        step(f"3 wait for the patient entry in the suggestion list, then click it (try {attempt})")
        entry_re = re.compile(rf"^\s*{pid}\s*-")
        entry = page.locator(SEL["patient_option"]).filter(has_text=entry_re)
        if entry.count() == 0:
            entry = page.get_by_role("option", name=entry_re)
        if entry.count() == 0:
            entry = page.get_by_text(entry_re)
        try:  # the list is filled from the server, so wait until the entry shows
            entry.first.wait_for(state="visible", timeout=15_000)
            click_robust(entry.first)
        except PWTimeout:
            print("    (patient entry not seen in the list; trying the keyboard: arrow down + Enter)")
            box.press("ArrowDown")
            box.press("Enter")
        page.wait_for_timeout(500)

        step(f"4 check the Patient box now shows '<ID> - <Name>' (try {attempt})")
        # Safety check: never upload under the wrong patient. The URL does not change (single-page app);
        # once the patient is selected the box shows "1765201 - Manoj RAGHAVAN (AGE: 59)". Only the typed ID
        # (no " - Name") does not match, so this passes only after a real selection of THIS patient.
        try:
            expect(box).to_have_value(shown, timeout=8_000)
            return
        except AssertionError:
            if attempt == 3:
                raise


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

    # The site validates the sheet in the preview ("Found 1 records with errors" + an "Error: ..." row,
    # Upload button disabled). Report that clearly instead of clicking a disabled button.
    err_banner = page.get_by_text(re.compile(r"Found\s+\d+\s+records?\s+with\s+errors", re.I))
    if err_banner.count() > 0:
        found = page.get_by_text(re.compile(r"^\s*Error:")).all_inner_texts()   # under the table, not always in a td
        msgs = []
        for t in found:
            t = re.sub(r"\s+", " ", t).strip()
            if t and t not in msgs:
                msgs.append(t)
        raise ReportRejectedError("the site rejected the sheet in the preview: "
                                  + (" | ".join(msgs) if msgs else err_banner.first.inner_text()))

    # Submit button sits at the bottom of the pop-up (Playwright scrolls to it by itself).
    step("9 click the submit button at the bottom of the pop-up")
    # Look only inside the open pop-up and match the button whose whole name is "Upload" (the page
    # behind it also has an "Upload Lab Report" button, which must not be clicked by mistake).
    dialog = page.locator(".v-dialog--active, [role=dialog]").last
    submit = dialog.get_by_role("button", name=re.compile(r"^\s*(upload|submit|save)\s*$", re.I)).last
    submit.wait_for(state="visible", timeout=TIMEOUT_MS)
    expect(submit).to_be_enabled(timeout=TIMEOUT_MS)
    submit.scroll_into_view_if_needed()
    for click_no in (1, 2, 3):
        # Real mouse click first; on later tries fire the click event directly, then press Enter on
        # the focused button. A click only counts once the pop-up reacts (confirmation shows, the
        # pop-up closes, or the button turns disabled/busy).
        if click_no == 1:
            submit.click(timeout=CLICK_TIMEOUT_MS)
        elif click_no == 2:
            print("    (Upload click had no effect, sending the click event directly)")
            submit.dispatch_event("click")
        else:
            print("    (Upload still had no effect, pressing Enter on the button)")
            submit.focus()
            submit.press("Enter")
        try:
            page.wait_for_function(
                """([c, d]) => {
                    const t = document.body.innerText || '';
                    if (new RegExp(c, 'i').test(t)) return true;
                    const dlg = document.querySelector(d);
                    if (!dlg || dlg.offsetParent === null) return true;
                    const b = [...dlg.querySelectorAll('button')].find(
                        x => /^\\s*(upload|submit|save)\\s*$/i.test(x.innerText));
                    return !b || b.disabled || b.classList.contains('v-btn--loading')
                        || b.classList.contains('v-btn--disabled');
                }""",
                arg=[SUCCESS_TEXT, ".v-dialog--active, [role=dialog]"], timeout=6_000)
            break  # the pop-up reacted: the click took
        except PWTimeout:
            continue  # nothing happened yet: try the next way of clicking
    step("10 wait for the 'Saved Investigation' confirmation")

    # From here on a retry could create a duplicate, so failures raise AfterSubmitError.
    try:
        confirm = page.get_by_text(re.compile(SUCCESS_TEXT, re.I)).first
        confirm.wait_for(timeout=TIMEOUT_MS)
        text = confirm.inner_text()
    except PWTimeout as e:
        try:  # say what the screen shows, so the reason is visible in results.csv
            seen = " | ".join(t.strip().replace("\n", " ") for t in
                              page.locator(".v-snack, .v-snackbar, [role=alert], .v-dialog--active").all_inner_texts())
        except Exception:  # noqa: BLE001
            seen = ""
        raise AfterSubmitError(
            "submitted but no 'Saved Investigation' confirmation appeared; check the clinic system "
            "before uploading this file again. Screen showed: " + (seen[:300] or "(nothing)")
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
# ----------------------------------------------------------------------------
# ATTACH THE PDF TO THE PATIENT'S "ATTACHMENTS" (after the Excel upload succeeded)
# ----------------------------------------------------------------------------
def find_pdf(excel_path):
    """pdfs/<same name as the Excel file>.pdf, or None."""
    if not PDF_DIR.is_dir():
        return None
    want = excel_path.stem.lower() + ".pdf"
    for p in PDF_DIR.iterdir():
        if p.is_file() and p.name.lower() == want:
            return p
    return None


def append_csv(path, header, row):
    path.parent.mkdir(parents=True, exist_ok=True)
    new_file = not path.exists()
    with path.open("a", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        if new_file:
            w.writerow(header)
        w.writerow(row)


def _attachments_card(page):
    """The 'Attachments' box on the patient profile (title + Upload button)."""
    upload_name = re.compile(r"^\s*upload\s*$", re.I)
    cards = page.locator(".v-card").filter(has_text="Attachments").filter(
        has=page.get_by_role("button", name=upload_name))
    return cards.last if cards.count() else page.locator("body")  # innermost box; else the whole page


def attach_pdf(page, pid, pdf_path):
    """Open the patient's profile, Attachments > Upload > category 'Lab Report' > choose the PDF > Save,
    then confirm the file name shows up in the Attachments list. Returns a short note."""
    name = pdf_path.name

    step("A1 open the patient's profile")
    page.goto(f"{BASE_URL}/ui/clinic/modules/patients/{pid}")
    heading = page.get_by_text(re.compile(rf"#\s*{pid}\b")).first      # "Hameeda Ayub #15545"
    try:
        heading.wait_for(state="visible", timeout=20_000)
    except PWTimeout:
        # fallback, the way a person does it: pick the patient in the top box, click Patients in the menu
        select_patient(page, pid)
        click_robust(page.get_by_text("Patients", exact=True).first)
        heading.wait_for(state="visible", timeout=TIMEOUT_MS)           # never attach to a different patient

    step("A2 scroll to the Attachments section")
    card = _attachments_card(page)
    page.get_by_text("Attachments", exact=True).first.wait_for(state="attached", timeout=TIMEOUT_MS)
    card.scroll_into_view_if_needed()
    try:  # let the list finish loading, so the "before" count below is right
        card.locator("tbody tr").first.wait_for(state="attached", timeout=15_000)
        page.wait_for_load_state("networkidle", timeout=10_000)
    except PWTimeout:
        pass
    before = page.get_by_text(name, exact=True).count()

    step("A3 click Upload in the Attachments section")
    upload_btn = card.get_by_role("button", name=re.compile(r"^\s*upload\s*$", re.I)).first
    click_robust(upload_btn)
    dialog = page.locator(".v-dialog--active, [role=dialog]").last
    dialog.get_by_text("Upload patient medical record file").wait_for(state="visible", timeout=TIMEOUT_MS)

    step(f"A4 choose the category '{ATTACH_CATEGORY}'")
    click_robust(dialog.locator(".v-select, [role=combobox]").first)
    page.get_by_text("Latest Prescription", exact=True).first.wait_for(state="visible", timeout=TIMEOUT_MS)  # list is open
    wanted = re.compile(rf"^\s*{re.escape(ATTACH_CATEGORY)}\s*$")
    option = page.get_by_role("option", name=wanted)
    if option.count() == 0:
        option = page.locator(".v-list-item, .v-list__tile, [role=listitem]").filter(has_text=wanted)
    if option.count() == 0:
        option = page.get_by_text(wanted)       # the open list sits last in the page
    click_robust(option.last)
    expect(dialog).to_contain_text(ATTACH_CATEGORY, timeout=10_000)   # the field now shows the chosen category

    step("A5 attach the PDF")
    file_input = dialog.locator("input[type=file]")
    if file_input.count() == 0:
        file_input = page.locator("input[type=file]")
    file_input = file_input.last
    file_input.wait_for(state="attached", timeout=TIMEOUT_MS)
    file_input.set_input_files(str(pdf_path.resolve()))
    try:
        expect(dialog).to_contain_text(name, timeout=10_000)      # the file name shows in the pop-up
    except AssertionError:
        print(f"    (the pop-up does not show '{name}' yet; continuing, the list check below decides)")

    step("A6 click Save")
    save = dialog.get_by_role("button", name=re.compile(r"^\s*save\s*$", re.I))
    expect(save).to_be_enabled(timeout=10_000)
    click_robust(save)

    # From here on a retry could attach the PDF twice, so any problem is reported, never retried.
    try:
        step("A7 wait for the pop-up to close")
        dialog.wait_for(state="hidden", timeout=TIMEOUT_MS)
        step("A8 wait for the file name to appear in the Attachments list")
        for _ in range(40):                                        # up to ~20 s
            if page.get_by_text(name, exact=True).count() > before:
                return "attached"
            page.wait_for_timeout(500)
        raise AfterSubmitError(f"saved, but '{name}' was not seen in the Attachments list; check the patient's attachments")
    except AfterSubmitError:
        raise
    except Exception as e:  # noqa: BLE001
        raise AfterSubmitError(f"saved, but could not confirm the attachment ({type(e).__name__}); check the patient's attachments") from e


def attach_with_retries(page, pid, pdf_path):
    """Returns (ok, detail, reason, safe_to_retry). Retries once within the run, but only if nothing was
    submitted yet. safe_to_retry is False when Save was clicked (the PDF may be attached already)."""
    detail = reason = ""
    safe = True
    for attempt in (1, 2):
        try:
            return True, attach_pdf(page, pid, pdf_path), "", True
        except Exception as e:  # noqa: BLE001
            lines = str(e).strip().splitlines() or [repr(e)]
            detail = f"attach attempt {attempt} at step '{CURRENT_STEP}': {lines[0][:400]}"
            reason = f"PDF not attached. {lines[0][:150]}" if isinstance(e, AfterSubmitError) \
                else f"PDF not attached. Stopped at step {CURRENT_STEP}: {lines[0][:150]}"
            print(f"  ! {pid} {detail}")
            try:
                SHOTS.mkdir(exist_ok=True)
                page.screenshot(path=str(SHOTS / f"attach_fail_{pid}_{datetime.now():%H%M%S}.png"))
            except Exception:  # noqa: BLE001
                pass
            if isinstance(e, AfterSubmitError):
                safe = False
                break
    return False, detail, reason, safe


def attach_and_record(page, pid, pdf):
    """Attach one PDF and record the outcome for the Apps Script, then remove the PDF from the repo folder.
    Success -> status/attached.csv. Failure -> status/attach_failed.csv (with a 'safe to retry' flag)."""
    ok, detail, reason, safe = attach_with_retries(page, pid, pdf)
    now = datetime.now().isoformat(timespec="seconds")
    if ok:
        log_row(pid, pdf.name, "ATTACHED", detail)
        append_csv(STATUS_DIR / "attached.csv", ["time", "pdf"], [now, pdf.name])
    else:
        log_row(pid, pdf.name, "ATTACH_FAIL", detail)
        append_csv(STATUS_DIR / "attach_failed.csv", ["time", "pdf", "safe_to_retry", "reason"],
                   [now, pdf.name, "yes" if safe else "no", reason])
    pdf.unlink()   # the Drive copy stays in "report pdf upload in process" until it is attached (then the Apps Script clears it)
    print(f"    PDF {pdf.name}: {'attached' if ok else 'NOT attached' + ('' if safe else ' (check the patient)')}")
    try:
        page.goto(BASE_URL)
    except Exception:  # noqa: BLE001
        pass
    return ok


def plain_reason(exc, step_name, pid):
    """A short readable reason for reasons.csv (no Playwright jargon)."""
    n = re.match(r"\s*(\d+)", step_name or "")
    n = int(n.group(1)) if n else 0
    if isinstance(exc, AfterSubmitError) or n >= 9:
        return (f"Upload was clicked but ICAPP did not confirm it. Check patient {pid}'s Investigations list "
                "before uploading again (it may already be saved).")
    if n <= 4:
        return (f"Patient {pid} could not be selected: the ID was typed in the Patient box but ICAPP did not "
                "open that patient. Check the ID exists in ICAPP and matches the file name.")
    if n <= 6:
        return "Could not open Investigations > Upload Lab Report for this patient."
    return "The Excel file could not be attached or the preview did not load."


def keep_screenshot(shot, filename):
    """Copy the failure screenshot to error/<file name>.png, next to the failed Excel file."""
    try:
        if shot and Path(shot).exists():
            FAILED.mkdir(exist_ok=True)
            shutil.copyfile(str(shot), str(FAILED / f"{Path(filename).stem}.png"))
    except Exception as e:  # noqa: BLE001 - never stop the run for a screenshot
        print(f"  (could not keep the screenshot for {filename}: {e})")


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
    global PDF_DIR, STATUS_DIR, PDF_RETRY_DIR
    PDF_DIR, STATUS_DIR, PDF_RETRY_DIR = work / "pdfs", work / "status", work / "pdfs_retry"
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
                append_csv(FAILED / "reasons.csv", ["time", "file", "reason"],
                           [datetime.now().isoformat(timespec="seconds"), f.name, f"Patient ID problem: {e}"])
                stray = find_pdf(f)
                if stray:
                    stray.unlink()   # not uploaded: the PDF stays only in Drive ("report pdf upload in process")
            continue
        jobs.append((f, pid, sheet_has_id, source))
    if args.limit:
        jobs = jobs[: args.limit]

    # PDFs whose attachment failed earlier (Excel already uploaded): tried again, no Excel needed
    retry_pdfs = sorted(PDF_RETRY_DIR.glob("*.pdf")) if PDF_RETRY_DIR.is_dir() else []

    if not jobs and not retry_pdfs:
        print("Nothing to upload.")
        return 0
    if args.dry_run:
        for f, pid, _has, source in jobs:
            print(f"would upload {f.name} -> patient {pid} ({source})")
        for p_ in retry_pdfs:
            print(f"would try to attach {p_.name} again")
        return 0

    ok = failed = drive_problems = attached = attach_failed = 0
    attach_ctx = attach_page = None   # separate login for attachments, only if ATTACH_USER / ATTACH_PASS are set
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headed, slow_mo=300 if args.headed else 0)
        try:
            ctx, page = new_session(browser, args.headed)
            for i, (path, pid, sheet_has_id, _source) in enumerate(jobs, 1):
                status, detail, last_reason = "FAIL", "", ""
                err_shot = None
                for attempt in (1, 2):
                    try:
                        saved, note = upload_one(page, pid, path, sheet_has_id)
                        status = "OK"
                        detail = "; ".join(x for x in (f"investigation {saved}" if saved else "", note) if x)
                        break
                    except Exception as e:  # noqa: BLE001
                        lines = str(e).strip().splitlines() or [repr(e)]
                        detail = f"attempt {attempt} at step '{CURRENT_STEP}': {lines[0][:600]}"
                        try:
                            where = page.url
                        except Exception:  # noqa: BLE001
                            where = "?"
                        print(f"  ! {pid} {detail} (page: {where})")
                        for extra in lines[1:8]:  # Playwright's call log: names the exact element it waited for
                            print("      " + extra[:220])
                        try:
                            err_shot = SHOTS / f"fail_{pid}_{datetime.now():%H%M%S}.png"
                            page.screenshot(path=str(err_shot))
                            if isinstance(e, ReportRejectedError):   # the site's own words: just the reason(s)
                                last_reason = lines[0].split(": Error:", 1)[-1].replace("the site rejected the sheet in the preview:", "")
                                last_reason = re.sub(r"^\s*Error:\s*", "", last_reason).strip()
                            else:
                                last_reason = plain_reason(e, CURRENT_STEP, pid)
                            page.goto(BASE_URL)
                        except Exception:  # noqa: BLE001
                            pass
                        if isinstance(e, (AfterSubmitError, ReportRejectedError)):
                            break  # never retry after submitting: it could duplicate the report

                log_row(pid, path.name, status, detail)
                if status != "OK":   # one short line per failed file: error/reasons.csv (moved to Drive by the Apps Script)
                    append_csv(FAILED / "reasons.csv", ["time", "file", "reason"],
                               [datetime.now().isoformat(timespec="seconds"), path.name, last_reason or detail[:200]])
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
                    keep_screenshot(err_shot, path.name)   # error/<name>.png: what ICAPP showed
                print(f"[{i}/{len(jobs)}] {pid} {status} {detail}")

                # ---- the PDF of this report (same name as the Excel file) ----
                pdf = find_pdf(path)
                if pdf is not None and status != "OK":
                    pdf.unlink()                      # the Excel was not uploaded: nothing is attached; the PDF stays only in Drive
                    print(f"    PDF {pdf.name} not attached (the Excel upload failed); it stays in Drive")
                elif pdf is not None:
                    if ATTACH_USER and ATTACH_PASS and attach_page is None:
                        attach_ctx, attach_page = new_session(browser, args.headed, ATTACH_USER, ATTACH_PASS,
                                                              work / "auth_attach.json")
                    if attach_and_record(attach_page or page, pid, pdf):
                        attached += 1
                    else:
                        attach_failed += 1

                if i % RECYCLE_EVERY == 0 and i < len(jobs):
                    ctx.close()
                    ctx, page = new_session(browser, args.headed)

            # ---- PDFs to try again ----
            for rp in retry_pdfs:
                m = NAME_ID_RE.match(rp.name)
                if not m:
                    log_row("", rp.name, "ATTACH_FAIL", "no patient ID at the start of the PDF name")
                    append_csv(STATUS_DIR / "attach_failed.csv", ["time", "pdf", "safe_to_retry", "reason"],
                               [datetime.now().isoformat(timespec="seconds"), rp.name, "no",
                                "PDF not attached. No patient ID at the start of the file name"])
                    rp.unlink()
                    attach_failed += 1
                    continue
                if ATTACH_USER and ATTACH_PASS and attach_page is None:
                    attach_ctx, attach_page = new_session(browser, args.headed, ATTACH_USER, ATTACH_PASS,
                                                          work / "auth_attach.json")
                if attach_and_record(attach_page or page, m.group(1), rp):
                    attached += 1
                else:
                    attach_failed += 1
        finally:
            browser.close()

    print(f"Done: {ok} uploaded, {failed} failed, {attached} PDF(s) attached, {attach_failed} PDF(s) not attached. See {LOG_CSV}")
    return 1 if (failed or drive_problems or attach_failed) else 0


if __name__ == "__main__":
    sys.exit(main())
