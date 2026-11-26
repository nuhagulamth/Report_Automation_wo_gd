# Lab report upload automation

Uploads the Excel files in `reports/` into the clinic system (ICAPP): select the patient, open Investigations,
click Upload Lab Report, attach the file, submit, and wait for "Saved Investigation(s): <id>".

## File names

Name each file `<patientID>-<first name>_<last name>.xlsx`, for example `5668101-Ranjana_Gupta.xlsx`.
The patient ID is the number before the first hyphen. If the sheet also has an ID under "Twins Patient ID"
(cell A2), it must match the file name, otherwise the file is not uploaded. A file whose name has no ID
falls back to the ID in the sheet.

## What happens to each file

- Upload confirmed: the file is **removed** from `reports/`.
- Anything that goes wrong (no or mismatched ID, patient not found, upload not confirmed): the file is
  **moved to `error/`**, and the reason is written to `results.csv`.
  After fixing it, move it back into `reports/` to retry.
- A file whose submit button was clicked but whose confirmation never appeared goes to `error/` too and is
  never retried automatically. Check the patient's Investigations list before uploading it again.

## Run on GitHub (Actions tab > Upload lab reports > Run workflow)

- Source `repo`, Dry run ticked: lists `would upload <file> -> patient <id>`, uploads nothing.
- Source `repo`, Dry run unticked, limit `1`: uploads one file for real. Limit `0` = all files.
- After a real run the workflow commits the removals and moves back to the repo, so `reports/` and `error/`
  always show the current state. Dry runs change nothing.
- `results.csv` and failure screenshots are in the run's `run-log` download (kept 30 days). It is the only
  record of what was uploaded, because uploaded files are removed.
- Secrets needed: `CLINIC_USER`, `CLINIC_PASS`. (`GDRIVE_SA_JSON` and `DRIVE_FOLDER_ID` only for Drive mode.)
- Do not start a second run while one is going; the workflow queues it.

## Run on your own computer

    pip install -r requirements.txt
    playwright install chromium
    python upload_lab_reports.py --dry-run
    python upload_lab_reports.py --headed --limit 1     # watch one upload; log in by hand if no .env
    python upload_lab_reports.py --keep-processed       # keep uploaded files in processed/ instead of deleting

Put the login in a `.env` file (copy `.env.example`). Never commit `.env` or `auth.json`.

## Log lines

`[1/3] 5668101 OK investigation 16704201` is a confirmed save. A failure prints the step it stopped at
(for example `3 pick the patient from the dropdown list`) and the element Playwright was waiting for.

## Schedule (not switched on)

The workflow has no schedule while files live in the repo. Switch one on only after the Drive mode works.
7, 9 and 11 AM IST is the cron `30 1,3,5 * * *` (GitHub uses UTC).
