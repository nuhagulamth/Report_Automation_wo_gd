# Run on GitHub Actions (7:00, 9:00, 11:00 AM IST)

Before you start: fix the login and patient-box selectors in `upload_lab_reports.py` on your laptop (README, step 2).
A headless GitHub run cannot show you where a wrong selector fails, it only saves a screenshot.

## A. Google service account (a robot login that can read your Drive folder)

1. Go to console.cloud.google.com and create a project (name: `lab-upload`).
2. APIs & Services > Library > search "Google Drive API" > Enable.
3. IAM & Admin > Service Accounts > Create service account. Name: `lab-upload-bot`. Skip the role steps. Click Done.
4. Open the new account > Keys tab > Add key > Create new key > JSON > Create. A `.json` file downloads. This file is a password: never put it in the repo.
5. Copy the account's email (looks like `lab-upload-bot@lab-upload-123456.iam.gserviceaccount.com`).
6. In Google Drive, right-click the `Reports Excel` folder > Share > paste that email > role **Editor** > untick "Notify people" > Share.
7. Open the `Reports Excel` folder. The URL ends with `/folders/XXXXXXXX`. Copy the `XXXXXXXX` part: that is the folder ID.

If step 4 is blocked ("Service account key creation is disabled"), your Google organisation has a policy against it. Ask your Workspace admin to allow it, or tell me and we use another route.

## B. GitHub secrets

Repo > Settings > Secrets and variables > Actions > New repository secret. Add four:

| Name | Value |
|---|---|
| `CLINIC_USER` | clinic login username |
| `CLINIC_PASS` | clinic login password |
| `GDRIVE_SA_JSON` | the whole contents of the downloaded `.json` file |
| `DRIVE_FOLDER_ID` | the folder ID from step A7 |

## C. Add the files to the repo

Upload or replace: `upload_lab_reports.py`, `drive_sync.py`, `requirements.txt`.
Create `.github/workflows/upload.yml`: Add file > Create new file, type the name `.github/workflows/upload.yml` (typing the slashes creates the folders), paste the contents, commit.

## D. Test

1. Repo > Actions tab > "Upload lab reports" > Run workflow. Leave "Dry run" ticked, limit `0`. Run.
2. Open the run > "Upload reports" step. You should see `Downloaded N Excel file(s) from Drive` and one `would upload ... -> patient ID` line per file.
3. Run again with Dry run unticked and limit `1`. Check the patient's Investigations in the clinic system.
4. If that works, run with limit `0` once, then leave the schedule on.

## How it behaves

- After a confirmed upload the file moves to a `processed` subfolder of `Reports Excel` in Drive. After two failed attempts it moves to `failed`. Files with no readable Twin ID stay where they are and are logged as SKIPPED.
- If any file fails, the run is marked failed and GitHub emails you.
- The log (`results.csv`) and failure screenshots are attached to each run for 3 days. Screenshots can show patient names, so only give repo access to people who may see them.
- Scheduled runs can start a few minutes late. GitHub pauses scheduled runs if the repo has no activity for 60 days; re-enable them in the Actions tab.
