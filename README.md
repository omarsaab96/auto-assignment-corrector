# Auto Assignment Corrector

Compares each student `.xlsx` file against the template solution workbook, highlights cells that differ, and saves the marked copies in a timestamped corrected folder.

## Setup

```powershell
python -m pip install -r requirements.txt
```

## Google Drive Cloud Run

Use this when the files are only in Google Drive cloud, not synced to a local `G:\My Drive` folder.

## Web Page

Start the local page:

```powershell
python web_app.py
```

Open:

```text
http://127.0.0.1:5000
```

Use the Drive browser in the page to open folders, select the students folder, and select the solution `.xlsx` file. Click `Start correcting` to scan submissions, upload marked copies, and upload `grading_summary.csv` into a timestamped Drive folder such as `corrected-20260613-103245`.

If the saved Google token has expired or been revoked, the page will trigger the normal Google sign-in flow again and save a fresh `token.json`.

Student submissions can be placed directly in the students folder or inside student subfolders. Subfolders are scanned recursively, so one student can have multiple `.xlsx` submissions. Corrected files keep the same student subfolder structure under the timestamped corrected folder.

## Render Deployment

The web app can run on Render with Google OAuth Web credentials.

1. Push this project to GitHub.

2. In Google Cloud Console:

- Enable the Google Drive API.
- Create OAuth credentials for a **Web application**.
- Add this authorized redirect URI after you create the Render service:

```text
https://YOUR_RENDER_SERVICE.onrender.com/oauth2callback
```

3. In Render:

- Create a new **Web Service** from the GitHub repo.
- Use this start command:

```bash
gunicorn web_app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 300
```

- Add environment variables:

```text
FLASK_SECRET_KEY=<any long random string>
GOOGLE_OAUTH_CLIENT_JSON=<full JSON contents from the Google Web OAuth client>
```

The included `render.yaml` sets the same build/start commands if you use Render Blueprint deploys.

Use one worker because correction progress is stored in memory. If you later scale beyond one worker, move `JOBS` into Redis or a database.

### One-Time Setup

1. Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

2. Create a Google Cloud OAuth client:

- Open Google Cloud Console.
- Create or select a project.
- Enable the Google Drive API.
- Configure the OAuth consent screen.
- Set publishing status to `Testing`.
- Add your Google account as a test user, for example `omar.saab.96@gmail.com`.
- Create OAuth credentials for a Desktop app.
- Download the JSON file and save it in this project as:

```text
credentials.json
```

If Google shows `Error 403: access_denied` and says the app has not completed verification, the signed-in account is not listed as a test user on the OAuth consent screen. Add the account, save, wait a minute, then run the command again.

3. Get the IDs from Google Drive URLs:

Students folder URL:

```text
https://drive.google.com/drive/folders/STUDENTS_FOLDER_ID
```

Template file URL:

```text
https://drive.google.com/file/d/TEMPLATE_FILE_ID/view
```

### Manual Cloud Run

```powershell
python grade_google_drive.py --students-folder-id "STUDENTS_FOLDER_ID" --template-file-id "TEMPLATE_FILE_ID"
```

The first run opens a browser so you can sign in and approve Drive access. After that, the saved `token.json` is reused.

The script creates a timestamped folder such as `corrected-20260613-103245` inside the Drive `students` folder, uploads each corrected workbook there, and uploads `grading_summary.csv`. Use `--corrected-folder-name` only when you want to force a specific output folder name.

## Local Folder Run

If the `students` folder and template file are both inside this project:

```powershell
python grade_assignments.py
```

If the students folder is in Google Drive and the template is somewhere else, pass both paths:

```powershell
python grade_assignments.py --students "G:\My Drive\students" --template "G:\My Drive\Jarir - 3 Statement - Assignment.Sol.xlsx"
```

If `--template` is not provided, the script searches for this file name:

```text
Jarir - 3 Statement - Assignment.Sol.xlsx
```

It checks these locations in order:

```text
./Jarir - 3 Statement - Assignment.Sol.xlsx
<students parent>\Jarir - 3 Statement - Assignment.Sol.xlsx
<students>\Jarir - 3 Statement - Assignment.Sol.xlsx
```

For example, if your Drive layout is:

```text
G:\My Drive\Jarir - 3 Statement - Assignment.Sol.xlsx
G:\My Drive\students\
```

you can run:

```powershell
python grade_assignments.py --students "G:\My Drive\students"
```

Outputs are written to:

```text
<students>\corrected
```

Each corrected workbook includes a `Correction Report` sheet, and the corrected output folder also gets a `grading_summary.csv`.

## Notes

- Yellow cells are different from the template.
- Red cells indicate extra student content where the template cell is blank.
- Temporary Excel lock files such as `~$file.xlsx` are ignored.
- Text comparison trims leading and trailing spaces by default. Use `--no-trim-text` to count those spaces as differences.
- Use `--ignore-case` if uppercase/lowercase differences should not count.
