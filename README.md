# Playwright Advantage Component

Python + Playwright script that adds **course components to FAME Advantage** from an Excel file.
It opens each course and grade book in Advantage (creating them if missing), creates missing components,
and adds the grade book rows (number, weight, min score, sequence, Required / Must Pass). Every saved row
is verified in Advantage.

> This is the **command-line script** version. The full web app with a UI (upload, preview, validation,
> live progress, run history) is in
> [advantage-setup-automation](https://github.com/ENGINEER-MUHAMMAD-SHAHZAIB/advantage-setup-automation).

## What it does

```
Configuration guide (CSV/XLSX) --convert_guide.py--> data/automation_ready.xlsx --add_components.py--> Advantage
```

* **`convert_guide.py`** turns the school's configuration guide into `automation_ready.xlsx`, with one row per component plus *Subjects* and *Needs Review* sheets.
* **`add_components.py`** reads the Excel and, for every row with `Run = Yes`:
  1. Logs in, opens the program version and goes to **Definition → ADD/VIEW COURSE**.
  2. Selects the course and grade book. If one is missing, it creates it.
  3. Creates components that aren't in the list yet (**New Component**).
  4. Adds the rows in batches, clicks **Save → Confirm**, reopens the grade book and checks every row is there.

## Safety

* **Test mode:** `DRY_RUN = True` fills the first row but never clicks Save.
* **Check-only mode:** `VERIFY_ONLY = True` compares Advantage with the Excel and adds nothing.
* **Live site:** `ALLOW_LIVE` must be `True` to save on the live site (`advantage.fameinc.com`).
* **No duplicates:** a done-cache remembers saved rows, and a row counts as saved only after it is seen in the grade book.
* **Login details:** read from `.env`. They are never written in the code.

## Setup (Windows)

Needs **Python 3.10+**.

```powershell
git clone https://github.com/ENGINEER-MUHAMMAD-SHAHZAIB/playwright-advantage-component.git
cd playwright-advantage-component
python -m venv .venv
.venv\Scripts\activate
pip install playwright pandas openpyxl python-dotenv
playwright install chromium
```

Create a file called **`.env`** in the project folder. It is ignored by git, so it is never committed.

```
SIS_URL=https://advantage.fameinc.com/login
SIS_USER=your.username
SIS_PASS=your-password
SIS_TENANT=royal
SIS_TENANT_OPTION=RoyalTreatmentBeautyAcademy
```

## Usage

```powershell
.venv\Scripts\activate

# 1. convert the guide (edit the SETTINGS at the top of the file first)
python convert_guide.py

# 2. add the components (edit the SETTINGS at the top of the file first)
python add_components.py
```

Stop a run with **Ctrl + C**. Run it again and it resumes: rows already saved are skipped.

### Main settings (top of `add_components.py`)

| Setting | Meaning |
|---------|---------|
| `EXCEL_FILE`, `SHEET` | Input file (`data/automation_ready.xlsx`, sheet `Components`) |
| `LIMIT` | Max new rows per run |
| `DRY_RUN` | `True` = fill 1 row, don't save |
| `VERIFY_ONLY` | `True` = only compare with Advantage |
| `ALLOW_LIVE` | Must be `True` to save on the live site |
| `PROGRAM_SEARCH_TEXT`, `PROGRAM_OPTION` | Which program version to open |
| `BATCH_SIZE` | Rows added before one Save |
| `CREATE_COURSE_IF_MISSING`, `CREATE_IF_MISSING` | Create missing course / grade book / component |
| `REDO_ROW_IDS` | Row IDs to add again (delete them in Advantage first) |
| `SLOW_MO`, `HEADLESS`, `BROWSER_CHANNEL` | Speed and browser |

More detail is in the `help` folder: `commands.txt`, `NEW_PROGRAM_GUIDE.txt` and the PDF guides.

## Project structure

```
add_components.py     Main automation script
convert_guide.py      Configuration guide -> automation_ready.xlsx
help/                 Run guide, setup guide, command list, new-program guide
.env                  Login details (NOT in git)
data/, guides/, logs/ Excel files, school guides, run logs and screenshots (NOT in git)
```

## Not in this repository (on purpose)

`.env`, `recorded.txt` (Playwright recordings contain the typed password), `data/`, `guides/`, `logs/`,
screenshots and `.venv/` are listed in `.gitignore`. Never commit school data or passwords.
