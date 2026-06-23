# F.A.S.T. Rescue — Course Roster Builder

Automatically generates print-ready Word course rosters from an Excel export. Upload a `.xlsx` file and download filled `.docx` rosters — one per course — with participant names sorted, a Sr. No. column, correct template selection, and clean formatting throughout.

---

## Features

- Detects the course type from the course name and picks the right template automatically
- Fills all info fields: Course Name, Company Name, Location, Contact, Date, Time, Course Material, Instructor, Notes
- Cleans account-number prefixes/suffixes from company names (e.g. `519817 CUPE Ontario` → `CUPE Ontario`)
- Normalizes course names — strips trailing city names, always ends with `Training`
- Formats dates: single day (`May 6, 2026`), two days (`May 28 & 29, 2026`), multi-session (`May 2, 3 & 5, 2026`)
- Formats times as `8:30 am to 4:00 pm`
- Parses participant lists in any export format (per-line, multi-line with emails, bulleted/grouped)
- Strips emails, fixes `Last, First` name order, collapses double spaces
- Sorts participant names A–Z (optional)
- Inserts a **Sr. No.** column — numbered 1 through N including blank rows
- Sizes the participant table to `max(20, participants + 5)` rows — always at least 20, always 5 blank rows at the bottom
- Skips cancelled courses (opt-in flag to include them)
- Output filename: `Instructor - First date - Company.docx`

---

## Repository Structure

```
roster-app/
├── app.py                  # Streamlit web UI
├── build_rosters.py        # Core logic (also runs as a CLI)
├── requirements.txt        # Python dependencies
├── README.md
├── templates/              # Blank Word templates (do not rename)
│   ├── Blended_First_Aid_Template.docx
│   ├── In_Class_First_Aid_Template.docx
│   ├── Recert_First_Aid_Template.docx
│   ├── Working_At_Heights_Template.docx
│   └── General_Template.docx
├── assets/
│   └── logo.png            # F.A.S.T. Rescue logo (used in the web app)
└── .streamlit/
    └── config.toml         # Theme and server settings
```

---

## Template Selection

The script picks a template based on keywords in the course name:

| Course name contains | Template used |
|---|---|
| `blended` | Blended First Aid |
| `recert` | Recertification First Aid |
| `in-class` or `in class` | In-Class First Aid |
| `working at heights` | Working At Heights |
| Anything else | General |

Cancelled courses (name contains `cancel`) are skipped by default.

---

## Running Locally

### Requirements

- Python 3.8+
- pip

### Install dependencies

```bash
pip install -r requirements.txt
```

### Run the web app

```bash
streamlit run app.py
```

Then open [http://localhost:8501](http://localhost:8501) in your browser.

### Run as a command-line script

```bash
python build_rosters.py EXPORT.xlsx --templates templates --out rosters
```

**Options:**

| Flag | Description |
|---|---|
| `--templates PATH` | Folder containing blank template `.docx` files (default: `.`) |
| `--out PATH` | Output folder for generated rosters (default: `rosters`) |
| `--keep-order` | Preserve export order instead of sorting participants A–Z |
| `--include-cancelled` | Also generate rosters for cancelled courses |

**Example:**

```bash
python build_rosters.py "May Export.xlsx" --templates templates --out "May Rosters"
```

---

## Deploying to Streamlit Community Cloud (Free)

1. **Fork or upload this repo to GitHub** — the repo must be public for the free tier.

2. **Go to [share.streamlit.io](https://share.streamlit.io)** and sign in with your GitHub account.

3. Click **Create app → From an existing repo**, select your repository, set the main file path to `app.py`, and click **Deploy**.

4. Streamlit builds and launches the app — usually takes about a minute.

5. **Optional — add a password:**
   In the Streamlit Cloud dashboard, go to your app's **Settings → Secrets** and add:
   ```toml
   password = "your-password-here"
   ```
   Save and the app will show a password prompt before anyone can use it.

The app will go to sleep after a period of inactivity on the free tier; the first visit after a quiet period takes about 30 seconds to wake up.

### Updating the app

Push changes to your GitHub repo (replace files via **Add file → Upload files**). Streamlit Cloud redeploys automatically within a minute.

Files you may need to update:
- `build_rosters.py` — when logic changes
- `app.py` — when the UI changes
- files in `templates/` — when blank templates change

---

## Adding a New Template

1. Add the blank `.docx` file to the `templates/` folder.
2. Update `detect_course_type()` in `build_rosters.py` to return a new type string for the relevant course-name keyword.
3. Update `classify_template()` to match that type string to your new file (by filename keyword).
4. No other changes needed.

---

## Excel Export Format

The script expects a single sheet with these column headers (exact names, case-insensitive):

| Column | Description |
|---|---|
| Name | Full course entry string including account ID, date, course name, city |
| Customer | Company name (may include account prefix) |
| Course Location | Address (may be prefixed with company or contact name) |
| Training Contact | Contact name (may include account prefix) |
| Training Contact Phone | Phone number |
| Course Date | Start date |
| Course End Date | End date |
| Other Date | Middle session date (for multi-session courses) |
| Time (from) | Start time |
| Time (To) | End time |
| Instructors | Instructor name |
| Participant Names | Newline-separated participant list (emails, bullets auto-handled) |
| Recertification Participants | Same format — used for recert courses |

---

## Dependencies

| Package | Purpose |
|---|---|
| `python-docx` | Reading and writing Word `.docx` files |
| `openpyxl` | Reading Excel `.xlsx` files |
| `streamlit` | Web interface |

Install all at once:
```bash
pip install python-docx openpyxl streamlit
```

---

## Notes

- The three First Aid templates (Blended, In-Class, Recert) have their In-Class and Blended filenames cross-named in the originals. The script identifies them by layout (presence of an "Online Part 1" column and "Manual" in Course Material), not filename — so renaming the files will not break detection.
- If a course has no instructor in the export, the filename uses `Instructor` as a placeholder.
- If a course has no company in the export (e.g. a public course with no Customer value), the company segment is omitted from the filename.
- The Notes field in the roster is always left blank — there is no Notes column in the export format.
