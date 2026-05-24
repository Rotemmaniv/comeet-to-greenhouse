# Comeet → Greenhouse Converter

Converts a Comeet (Recruit) candidate export into the exact format required for Greenhouse bulk import.

---

## Setup (one time)

### 1. Install Python
- **macOS** — already installed. Open Terminal and run `python3 --version` to confirm.
- **Windows** — download from [python.org/downloads](https://www.python.org/downloads/). During install, check **"Add Python to PATH"**.

### 2. Download the two script files
Download both files from this repo and save them **in the same folder** (e.g. your Desktop):
- `transform.py`
- `run.py`

The first time you run the converter it will automatically install the required libraries.

---

## Every time you use it

1. **Export candidates from Comeet** as `.xlsx` (Windows) or `.numbers` / `.xlsx` (macOS)
2. **Double-click `run.py`** — a window will appear
3. Click **Select Comeet Export File…** and pick your export
4. `greenhouse_import.xlsx` will appear in the same folder as your export
5. Upload that file to Greenhouse

> **Windows note:** if double-clicking `run.py` opens Notepad instead of running it,
> right-click → Open with → Python.

---

## Output columns

| Greenhouse column | Source |
|---|---|
| First Name / Last Name | `Name` (split) |
| Company | `Current company` |
| Title | `Current position` |
| Email | `Email` |
| Phone | `Phone #1` |
| Website | `Candidate LinkedIn URL` |
| Address | City + State + Country |
| Source | `Source Name` |
| Who gets credit | `Recruiter(s)` — first name only |
| Job | `Position Name` |
| Milestone | Mapped from `Current stage` (see below) |
| Education | Degree + Institution + Level |
| Notes | Skills, Languages, Availability, Salary, interview & questionnaire answers |

### Milestone mapping

| Comeet stage | Greenhouse milestone |
|---|---|
| Application screening, CV Screen | Application |
| Phone screen, Interview, Technical, Home assignment | Assessment |
| Face to face, In-person, On-site | Face to Face |
| Offer | Offer |
| Hired, Onboarding | Hired |
