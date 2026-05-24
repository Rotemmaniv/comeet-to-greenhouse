---
name: comeet-to-greenhouse
description: >-
  Converts a Comeet (Recruit) export file into a Greenhouse bulk-import file.
  Accepts the export as-is in .numbers, .csv, or .xlsx format and produces
  greenhouse_import.xlsx with the exact columns Greenhouse expects: First Name,
  Last Name, Company, Title, Notes, Email, Phone, Social Media, Website,
  Address, Source, Who gets credit, Job, Milestone. Notes includes Skills,
  Languages, Education, Availability, Salary expectations, Disposition Notes,
  and all filled questionnaire/interview answers. Use when a team member
  provides a Comeet export and wants a ready-to-upload Greenhouse import file.
---

# Comeet → Greenhouse converter

## What this skill does

Transforms a Comeet (Recruit) candidate export into the exact format required
by Greenhouse's bulk-import template. The team member provides their downloaded
Comeet export and receives `greenhouse_import.xlsx` — ready to drag into
Greenhouse without any manual column work.

---

## Single-shot mode (default)

When the user provides a Comeet export file (absolute path or `@`-attached),
run the full pipeline without asking intermediate questions:

1. Resolve the file path (Step 0)
2. Install dependencies if needed (Step 1)
3. Run the transformation (Step 2)
4. Report the result (Step 3)

---

## Step 0 — Resolve the input file (BLOCKING)

Do **not** run the script until a file is resolved.

- If the message contains an absolute path to a `.numbers`, `.csv`, `.xlsx`,
  or `.xlsm` file → treat as resolved.
- If a file is `@`-attached → use that path.
- Otherwise ask once:  
  > "Please provide the path to your Comeet export file (.numbers, .csv, or
  > .xlsx), or @-attach it."

---

## Step 1 — Ensure dependencies

Run this check once before calling the script:

```bash
python3 -c "import pandas, openpyxl, numbers_parser" 2>/dev/null || \
  pip3 install --quiet pandas openpyxl numbers-parser
```

---

## Step 2 — Run the transformation

```bash
python3 "$HOME/.cursor/skills/comeet-to-greenhouse/scripts/transform.py" \
  "<ABSOLUTE_PATH_TO_COMEET_EXPORT>"
```

- The output file is written to the **same folder** as the input file, named
  `greenhouse_import.xlsx`.
- To write elsewhere, add `--out <path>`.
- The script prints progress lines and a final `✓ Done` summary with the
  output path.

### What the script maps

| Greenhouse column | Source in Comeet export |
|---|---|
| First Name | `Name` — first word |
| Last Name | `Name` — remainder |
| Company | `Current company` |
| Title | `Current position` |
| Notes | Skills, Languages, Availability, Salary expectations, Disposition Notes, **all non-empty questionnaire / interview answers** |
| Email | `Email` |
| Phone | `Phone #1` (float `.0` suffix stripped) |
| Social Media | `Candidate LinkedIn URL` |
| Website | `Candidate URL` (Comeet profile link) |
| Address | `City, State, Country` (joined, blanks skipped) |
| Source | `Source Name` |
| Who gets credit | `Recruiter(s)` |
| Job | `Position Name` |
| Milestone | `Current stage` — **mapped** to the nearest valid Greenhouse value (see below) |
| Education | `Degree \| Educational Institution \| Education level` |

### Milestone mapping

Comeet uses free-text stage names; Greenhouse accepts only five values.
The script maps by keyword (first match wins):

| Comeet stage contains… | → Greenhouse Milestone |
|---|---|
| "hired", "onboard" | **Hired** |
| "offer" | **Offer** |
| "face to face", "in person", "in-person", "onsite", "on-site" | **Face to Face** |
| "phone", "video", "introduct", "interview", "technical", "assignment", "exercise", "test", "task", "home assign", "take home" | **Assessment** |
| *(anything else)* | **Application** |

Questionnaire answers in Notes are grouped by form name:

```
[Phone interview Lydia 1]
  Q: Why are you interested in exploring a new job opportunity?
  A: I'm looking for a bigger scope of work...

[Candidate Questionnaire 1]
  Q: Choose 5 values you most identify with...
  A: Ownership, Curiosity, ...
```

---

## Step 3 — Report

Print the output path and candidate count. Example:

```
✓  Done — 47 candidate(s) written to:
   /Users/someone/Downloads/greenhouse_import.xlsx
```

If the script exits with a non-zero code, show the error output clearly so the
user can fix the issue (most likely a missing dependency or an unsupported file
format).

---

## Error handling

| Problem | What to do |
|---|---|
| `ModuleNotFoundError` | Re-run Step 1, then retry Step 2 |
| `File not found` | Ask the user to confirm the correct path |
| Unsupported file type | Tell the user to export from Comeet as CSV or re-save the .numbers as .xlsx |
| Empty output (0 rows) | Warn the user — the export may only contain a header row |

---

## Notes for the team

- **Input formats accepted:** Apple Numbers (`.numbers`), CSV (`.csv`), Excel (`.xlsx` / `.xlsm`)
- **Output:** `greenhouse_import.xlsx` — sheet named `Sheet people`, header frozen and styled, Arial 10pt
- **Milestone:** automatically mapped from Comeet's `Current stage` to the nearest valid Greenhouse value (`Application`, `Assessment`, `Face to Face`, `Offer`, `Hired`). The script prints a mapping summary so you can spot any unexpected mappings before uploading.
- **Who gets credit:** copied from `Recruiter(s)` (comma-separated list as Comeet exports it). Greenhouse accepts a single name; if the field imports blank, fill it manually.
- **The script is safe to re-run** — it always overwrites `greenhouse_import.xlsx` in place.
