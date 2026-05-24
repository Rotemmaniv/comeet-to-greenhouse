#!/usr/bin/env python3
"""
Comeet → Greenhouse candidate import transformer.

Usage:
    python3 transform.py <comeet_export_file> [--out <output_file>]

Accepts:  .numbers (Apple Numbers), .csv, .xlsx / .xlsm
Outputs:  greenhouse_import.xlsx  (same folder as input unless --out is given)
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing dependencies. Run: pip3 install pandas openpyxl numbers-parser")

GREENHOUSE_COLUMNS = [
    "First Name", "Last Name", "Company", "Title", "Notes",
    "Email", "Phone", "Social Media", "Website", "Address",
    "Source", "Who gets credit", "Job", "Milestone", "Education",
]

# Structured candidate fields included in Notes (in display order)
# Education is intentionally excluded — it has its own Greenhouse column.
NOTES_STRUCTURED = [
    ("Skills",              "Skills"),
    ("Languages",           "Languages"),
    ("Availability",        "Availability"),
    ("Salary expectations", "Salary expectations"),
    ("Notes",               "Disposition Notes"),
]

# ── Milestone mapping ─────────────────────────────────────────────────────────
# Maps Comeet's free-text stage names to Greenhouse's fixed milestone values.
# Rules are checked in order; first match wins. Default: "Application".
_MILESTONE_RULES = [
    ("Hired",        ["hired", "onboard"]),
    ("Offer",        ["offer"]),
    ("Face to Face", ["face to face", "in person", "in-person", "onsite", "on-site"]),
    ("Assessment",   ["phone", "video", "introduct", "interview", "technical",
                      "assignment", "exercise", "test", "task", "home assign", "take home"]),
]

def normalize_milestone(stage: str) -> str:
    """Map a Comeet stage name to the nearest valid Greenhouse milestone."""
    if not stage:
        return "Application"
    lower = stage.lower()
    for milestone, keywords in _MILESTONE_RULES:
        if any(kw in lower for kw in keywords):
            return milestone
    return "Application"


# ── helpers ──────────────────────────────────────────────────────────────────

def _val(row, col):
    """Return stripped string value or '' for None/NaN/empty.

    Handles duplicate column names: Comeet exports repeat some questionnaire
    headers, causing pandas to return a Series instead of a scalar.
    In that case we take the first non-null value.
    """
    import math

    v = row.get(col) if isinstance(row, dict) else (
        row[col] if col in row.index else None
    )

    # Duplicate column → pandas returns a Series; pick first non-null value
    if isinstance(v, pd.Series):
        non_null = v.dropna()
        v = non_null.iloc[0] if len(non_null) > 0 else None

    if v is None:
        return ""
    try:
        if isinstance(v, float) and math.isnan(v):
            return ""
    except Exception:
        pass
    s = str(v).strip()
    return "" if s.lower() in ("none", "nan") else s


def clean_phone(val) -> str:
    if val is None:
        return ""
    import math
    try:
        if isinstance(val, float) and math.isnan(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s.lower() == "none":
        return ""
    # Strip .0 suffix produced by float conversion
    if s.endswith(".0"):
        s = s[:-2]
    return s


def split_name(name: str):
    if not name:
        return "", ""
    parts = name.strip().split(" ", 1)
    return parts[0], (parts[1] if len(parts) > 1 else "")


def build_education(row) -> str:
    level = _val(row, "Education level")
    inst  = _val(row, "Educational Institution")
    deg   = _val(row, "Degree")
    parts = [p for p in [deg, inst, level] if p]
    return " | ".join(parts)


def build_notes(row, all_columns: list) -> str:
    blocks = []

    # ── 1. Structured fields ────────────────────────────────────────────────
    for label, col in NOTES_STRUCTURED:
        v = _val(row, col)
        if v:
            blocks.append(f"{label}: {v}")

    # ── 2. Questionnaire / interview answers ────────────────────────────────
    current_form = None
    form_lines = []
    seen_q_cols: set = set()   # guard against duplicate column names

    def flush_form():
        nonlocal current_form, form_lines
        if current_form and form_lines:
            blocks.append(f"[{current_form}]\n" + "\n".join(form_lines))
        current_form = None
        form_lines = []

    for col in all_columns:
        if not (col and col.startswith("[")):
            continue
        if col in seen_q_cols:          # skip duplicate header names
            continue
        seen_q_cols.add(col)
        v = _val(row, col)
        if not v:
            continue

        m = re.match(r"\[([^\]]+)\](.*)", col)
        if not m:
            continue
        form_name = m.group(1).strip()
        question  = m.group(2).strip().lstrip("•").strip()

        if form_name != current_form:
            flush_form()
            current_form = form_name

        if question:
            form_lines.append(f"  Q: {question}\n  A: {v}")
        else:
            form_lines.append(f"  {v}")

    flush_form()

    return "\n\n".join(blocks)


# ── file readers ──────────────────────────────────────────────────────────────

def read_numbers(path: Path) -> pd.DataFrame:
    try:
        from numbers_parser import Document
    except ImportError:
        sys.exit("numbers-parser not installed. Run: pip3 install numbers-parser")
    doc = Document(str(path))
    # Pick the table with the most rows across all sheets — handles cases where
    # Numbers adds an empty table when the file is opened on macOS.
    best_table = None
    best_row_count = -1
    for sheet in doc.sheets:
        for table in sheet.tables:
            if table.num_rows > best_row_count:
                best_row_count = table.num_rows
                best_table = table
    if best_table is None:
        sys.exit("No tables found in the .numbers file.")
    rows    = list(best_table.iter_rows())
    headers = [c.value for c in rows[0]]
    data    = [[c.value for c in row] for row in rows[1:]]
    return pd.DataFrame(data, columns=headers)


def read_input(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext == ".numbers":
        return read_numbers(path)
    elif ext == ".csv":
        return pd.read_csv(path, dtype=str)
    elif ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=str)
    else:
        sys.exit(f"Unsupported file type: {ext!r}. Expected .numbers, .csv, or .xlsx")


# ── transform ─────────────────────────────────────────────────────────────────

def transform(df: pd.DataFrame) -> pd.DataFrame:
    all_cols = list(df.columns)
    output_rows = []

    for _, row in df.iterrows():
        first, last = split_name(_val(row, "Name"))

        city    = _val(row, "City")
        state   = _val(row, "State")
        country = _val(row, "Country")
        addr_parts = [p for p in [city, state, country] if p]
        address = ", ".join(addr_parts)

        output_rows.append({
            "First Name":     first,
            "Last Name":      last,
            "Company":        _val(row, "Current company"),
            "Title":          _val(row, "Current position"),
            "Notes":          build_notes(row, all_cols),
            "Email":          _val(row, "Email"),
            "Phone":          clean_phone(
                                  row.get("Phone #1") if isinstance(row, dict)
                                  else (row["Phone #1"] if "Phone #1" in row.index else None)
                              ),
            "Social Media":   "",
            "Website":        _val(row, "Candidate LinkedIn URL"),
            "Address":        address,
            "Source":         _val(row, "Source Name"),
            "Who gets credit": _val(row, "Recruiter(s)").split(",")[0].strip(),
            "Job":            _val(row, "Position Name"),
            "Milestone":      normalize_milestone(_val(row, "Current stage")),
            "Education":      build_education(row),
        })

    return pd.DataFrame(output_rows, columns=GREENHOUSE_COLUMNS)


# ── Excel formatting ──────────────────────────────────────────────────────────

COL_WIDTHS = {
    "First Name": 18, "Last Name": 20, "Company": 28, "Title": 30,
    "Notes": 60,      "Email": 30,     "Phone": 18,   "Social Media": 35,
    "Website": 40,    "Address": 28,   "Source": 18,  "Who gets credit": 22,
    "Job": 30,        "Milestone": 18, "Education": 40,
}

HEADER_BG   = "1F4E79"   # dark blue
HEADER_FONT = "FFFFFF"   # white


def format_workbook(path: Path):
    wb = load_workbook(str(path))
    ws = wb["Sheet people"]

    # Header row styling
    for cell in ws[1]:
        cell.font      = Font(name="Arial", bold=True, color=HEADER_FONT, size=10)
        cell.fill      = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

    # Data rows
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font      = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Column widths
    for col_idx, col_name in enumerate(GREENHOUSE_COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = COL_WIDTHS.get(col_name, 20)

    # Freeze header
    ws.freeze_panes = "A2"

    # Row height for header
    ws.row_dimensions[1].height = 22

    wb.save(str(path))


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert a Comeet export to Greenhouse import format."
    )
    parser.add_argument("input",  help="Comeet export file (.numbers / .csv / .xlsx)")
    parser.add_argument("--out",  default=None,
                        help="Output path (default: greenhouse_import.xlsx next to input)")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        sys.exit(f"File not found: {input_path}")

    out_path = (
        Path(args.out).expanduser().resolve()
        if args.out
        else input_path.parent / "greenhouse_import.xlsx"
    )

    print(f"Reading:      {input_path}")
    df = read_input(input_path)
    print(f"              {len(df)} candidate(s), {len(df.columns)} column(s) found")

    print("Transforming...")
    gh_df = transform(df)

    # Show milestone mapping summary
    if "Milestone" in gh_df.columns and "Current stage" in df.columns:
        stage_col = df["Current stage"] if isinstance(df["Current stage"], pd.Series) else df.iloc[:, list(df.columns).index("Current stage")]
        mapping = {}
        for raw, mapped in zip(stage_col, gh_df["Milestone"]):
            raw_s = str(raw).strip() if raw is not None else ""
            if raw_s and raw_s.lower() not in ("none", "nan"):
                mapping[raw_s] = mapped
        if mapping:
            print("  Milestone mapping:")
            for src, dst in sorted(mapping.items()):
                print(f"    {src!r:35s} → {dst!r}")

    print(f"Writing:      {out_path}")
    gh_df.to_excel(str(out_path), index=False, sheet_name="Sheet people")

    print("Formatting...")
    format_workbook(out_path)

    print(f"\n✓  Done — {len(gh_df)} candidate(s) written to:\n   {out_path}")


if __name__ == "__main__":
    main()
