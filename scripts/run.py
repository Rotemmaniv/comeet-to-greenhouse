#!/usr/bin/env python3
"""
Comeet → Greenhouse converter — GUI launcher.
Double-click this file to open the converter.
Requires transform.py in the same folder.
"""

import os
import subprocess
import sys
import threading
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    print("tkinter not available. Run from command line instead:")
    print(f"  python3 transform.py <comeet_export_file>")
    sys.exit(1)

TRANSFORM_SCRIPT = Path(__file__).parent / "transform.py"

FILETYPES = [
    ("Comeet exports", "*.numbers *.xlsx *.xlsm *.csv"),
    ("Apple Numbers",  "*.numbers"),
    ("Excel",          "*.xlsx *.xlsm"),
    ("CSV",            "*.csv"),
    ("All files",      "*.*"),
]

BG        = "#F5F5F5"
ACCENT    = "#1F4E79"
BTN_FG    = "#FFFFFF"
OK_COLOR  = "#2E7D32"
ERR_COLOR = "#C62828"
DIM       = "#777777"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comeet → Greenhouse")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._output_path = None
        self._build()
        self._center()

    def _build(self):
        # ── Header ──────────────────────────────────────────────────────────
        tk.Label(
            self, text="Comeet → Greenhouse",
            font=("Arial", 16, "bold"), bg=BG, fg=ACCENT,
        ).pack(padx=32, pady=(24, 4))

        tk.Label(
            self,
            text=(
                "Select your Comeet export file\n"
                "(.numbers, .xlsx, or .csv)\n"
                "and receive a ready-to-upload Greenhouse file."
            ),
            font=("Arial", 10), bg=BG, fg=DIM, justify="center",
        ).pack(padx=32, pady=(0, 20))

        # ── Main button ──────────────────────────────────────────────────────
        self.btn = tk.Button(
            self, text="Select Comeet Export File…",
            font=("Arial", 12, "bold"),
            bg=ACCENT, fg=BTN_FG,
            activebackground="#163A5A", activeforeground=BTN_FG,
            relief="flat", cursor="hand2", padx=20, pady=10,
            command=self._pick,
        )
        self.btn.pack(padx=32)

        # ── Status label ─────────────────────────────────────────────────────
        self._status = tk.StringVar(value="")
        self._status_lbl = tk.Label(
            self, textvariable=self._status,
            font=("Arial", 10), bg=BG, fg=DIM,
            wraplength=340, justify="center",
        )
        self._status_lbl.pack(padx=32, pady=(12, 4))

        # ── Reveal button (hidden until success) ─────────────────────────────
        label = "Show in Finder" if sys.platform == "darwin" else "Show in Explorer"
        self._reveal_btn = tk.Button(
            self, text=label,
            font=("Arial", 10), bg="#E0E0E0", fg="#333333",
            relief="flat", cursor="hand2", padx=12, pady=6,
            command=self._reveal,
        )

        # ── Divider + Quit ───────────────────────────────────────────────────
        tk.Frame(self, bg="#DDDDDD", height=1).pack(fill="x", padx=24, pady=(16, 0))
        tk.Button(
            self, text="Quit", font=("Arial", 9), bg=BG, fg="#AAAAAA",
            relief="flat", cursor="hand2", command=self.destroy,
        ).pack(pady=(4, 16))

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _pick(self):
        path = filedialog.askopenfilename(
            title="Select Comeet Export", filetypes=FILETYPES
        )
        if path:
            self._run(Path(path))

    def _run(self, input_path: Path):
        self.btn.config(state="disabled", text="Converting…")
        self._reveal_btn.pack_forget()
        self._set_status("Converting…", DIM)

        def worker():
            result = subprocess.run(
                [sys.executable, str(TRANSFORM_SCRIPT), str(input_path)],
                capture_output=True, text=True,
            )
            self.after(0, self._done, result, input_path)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, result, input_path: Path):
        self.btn.config(state="normal", text="Select Comeet Export File…")

        if result.returncode == 0:
            out = input_path.parent / "greenhouse_import.xlsx"
            self._output_path = out

            # Parse candidate count from script output
            count_str = ""
            for line in result.stdout.splitlines():
                if "Done" in line and "candidate" in line:
                    try:
                        count_str = line.split("—")[1].split("written")[0].strip()
                    except IndexError:
                        pass
                    break

            self._set_status(
                f"✓  Done — {count_str}\nSaved as: {out.name}\nin: {out.parent}",
                OK_COLOR,
            )
            self._reveal_btn.pack(pady=(0, 4))
        else:
            err = (result.stderr or result.stdout or "Unknown error").strip()
            self._set_status(f"✗  Error:\n{err[:300]}", ERR_COLOR)

    def _set_status(self, text: str, color: str):
        self._status.set(text)
        self._status_lbl.config(fg=color)

    def _reveal(self):
        if not self._output_path:
            return
        if sys.platform == "darwin":
            subprocess.run(["open", "-R", str(self._output_path)])
        else:
            # Windows: open Explorer with the file selected
            subprocess.run(["explorer", "/select,", str(self._output_path)])


if __name__ == "__main__":
    App().mainloop()
