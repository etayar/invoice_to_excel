"""
invoice_to_excel — main entry point.
Run with: python app.py
"""
from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

import pandas as pd

# Make sure the src package is importable whether running from source or from
# a PyInstaller bundle (where __file__ is the exe path).
_BASE = os.path.dirname(os.path.abspath(__file__))
if _BASE not in sys.path:
    sys.path.insert(0, _BASE)

from src.file_loader import load_file, FileLoadError
from src.mapper import TARGET_COLUMNS, NO_MAP_OPTION, apply_mapping, validate_mapping
from src.exporter import export_to_excel

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
APP_TITLE = "חשבונית לאקסל"
WIN_SIZE = "980x760"
WIN_MIN = (800, 600)
FONT = ("Arial", 10)
FONT_BOLD = ("Arial", 10, "bold")
PAD = 6


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _auto_match(target: str, source_cols: list) -> str:
    """Return the first source column whose name matches *target* (case-insensitive),
    or NO_MAP_OPTION when there is no match."""
    needle = target.strip().lower()
    for col in source_cols:
        if str(col).strip().lower() == needle:
            return col
    return NO_MAP_OPTION


def _cell_text(value) -> str:
    """Convert a cell value to a display string, treating NaN as empty."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value)


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(WIN_SIZE)
        self.minsize(*WIN_MIN)

        # Apply Arial globally so Hebrew characters render properly.
        for name in ("TkDefaultFont", "TkTextFont", "TkFixedFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(family="Arial", size=10)
            except Exception:
                pass

        # State
        self.source_df: pd.DataFrame | None = None
        self.current_file: str | None = None
        self.sheet_names: list | None = None
        self.mapping_vars: dict[str, tk.StringVar] = {}

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)   # preview row grows

        # ── Row 0: file chooser ────────────────────────────────────────
        file_lf = ttk.LabelFrame(main, text=" בחירת קובץ ", padding=8)
        file_lf.grid(row=0, column=0, sticky="ew", pady=(0, PAD))

        ttk.Button(file_lf, text="בחר קובץ", command=self._on_choose_file).pack(side="left")
        self.lbl_file = ttk.Label(
            file_lf, text="לא נבחר קובץ", foreground="gray", font=FONT
        )
        self.lbl_file.pack(side="left", padx=10)

        # ── Row 1: sheet selector (hidden until an Excel file with >1 sheet) ─
        self.sheet_lf = ttk.LabelFrame(main, text=" בחירת גיליון ", padding=8)
        self.sheet_lf.grid(row=1, column=0, sticky="ew", pady=(0, PAD))
        self.sheet_lf.grid_remove()   # hidden by default

        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(
            self.sheet_lf, textvariable=self.sheet_var,
            state="readonly", width=36, font=FONT,
        )
        self.sheet_combo.pack(side="left")
        ttk.Button(
            self.sheet_lf, text="טען גיליון", command=self._on_load_sheet
        ).pack(side="left", padx=8)

        # ── Row 2: data preview ────────────────────────────────────────
        preview_lf = ttk.LabelFrame(
            main, text=" תצוגה מקדימה (עד 100 שורות) ", padding=8
        )
        preview_lf.grid(row=2, column=0, sticky="nsew", pady=(0, PAD))
        preview_lf.rowconfigure(0, weight=1)
        preview_lf.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(preview_lf, show="headings", height=8)
        vsb = ttk.Scrollbar(preview_lf, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(preview_lf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # ── Row 3: column mapping ──────────────────────────────────────
        self.mapping_lf = ttk.LabelFrame(main, text=" מיפוי עמודות ", padding=8)
        self.mapping_lf.grid(row=3, column=0, sticky="ew", pady=(0, PAD))

        self._lbl_no_file = ttk.Label(
            self.mapping_lf,
            text="טען קובץ כדי להגדיר מיפוי.",
            foreground="gray", font=FONT,
        )
        self._lbl_no_file.pack()

        # ── Row 4: export button ───────────────────────────────────────
        btn_row = ttk.Frame(main)
        btn_row.grid(row=4, column=0, sticky="ew")

        self.btn_export = ttk.Button(
            btn_row, text="ייצא לאקסל",
            command=self._on_export, state="disabled",
        )
        self.btn_export.pack(side="right", ipadx=14, ipady=5)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _on_choose_file(self):
        path = filedialog.askopenfilename(
            title="בחר קובץ קלט",
            filetypes=[
                ("קבצים נתמכים", "*.pdf *.xlsx *.xls *.csv"),
                ("PDF",          "*.pdf"),
                ("Excel",        "*.xlsx *.xls"),
                ("CSV",          "*.csv"),
                ("כל הקבצים",   "*.*"),
            ],
        )
        if path:
            self.current_file = path
            self._load_file(path, sheet_name=None)

    def _on_load_sheet(self):
        if self.current_file:
            self._load_file(self.current_file, sheet_name=self.sheet_var.get())

    def _on_export(self):
        if self.source_df is None:
            messagebox.showwarning("אזהרה", "אנא טען קובץ תחילה.")
            return

        mapping = {t: v.get() for t, v in self.mapping_vars.items()}

        # Validate and warn — but always let the user continue.
        warnings = validate_mapping(self.source_df, mapping)
        if warnings:
            bullet_list = "\n".join(f"  • {w}" for w in warnings)
            proceed = messagebox.askyesno(
                "אזהרה לפני ייצוא",
                f"נמצאו בעיות במיפוי:\n\n{bullet_list}\n\n"
                "ייתכן שחלק מהנתונים יהיו ריקים בקובץ המיוצא.\n"
                "האם ברצונך להמשיך בכל זאת?",
                icon="warning",
            )
            if not proceed:
                return

        save_path = filedialog.asksaveasfilename(
            title="שמור קובץ אקסל",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="output.xlsx",
        )
        if not save_path:
            return

        try:
            result_df = apply_mapping(self.source_df, mapping)
            export_to_excel(result_df, save_path)
            messagebox.showinfo("הצלחה", f"הקובץ נשמר בהצלחה:\n{save_path}")
        except PermissionError:
            messagebox.showerror(
                "שגיאת הרשאות",
                f"לא ניתן לשמור את הקובץ.\n"
                f"ודא שהקובץ אינו פתוח בתוכנה אחרת:\n{save_path}",
            )
        except Exception as exc:
            messagebox.showerror("שגיאה בייצוא", f"אירעה שגיאה בעת הייצוא:\n{exc}")

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _load_file(self, path: str, sheet_name=None):
        try:
            df, sheet_names = load_file(path, sheet_name=sheet_name)
        except FileLoadError as exc:
            messagebox.showerror("שגיאה בטעינת קובץ", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("שגיאה בלתי צפויה", f"אירעה שגיאה:\n{exc}")
            return

        self.source_df = df
        self.sheet_names = sheet_names
        self.lbl_file.configure(
            text=os.path.basename(path), foreground="black"
        )

        # Show / hide sheet selector.
        if sheet_names and len(sheet_names) > 1:
            self.sheet_combo["values"] = sheet_names
            if sheet_name is None:
                self.sheet_var.set(sheet_names[0])
            self.sheet_lf.grid()
        else:
            self.sheet_lf.grid_remove()

        self._refresh_preview()
        self._build_mapping()
        self.btn_export.configure(state="normal")

    def _refresh_preview(self):
        df = self.source_df.head(100)
        cols = list(df.columns)

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols

        for col in cols:
            text = str(col)
            width = max(80, min(len(text) * 9, 220))
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, minwidth=50, stretch=True)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[_cell_text(v) for v in row])

    def _build_mapping(self):
        for w in self.mapping_lf.winfo_children():
            w.destroy()

        self.mapping_vars = {}
        source_options = [NO_MAP_OPTION] + list(self.source_df.columns)

        container = ttk.Frame(self.mapping_lf)
        container.pack(anchor="w")

        # Header row
        ttk.Label(container, text="עמודת יעד",   font=FONT_BOLD).grid(
            row=0, column=0, padx=(0, 12), pady=(0, 4), sticky="e"
        )
        ttk.Label(container, text="עמודת מקור", font=FONT_BOLD).grid(
            row=0, column=2, padx=(20, 0), pady=(0, 4), sticky="w"
        )
        ttk.Separator(container, orient="horizontal").grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=3
        )

        for i, target in enumerate(TARGET_COLUMNS):
            var = tk.StringVar(
                value=_auto_match(target, list(self.source_df.columns))
            )
            self.mapping_vars[target] = var

            ttk.Label(
                container, text=target, font=FONT, width=28, anchor="e"
            ).grid(row=i + 2, column=0, padx=(0, 8), pady=3, sticky="e")

            ttk.Label(container, text="→").grid(row=i + 2, column=1, padx=4)

            ttk.Combobox(
                container,
                textvariable=var,
                values=source_options,
                state="readonly",
                width=38,
                font=FONT,
            ).grid(row=i + 2, column=2, padx=(8, 0), pady=3, sticky="w")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
