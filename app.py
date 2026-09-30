"""ITSRR department selector desktop application."""

from __future__ import annotations

import json
import re
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


DEFAULT_DEPARTMENTS = {
    "HR": "HR",
    "Sales": "S",
    "Nurse": "NRS",
    "Safety Officer": "SO",
    "Jobin QM": "JQM",
    "Jobin Maintenance": "JM",
    "Jobin Production": "JP",
    "Daichi Warehouse": "DWHS",
    "Cavite Warehouse": "CWHS",
    "Cavite CoEx": "CCoEx",
    "Cavite Office": "CO",
    "Cavite Production": "CP",
    "Grace Office": "GRC",
    "Makati": "MKT",
}

APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
COUNTER_FILE = APP_DIR / "itsrr_counters.json"
SUMMARY_FILE = APP_DIR / "ITSRR Summary.xlsx"
CUSTOM_DEPARTMENTS_FILE = APP_DIR / "custom_departments.json"

FORM_FIELDS = (
    "ITSRR Number",
    "Terminal User",
    "Date and Time Reported",
    "Reported By",
    "Problem",
    "Date and Time Resolved",
    "Cause",
    "Solution",
    "Status",
)


class DepartmentSelector(tk.Tk):
    """Main window for choosing which department form to generate."""

    BACKGROUND = "#F4F7FB"
    SURFACE = "#FFFFFF"
    PRIMARY = "#174EA6"
    PRIMARY_HOVER = "#0F3D85"
    TEXT = "#172033"
    MUTED = "#667085"
    BORDER = "#D9E2F1"

    def __init__(self) -> None:
        super().__init__()
        self.title("ITSRR Form Code Generator")
        self.geometry("1040x680")
        self.minsize(720, 560)
        self.configure(bg=self.BACKGROUND)

        self.departments = self._load_departments()
        self.selected_department = tk.StringVar(value="")
        self.generated_code = ""
        self.department_buttons: list[tk.Button] = []
        self._resize_job: str | None = None

        self._build_interface()
        self.bind("<Configure>", self._schedule_grid_update)
        self.bind("<Escape>", lambda _event: self.destroy())
        self.after_idle(self._update_department_grid)

    def _build_interface(self) -> None:
        header = tk.Frame(self, bg=self.PRIMARY, padx=42, pady=30)
        header.pack(fill="x")

        tk.Label(
            header,
            text="ITSRR FORM CODE GENERATOR",
            bg=self.PRIMARY,
            fg="#CFE0FF",
            font=("Segoe UI Semibold", 10),
        ).pack(anchor="w")

        tk.Label(
            header,
            text="Choose a department",
            bg=self.PRIMARY,
            fg="white",
            font=("Segoe UI Semibold", 26),
        ).pack(anchor="w", pady=(5, 2))

        tk.Label(
            header,
            text="Select where the form will be used to begin.",
            bg=self.PRIMARY,
            fg="#E7EFFF",
            font=("Segoe UI", 11),
        ).pack(anchor="w")

        content = tk.Frame(self, bg=self.BACKGROUND, padx=42, pady=30)
        content.pack(fill="both", expand=True)

        top_row = tk.Frame(content, bg=self.BACKGROUND)
        top_row.pack(fill="x", pady=(0, 16))

        tk.Label(
            top_row,
            text="Departments",
            bg=self.BACKGROUND,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 15),
        ).pack(side="left")

        self.department_count_label = tk.Label(
            top_row,
            text=f"{len(self.departments)} available",
            bg=self.BACKGROUND,
            fg=self.MUTED,
            font=("Segoe UI", 10),
        )
        self.department_count_label.pack(side="right")

        self.grid_container = tk.Frame(
            content,
            bg=self.SURFACE,
            highlightbackground=self.BORDER,
            highlightthickness=1,
        )
        self.grid_container.pack(fill="both", expand=True)

        self.department_canvas = tk.Canvas(
            self.grid_container,
            bg=self.SURFACE,
            highlightthickness=0,
        )
        department_scrollbar = tk.Scrollbar(
            self.grid_container,
            orient="vertical",
            command=self.department_canvas.yview,
        )
        self.department_canvas.configure(
            yscrollcommand=department_scrollbar.set
        )
        department_scrollbar.pack(side="right", fill="y")
        self.department_canvas.pack(side="left", fill="both", expand=True)

        self.grid_frame = tk.Frame(
            self.department_canvas,
            bg=self.SURFACE,
            padx=20,
            pady=20,
        )
        self.grid_window = self.department_canvas.create_window(
            (0, 0),
            window=self.grid_frame,
            anchor="nw",
        )
        self.grid_frame.bind(
            "<Configure>",
            self._update_scroll_region,
        )
        self.department_canvas.bind(
            "<Configure>",
            self._resize_scroll_content,
        )
        self.department_canvas.bind(
            "<MouseWheel>",
            self._scroll_departments,
        )
        self.grid_frame.bind(
            "<MouseWheel>",
            self._scroll_departments,
        )

        for department in self.departments:
            self._add_department_button_widget(department)

        footer = tk.Frame(content, bg=self.BACKGROUND)
        footer.pack(
            side="bottom",
            fill="x",
            pady=(18, 0),
            before=self.grid_container,
        )

        status_row = tk.Frame(footer, bg=self.BACKGROUND)
        status_row.pack(fill="x")

        self.status_label = tk.Label(
            status_row,
            text="No department selected",
            bg=self.BACKGROUND,
            fg=self.MUTED,
            font=("Segoe UI", 10),
        )
        self.status_label.pack(side="left")

        tk.Label(
            status_row,
            text="Press Esc to close",
            bg=self.BACKGROUND,
            fg="#98A2B3",
            font=("Segoe UI", 9),
        ).pack(side="right")

        tk.Button(
            status_row,
            text="Counter Settings",
            command=self._open_counter_settings,
            bg="#E5E7EB",
            fg=self.TEXT,
            activebackground="#D1D5DB",
            relief="solid",
            bd=1,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            padx=12,
            pady=5,
        ).pack(side="right", padx=(0, 16))

        tk.Button(
            status_row,
            text="Compile Response",
            command=self._open_compile_dialog,
            bg=self.PRIMARY,
            fg="white",
            activebackground=self.PRIMARY_HOVER,
            activeforeground="white",
            relief="solid",
            bd=1,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            padx=12,
            pady=5,
        ).pack(side="right", padx=(0, 8))

        tk.Button(
            status_row,
            text="Add Department",
            command=self._open_add_department_dialog,
            bg="#EAF1FC",
            fg=self.PRIMARY,
            activebackground="#D8E6FA",
            activeforeground=self.PRIMARY_HOVER,
            relief="solid",
            bd=1,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            padx=12,
            pady=5,
        ).pack(side="right", padx=(0, 8))

    def _add_department_button_widget(self, department: str) -> None:
        button = tk.Button(
            self.grid_frame,
            text=department,
            command=lambda name=department: self._select_department(name),
            bg=self.SURFACE,
            fg=self.TEXT,
            activebackground="#EAF1FC",
            activeforeground=self.PRIMARY,
            relief="solid",
            bd=1,
            highlightbackground=self.BORDER,
            highlightthickness=1,
            cursor="hand2",
            font=("Segoe UI Semibold", 11),
            padx=16,
            pady=15,
        )
        button.bind(
            "<Enter>",
            lambda _event, item=button: self._set_hover(item, True),
        )
        button.bind(
            "<Leave>",
            lambda _event, item=button: self._set_hover(item, False),
        )
        button.bind("<MouseWheel>", self._scroll_departments)
        self.department_buttons.append(button)

    def _load_departments(self) -> dict[str, str]:
        departments = dict(DEFAULT_DEPARTMENTS)
        if CUSTOM_DEPARTMENTS_FILE.exists():
            try:
                saved = json.loads(
                    CUSTOM_DEPARTMENTS_FILE.read_text(encoding="utf-8")
                )
                if isinstance(saved, dict):
                    for name, code in saved.items():
                        if (
                            isinstance(name, str)
                            and isinstance(code, str)
                            and name.strip()
                            and code.strip()
                        ):
                            departments[name.strip()] = code.strip()
            except (OSError, ValueError, TypeError):
                pass
        return departments

    def _save_custom_departments(self) -> None:
        custom_departments = {
            name: code
            for name, code in self.departments.items()
            if name not in DEFAULT_DEPARTMENTS
        }
        temporary_file = CUSTOM_DEPARTMENTS_FILE.with_suffix(".tmp")
        temporary_file.write_text(
            json.dumps(custom_departments, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        temporary_file.replace(CUSTOM_DEPARTMENTS_FILE)

    def _open_add_department_dialog(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Add Department")
        dialog.geometry("480x310")
        dialog.resizable(False, False)
        dialog.configure(bg=self.SURFACE)
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(
            dialog,
            text="Add a Department",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 18),
        ).pack(anchor="w", padx=30, pady=(24, 4))

        tk.Label(
            dialog,
            text="Enter the button name and the code used in its ITSRR number.",
            bg=self.SURFACE,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=30, pady=(0, 18))

        form = tk.Frame(dialog, bg=self.SURFACE)
        form.pack(fill="x", padx=30)
        form.grid_columnconfigure(1, weight=1)

        tk.Label(
            form,
            text="Department name:",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 10),
        ).grid(row=0, column=0, sticky="w", padx=(0, 14), pady=8)

        name_entry = tk.Entry(
            form,
            bg="#FFFFFF",
            fg=self.TEXT,
            font=("Segoe UI", 11),
            relief="solid",
            bd=1,
        )
        name_entry.grid(row=0, column=1, sticky="ew", pady=8, ipady=6)

        tk.Label(
            form,
            text="Department code:",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 10),
        ).grid(row=1, column=0, sticky="w", padx=(0, 14), pady=8)

        code_entry = tk.Entry(
            form,
            bg="#FFFFFF",
            fg=self.TEXT,
            font=("Consolas", 12, "bold"),
            relief="solid",
            bd=1,
        )
        code_entry.grid(row=1, column=1, sticky="ew", pady=8, ipady=6)

        def add_department() -> None:
            name = name_entry.get().strip()
            code = code_entry.get().strip()

            if not name or not code:
                messagebox.showerror(
                    "Missing information",
                    "Enter both a department name and department code.",
                    parent=dialog,
                )
                return
            if not re.fullmatch(r"[A-Za-z0-9]+", code):
                messagebox.showerror(
                    "Invalid code",
                    "The department code may only contain letters and numbers.",
                    parent=dialog,
                )
                return
            if any(existing.casefold() == name.casefold() for existing in self.departments):
                messagebox.showerror(
                    "Department already exists",
                    "A department with this name already exists.",
                    parent=dialog,
                )
                return
            if any(existing.casefold() == code.casefold() for existing in self.departments.values()):
                messagebox.showerror(
                    "Code already exists",
                    "This department code is already assigned.",
                    parent=dialog,
                )
                return

            self.departments[name] = code
            try:
                self._save_custom_departments()
            except OSError as error:
                del self.departments[name]
                messagebox.showerror(
                    "Could not save",
                    f"The department could not be saved:\n{error}",
                    parent=dialog,
                )
                return

            self._add_department_button_widget(name)
            self.department_count_label.configure(
                text=f"{len(self.departments)} available"
            )
            self._update_department_grid()
            dialog.destroy()

        actions = tk.Frame(dialog, bg=self.SURFACE)
        actions.pack(fill="x", padx=30, pady=(22, 0))

        tk.Button(
            actions,
            text="Cancel",
            command=dialog.destroy,
            bg="#E5E7EB",
            fg=self.TEXT,
            activebackground="#D1D5DB",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=18,
            pady=8,
        ).pack(side="right")

        tk.Button(
            actions,
            text="Add Department",
            command=add_department,
            bg=self.PRIMARY,
            fg="white",
            activebackground=self.PRIMARY_HOVER,
            activeforeground="white",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=18,
            pady=8,
        ).pack(side="right", padx=(0, 10))

        name_entry.focus_set()
        dialog.bind("<Return>", lambda _event: add_department())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _update_scroll_region(self, _event: tk.Event) -> None:
        self.department_canvas.configure(
            scrollregion=self.department_canvas.bbox("all")
        )

    def _resize_scroll_content(self, event: tk.Event) -> None:
        self.department_canvas.itemconfigure(
            self.grid_window,
            width=event.width,
        )

    def _scroll_departments(self, event: tk.Event) -> None:
        self.department_canvas.yview_scroll(
            int(-event.delta / 120),
            "units",
        )

    def _schedule_grid_update(self, _event: tk.Event) -> None:
        if self._resize_job is not None:
            self.after_cancel(self._resize_job)
        self._resize_job = self.after(80, self._update_department_grid)

    def _update_department_grid(self) -> None:
        self._resize_job = None
        width = self.winfo_width()
        columns = 3 if width >= 960 else 2 if width >= 650 else 1

        for column in range(3):
            self.grid_frame.grid_columnconfigure(column, weight=0)
        for column in range(columns):
            self.grid_frame.grid_columnconfigure(column, weight=1, uniform="department")

        for index, button in enumerate(self.department_buttons):
            button.grid_forget()
            row, column = divmod(index, columns)
            button.grid(
                row=row,
                column=column,
                sticky="nsew",
                padx=7,
                pady=7,
            )

    def _set_hover(self, button: tk.Button, hovering: bool) -> None:
        selected = button.cget("text") == self.selected_department.get()
        if selected:
            return
        button.configure(
            bg="#F0F5FD" if hovering else self.SURFACE,
            fg=self.PRIMARY if hovering else self.TEXT,
        )

    def _select_department(self, department: str) -> None:
        self.selected_department.set(department)
        for button in self.department_buttons:
            selected = button.cget("text") == department
            button.configure(
                bg=self.PRIMARY if selected else self.SURFACE,
                fg="white" if selected else self.TEXT,
                activebackground=self.PRIMARY_HOVER if selected else "#EAF1FC",
                activeforeground="white" if selected else self.PRIMARY,
            )

        self.status_label.configure(
            text=f"Selected: {department}",
            fg=self.PRIMARY,
            font=("Segoe UI Semibold", 10),
        )
        month_part = datetime.now().strftime("%m")
        counter_period = datetime.now().strftime("%Y-%m")
        prefix = self.departments[department]
        sequence = self._next_sequence(prefix, counter_period)
        self.generated_code = f"{prefix}-{month_part}-{sequence:02d}"
        self.status_label.configure(text=f"Generated: {self.generated_code}")
        self._show_code_dialog(department)

    def _load_counters(self) -> dict[str, int]:
        counters: dict[str, int] = {}
        if COUNTER_FILE.exists():
            try:
                saved = json.loads(COUNTER_FILE.read_text(encoding="utf-8"))
                if isinstance(saved, dict):
                    counters = {
                        str(key): int(value)
                        for key, value in saved.items()
                        if isinstance(value, int) and value >= 0
                    }
            except (OSError, ValueError, TypeError):
                counters = {}
        return counters

    def _save_counters(self, counters: dict[str, int]) -> None:
        temporary_file = COUNTER_FILE.with_suffix(".tmp")
        temporary_file.write_text(
            json.dumps(counters, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        temporary_file.replace(COUNTER_FILE)

    def _next_sequence(self, prefix: str, counter_period: str) -> int:
        """Return and persist the next number for this department and month."""
        counters = self._load_counters()

        counter_key = f"{prefix}-{counter_period}"
        next_number = counters.get(counter_key, 0) + 1
        counters[counter_key] = next_number
        self._save_counters(counters)
        return next_number

    def _open_compile_dialog(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Compile Client Response")
        dialog.geometry("700x680")
        dialog.minsize(560, 500)
        dialog.configure(bg=self.SURFACE)
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(
            dialog,
            text="Compile Client Response",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 20),
        ).pack(anchor="w", padx=34, pady=(24, 3))

        tk.Label(
            dialog,
            text="Paste the completed ITSRR template below, then add it to Excel.",
            bg=self.SURFACE,
            fg=self.MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=34, pady=(0, 16))

        actions = tk.Frame(dialog, bg=self.SURFACE)
        actions.pack(side="bottom", fill="x", padx=34, pady=(16, 24))

        response_text = tk.Text(
            dialog,
            wrap="word",
            bg="#FFFFFF",
            fg=self.TEXT,
            insertbackground=self.TEXT,
            selectbackground=self.PRIMARY,
            selectforeground="white",
            font=("Segoe UI", 11),
            relief="solid",
            bd=1,
            padx=16,
            pady=14,
        )
        response_text.pack(fill="both", expand=True, padx=34)
        response_text.focus_set()

        def compile_response() -> None:
            response = response_text.get("1.0", "end-1c")
            record = self._parse_response(response)
            if not record["ITSRR Number"]:
                messagebox.showerror(
                    "ITSRR number missing",
                    "Paste a completed template containing an ITSRR Number.",
                    parent=dialog,
                )
                return

            try:
                action = self._save_response_to_excel(record)
            except PermissionError:
                messagebox.showerror(
                    "Excel file is open",
                    f"Close {SUMMARY_FILE.name} in Excel and try again.",
                    parent=dialog,
                )
                return
            except OSError as error:
                messagebox.showerror(
                    "Could not save",
                    f"The Excel summary could not be saved:\n{error}",
                    parent=dialog,
                )
                return

            messagebox.showinfo(
                "Summary updated",
                f"{record['ITSRR Number']} was {action} in:\n"
                f"{SUMMARY_FILE.name}",
                parent=dialog,
            )
            response_text.delete("1.0", tk.END)
            response_text.focus_set()

        tk.Button(
            actions,
            text="Add to Excel",
            command=compile_response,
            bg=self.PRIMARY,
            fg="white",
            activebackground=self.PRIMARY_HOVER,
            activeforeground="white",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=20,
            pady=9,
        ).pack(side="right")

        tk.Button(
            actions,
            text="Clear",
            command=lambda: response_text.delete("1.0", tk.END),
            bg="#E5E7EB",
            fg=self.TEXT,
            activebackground="#D1D5DB",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=20,
            pady=9,
        ).pack(side="right", padx=(0, 10))

        dialog.bind("<Control-Return>", lambda _event: compile_response())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _parse_response(self, response: str) -> dict[str, str]:
        """Extract form values while preserving multiline answers."""
        labels = "|".join(re.escape(field) for field in FORM_FIELDS)
        record: dict[str, str] = {}
        for field in FORM_FIELDS:
            pattern = (
                rf"(?ims)^\s*{re.escape(field)}\s*:\s*(.*?)"
                rf"(?=^\s*(?:{labels})\s*:|\Z)"
            )
            match = re.search(pattern, response)
            record[field] = match.group(1).strip() if match else ""
        return record

    def _save_response_to_excel(self, record: dict[str, str]) -> str:
        """Append a new ITSRR row, or update the row with the same number."""
        headers = (*FORM_FIELDS, "Compiled On")

        if SUMMARY_FILE.exists():
            workbook = load_workbook(SUMMARY_FILE)
            worksheet = workbook["ITSRR Summary"]
        else:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "ITSRR Summary"
            worksheet.append(headers)
            header_fill = PatternFill("solid", fgColor="174EA6")
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = f"A1:J1"

        itsrr_number = record["ITSRR Number"]
        target_row = None
        for row in range(2, worksheet.max_row + 1):
            if str(worksheet.cell(row=row, column=1).value or "").strip() == itsrr_number:
                target_row = row
                break

        action = "updated" if target_row else "added"
        if target_row is None:
            target_row = worksheet.max_row + 1

        values = [record[field] for field in FORM_FIELDS]
        values.append(datetime.now().strftime("%m/%d/%Y %I:%M %p"))
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row=target_row, column=column, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        widths = (24, 22, 24, 22, 38, 24, 38, 38, 16, 24)
        for column, width in enumerate(widths, start=1):
            worksheet.column_dimensions[
                worksheet.cell(row=1, column=column).column_letter
            ].width = width

        worksheet.auto_filter.ref = f"A1:J{worksheet.max_row}"
        workbook.save(SUMMARY_FILE)
        return action

    def _open_counter_settings(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Counter Settings")
        dialog.geometry("480x330")
        dialog.resizable(False, False)
        dialog.configure(bg=self.SURFACE)
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(
            dialog,
            text="Counter Settings",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 18),
        ).pack(anchor="w", padx=30, pady=(24, 4))

        tk.Label(
            dialog,
            text="Set the next number for a department this month, or reset all counters.",
            bg=self.SURFACE,
            fg=self.MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=30, pady=(0, 18))

        form = tk.Frame(dialog, bg=self.SURFACE)
        form.pack(fill="x", padx=30)
        form.grid_columnconfigure(1, weight=1)

        tk.Label(
            form,
            text="Department:",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 10),
        ).grid(row=0, column=0, sticky="w", padx=(0, 14), pady=8)

        department_var = tk.StringVar(value=next(iter(self.departments)))
        department_menu = tk.OptionMenu(
            form,
            department_var,
            *self.departments.keys(),
        )
        department_menu.configure(
            bg="#FFFFFF",
            fg=self.TEXT,
            activebackground="#EAF1FC",
            relief="solid",
            bd=1,
            highlightthickness=0,
            font=("Segoe UI", 10),
        )
        department_menu.grid(row=0, column=1, sticky="ew", pady=8)

        tk.Label(
            form,
            text="Next counter:",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 10),
        ).grid(row=1, column=0, sticky="w", padx=(0, 14), pady=8)

        next_counter_var = tk.StringVar()
        counter_entry = tk.Entry(
            form,
            textvariable=next_counter_var,
            bg="#FFFFFF",
            fg=self.TEXT,
            font=("Consolas", 12, "bold"),
            relief="solid",
            bd=1,
        )
        counter_entry.grid(row=1, column=1, sticky="ew", pady=8, ipady=6)

        counter_period = datetime.now().strftime("%Y-%m")

        def refresh_next_counter(*_args: object) -> None:
            prefix = self.departments[department_var.get()]
            last_used = self._load_counters().get(f"{prefix}-{counter_period}", 0)
            next_counter_var.set(f"{last_used + 1:02d}")

        def save_next_counter() -> None:
            try:
                next_number = int(next_counter_var.get())
                if next_number < 1:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "Invalid counter",
                    "Enter a whole number of 1 or higher.",
                    parent=dialog,
                )
                return

            prefix = self.departments[department_var.get()]
            counters = self._load_counters()
            counters[f"{prefix}-{counter_period}"] = next_number - 1
            self._save_counters(counters)
            next_counter_var.set(f"{next_number:02d}")
            messagebox.showinfo(
                "Counter saved",
                f"The next {department_var.get()} code will end in "
                f"{next_number:02d}.",
                parent=dialog,
            )

        def reset_all_counters() -> None:
            confirmed = messagebox.askyesno(
                "Reset all counters?",
                "The next code for every department will start at 01. Continue?",
                parent=dialog,
            )
            if confirmed:
                self._save_counters({})
                refresh_next_counter()

        department_var.trace_add("write", refresh_next_counter)
        refresh_next_counter()

        actions = tk.Frame(dialog, bg=self.SURFACE)
        actions.pack(fill="x", padx=30, pady=(22, 0))

        tk.Button(
            actions,
            text="Reset All",
            command=reset_all_counters,
            bg="#FEE2E2",
            fg="#B42318",
            activebackground="#FECACA",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=16,
            pady=8,
        ).pack(side="left")

        tk.Button(
            actions,
            text="Save",
            command=save_next_counter,
            bg=self.PRIMARY,
            fg="white",
            activebackground=self.PRIMARY_HOVER,
            activeforeground="white",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=20,
            pady=8,
        ).pack(side="right")

        counter_entry.focus_set()
        counter_entry.selection_range(0, tk.END)
        dialog.bind("<Return>", lambda _event: save_next_counter())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _show_code_dialog(self, department: str) -> None:
        dialog = tk.Toplevel(self)
        dialog.title(f"ITSRR Template - {department}")
        dialog.geometry("650x660")
        dialog.minsize(520, 540)
        dialog.resizable(True, True)
        dialog.configure(bg=self.SURFACE)
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(
            dialog,
            text="ITSRR Client Template",
            bg=self.SURFACE,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 20),
        ).pack(anchor="w", padx=34, pady=(24, 3))

        tk.Label(
            dialog,
            text="Copy this message and send it to the client to complete.",
            bg=self.SURFACE,
            fg=self.MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=34, pady=(0, 18))

        template = (
            f"ITSRR Number: {self.generated_code}\n\n"
            "Terminal User: \n\n"
            "Date and Time Reported: \n\n"
            "Reported By: \n\n"
            "Problem: \n\n"
            "Date and Time Resolved: \n\n"
            "Cause: \n\n"
            "Solution: \n\n"
            "Status: "
        )

        template_text = tk.Text(
            dialog,
            wrap="word",
            bg="#FFFFFF",
            fg=self.TEXT,
            insertbackground=self.TEXT,
            selectbackground=self.PRIMARY,
            selectforeground="white",
            font=("Segoe UI", 11),
            relief="solid",
            bd=1,
            padx=16,
            pady=14,
        )
        template_text.pack(fill="both", expand=True, padx=34)
        template_text.insert("1.0", template)
        template_text.focus_set()
        template_text.tag_add("itsrr_number", "1.0", "1.end")
        template_text.tag_configure(
            "itsrr_number",
            foreground="#0B3D91",
            font=("Segoe UI Semibold", 11),
        )

        actions = tk.Frame(dialog, bg=self.SURFACE)
        actions.pack(
            side="bottom",
            fill="x",
            padx=34,
            pady=(16, 24),
            before=template_text,
        )

        copy_button = tk.Button(
            actions,
            text="Copy",
            command=lambda: self._copy_from_dialog(copy_button, template_text),
            bg=self.PRIMARY,
            fg="white",
            activebackground=self.PRIMARY_HOVER,
            activeforeground="white",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=22,
            pady=9,
        )
        copy_button.pack(side="right")

        tk.Button(
            actions,
            text="Close",
            command=dialog.destroy,
            bg="#E5E7EB",
            fg=self.TEXT,
            activebackground="#D1D5DB",
            relief="solid",
            bd=1,
            font=("Segoe UI Semibold", 10),
            padx=22,
            pady=9,
        ).pack(side="right", padx=(0, 10))

        dialog.bind(
            "<Control-Return>",
            lambda _event: self._copy_from_dialog(copy_button, template_text),
        )
        dialog.bind("<Escape>", lambda _event: dialog.destroy())

    def _copy_from_dialog(self, button: tk.Button, template_text: tk.Text) -> None:
        template = template_text.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(template)
        self.update()
        button.configure(text="Copied!")

if __name__ == "__main__":
    DepartmentSelector().mainloop()
