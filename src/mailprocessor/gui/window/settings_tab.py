"""The Einstellungen tab: the Excel file, language, limits, maintenance and technical settings."""

from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from mailprocessor.gui.config_files import output_file_path, path_setting
from mailprocessor.gui.shortcuts import shortcuts

LANGUAGES = {"de": "Deutsch", "en": "English"}


class SettingsTab:
    """Settings that are rarely changed."""

    def _build_settings_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "settings.title", style="Title.TLabel").pack(anchor="w", pady=(0, 12))
        columns = ttk.Frame(tab)
        columns.pack(fill=tk.BOTH, expand=True)
        left, right = ttk.Frame(columns), ttk.Frame(columns)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 12))
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))

        self._build_excel_section(left)

        language = self._section(left, "settings.language")
        self.language_box = ttk.Combobox(language, state="readonly", values=list(LANGUAGES.values()), width=20)
        self.language_box.set(LANGUAGES[self.lang])
        self.language_box.grid(row=0, column=0, sticky="w")
        self.language_box.bind("<<ComboboxSelected>>", lambda _event: self.on_language_selected())

        limits = self._section(left, "settings.limits")
        self._setting_row(limits, 0, "max_age_days", "settings.max_age_days", "0", width=8)
        self._setting_row(limits, 2, "max_messages", "settings.max_messages", "0", width=8)

        maintenance = self._section(right, "settings.maintenance")
        self.start_over_button = self.button(maintenance, "button.start_over", self.on_start_over)
        self.start_over_button.grid(row=0, column=0, sticky="w")
        self.hint(maintenance, "settings.start_over_hint", wrap=420).grid(row=1, column=0, sticky="w", pady=(2, 10))
        self.button(maintenance, "button.open_log", self.open_log, tip="tip.open_log").grid(row=2, column=0, sticky="w")
        self.button(maintenance, "button.open_config_folder", lambda: self.open_path(self.config_dir)).grid(
            row=3, column=0, sticky="w", pady=(6, 0)
        )

        technical = self._section(right, "settings.technical")
        self.var("log_level", "INFO")
        self.label(technical, "label.log_level").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Combobox(
            technical,
            textvariable=self.form["log_level"],
            values=["DEBUG", "INFO", "WARNING", "ERROR"],
            state="readonly",
            width=12,
        ).grid(row=0, column=1, sticky="w")
        self._setting_row(technical, 1, "sqlite_path", "label.sqlite_path", "")
        self.files_label = ttk.Label(technical, style="Hint.TLabel", wraplength=self.px(440), justify=tk.LEFT)
        self.files_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        self.var("dry_run", False)  # not shown: "Testlauf" decides per run; kept so the file value survives

    def _build_excel_section(self, parent: ttk.Frame) -> None:
        excel = self._section(parent, "settings.excel")
        self.var("output_xlsx", "")
        file_row = ttk.Frame(excel)
        file_row.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.entry(file_row, "output_xlsx").pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.button(file_row, "button.browse", self.choose_output_file).pack(side=tk.LEFT, padx=(8, 0))
        self.error_label(excel, "output_xlsx").grid(row=1, column=0, columnspan=2, sticky="w")
        self.hint(excel, "settings.excel_hint", wrap=440).grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 8))
        self._setting_row(excel, 3, "sheet_data", "settings.sheet_data", "daten")
        self._setting_row(excel, 5, "sheet_errors", "settings.sheet_errors", "fehler")
        self.var("profile_sheets", "shared")
        self.label(excel, "settings.profile_sheets").grid(row=7, column=0, sticky="nw", padx=(0, 8), pady=(8, 0))
        layout = ttk.Frame(excel)
        layout.grid(row=7, column=1, sticky="w", pady=(8, 0))
        for value, key in (
            ("shared", "settings.profile_sheets_shared"),
            ("per_profile", "settings.profile_sheets_per_profile"),
        ):
            choice = ttk.Radiobutton(layout, variable=self.form["profile_sheets"], value=value)
            self._translate(choice, key).pack(anchor="w")
        self.hint(excel, "settings.profile_sheets_hint", wrap=440).grid(
            row=8, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

    def choose_output_file(self) -> None:
        current = output_file_path(str(self.form["output_xlsx"].get()).strip(), self.config_dir)
        chosen = filedialog.asksaveasfilename(
            parent=self.root,
            title=self.tr("dialog.choose_output_xlsx"),
            initialdir=str(current.parent if current.parent.is_dir() else self.config_dir),
            initialfile=current.name,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            confirmoverwrite=False,  # new rows are appended; the file is never replaced
        )
        if chosen:
            self.form["output_xlsx"].set(path_setting(Path(chosen), self.config_dir))

    def on_language_selected(self) -> None:
        names = {name: code for code, name in LANGUAGES.items()}
        self.lang = names.get(self.language_box.get(), "de")
        self.shortcuts = shortcuts(sys.platform, self.lang)
        self.apply_language()
