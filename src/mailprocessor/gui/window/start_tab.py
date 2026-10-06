"""The Start tab: status cards, the run buttons, the result and the problems of the last run."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import ttk

from mailprocessor.gui.config_files import output_file_path
from mailprocessor.gui.messages import problem_text
from mailprocessor.gui.preview import SampleMail
from mailprocessor.gui.status_cards import CardStatus, excel_status, fields_status, mails_status, profiles_status
from mailprocessor.gui.window.widgets import AMBER, FIELDS_TAB, GRAY, GREEN, MAILS_TAB, START_TAB
from mailprocessor.logfile import log_file_path
from mailprocessor.parser import normalize_body
from mailprocessor.run_summary import Problem


class StartTab:
    """Status cards, run buttons, result, problem list and log details."""

    def _build_start_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "start.title", style="Title.TLabel").pack(anchor="w")
        self.hint(tab, "start.intro", wrap=760).pack(anchor="w", pady=(4, 16))
        self._build_cards(tab)
        ttk.Separator(tab).pack(fill=tk.X, pady=16)
        self._build_run_buttons(tab)
        self._build_progress_row(tab)
        self.result_label = ttk.Label(tab, wraplength=self.px(900), justify=tk.LEFT, font=self.bold_font)
        self.result_label.pack(anchor="w", pady=(16, 0))
        self._build_problem_list(tab)
        self._build_details(tab)

    def _build_cards(self, tab: ttk.Frame) -> None:
        cards = ttk.Frame(tab)
        cards.pack(fill=tk.X)
        cards.columnconfigure(2, weight=1)
        self.cards: dict[str, tuple[ttk.Label, ttk.Label]] = {}

        def show(tab: int) -> Callable[[], None]:
            return lambda: self.notebook.select(tab)

        card_rows = (
            ("mails", "card.mails.title", "button.change", show(MAILS_TAB), "tip.card_mails", None),
            ("fields", "card.fields.title", "button.edit", show(FIELDS_TAB), "tip.card_fields", None),
            ("excel", "card.excel.title", "button.open_excel", self.open_output, "tip.open_excel", "open_excel"),
        )
        for row, (name, title_key, action_key, action, tip, shortcut) in enumerate(card_rows):
            icon = ttk.Label(cards, style="CardIcon.TLabel", width=2)
            icon.grid(row=row, column=0, sticky="w", pady=6)
            self.label(cards, title_key, style="Bold.TLabel").grid(row=row, column=1, sticky="w", padx=(4, 16))
            text = ttk.Label(cards, wraplength=self.px(640), justify=tk.LEFT)
            text.grid(row=row, column=2, sticky="w")
            self.button(cards, action_key, action, tip=tip, shortcut=shortcut).grid(
                row=row, column=3, sticky="e", padx=(12, 0)
            )
            self.cards[name] = (icon, text)

    def _build_run_buttons(self, tab: ttk.Frame) -> None:
        actions = ttk.Frame(tab)
        actions.pack(fill=tk.X)
        self.run_button = self.button(
            actions,
            "button.run",
            lambda: self.execute_run(dry_run=False),
            style="Big.TButton",
            tip="tip.run",
            shortcut="run",
        )
        self.run_button.pack(side=tk.LEFT)
        self.test_run_button = self.button(
            actions, "button.test_run", lambda: self.execute_run(dry_run=True), tip="tip.test_run", shortcut="test_run"
        )
        self.test_run_button.pack(side=tk.LEFT, padx=(12, 0))
        self.hint(tab, "start.test_run_hint", wrap=760).pack(anchor="w", pady=(6, 0))

    def _build_problem_list(self, tab: ttk.Frame) -> None:
        """Mails that failed in the last run, each with a direct way to the rule that did not match."""
        self.problems_frame = ttk.Frame(tab)
        self.label(self.problems_frame, "problems.title", style="Heading.TLabel").pack(anchor="w")
        self.hint(self.problems_frame, "problems.hint", wrap=900).pack(anchor="w", pady=(2, 6))
        problem_list = ttk.Frame(self.problems_frame)
        problem_list.pack(fill=tk.X)
        self.problems_view = ttk.Treeview(
            problem_list, columns=("mail", "problem"), show="headings", height=5, selectmode="browse"
        )
        self.problems_view.column("mail", width=self.px(320), anchor="w")
        self.problems_view.column("problem", width=self.px(520), anchor="w")
        problems_scroll = ttk.Scrollbar(problem_list, orient=tk.VERTICAL, command=self.problems_view.yview)
        self.problems_view.configure(yscrollcommand=problems_scroll.set)
        self.problems_view.pack(side=tk.LEFT, fill=tk.X, expand=True)
        problems_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.problems_view.bind("<Double-1>", lambda _event: self.open_problem())
        self.problems_view.bind("<Return>", lambda _event: self.open_problem())
        self.problems_view.bind("<<TreeviewSelect>>", lambda _event: self._update_problem_button())
        self.open_problem_button = self.button(
            self.problems_frame, "button.open_problem", self.open_problem, tip="tip.open_problem"
        )
        self.open_problem_button.pack(anchor="w", pady=(6, 0))

    def _build_details(self, tab: ttk.Frame) -> None:
        details_row = ttk.Frame(tab)
        details_row.pack(fill=tk.X, pady=(12, 0))
        self.details_row = details_row
        self.details_button = ttk.Button(details_row, command=self.toggle_details)
        self.tip(self.details_button, "tip.details")
        self.details_button.pack(side=tk.LEFT)
        self.button(details_row, "button.open_log", self.open_log, tip="tip.open_log").pack(side=tk.LEFT, padx=(8, 0))
        self.details_frame = ttk.Frame(tab)
        self.log_text = tk.Text(self.details_frame, height=10, wrap="word", state=tk.DISABLED)
        log_scroll = ttk.Scrollbar(self.details_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.details_visible = False
        self._update_details_button()

    def on_tab_changed(self) -> None:
        if self.notebook.index("current") == START_TAB:
            self.refresh_cards()

    def _set_card(self, name: str, status: CardStatus) -> None:
        icon, text = self.cards[name]
        symbol, color = {True: ("✓", GREEN), False: ("⚠", AMBER), None: ("•", GRAY)}[status.ok]
        icon.configure(text=symbol, foreground=color)
        text.configure(text=status.text)

    def refresh_cards(self) -> None:
        values = self.values()
        self._set_card("mails", mails_status(values, self.config_dir, bool(self.password.get()), self.lang))
        self._set_fields_card()
        excel_path = output_file_path(str(values.get("output_xlsx", "")), self.config_dir)
        if values.get("profile_sheets") == "per_profile":
            data_sheets = [profile.name for profile in self.profiles.profiles()]
        else:
            data_sheets = [str(values.get("sheet_data") or "daten")]
        error_sheet = str(values.get("sheet_errors") or "fehler")
        self._set_card("excel", excel_status(excel_path, data_sheets, error_sheet, self.lang))

    def _set_fields_card(self) -> None:
        if len(self.profiles) > 1:
            status = profiles_status(len(self.profiles), self.profiles.field_count, self.best, self.lang)
        else:
            status = fields_status(self.fields, self.previews, self.lang)
        self._set_card("fields", status)

    def toggle_details(self) -> None:
        self.details_visible = not self.details_visible
        if self.details_visible:
            self.details_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        else:
            self.details_frame.pack_forget()
        self._update_details_button()

    def _update_details_button(self) -> None:
        self.details_button.configure(
            text=self.tr("button.details_hide" if self.details_visible else "button.details_show")
        )

    def append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def show_result(self, text: str, color: str) -> None:
        self.result_label.configure(text=text, foreground=color)

    def show_problems(self, problems: tuple[Problem, ...]) -> None:
        self.problems = problems
        self.problems_view.delete(*self.problems_view.get_children())
        for index, problem in enumerate(problems):
            self.problems_view.insert(
                "", tk.END, iid=str(index), values=(problem.name, problem_text(problem, self.lang))
            )
        if problems:
            self.problems_frame.pack(fill=tk.X, pady=(12, 0), before=self.details_row)
            self.problems_view.selection_set("0")
        else:
            self.problems_frame.pack_forget()
        self._update_problem_button()

    def _update_problem_button(self) -> None:
        selection = self.problems_view.selection()
        readable = bool(selection) and self.problems[int(selection[0])].body is not None
        self.open_problem_button.state(["!disabled"] if readable else ["disabled"])

    def open_problem(self) -> None:
        """Show the failed mail in the Felder tab and select the first field that was not found."""
        selection = self.problems_view.selection()
        if not selection:
            return
        problem = self.problems[int(selection[0])]
        if problem.body is None:
            self.show_info(self.tr("problems.unreadable").format(reason=problem.reason))
            return
        self.show_sample(
            SampleMail(title=problem.name, body=normalize_body(problem.body), header_text=problem.header_text)
        )
        self.notebook.select(FIELDS_TAB)
        profile_index = self.profiles.find(problem.profile) if problem.profile else None
        if profile_index is not None and profile_index != self.profiles.index:
            self.show_profile(profile_index)
        columns = self.fields.columns()
        missing = [column for column in problem.missing if column in columns]
        if missing:
            self.flush_editor()
            self.select_field(columns.index(missing[0]))
        self.set_status(self.tr("problems.opened").format(name=problem.name, columns=", ".join(problem.missing)))

    def open_output(self) -> None:
        self.open_path(
            output_file_path(str(self.form["output_xlsx"].get()).strip(), self.config_dir), "info.no_excel_yet"
        )

    def open_log(self) -> None:
        self.open_path(log_file_path(self.config_dir), "info.no_log_yet")
