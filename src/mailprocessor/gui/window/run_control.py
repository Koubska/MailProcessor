"""Running the pipeline on a worker thread, with progress, stopping and "Alles neu exportieren"."""

from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from mailprocessor.config import ParsingRules, load_app_config
from mailprocessor.gui.messages import friendly_error, run_summary_text
from mailprocessor.gui.window.widgets import AMBER, FIELDS_TAB, GRAY, GREEN, RED, START_TAB
from mailprocessor.logfile import attach_log_file
from mailprocessor.processor import run_pipeline, start_over


class RunControl:
    """Starts runs off the UI thread and shows their progress and result (via `run_results`)."""

    def _build_progress_row(self, tab: ttk.Frame) -> None:
        """Shown only while a run is going on."""
        self.progress_row = ttk.Frame(tab)
        self.progress_bar = ttk.Progressbar(self.progress_row, mode="determinate", length=self.px(320))
        self.progress_bar.pack(side=tk.LEFT)
        self.progress_label = ttk.Label(self.progress_row)
        self.progress_label.pack(side=tk.LEFT, padx=(12, 0))
        self.stop_button = self.button(self.progress_row, "button.stop", self.stop_run, tip="tip.stop")
        self.stop_button.pack(side=tk.LEFT, padx=(12, 0))

    def set_busy(self, busy: bool) -> None:
        self.running = busy
        state = tk.DISABLED if busy else tk.NORMAL
        for button in (self.run_button, self.test_run_button, self.start_over_button):
            button.configure(state=state)
        if busy:
            self.progress_bar.configure(value=0, maximum=1)
            self.progress_label.configure(text=self.tr("label.status.running"))
            self.stop_button.configure(state=tk.NORMAL)
            self.progress_row.pack(fill=tk.X, pady=(16, 0), before=self.result_label)
        else:
            self.progress_row.pack_forget()

    def execute_run(self, dry_run: bool) -> None:
        if self.running:
            return
        self.show_problems(())
        if not self.flush_saves():
            self.show_error(self.tr("error.fix_inputs"))
            self.jump_to_first_error()
            return
        if not self.profiles.field_count:
            self.show_error(self.tr("card.fields.none"))
            self.notebook.select(FIELDS_TAB)
            return
        try:
            app_config = load_app_config(self.cfg_path)
            parsing_rules = self.profiles.parsing_rules()
        except (OSError, ValueError) as exc:
            self.show_error(friendly_error(exc, self.lang))
            return
        if app_config.source.imap is not None:
            app_config.source.imap.password = self.password.get() or None
        app_config.app.dry_run = dry_run
        self.package_logger.setLevel(app_config.app.log_level)
        attach_log_file(self.config_dir)
        self.notebook.select(START_TAB)
        self.append_log(f"--- {time.strftime('%H:%M:%S')} ---")
        self.show_result("", GRAY)
        self.cancel_event.clear()
        self.set_busy(True)
        threading.Thread(target=self._run_worker, args=(app_config, parsing_rules), daemon=True).start()
        self.root.after(100, self.poll_run)

    def _run_worker(self, app_config, parsing_rules: ParsingRules) -> None:
        # Runs off the UI thread; must not touch tkinter objects.
        def report_progress(current: int, total: int) -> None:
            self.run_results.put(("progress", (current, total)))

        try:
            summary = run_pipeline(app_config, parsing_rules, progress=report_progress, cancel=self.cancel_event)
            self.run_results.put(("ok", (app_config, summary)))
        except Exception as exc:
            if not isinstance(exc, (ValueError, OSError)):
                self.package_logger.debug("Unexpected error", exc_info=True)
            self.run_results.put(("error", exc))

    def poll_run(self) -> None:
        # Log lines and progress first; the run's result ("ok"/"error") is always the last item.
        while True:
            try:
                kind, payload = self.run_results.get_nowait()
            except queue.Empty:
                self.root.after(100, self.poll_run)
                return
            if kind == "log":
                self.append_log(payload)
            elif kind == "progress":
                current, total = payload
                self.progress_bar.configure(maximum=max(total, 1), value=current)
                if not self.cancel_event.is_set():
                    self.progress_label.configure(
                        text=self.tr("label.status.progress").format(current=current, total=total)
                    )
            else:
                break
        self.set_busy(False)
        if kind == "ok":
            app_config, summary = payload
            text = run_summary_text(
                summary,
                Path(app_config.app.output_xlsx).name,
                app_config.app.sheet_errors,
                app_config.app.dry_run,
                self.lang,
            )
            self.append_log(text)
            self.show_result(text, AMBER if summary.failed or summary.cancelled else GREEN)
            self.show_problems(summary.problems)
        else:
            message = friendly_error(payload, self.lang)
            self.append_log(f"ERROR   {message}")
            self.show_result(message, RED)
            self.show_error(message)
        self.refresh_cards()

    def stop_run(self) -> None:
        self.cancel_event.set()
        self.stop_button.configure(state=tk.DISABLED)
        self.progress_label.configure(text=self.tr("label.status.stopping"))

    def on_start_over(self) -> None:
        if self.running or not self.confirm(self.tr("confirm.start_over"), icon="warning"):
            return
        if not self.flush_saves():
            self.show_error(self.tr("error.fix_inputs"))
            self.jump_to_first_error()
            return
        try:
            app_config = load_app_config(self.cfg_path)
            attach_log_file(self.config_dir)
            backup = start_over(app_config)
        except (OSError, ValueError) as exc:
            self.show_error(friendly_error(exc, self.lang))
            return
        if backup is not None:
            self.append_log(self.tr("info.backup_created").format(file=backup.name))
        self.execute_run(dry_run=False)
