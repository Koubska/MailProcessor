"""The E-Mails tab: the source (.eml folder or IMAP mailbox) and the filter."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from mailprocessor.config import load_app_config
from mailprocessor.gui.config_files import path_setting, setting_path
from mailprocessor.gui.form import DEFAULT_IMAP_PORTS
from mailprocessor.gui.messages import friendly_error
from mailprocessor.gui.status_cards import mail_count_in_folder
from mailprocessor.gui.window.widgets import AMBER, GRAY, GREEN, RED
from mailprocessor.sources.imap_source import check_imap_connection


class MailsTab:
    """Where the mails come from, and which of them count."""

    def _build_mails_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "mails.title", style="Title.TLabel").pack(anchor="w")
        self.hint(tab, "mails.intro", wrap=760).pack(anchor="w", pady=(4, 12))
        source = self.var("source_type", "eml")
        self._build_eml_section(tab, source)
        self._build_imap_section(tab, source)
        self.password.trace_add("write", lambda *_args: self.refresh_cards())
        self._build_filter_section(tab)

    def _build_eml_section(self, tab: ttk.Frame, source: tk.Variable) -> None:
        eml_choice = ttk.Radiobutton(tab, variable=source, value="eml", style="Heading.TRadiobutton")
        self._translate(eml_choice, "mails.eml.choice").pack(anchor="w")
        self.eml_frame = ttk.Frame(tab, padding=(28, 4, 0, 12))
        self.eml_frame.pack(fill=tk.X)
        self.hint(self.eml_frame, "mails.eml.hint").grid(row=0, column=0, columnspan=4, sticky="w")
        self.var("eml_glob", "*.eml")  # not shown; kept so a custom pattern in config.toml survives
        self.var("eml_folder", "")
        self.entry(self.eml_frame, "eml_folder", width=50).grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.button(self.eml_frame, "button.browse", self.choose_eml_folder).grid(
            row=1, column=1, padx=(8, 0), pady=(6, 0)
        )
        self.button(self.eml_frame, "button.open_folder", self.open_eml_folder, tip="tip.open_folder").grid(
            row=1, column=2, padx=(8, 0), pady=(6, 0)
        )
        self.eml_count = ttk.Label(self.eml_frame)
        self.eml_count.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))
        self.error_label(self.eml_frame, "eml_folder").grid(row=3, column=0, columnspan=4, sticky="w")
        self.eml_frame.columnconfigure(0, weight=1)

    def _build_imap_section(self, tab: ttk.Frame, source: tk.Variable) -> None:
        imap_choice = ttk.Radiobutton(tab, variable=source, value="imap", style="Heading.TRadiobutton")
        self._translate(imap_choice, "mails.imap.choice").pack(anchor="w", pady=(8, 0))
        self.imap_frame = ttk.Frame(tab, padding=(28, 4, 0, 0))
        self.imap_frame.pack(fill=tk.X)
        self.hint(self.imap_frame, "mails.imap.hint").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        rows = (
            ("imap_host", "mails.imap.host", "mails.imap.host_hint"),
            ("imap_username", "mails.imap.username", "mails.imap.username_hint"),
            (None, "mails.imap.password", "mails.imap.password_hint"),
            ("imap_mailbox", "mails.imap.mailbox", "mails.imap.mailbox_hint"),
        )
        for key, default in (("imap_host", ""), ("imap_username", ""), ("imap_mailbox", "INBOX")):
            self.var(key, default)
        self.var("imap_port", "993")
        self.var("imap_use_ssl", True)
        for row, (key, label_key, hint_key) in enumerate(rows, start=1):
            self.label(self.imap_frame, label_key).grid(
                row=row * 2 - 1, column=0, sticky="w", padx=(0, 12), pady=(4, 0)
            )
            if key is None:
                self.password_entry = ttk.Entry(self.imap_frame, textvariable=self.password, show="•", width=40)
                widget = self.password_entry
            else:
                widget = self.entry(self.imap_frame, key)
            widget.grid(row=row * 2 - 1, column=1, sticky="ew", pady=(4, 0))
            self.hint(self.imap_frame, hint_key, wrap=440).grid(row=row * 2 - 1, column=2, sticky="w", padx=(12, 0))
            if key is not None:
                self.error_label(self.imap_frame, key).grid(row=row * 2, column=1, columnspan=2, sticky="w")
        security = ttk.Frame(self.imap_frame)
        security.grid(row=12, column=1, columnspan=2, sticky="w", pady=(8, 0))
        ssl_box = ttk.Checkbutton(security, variable=self.form["imap_use_ssl"], command=self.on_ssl_toggled)
        self._translate(ssl_box, "mails.imap.ssl").pack(side=tk.LEFT)
        self.tip(ssl_box, "tip.ssl")
        self.label(security, "mails.imap.port").pack(side=tk.LEFT, padx=(16, 6))
        self.entry(security, "imap_port", width=7).pack(side=tk.LEFT)
        self.error_label(self.imap_frame, "imap_port").grid(row=13, column=1, columnspan=2, sticky="w")
        test_row = ttk.Frame(self.imap_frame)
        test_row.grid(row=14, column=1, columnspan=2, sticky="w", pady=(12, 0))
        self.connection_button = self.button(
            test_row, "button.test_connection", self.test_connection, tip="tip.test_connection"
        )
        self.connection_button.pack(side=tk.LEFT)
        self.connection_result = ttk.Label(test_row, wraplength=self.px(560), justify=tk.LEFT)
        self.connection_result.pack(side=tk.LEFT, padx=(12, 0))
        self.imap_frame.columnconfigure(1, weight=1)

    def _build_filter_section(self, tab: ttk.Frame) -> None:
        """Which mails count, for both sources: others are left out instead of landing on the error sheet."""
        self.label(tab, "mails.filter.title", style="Heading.TLabel").pack(anchor="w", pady=(16, 0))
        self.hint(tab, "mails.filter.hint", wrap=760).pack(anchor="w", pady=(2, 4))
        frame = ttk.Frame(tab, padding=(28, 0, 0, 0))
        frame.pack(fill=tk.X)
        for row, (key, label_key, hint_key) in enumerate(
            (
                ("filter_subject", "mails.filter.subject", "mails.filter.subject_hint"),
                ("filter_sender", "mails.filter.sender", "mails.filter.sender_hint"),
            )
        ):
            self.var(key, "")
            self.label(frame, label_key).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=(4, 0))
            self.entry(frame, key).grid(row=row, column=1, sticky="ew", pady=(4, 0))
            self.hint(frame, hint_key, wrap=440).grid(row=row, column=2, sticky="w", padx=(12, 0), pady=(4, 0))
        frame.columnconfigure(1, weight=1)

    def on_source_changed(self) -> None:
        use_imap = self.form["source_type"].get() == "imap"
        self._set_enabled(self.eml_frame, not use_imap)
        self._set_enabled(self.imap_frame, use_imap)

    def update_mail_count(self) -> None:
        count = mail_count_in_folder(self.values(), self.config_dir)
        if count is None:
            self.eml_count.configure(text=self.tr("mails.eml.missing"), foreground=AMBER)
        else:
            self.eml_count.configure(
                text=self.tr("mails.eml.count").format(count=count), foreground=GREEN if count else AMBER
            )

    def _set_enabled(self, frame: tk.Widget, enabled: bool) -> None:
        for child in frame.winfo_children():
            if isinstance(child, (ttk.Entry, ttk.Button, ttk.Checkbutton)):
                child.state(["!disabled"] if enabled else ["disabled"])
            elif isinstance(child, ttk.Frame):
                self._set_enabled(child, enabled)

    def on_ssl_toggled(self) -> None:
        # Follow the default port of the chosen security, unless the user typed a custom one.
        use_ssl = bool(self.form["imap_use_ssl"].get())
        if self.form["imap_port"].get().strip() in ("", DEFAULT_IMAP_PORTS[not use_ssl]):
            self.form["imap_port"].set(DEFAULT_IMAP_PORTS[use_ssl])

    def test_connection(self) -> None:
        if not self.save_config():
            self.connection_result.configure(text=self.tr("error.fix_inputs"), foreground=RED)
            return
        config = load_app_config(self.cfg_path).source.imap
        if config is None:
            return
        config.password = self.password.get() or None
        self.connection_button.state(["disabled"])
        self.connection_result.configure(text=self.tr("mails.imap.testing"), foreground=GRAY)
        results: queue.Queue = queue.Queue()

        def worker() -> None:
            try:
                results.put(("ok", check_imap_connection(config)))
            except Exception as exc:
                results.put(("error", exc))

        def poll() -> None:
            try:
                kind, payload = results.get_nowait()
            except queue.Empty:
                self.root.after(150, poll)
                return
            self.connection_button.state(["!disabled"])
            if kind == "ok":
                key = "mails.imap.connected" if payload is None else "mails.imap.connected_count"
                text = self.tr(key).format(count=payload, mailbox=config.mailbox)
                self.connection_result.configure(text=text, foreground=GREEN)
            else:
                self.connection_result.configure(
                    text=friendly_error(payload, self.lang).split("\n\n")[0], foreground=RED
                )

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(150, poll)

    def eml_folder(self) -> Path:
        return setting_path(str(self.form["eml_folder"].get()).strip() or ".", self.config_dir)

    def open_eml_folder(self) -> None:
        self.open_path(self.eml_folder(), "info.folder_missing")

    def choose_eml_folder(self) -> None:
        current = self.eml_folder()
        chosen = filedialog.askdirectory(
            parent=self.root,
            title=self.tr("dialog.choose_eml_folder"),
            initialdir=str(current if current.is_dir() else self.config_dir),
            mustexist=True,
        )
        if chosen:
            self.form["eml_folder"].set(path_setting(Path(chosen), self.config_dir))
            self.refresh_sample_files()
