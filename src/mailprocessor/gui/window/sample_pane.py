"""The sample mail next to the field list, and the live preview of every rule on it."""

from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, ttk

from mailprocessor.gui.preview import (
    SampleMail,
    best_profile,
    load_sample,
    preview_rule,
    profile_summary_text,
    result_text,
    sample_files,
    summary_text,
)
from mailprocessor.gui.window.widgets import GRAY, GREEN, HIGHLIGHT, RED
from mailprocessor.parser import normalize_body


@dataclass
class SampleState:
    """The sample mail shown; its body is the text widget's content, so it can be edited."""

    files: list[Path] = field(default_factory=list)  # mails of the configured folder, to step through
    index: int = -1  # position of the shown mail in `files`; -1 if it is not one of them
    title: str = ""
    header_text: str = ""


class SamplePane:
    """Shows a sample mail and what each rule finds in it, updated while typing."""

    def _build_sample_pane(self, pane: ttk.Frame) -> None:
        self.label(pane, "preview.title", style="Heading.TLabel").pack(anchor="w")
        nav = ttk.Frame(pane)
        nav.pack(fill=tk.X, pady=(6, 0))
        self.prev_button = ttk.Button(nav, text="◀", width=3, command=lambda: self.step_sample(-1))
        self.tip(self.prev_button, "tip.sample_previous")
        self.prev_button.pack(side=tk.LEFT)
        self.next_button = ttk.Button(nav, text="▶", width=3, command=lambda: self.step_sample(1))
        self.tip(self.next_button, "tip.sample_next")
        self.next_button.pack(side=tk.LEFT, padx=(4, 0))
        self.sample_title = ttk.Label(nav)
        self.sample_title.pack(side=tk.LEFT, padx=(8, 0))
        sample_actions = ttk.Frame(pane)
        sample_actions.pack(fill=tk.X, pady=(6, 6))
        self.button(sample_actions, "button.sample_load", self.choose_sample_file, tip="tip.sample_load").pack(
            side=tk.LEFT
        )
        self.button(sample_actions, "button.sample_paste", self.paste_sample, tip="tip.sample_paste").pack(
            side=tk.LEFT, padx=(8, 0)
        )
        self.sample_summary = ttk.Label(pane, wraplength=self.px(440), justify=tk.LEFT, font=self.bold_font)
        self.sample_summary.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        self.hint(pane, "preview.header_note", wrap=440).pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        text_frame = ttk.Frame(pane)
        text_frame.pack(fill=tk.BOTH, expand=True)
        self.sample_text = tk.Text(text_frame, wrap="word", height=12, undo=True)
        sample_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.sample_text.yview)
        self.sample_text.configure(yscrollcommand=sample_scroll.set)
        self.sample_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sample_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.sample_text.tag_configure("match", background=HIGHLIGHT, foreground="black")
        self.sample_text.bind("<<Modified>>", lambda _event: self.on_sample_edited())

    def current_sample(self) -> SampleMail | None:
        body = self.sample_text.get("1.0", "end-1c")
        if not body.strip() and not self.sample.header_text:
            return None
        return SampleMail(title=self.sample.title, body=body, header_text=self.sample.header_text)

    def update_sample_title(self) -> None:
        files, index = self.sample.files, self.sample.index
        if 0 <= index < len(files):
            text = self.tr("preview.position").format(name=self.sample.title, index=index + 1, total=len(files))
        else:
            text = self.sample.title or self.tr("preview.none")
        self.sample_title.configure(text=text)
        self.prev_button.state(["!disabled"] if index > 0 else ["disabled"])
        self.next_button.state(["!disabled"] if len(files) > 1 and index < len(files) - 1 else ["disabled"])

    def show_sample(self, sample: SampleMail, index: int = -1) -> None:
        self.sample.title, self.sample.header_text, self.sample.index = sample.title, sample.header_text, index
        self.sample_text.delete("1.0", tk.END)
        self.sample_text.insert("1.0", sample.body)
        self.sample_text.edit_reset()
        self.update_sample_title()
        self.schedule_preview()

    def show_sample_file(self, path: Path, index: int = -1) -> None:
        try:
            self.show_sample(load_sample(path), index)
        except OSError as exc:
            self.show_error(self.tr("error.sample_unreadable").format(file=path, error=exc))

    def refresh_sample_files(self) -> None:
        """Offer the mails of the configured folder; show the first one if nothing is shown yet."""
        old_files, old_index = self.sample.files, self.sample.index
        current = old_files[old_index] if 0 <= old_index < len(old_files) else None
        files = sample_files(self.eml_folder(), str(self.form["eml_glob"].get()).strip() or "*.eml")
        self.sample.files = files
        if current in files:
            self.sample.index = files.index(current)
        elif files and self.current_sample() is None:
            self.show_sample_file(files[0], 0)
            return
        else:
            self.sample.index = -1
        self.update_sample_title()
        self.schedule_preview()

    def step_sample(self, offset: int) -> None:
        files, index = self.sample.files, self.sample.index
        new_index = index + offset if index >= 0 else 0
        if 0 <= new_index < len(files):
            self.show_sample_file(files[new_index], new_index)

    def choose_sample_file(self) -> None:
        folder = self.eml_folder()
        chosen = filedialog.askopenfilename(
            parent=self.root,
            title=self.tr("dialog.choose_sample"),
            initialdir=str(folder if folder.is_dir() else self.config_dir),
            filetypes=[(self.tr("dialog.eml_files"), "*.eml"), (self.tr("dialog.all_files"), "*")],
        )
        if chosen:
            path = Path(chosen)
            files = self.sample.files
            self.show_sample_file(path, files.index(path) if path in files else -1)

    def paste_sample(self) -> None:
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""
        if not text.strip():
            self.show_info(self.tr("error.clipboard_empty"))
            return
        self.show_sample(SampleMail(title=self.tr("preview.pasted"), body=normalize_body(text)))

    def on_sample_edited(self) -> None:
        if self.sample_text.edit_modified():
            self.sample_text.edit_modified(False)
            self.schedule_preview()

    def schedule_preview(self) -> None:
        # Typing triggers many updates; compute once the input settles.
        if self.preview_job is not None:
            self.root.after_cancel(self.preview_job)
        self.preview_job = self.root.after(120, self.update_preview)

    def update_preview(self) -> None:
        self.preview_job = None
        sample = self.current_sample()
        self.previews = [preview_rule(rule, sample) for rule in self.fields] if sample else None
        for index, rule in enumerate(self.fields):
            iid = str(index)
            if not self.fields_view.exists(iid):
                continue
            preview = self.previews[index] if self.previews else None
            self.fields_view.set(iid, "result", result_text(preview, self.lang))
            tags: tuple[str, ...] = ()
            if preview is not None and not preview.found:
                tags = ("missing_required",) if rule.required else ("missing_optional",)
            self.fields_view.item(iid, tags=tags)
        self.best = None
        if sample is not None and len(self.profiles) > 1:
            try:
                self.best = best_profile(self.profiles.parsing_rules(), sample)
            except ValueError:  # no profile has fields yet
                self.best = None
        sheet = str(self.form["sheet_errors"].get()).strip() or "fehler"
        if sample is None:
            self.sample_summary.configure(text=self.tr("preview.no_sample_summary"), foreground=GRAY)
        elif self.best is not None:
            text = profile_summary_text(self.best, sheet, self.lang)
            self.sample_summary.configure(text=text, foreground=GREEN if text.startswith("✓") else RED)
        elif self.fields and self.previews is not None:
            text = summary_text(self.fields, self.previews, sheet, self.lang)
            self.sample_summary.configure(text=text, foreground=GREEN if text.startswith("✓") else RED)
        else:
            self.sample_summary.configure(text="")
        self._update_editor_preview(sample)
        self._set_fields_card()
