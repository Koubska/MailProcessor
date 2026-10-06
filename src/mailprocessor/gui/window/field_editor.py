"""The editor below the field list: one rule's inputs, applied as soon as they are valid."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from mailprocessor.config import FieldRule
from mailprocessor.gui.preview import SampleMail, preview_rule
from mailprocessor.gui.rule_inputs import rule_from_inputs
from mailprocessor.gui.rule_lists import column_note, complete_column
from mailprocessor.gui.window.widgets import AMBER, GRAY, GREEN, RED
from mailprocessor.rule_patterns import LABEL_TYPES, RULE_TYPES, RuleType

# Keys that must not trigger inline completion of a column name (deleting, moving, modifiers).
COMPLETION_IGNORED_KEYS = {"BackSpace", "Delete", "Left", "Right", "Up", "Down", "Home", "End", "Tab", "Return"}


class FieldEditor:
    """Edits the selected field (or a new one) and shows what it finds in the sample mail."""

    def _build_field_editor(self, parent: ttk.Frame) -> None:
        self.editor = ttk.LabelFrame(parent, padding=10)
        self.editor.pack(fill=tk.X)
        self.editor.columnconfigure(1, weight=1)
        self.field_vars = {name: tk.StringVar() for name in ("column", "labels", "start", "end", "pattern")}
        self.field_required = tk.BooleanVar(value=True)
        for variable in (*self.field_vars.values(), self.field_required):
            variable.trace_add("write", lambda *_args: self.on_editor_changed())

        self.label(self.editor, "label.field_column").grid(row=0, column=0, sticky="w", padx=(0, 8))
        # Editable: suggests the columns of the other profiles (same name = same Excel column).
        self.column_entry = ttk.Combobox(
            self.editor, textvariable=self.field_vars["column"], postcommand=self.update_column_choices
        )
        self.column_entry.grid(row=0, column=1, sticky="ew")
        self.tip(self.column_entry, "tip.field_column")
        self.column_entry.bind("<KeyRelease>", self.on_column_typed)
        self.column_entry.bind("<<ComboboxSelected>>", lambda _event: self.on_column_chosen())
        self.column_entry.bind("<Return>", lambda _event: self.on_column_chosen())
        required_box = ttk.Checkbutton(self.editor, variable=self.field_required)
        self._translate(required_box, "label.field_required").grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.tip(required_box, "tip.field_required")
        self.column_hint = ttk.Label(self.editor, style="Hint.TLabel", wraplength=self.px(520), justify=tk.LEFT)
        self.column_hint.grid(row=1, column=1, columnspan=2, sticky="w", pady=(2, 0))
        self.label(self.editor, "label.field_type").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=(8, 0))
        self.type_box = ttk.Combobox(self.editor, state="readonly")
        self.tip(self.type_box, "tip.field_type")
        self.type_box.grid(row=2, column=1, sticky="ew", pady=(8, 0))
        self.type_box.bind("<<ComboboxSelected>>", lambda _event: self.on_type_selected())

        self.inputs = ttk.Frame(self.editor)
        self.inputs.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.inputs.columnconfigure(1, weight=1)
        self.input_widgets = {
            "labels": (ttk.Label(self.inputs), ttk.Entry(self.inputs, textvariable=self.field_vars["labels"])),
            "start": (
                self.label(self.inputs, "label.field_start"),
                ttk.Entry(self.inputs, textvariable=self.field_vars["start"]),
            ),
            "end": (
                self.label(self.inputs, "label.field_end"),
                ttk.Entry(self.inputs, textvariable=self.field_vars["end"]),
            ),
            "pattern": (
                self.label(self.inputs, "label.field_pattern"),
                ttk.Entry(self.inputs, textvariable=self.field_vars["pattern"]),
            ),
        }
        self.type_hint = ttk.Label(self.editor, style="Hint.TLabel", wraplength=self.px(560), justify=tk.LEFT)
        self.type_hint.grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.editor_result = ttk.Label(self.editor, wraplength=self.px(560), justify=tk.LEFT, font=self.bold_font)
        self.editor_result.grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 0))
        self.editor_problem = ttk.Label(self.editor, wraplength=self.px(560), justify=tk.LEFT, foreground=AMBER)
        self.editor_problem.grid(row=6, column=0, columnspan=3, sticky="w")
        editor_buttons = ttk.Frame(self.editor)
        editor_buttons.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self.add_button = self.button(editor_buttons, "button.field_add", self.add_field)
        self.cancel_new_button = self.button(editor_buttons, "button.cancel", self.cancel_new_field)
        self.as_regex_button = self.button(editor_buttons, "button.as_regex", self.convert_to_regex, tip="tip.as_regex")

    def on_field_selected(self) -> None:
        index = self.selected_index()
        if index is None or (index == self.editing_index and self.editor_mode == "edit"):
            return
        self.flush_editor()  # pending changes belong to the previously selected field
        rule = self.fields[index]
        self.editor_mode = "edit"
        self.editing_index = index
        self._fill_editor(rule)
        self._update_editor_buttons()

    def _fill_editor(self, rule: FieldRule | None) -> None:
        self.loading_editor = True
        self.field_vars["column"].set(rule.column if rule else "")
        self.field_vars["labels"].set("; ".join(rule.label) if rule else "")
        self.field_vars["start"].set((rule.start or "") if rule else "")
        self.field_vars["end"].set((rule.end or "") if rule else "")
        self.field_vars["pattern"].set((rule.pattern or "") if rule else "")
        self.field_required.set(rule.required if rule else True)
        self.loading_editor = False
        self.set_field_type(rule.type if rule else "label")

    def _update_editor_buttons(self) -> None:
        new = self.editor_mode == "new"
        self.editor.configure(text=self.tr("fields.editor_new" if new else "fields.editor_edit"))
        for widget in (self.add_button, self.cancel_new_button, self.as_regex_button):
            widget.pack_forget()
        if new:
            self.add_button.pack(side=tk.LEFT)
            if self.fields:
                self.cancel_new_button.pack(side=tk.LEFT, padx=(8, 0))
        if self.field_type != "regex":
            self.as_regex_button.pack(side=tk.RIGHT)

    def set_field_type(self, rule_type: RuleType) -> None:
        self.field_type = rule_type
        self.type_box.current(RULE_TYPES.index(rule_type))
        for label_widget, entry_widget in self.input_widgets.values():
            label_widget.grid_forget()
            entry_widget.grid_forget()
        if rule_type in LABEL_TYPES:
            shown = ["labels"]
            key = "label.field_labels_optional" if rule_type == "email" else "label.field_labels"
            self.input_widgets["labels"][0].configure(text=self.tr(key))
        elif rule_type == "between":
            shown = ["start", "end"]
        else:
            shown = ["pattern"]
        for row, name in enumerate(shown):
            label_widget, entry_widget = self.input_widgets[name]
            label_widget.grid(row=row, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
            entry_widget.grid(row=row, column=1, sticky="ew", pady=(0, 4))
        self.type_hint.configure(text=self.tr(f"rule.hint.{rule_type}"))
        self._update_editor_buttons()
        self.on_editor_changed()

    def on_type_selected(self) -> None:
        self.set_field_type(RULE_TYPES[self.type_box.current()])

    @property
    def _edited_index(self) -> int | None:
        """The position of the field being edited; None for a new field, whose name must differ from all others."""
        return self.editing_index if self.editor_mode == "edit" else None

    def editor_rule(self, for_preview: bool = False) -> FieldRule:
        values = {name: variable.get() for name, variable in self.field_vars.items()}
        return rule_from_inputs(
            column=values["column"] or ("?" if for_preview else ""),
            rule_type=self.field_type,
            labels_text=values["labels"],
            start=values["start"],
            end=values["end"],
            pattern=values["pattern"],
            required=self.field_required.get(),
            other_columns=None if for_preview else self.fields.columns(except_index=self._edited_index),
            lang=self.lang,
        )

    def _column_suggestions(self, typed: str) -> list[str]:
        return self.profiles.column_suggestions(typed, except_index=self._edited_index)

    def update_column_choices(self) -> None:
        """Fill the drop-down just before it opens: other profiles' columns that match what was typed."""
        typed = self.field_vars["column"].get()
        choices = self._column_suggestions(typed)
        # Exactly one of them typed: show them all, so the list also works for switching to another name.
        self.column_entry.configure(values=choices if choices != [typed] else self.profiles.column_suggestions(""))

    def on_column_typed(self, event: tk.Event) -> None:
        """Complete a new field's name inline from the other profiles; the completed part stays selected."""
        if self.editor_mode != "new" or event.keysym in COMPLETION_IGNORED_KEYS or len(event.char) != 1:
            return
        typed = self.column_entry.get()[: self.column_entry.index(tk.INSERT)]
        completion = complete_column(typed, self._column_suggestions(self.field_vars["column"].get()))
        if completion is None:
            return
        self.field_vars["column"].set(completion)
        self.column_entry.icursor(len(typed))
        self.column_entry.selection_range(len(typed), tk.END)

    def on_column_chosen(self) -> None:
        """A new field named like another profile's column starts from that profile's rule (if nothing typed yet)."""
        self.column_entry.selection_clear()
        self.column_entry.icursor(tk.END)
        column = self.field_vars["column"].get().strip()
        rule = self.profiles.rule_for_column(column)
        inputs = ("labels", "start", "end", "pattern")
        if self.editor_mode != "new" or rule is None or any(self.field_vars[name].get().strip() for name in inputs):
            return
        self._fill_editor(rule)
        profile = next(iter(self.profiles.profiles_with_column(column)), "")
        self.set_status(self.tr("status.rule_copied").format(column=self.quote(column), profile=self.quote(profile)))

    def update_column_note(self) -> None:
        if not hasattr(self, "column_hint"):
            return
        per_profile = self.form["profile_sheets"].get() == "per_profile" if "profile_sheets" in self.form else False
        note = column_note(self.profiles, self.field_vars["column"].get(), per_profile, self.lang)
        self.column_hint.configure(text=note.text, foreground=AMBER if note.warning else GRAY)
        if note.text:
            self.column_hint.grid()
        else:
            self.column_hint.grid_remove()

    def on_editor_changed(self) -> None:
        if self.loading_editor:
            return
        self.schedule_preview()
        if self.editor_mode == "edit" and self.editing_index is not None:
            if self.apply_job is not None:
                self.root.after_cancel(self.apply_job)
            self.apply_job = self.root.after(500, self.apply_editor)

    def apply_editor(self) -> None:
        """Changes to the selected field take effect as soon as they are valid; no extra button."""
        self.apply_job = None
        index = self.editing_index
        if self.editor_mode != "edit" or index is None or index >= len(self.fields):
            return
        try:
            changed = self.fields.replace(index, self.editor_rule())
        except ValueError:
            return  # the problem is shown below the editor; the list keeps the last valid version
        if not changed:
            return
        self._update_undo_button()
        self.fields_view.item(str(index), values=self._row_values(self.fields[index]))
        self.save_rules()
        self.schedule_preview()

    def convert_to_regex(self) -> None:
        try:
            rule = self.editor_rule(for_preview=self.editor_mode == "new")
        except ValueError as exc:
            self.show_error(str(exc))
            return
        self.loading_editor = True
        self.field_vars["pattern"].set(rule.regex)
        self.loading_editor = False
        self.set_field_type("regex")

    def _update_editor_preview(self, sample: SampleMail | None) -> None:
        """Result of the rule being edited, and what is still wrong with its inputs (e.g. a duplicate column)."""
        self.update_column_note()
        self.sample_text.tag_remove("match", "1.0", tk.END)
        self.sample_header.tag_remove("match", "1.0", tk.END)
        self.editor_problem.configure(text="")
        inputs = {"between": ("start", "end"), "regex": ("pattern",)}.get(self.field_type, ("labels",))
        if self.field_type != "email" and not any(self.field_vars[name].get().strip() for name in inputs):
            self.editor_result.configure(text=self.tr("preview.editor.waiting"), foreground=GRAY)
            return
        try:
            rule = self.editor_rule(for_preview=True)
        except ValueError as exc:
            self.editor_result.configure(text=f"⚠ {exc}", foreground=AMBER)
            return
        try:
            self.editor_rule()
        except ValueError as exc:
            self.editor_problem.configure(text=f"⚠ {exc}")
        if sample is None:
            self.editor_result.configure(text=self.tr("preview.editor.no_sample"), foreground=GRAY)
            return
        preview = preview_rule(rule, sample)
        if not preview.found:
            self.editor_result.configure(text=self.tr("preview.editor.not_found"), foreground=RED)
            return
        key = "preview.editor.from_header" if preview.from_header else "preview.editor.found"
        self.editor_result.configure(text=self.tr(key).format(value=preview.value), foreground=GREEN)
        if preview.span is not None:
            start, end = (f"1.0 + {offset} chars" for offset in preview.span)
            self.sample_text.tag_add("match", start, end)
            self.sample_text.see(start)
        elif preview.header_span is not None:
            start, end = (f"1.0 + {offset} chars" for offset in preview.header_span)
            self.sample_header.tag_add("match", start, end)
            self.sample_header.see(start)
