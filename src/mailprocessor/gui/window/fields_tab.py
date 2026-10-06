"""The Felder tab's list of fields: adding, removing (with undo) and reordering them."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from mailprocessor.config import FieldRule
from mailprocessor.gui.rule_inputs import describe_rule
from mailprocessor.gui.window.widgets import AMBER, GRAY, RED


class FieldsTab:
    """The field list of the shown profile; the editor and the sample mail are separate mixins."""

    def _build_fields_tab(self, tab: ttk.Frame) -> None:
        self.label(tab, "fields.title", style="Title.TLabel").pack(anchor="w")
        self.hint(tab, "fields.intro", wrap=900).pack(anchor="w", pady=(4, 10))
        paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        rules_pane = ttk.Frame(paned, padding=(0, 0, 12, 0))
        sample_pane = ttk.Frame(paned, padding=(12, 0, 0, 0))
        paned.add(rules_pane, weight=3)
        paned.add(sample_pane, weight=2)
        self._build_profile_row(rules_pane)
        self._build_field_list(rules_pane)
        self._build_field_editor(rules_pane)
        self._build_sample_pane(sample_pane)

    def _build_field_list(self, parent: ttk.Frame) -> None:
        columns = ("column", "type", "description", "required", "result")
        self.fields_view = ttk.Treeview(parent, columns=columns, show="headings", height=7, selectmode="browse")
        for name, width, anchor in (
            ("column", 120, "w"),
            ("type", 150, "w"),
            ("description", 220, "w"),
            ("required", 50, "center"),
            ("result", 190, "w"),
        ):
            self.fields_view.column(name, width=self.px(width), anchor=anchor)
        self.fields_view.tag_configure("missing_required", foreground=RED)
        self.fields_view.tag_configure("missing_optional", foreground=GRAY)
        self.fields_view.pack(fill=tk.BOTH, expand=True)
        self.fields_view.bind("<<TreeviewSelect>>", lambda _event: self.on_field_selected())

        list_buttons = ttk.Frame(parent)
        list_buttons.pack(fill=tk.X, pady=(6, 12))
        list_actions = (
            ("button.field_new", self.new_field, "tip.field_new"),
            ("button.field_remove", self.remove_field, "tip.field_remove"),
            ("button.field_up", lambda: self.move_field(-1), "tip.field_move"),
            ("button.field_down", lambda: self.move_field(1), "tip.field_move"),
        )
        for position, (key, command, tip) in enumerate(list_actions):
            self.button(list_buttons, key, command, tip=tip).pack(side=tk.LEFT, padx=(8 if position else 0, 0))

    def refresh_fields_view(self, keep_selection: bool = False) -> None:
        selected = self.selected_index() if keep_selection else None
        self.fields_view.delete(*self.fields_view.get_children())
        for index, rule in enumerate(self.fields):
            self.fields_view.insert("", tk.END, iid=str(index), values=self._row_values(rule))
        if selected is not None and selected < len(self.fields):
            self.fields_view.selection_set(str(selected))
        self.schedule_preview()

    def _row_values(self, rule: FieldRule) -> tuple[str, ...]:
        required = self.tr("view.field.required_yes") if rule.required else self.tr("view.field.required_no")
        return (rule.column, self.tr(f"rule.type.{rule.type}"), describe_rule(rule, self.lang), required, "–")

    def selected_index(self) -> int | None:
        selection = self.fields_view.selection()
        return int(selection[0]) if selection else None

    def select_field(self, index: int | None) -> None:
        if index is None:
            self.new_field()
            return
        self.fields_view.selection_set(str(index))
        self.fields_view.see(str(index))

    def new_field(self) -> None:
        self.flush_editor()
        self.editor_mode = "new"
        self.editing_index = None
        self.fields_view.selection_remove(*self.fields_view.selection())
        self._fill_editor(None)
        self._update_editor_buttons()
        self.column_entry.focus_set()

    def cancel_new_field(self) -> None:
        self.select_field(0 if self.fields else None)

    def add_field(self) -> None:
        try:
            rule = self.editor_rule()
            index = self.fields.add(rule)
        except ValueError as exc:
            self.show_error(str(exc))
            return
        self.editing_index = None
        self._update_undo_button()
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(index)
        self.set_status(self.tr("label.status.field_added").format(column=rule.column))

    def remove_field(self) -> None:
        self.flush_editor()
        index = self.selected_index()
        if index is None:
            self.show_info(self.tr("error.select_field_remove"))
            return
        removed = self.fields.remove(index)
        self.editing_index = None
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(min(index, len(self.fields) - 1) if self.fields else None)
        # No confirmation dialog: the deletion can be undone from the status bar instead.
        self.set_status(self.tr("status.field_removed").format(column=removed.column))
        self._update_undo_button()

    def undo_remove_field(self) -> None:
        self.flush_editor()
        index = self.fields.undo_remove()
        self._update_undo_button()
        if index is None:
            self.set_status(self.tr("status.undo_failed"), AMBER)
            return
        self.editing_index = None
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(index)

    def _update_undo_button(self) -> None:
        if self.fields.can_undo:
            self.undo_button.pack(side=tk.LEFT, padx=(12, 0))
        else:
            self.undo_button.pack_forget()

    def move_field(self, offset: int) -> None:
        self.flush_editor()
        index = self.selected_index()
        if index is None:
            self.show_info(self.tr("error.select_field_move"))
            return
        new_index = self.fields.move(index, offset)
        if new_index is None:
            return
        self.editing_index = None
        self._update_undo_button()
        self.save_rules()
        self.refresh_fields_view()
        self.select_field(new_index)
