"""Profiles in the Felder tab: choosing, adding, renaming and removing them."""

from __future__ import annotations

import tkinter as tk
from tkinter import simpledialog, ttk


class ProfileActions:
    """The profile row above the field list."""

    def _build_profile_row(self, parent: ttk.Frame) -> None:
        profile_row = ttk.Frame(parent)
        profile_row.pack(fill=tk.X, pady=(0, 8))
        profile_label = self.label(profile_row, "label.profile", style="Heading.TLabel")
        profile_label.pack(side=tk.LEFT)
        self.tip(profile_label, "tip.profile")
        self.profile_box = ttk.Combobox(profile_row, state="readonly", width=18)
        self.profile_box.pack(side=tk.LEFT, padx=(8, 0))
        self.tip(self.profile_box, "tip.profile")
        self.profile_box.bind("<<ComboboxSelected>>", lambda _event: self.on_profile_selected())
        self.button(profile_row, "button.profile_new", self.new_profile, tip="tip.profile_new").pack(
            side=tk.LEFT, padx=(8, 0)
        )
        self.button(profile_row, "button.profile_rename", self.rename_profile, tip="tip.profile_rename").pack(
            side=tk.LEFT, padx=(8, 0)
        )
        self.remove_profile_button = self.button(
            profile_row, "button.profile_remove", self.remove_profile, tip="tip.profile_remove"
        )
        self.remove_profile_button.pack(side=tk.LEFT, padx=(8, 0))

    def refresh_profile_box(self) -> None:
        self.profile_box.configure(values=self.profiles.names)
        self.profile_box.current(self.profiles.index)
        self.remove_profile_button.state(["!disabled"] if len(self.profiles) > 1 else ["disabled"])

    def show_profile(self, index: int) -> None:
        """Show another profile's fields; pending edits are applied to the profile shown so far."""
        self.flush_editor()
        self.profiles.select(index)
        self.editing_index = None
        self.refresh_profile_box()
        self.refresh_fields_view()
        self.select_field(0 if self.fields else None)
        self._update_undo_button()

    def on_profile_selected(self) -> None:
        index = self.profile_box.current()
        if index != self.profiles.index:
            self.show_profile(index)

    def ask_profile_name(self, title_key: str, initial: str = "") -> str | None:
        return simpledialog.askstring(
            self.tr(title_key), self.tr("dialog.profile_name"), initialvalue=initial, parent=self.root
        )

    def new_profile(self) -> None:
        self.flush_editor()
        name = ""
        while True:
            name = self.ask_profile_name("dialog.profile_new", name)
            if name is None:
                return
            try:
                index = self.profiles.add(name, self.lang)
                break
            except ValueError as exc:
                self.show_error(str(exc))
        # Not saved yet: a profile without fields is left out of the rules file until it has one.
        self.show_profile(index)
        self.set_status(self.tr("status.profile_added").format(name=self.quote(self.profiles.current_name)))

    def rename_profile(self) -> None:
        self.flush_editor()
        name = self.profiles.current_name
        while True:
            name = self.ask_profile_name("dialog.profile_rename", name)
            if name is None or name.strip() == self.profiles.current_name:
                return
            try:
                self.profiles.rename(name, self.lang)
                break
            except ValueError as exc:
                self.show_error(str(exc))
        self.save_rules()
        self.refresh_profile_box()
        self.schedule_preview()
        self.set_status(self.tr("status.profile_renamed").format(name=self.quote(self.profiles.current_name)))

    def remove_profile(self) -> None:
        if len(self.profiles) < 2:
            return
        self.flush_editor()
        name = self.quote(self.profiles.current_name)
        question = self.tr("confirm.profile_remove").format(name=name, count=len(self.fields))
        if not self.confirm(question):
            return
        self.profiles.remove()
        self.save_rules()
        self.show_profile(self.profiles.index)
        self.set_status(self.tr("status.profile_removed").format(name=name))
