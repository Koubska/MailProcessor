"""The profiles and fields being edited in the Felder tab, validated like the rules file (no tkinter)."""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass

from mailprocessor.config import (
    DEFAULT_PROFILE_NAME,
    SHEET_NAME_MAX_LENGTH,
    FieldRule,
    ParsingRules,
    Profile,
    profile_name_problem,
)
from mailprocessor.i18n import quote, t


class RuleList:
    """The fields edited in the GUI. Every change is validated like the rules file; a removal can be undone."""

    def __init__(self, rules: list[FieldRule]) -> None:
        self.rules = list(rules)
        self._removed: tuple[int, FieldRule] | None = None

    def __len__(self) -> int:
        return len(self.rules)

    def __iter__(self) -> Iterator[FieldRule]:
        return iter(self.rules)

    def __getitem__(self, index: int) -> FieldRule:
        return self.rules[index]

    def columns(self, except_index: int | None = None) -> list[str]:
        return [rule.column for index, rule in enumerate(self.rules) if index != except_index]

    def parsing_rules(self) -> ParsingRules:
        return ParsingRules(fields=self.rules)

    def _commit(self, rules: list[FieldRule]) -> None:
        if rules:
            ParsingRules(fields=rules)  # raises ValueError, e.g. for duplicates or the reserved column name
        self.rules = rules
        self._removed = None

    def add(self, rule: FieldRule) -> int:
        self._commit([*self.rules, rule])
        return len(self.rules) - 1

    def replace(self, index: int, rule: FieldRule) -> bool:
        """Put `rule` at `index`; returns False (and changes nothing) if it is the same rule."""
        if self.rules[index] == rule:
            return False
        self._commit([*self.rules[:index], rule, *self.rules[index + 1 :]])
        return True

    def move(self, index: int, offset: int) -> int | None:
        new_index = index + offset
        if not 0 <= new_index < len(self.rules):
            return None
        rules = list(self.rules)
        rules[index], rules[new_index] = rules[new_index], rules[index]
        self._commit(rules)
        return new_index

    def remove(self, index: int) -> FieldRule:
        removed = self.rules[index]
        self.rules = [*self.rules[:index], *self.rules[index + 1 :]]
        self._removed = (index, removed)
        return removed

    @property
    def can_undo(self) -> bool:
        return self._removed is not None

    def undo_remove(self) -> int | None:
        """Put the last removed field back where it was; None if nothing to undo or it would be invalid now."""
        if self._removed is None:
            return None
        index, rule = self._removed
        index = min(index, len(self.rules))
        try:
            self._commit([*self.rules[:index], rule, *self.rules[index:]])
        except ValueError:
            self._removed = None
            return None
        return index


class ProfileList:
    """The profiles edited in the GUI, each with its own `RuleList`. The Felder tab shows one at a time.

    A profile without fields is kept while editing but left out when saving or running.
    """

    def __init__(self, profiles: list[Profile]) -> None:
        self.names = [profile.name for profile in profiles] or [DEFAULT_PROFILE_NAME]
        self.fields = [RuleList(profile.fields) for profile in profiles] or [RuleList([])]
        self.index = 0

    def __len__(self) -> int:
        return len(self.names)

    @property
    def current(self) -> RuleList:
        return self.fields[self.index]

    @property
    def current_name(self) -> str:
        return self.names[self.index]

    @property
    def field_count(self) -> int:
        return sum(len(rules) for rules in self.fields)

    def select(self, index: int) -> None:
        if 0 <= index < len(self.names):
            self.index = index

    def find(self, name: str) -> int | None:
        return self.names.index(name) if name in self.names else None

    def name_problem(self, name: str, lang: str, except_index: int | None = None) -> str | None:
        """Why `name` cannot be used for a new or renamed profile, in the user's language; None if it can."""
        name = name.strip()
        if not name:
            return t("error.profile.name_empty", lang)
        if len(name) > SHEET_NAME_MAX_LENGTH:
            return t("error.profile.name_long", lang).format(max=SHEET_NAME_MAX_LENGTH)
        if profile_name_problem(name):
            return t("error.profile.name_chars", lang)
        taken = [other for index, other in enumerate(self.names) if index != except_index]
        if name.casefold() in (other.casefold() for other in taken):
            return t("error.profile.name_taken", lang).format(name=quote(name, lang))
        return None

    def add(self, name: str, lang: str) -> int:
        """Add an empty profile and show it; raises ValueError with a message for the user."""
        problem = self.name_problem(name, lang)
        if problem:
            raise ValueError(problem)
        self.names.append(name.strip())
        self.fields.append(RuleList([]))
        self.index = len(self.names) - 1
        return self.index

    def rename(self, name: str, lang: str) -> None:
        problem = self.name_problem(name, lang, except_index=self.index)
        if problem:
            raise ValueError(problem)
        self.names[self.index] = name.strip()

    def remove(self) -> str:
        """Remove the shown profile (never the last one) and show its neighbour."""
        if len(self.names) == 1:
            raise ValueError("the last profile cannot be removed")
        name = self.names.pop(self.index)
        self.fields.pop(self.index)
        self.index = min(self.index, len(self.names) - 1)
        return name

    def other_columns(self) -> list[str]:
        """Columns of the other profiles, each once, in profile order."""
        columns = (rule.column for index, rules in enumerate(self.fields) if index != self.index for rule in rules)
        return list(dict.fromkeys(columns))

    def column_suggestions(self, typed: str, except_index: int | None = None) -> list[str]:
        """Column names of other profiles to offer in the shown profile: not used there yet, containing `typed`.

        `except_index` is the field being edited; its own name stays available.
        """
        own = set(self.current.columns(except_index=except_index))
        wanted = typed.strip().casefold()
        return [column for column in self.other_columns() if column not in own and wanted in column.casefold()]

    def profiles_with_column(self, column: str) -> list[str]:
        """The other profiles that have exactly this column."""
        return [
            name
            for index, (name, rules) in enumerate(zip(self.names, self.fields, strict=True))
            if index != self.index and column in rules.columns()
        ]

    def rule_for_column(self, column: str) -> FieldRule | None:
        """How another profile finds this column (the first that has it), to start from in the shown profile."""
        for index, rules in enumerate(self.fields):
            if index != self.index:
                for rule in rules:
                    if rule.column == column:
                        return rule
        return None

    def profiles(self) -> list[Profile]:
        """The profiles to save: all that have fields."""
        return [
            Profile(name=name, fields=rules.rules) for name, rules in zip(self.names, self.fields, strict=True) if rules
        ]

    def parsing_rules(self) -> ParsingRules:
        """Raises ValueError if no profile has fields yet."""
        return ParsingRules(profiles=self.profiles())


def complete_column(typed: str, suggestions: list[str]) -> str | None:
    """The first suggestion that continues what was typed (ignoring case), for inline completion."""
    wanted = typed.casefold()
    if not wanted:
        return None
    for suggestion in suggestions:
        if suggestion.casefold().startswith(wanted) and len(suggestion) > len(typed):
            return suggestion
    return None


def _column_key(column: str) -> str:
    """Spelling-insensitive form of a column name, to spot near-duplicates such as "Telefon-Nr" / "telefon nr"."""
    return re.sub(r"[\s\-_.:/]+", "", column.casefold())


@dataclass(frozen=True)
class ColumnNote:
    text: str
    warning: bool = False


def column_note(profiles: ProfileList, column: str, per_profile_sheets: bool, lang: str) -> ColumnNote:
    """Where the column ends up in Excel, shown under the column name when there are several profiles."""
    column = column.strip()
    if len(profiles) < 2 or not column:
        return ColumnNote("")
    if per_profile_sheets:
        return ColumnNote(t("column.note.per_profile", lang).format(sheet=quote(profiles.current_name, lang)))
    same = profiles.profiles_with_column(column)
    if same:
        key = "column.note.shared_one" if len(same) == 1 else "column.note.shared_many"
        return ColumnNote(t(key, lang).format(profiles=", ".join(quote(name, lang) for name in same)))
    similar = [other for other in profiles.other_columns() if _column_key(other) == _column_key(column)]
    if similar:
        return ColumnNote(
            t("column.note.similar", lang).format(
                column=quote(similar[0], lang),
                profiles=", ".join(quote(name, lang) for name in profiles.profiles_with_column(similar[0])),
            ),
            warning=True,
        )
    return ColumnNote(t("column.note.own", lang))
