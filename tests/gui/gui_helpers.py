"""Rules shared by the GUI tests."""

from mailprocessor.config import FieldRule

SIMPLE_RULES = [
    FieldRule(column="Mail-Adresse", type="email", label=["Von", "From"]),
    FieldRule(column="Absender", type="email"),
    FieldRule(column="Name", type="between", start="mein Kind", end="für folgendes Angebot"),
    FieldRule(column="Kurs", type="label", label="Angebot:", required=False),
    FieldRule(column="Bemerkung", type="next_line", label='Say "hi"'),
    FieldRule(column="Experte", pattern=r"(?m)^X:\s*(.+)$"),
]
