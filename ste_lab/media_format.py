"""Excel formatting for Media pp / mf / ff columns on the Media sheet.

Blue = strings, green = brass, orange = woodwinds. Cells are filled immediately
and the same colours are stored as Excel conditional-formatting rules.
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .catalog import orchestral_group

FAMILY_MEDIA_COLORS = {
    "strings": "BDD7EE",  # blue
    "brass": "C6EFCE",  # green
    "woodwinds": "F8CBAD",  # orange
}

_HEADER_NAMES = (
    "media pp",
    "media mf",
    "media ff",
    "média pp",
    "média mf",
    "média ff",
)


def find_media_value_columns(ws: Worksheet) -> list[int]:
    """Column indices for Media / Média pp, mf, ff (not ln Media, not origins)."""
    cols: list[int] = []
    for col in range(1, (ws.max_column or 1) + 1):
        raw = ws.cell(1, col).value
        if raw is None:
            continue
        header = str(raw).strip().lower()
        if header.startswith("ln "):
            continue
        if any(header == name or header.startswith(name + " ") or header.startswith(name + "=") for name in _HEADER_NAMES):
            cols.append(col)
    return cols


def apply_media_family_conditional_format(
    ws: Worksheet,
    media_columns: list[int] | None,
    first_row: int,
    last_row: int,
    instrument: str,
) -> None:
    """Fill Media pp / mf / ff on this sheet and attach matching CF rules."""
    cols = [c for c in (media_columns or find_media_value_columns(ws)) if c >= 1]
    if not cols or last_row < first_row:
        return

    group = orchestral_group(instrument)
    color = FAMILY_MEDIA_COLORS.get(group)
    if not color:
        return
    fill = PatternFill("solid", fgColor=color)
    header_font = Font(name="Calibri", size=11, bold=True, color="1F4E79")
    legend = (
        f"Media pp / mf / ff: blue = strings, green = brass, orange = woodwinds. "
        f"This book is {group} ({instrument})."
    )

    for col in cols:
        head = ws.cell(1, col)
        head.fill = fill
        head.font = header_font
        head.alignment = Alignment(wrap_text=True, vertical="center")
        for row in range(first_row, last_row + 1):
            ws.cell(row, col).fill = fill

    key_col = max((ws.max_column or 1), max(cols)) + 2
    ws.cell(1, key_col, "orchestral_group")
    ws.cell(2, key_col, group)
    ws.cell(1, key_col).font = Font(name="Calibri", size=8, italic=True, color="808080")
    ws.cell(2, key_col).font = Font(name="Calibri", size=8, italic=True, color="808080")
    key_ref = f"${get_column_letter(key_col)}$2"

    for name, hex_color in FAMILY_MEDIA_COLORS.items():
        rule = FormulaRule(
            formula=[f'{key_ref}="{name}"'],
            fill=PatternFill("solid", fgColor=hex_color),
        )
        for col in cols:
            letter = get_column_letter(col)
            ws.conditional_formatting.add(f"{letter}{first_row}:{letter}{last_row}", rule)

    note = Comment(legend, "STE Lab")
    note.width = "280pt"
    note.height = "80pt"
    if ws.cell(1, cols[0]).comment is None:
        ws.cell(1, cols[0]).comment = note

    legend_row = last_row + 2
    ws.cell(legend_row, 1, legend)
    ws.cell(legend_row, 1).font = Font(name="Calibri", size=9, italic=True, color="1F4E79")
    ws.merge_cells(start_row=legend_row, start_column=1, end_row=legend_row, end_column=min(10, max(cols)))
