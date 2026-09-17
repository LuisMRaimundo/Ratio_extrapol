"""Write transfer-relation inventory / validation sheets and stamp PI95."""

from __future__ import annotations

import math
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.worksheet import Worksheet

from .notes import midi_to_label
from .relations import Relation
from .session import Cell, Layer, Project
from .validation import (
    PAIRWISE_COLUMNS,
    SPREAD_COLUMNS,
    relation_level_sd,
    spread_sd_at,
    summary_validation_lines,
)


def _style_sheet(ws: Worksheet, start: int = 1) -> None:
    fill = PatternFill("solid", fgColor="1F4E79")
    hf = Font(name="Calibri", size=10, color="FFFFFF", bold=True)
    for cell in ws[start]:
        cell.fill = fill
        cell.font = hf
    for col in ws.columns:
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = min(52, max(12, len(str(col[0].value or "")) + 3))


def write_df_sheet(
    wb: Workbook,
    name: str,
    df: pd.DataFrame,
    note: str = "",
    columns: Optional[list[str]] = None,
) -> Worksheet:
    ws = wb.create_sheet(name[:31])
    start = 1
    if note:
        ws["A1"] = note
        ws["A1"].font = Font(name="Calibri", size=10, italic=True, color="5C5A54")
        start = 3
    if df is None:
        df = pd.DataFrame()
    if columns is not None and df.empty:
        df = pd.DataFrame(columns=columns)
    elif columns is not None:
        for col in columns:
            if col not in df.columns:
                df[col] = None
        df = df[columns]
    if df.empty and not list(df.columns):
        ws.cell(start, 1, "(empty)")
        return ws
    for r_i, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=start):
        for c_i, val in enumerate(row, start=1):
            cell = ws.cell(r_i, c_i, val)
            if isinstance(val, float) and val == val:
                cell.number_format = "0.000000"
    _style_sheet(ws, start)
    return ws


def attach_relation_sheets(
    wb: Workbook,
    *,
    pairwise: Optional[pd.DataFrame] = None,
    spread: Optional[pd.DataFrame] = None,
    summary_lines: Optional[list[str]] = None,
) -> None:
    write_df_sheet(
        wb,
        "L_Validation",
        pairwise if pairwise is not None else pd.DataFrame(columns=PAIRWISE_COLUMNS),
        "Pairwise check of the transfer assumption on overlapping anchors. Non-destructive.",
        PAIRWISE_COLUMNS,
    )
    write_df_sheet(
        wb,
        "L_Spread",
        spread if spread is not None else pd.DataFrame(columns=SPREAD_COLUMNS),
        "Per-note min/max/range of L where ≥2 collections overlap.",
        SPREAD_COLUMNS,
    )
    lines = summary_lines if summary_lines is not None else summary_validation_lines(pairwise)
    ws = wb.create_sheet("Summary_Validation")
    ws.append(["line"])
    _style_sheet(ws, 1)
    if not lines:
        ws.append(["No multi-collection transfer relation to check."])
    else:
        for line in lines:
            ws.append([line])
    ws.column_dimensions["A"].width = 120


def write_validation_csv(path: Path, pairwise: pd.DataFrame) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pairwise if pairwise is not None else pd.DataFrame(columns=PAIRWISE_COLUMNS)
    if df.empty:
        df = pd.DataFrame(columns=PAIRWISE_COLUMNS)
    df.to_csv(path, index=False)
    return path


def apply_empirical_spread_pi95(
    layers: Iterable[Layer],
    *,
    spread: pd.DataFrame,
    pairwise: pd.DataFrame,
    kind: str,
    instrument: str,
    source_cond: str,
    target_cond: str,
    dynamic: Optional[str] = None,
    fallback_sd: Optional[float] = None,
) -> int:
    """Fill PI95 on modelled cells only where ≥2 collections overlap at that MIDI.

    Leaves measured cells and non-overlap modelled cells untouched.
    """
    stamped = 0
    for layer in layers:
        if dynamic is not None and layer.dynamic != dynamic:
            continue
        for cell in layer.cells.values():
            if (cell.origin or "").strip().lower() == "measured":
                continue
            sd = spread_sd_at(
                spread,
                kind=kind,
                instrument=instrument,
                source_cond=source_cond,
                target_cond=target_cond,
                dynamic=layer.dynamic,
                midi=int(cell.midi),
                fallback_sd=fallback_sd,
            )
            if sd is None or not math.isfinite(sd):
                continue
            value = float(cell.value)
            if value <= 0 or not math.isfinite(value):
                continue
            cell.pi95_low = value * math.exp(-1.96 * sd)
            cell.pi95_high = value * math.exp(1.96 * sd)
            cell.uncertainty = "empirical_spread"
            stamped += 1
    return stamped


def apply_pi95_for_technique(
    project: Project,
    *,
    instrument: str,
    technique: str,
    spread: pd.DataFrame,
    pairwise: pd.DataFrame,
    source_cond: str = "ordinario",
) -> int:
    dummy = Relation(
        kind="technique",
        instrument=instrument,
        source_cond=source_cond,
        target_cond=technique,
        collection="",
        dynamic="",
    )
    fallback = relation_level_sd(pairwise, dummy)
    n = 0
    for dyn in ("pp", "mf", "ff", "p", "mp", "f"):
        n += apply_empirical_spread_pi95(
            [lg for lg in project.layers if lg.technique == technique and lg.dynamic == dyn],
            spread=spread,
            pairwise=pairwise,
            kind="technique",
            instrument=instrument,
            source_cond=source_cond,
            target_cond=technique,
            dynamic=dyn,
            fallback_sd=fallback,
        )
    return n


def apply_pi95_for_instrument(
    project: Project,
    *,
    instrument: str,
    source_cond: str,
    spread: pd.DataFrame,
    pairwise: pd.DataFrame,
) -> int:
    dummy = Relation(
        kind="instrument",
        instrument=instrument,
        source_cond=source_cond,
        target_cond=instrument,
        collection="",
        dynamic="",
    )
    fallback = relation_level_sd(pairwise, dummy)
    return apply_empirical_spread_pi95(
        [lg for lg in project.layers if (lg.technique or "").lower() in {"ordinario", "arco"}],
        spread=spread,
        pairwise=pairwise,
        kind="instrument",
        instrument=instrument,
        source_cond=source_cond,
        target_cond=instrument,
        fallback_sd=fallback,
    )


def contributor_log_lines(
    *,
    technique: str,
    dynamic: str,
    contributors: dict[int, list[tuple[str, float, float]]],
    mode: str,
) -> list[str]:
    lines = []
    for midi in sorted(contributors):
        hits = contributors[midi]
        if not hits:
            continue
        parts = [f"{coll} (w={w:g}, L={L:.4f})" for coll, L, w in hits]
        note = midi_to_label(int(midi))
        lines.append(
            f"L[{note}/{midi}] {technique} {dynamic} mode={mode}: " + "; ".join(parts)
        )
    return lines


def attach_project_validation(
    project: Project,
    *,
    pairwise: pd.DataFrame,
    spread: pd.DataFrame,
) -> None:
    project.validation_pairwise = pairwise
    project.validation_spread = spread
    project.validation_summary = summary_validation_lines(pairwise)
    for line in project.validation_summary:
        project.log.append(line)
