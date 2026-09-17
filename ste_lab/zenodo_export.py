"""Zenodo collection workbook: one instrument × one technique, IOWA + ORCH + Media.

Matches the sheet set and formula layout of
`Violin_Zenodo_collections_con_sordino (2).xlsx` (schema 1.2.0), which is the
current violin sibling of `Viola_Zenodo_collections_harmonics.xlsx`.
"""

from __future__ import annotations

import math
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .calibration import CalibrationConfig, load_config, normalize_combination_method
from .catalog import DEFAULT_MEASURED_CEILING_MIDI, orchestral_group, resolve_instrument
from .media_policy import resolve_iowa_orch_pair, uses_woodwind_combination
from .evidence import (
    SHARED_L_MEDIA_CAVEAT,
    evidence_map_rows,
    is_effect,
    principal_empirical_rows,
    principal_rows_are_iowa,
    principal_sheet_subtitle,
    principal_sheet_title,
    project_is_woodwind,
    WOODWIND_PRINCIPAL_SUBTITLE,
)
from .media_format import apply_media_family_conditional_format
from .notes import midi_to_label
from .session import Layer, Project

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
TITLE_FONT = Font(bold=True, name="Calibri", size=14, color="1F4E79")
CEILING_FILL = PatternFill("solid", fgColor="F4B183")
MEASURED_FILL = PatternFill("solid", fgColor="C6EFCE")
MODELLED_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

def collection_sheet_names(instrument: str = "violin") -> dict[tuple[str, str], str]:
    """Published viola arco book uses VIOLA_* (caps); violin uses Violin_*."""
    spec = resolve_instrument(instrument)
    iid = spec.instrument_id if spec else "violin"
    prefix = {"viola": "VIOLA", "violin": "Violin", "double_bass": "DBASS"}.get(
        iid, spec.display_name if spec else "Violin"
    )
    return {
        ("IOWA", "pp"): f"{prefix}_IOWA_pp",
        ("IOWA", "mf"): f"{prefix}__IOWA_mf",
        ("IOWA", "ff"): f"{prefix}_IOWA_ff",
        ("ORCH", "pp"): f"{prefix}__ORCH_pp",
        ("ORCH", "mf"): f"{prefix}__ORCH_mf",
        ("ORCH", "ff"): f"{prefix}__ORCH_ff",
    }


def _sheet_ref(name: str) -> str:
    """Quote an Excel sheet name when it is not a simple identifier (spaces, etc.)."""
    raw = name or ""
    simple = raw and raw[0].isalpha() and all(ch.isalnum() or ch == "_" for ch in raw)
    if simple:
        return raw
    return "'" + raw.replace("'", "''") + "'"


def media_sheet_name(instrument: str = "violin") -> str:
    spec = resolve_instrument(instrument)
    iid = spec.instrument_id if spec else "violin"
    if iid == "viola":
        return "VIOLA_Media"
    if iid == "violin":
        return "Violin_Media"
    if iid == "double_bass":
        return "DBass_Media"
    return f"{spec.display_name}_Media" if spec else "Violin_Media"


# Historical sheet-name quirk from the published violin / viola Zenodo books.
SHEET_NAMES = collection_sheet_names("violin")
CDM_CELL = {
    ("IOWA", "pp"): "G",
    ("IOWA", "mf"): "H",
    ("IOWA", "ff"): "H",
    ("ORCH", "pp"): "F",
    ("ORCH", "mf"): "F",
    ("ORCH", "ff"): "F",
}
MEDIA_COL = {"pp": "D", "mf": "G", "ff": "J"}
CORE_DYNS = ("pp", "mf", "ff")


def _style_header(ws: Worksheet, ncol: int) -> None:
    for col in range(1, ncol + 1):
        cell = ws.cell(1, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncol)}1"
    ws.row_dimensions[1].height = 28


def _autosize(ws: Worksheet, max_width: int = 42) -> None:
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = 12
        for cell in col[:80]:
            if cell.value is not None:
                width = max(width, min(max_width, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = width


def _paint_estimate(cell, origin: str) -> None:
    origin_l = (origin or "").lower()
    if origin_l == "measured":
        cell.fill = MEASURED_FILL
    elif origin_l:
        cell.fill = MODELLED_FILL
    cell.border = THIN


def _short_tech(technique: str) -> str:
    return {
        "con sordino": "con sord",
        "sul ponticello": "pont",
        "sul tasto": "tasto",
        "harmonics": "harm",
    }.get(technique, technique)


def _slug_tech(technique: str) -> str:
    return technique.strip().replace(" ", "_")


def _status(midi: int, ceiling: int) -> str:
    return "above_measured_ceiling" if midi > ceiling else "measured_range"


def _iowa_orch_layers(project: Project, technique: str) -> dict[tuple[str, str], Layer]:
    out: dict[tuple[str, str], Layer] = {}
    for layer in project.layers:
        if layer.technique != technique:
            continue
        coll = layer.collection.upper()
        if coll in {"ORCHIDEA", "ORCH."}:
            coll = "ORCH"
        if coll in {"IOWA", "ORCH"} and layer.dynamic in CORE_DYNS:
            out[(coll, layer.dynamic)] = layer
    return out


def _note_grid(
    layers: dict[tuple[str, str], Layer],
    technique: str,
    instrument: str = "",
) -> list[int]:
    """Chromatic grid from the instrument spec ∪ measured extrema (not the violin 55–107 box)."""
    spec = resolve_instrument(instrument) if instrument else None
    measured = [
        m
        for lg in layers.values()
        for m, c in lg.cells.items()
        if (c.origin or "").lower() == "measured"
    ]
    present = [m for lg in layers.values() for m in lg.cells]
    if technique == "harmonics" and present:
        return list(range(min(present), max(present) + 1))
    lo = spec.sounding_low if spec else (min(present) if present else 55)
    hi = spec.sounding_high if spec else (max(present) if present else 107)
    extra = measured or present
    if extra:
        lo = min(lo, min(extra))
        hi = max(hi, max(extra))
    return list(range(lo, hi + 1))


def measured_ceiling_midi(layers: dict[tuple[str, str], Layer], instrument: str = "") -> int:
    measured = [
        m
        for lg in layers.values()
        for m, c in lg.cells.items()
        if (c.origin or "").lower() == "measured"
    ]
    if measured:
        return max(measured)
    spec = resolve_instrument(instrument) if instrument else None
    return spec.sounding_high if spec else DEFAULT_MEASURED_CEILING_MIDI


def _write_kv(ws: Worksheet, headers: tuple[str, str], rows: list[tuple[str, object]]) -> None:
    ws.append(list(headers))
    _style_header(ws, 2)
    for key, value in rows:
        ws.append([key, value])
    _autosize(ws, 80)


def _write_iowa_sheet(
    ws: Worksheet,
    layer: Optional[Layer],
    midis: list[int],
    technique: str,
    dynamic: str,
    ceiling: int,
    layout: str,
    display_name: str = "Violin",
) -> None:
    if layout == "pp":
        headers = [
            "Instrument",
            "Collection",
            "Technique/state",
            "Dynamic",
            None,
            "Notes",
            "CDM - Media",
            "Estimate",
            "PI95 low",
            "PI95 high",
            "value_kind",
            "anchor_source",
            "reporting_status",
            "source_workbook",
            "MIDI",
        ]
    else:
        headers = [
            "Instrument",
            "Collection",
            "Technique/state",
            "Dynamic",
            "Source note",
            None,
            "Notes",
            "CDM - Media",
            "Estimate",
            "PI95 low",
            "PI95 high",
            "value_kind",
            "anchor_source",
            "reporting_status",
            "source_workbook",
            "MIDI",
        ]
    ws.append(headers)
    _style_header(ws, len(headers))
    for midi in midis:
        cell = layer.cells.get(midi) if layer else None
        note = midi_to_label(midi)
        origin = cell.origin if cell else ""
        status = _status(midi, ceiling)
        if layout == "pp":
            row = [
                display_name,
                "IOWA",
                technique,
                dynamic,
                None,
                note,
                cell.value if cell else None,
                origin or None,
                cell.pi95_low if cell else None,
                cell.pi95_high if cell else None,
                origin or None,
                (cell.anchor_source if cell else None) or None,
                status,
                (cell.source_workbook if cell else None) or None,
                midi,
            ]
            est_col = 8
            status_col = 13
        else:
            row = [
                display_name,
                "IOWA",
                technique,
                dynamic,
                note,
                None,
                note,
                cell.value if cell else None,
                origin or None,
                cell.pi95_low if cell else None,
                cell.pi95_high if cell else None,
                origin or None,
                (cell.anchor_source if cell else None) or None,
                status,
                (cell.source_workbook if cell else None) or None,
                midi,
            ]
            est_col = 9
            status_col = 14
        ws.append(row)
        if origin:
            _paint_estimate(ws.cell(ws.max_row, est_col), origin)
        if status == "above_measured_ceiling":
            ws.cell(ws.max_row, status_col).fill = CEILING_FILL
    _autosize(ws)


def _write_orch_sheet(
    ws: Worksheet,
    layer: Optional[Layer],
    midis: list[int],
    technique: str,
    dynamic: str,
    ceiling: int,
    display_name: str = "Violin",
) -> None:
    headers = [
        "Instrument",
        "Collection",
        "Technique/state",
        "Dynamic",
        "Source note",
        "Combined density metric",
        "Estimate",
        "PI95 low",
        "PI95 high",
        "value_kind",
        "anchor_source",
        "reporting_status",
        "source_workbook",
        "MIDI",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    for midi in midis:
        cell = layer.cells.get(midi) if layer else None
        note = midi_to_label(midi)
        origin = cell.origin if cell else ""
        status = _status(midi, ceiling)
        ws.append(
            [
                display_name,
                "ORCH",
                technique,
                dynamic,
                note,
                cell.value if cell else None,
                origin or None,
                cell.pi95_low if cell else None,
                cell.pi95_high if cell else None,
                origin or None,
                (cell.anchor_source if cell else None) or None,
                status,
                (cell.source_workbook if cell else None) or None,
                midi,
            ]
        )
        if origin:
            _paint_estimate(ws.cell(ws.max_row, 7), origin)
        if status == "above_measured_ceiling":
            ws.cell(ws.max_row, 12).fill = CEILING_FILL
    _autosize(ws)


def _layer_cdm(layers: Optional[dict], collection: str, dynamic: str, midi: int) -> Optional[float]:
    if not layers:
        return None
    layer = layers.get((collection, dynamic))
    cell = layer.cells.get(midi) if layer else None
    if cell is None or cell.value is None:
        return None
    try:
        value = float(cell.value)
    except (TypeError, ValueError):
        return None
    return value if value > 0 else None


def _mean_available(*vals: Optional[float]) -> Optional[float]:
    xs = [v for v in vals if v is not None]
    return (sum(xs) / len(xs)) if xs else None


def _media_sheet_formula(
    iowa_col: str,
    orch_col: str,
    row: str,
    instrument: str,
    technique: str,
    cfg: Optional[CalibrationConfig],
) -> str:
    """Excel formula matching the production IOWA/ORCH resolver."""
    iowa = f"{iowa_col}{row}"
    orch = f"{orch_col}{row}"
    if uses_woodwind_combination(instrument, technique):
        cfg = cfg or load_config()
        method = normalize_combination_method(cfg.combination_method)
        if method == "empirical_only":
            return f'=IF({iowa}<>"",{iowa},IF({orch}<>"",{orch},""))'
        if method == "legacy_equal_weight":
            we = float(cfg.empirical_weight)
            wt = 1.0 - we
            return (
                f'=IF(AND({iowa}<>"",{orch}<>""),{we}*{iowa}+{wt}*{orch},'
                f'IF({iowa}<>"",{iowa},IF({orch}<>"",{orch},"")))'
            )
    return f'=IF(COUNT({iowa}:{orch})=0,"",AVERAGE({iowa}:{orch}))'


def _resolved_media(
    iowa: Optional[float],
    orch: Optional[float],
    instrument: str,
    technique: str,
    cfg: Optional[CalibrationConfig],
) -> Optional[float]:
    value, _, _, _ = resolve_iowa_orch_pair(
        iowa, orch, instrument=instrument, technique=technique, cfg=cfg
    )
    return value


def _write_media(
    ws: Worksheet,
    midis: list[int],
    ceiling: int,
    instrument: str = "violin",
    sheet_names: Optional[dict[tuple[str, str], str]] = None,
    layers: Optional[dict] = None,
    technique: str = "ordinario",
    cfg: Optional[CalibrationConfig] = None,
) -> None:
    headers = [
        "Note",
        "IOWA pp",
        "ORCH pp",
        "Média pp = (IOWA pp + ORCH pp) / 2",
        "IOWA mf",
        "ORCH mf",
        "Média mf = (IOWA mf + ORCH mf) / 2",
        "IOWA ff",
        "ORCH ff",
        "Média ff = (IOWA ff + ORCH ff) / 2",
        None,
        " note",
        "Media pp",
        "Media mf",
        "Media ff",
        "IOWA -mf-pp",
        "IOWA -ff-mf",
        "IOWA -ff-pp",
        "ORCH -mf-pp",
        "ORCH -ff-mf",
        "ORCH -ff-pp",
        "FINAL_mf_minus_pp",
        "FINAL_ff_minus_mf",
        "FINAL_ff_minus_pp",
        "pp_halfgap_IOWA_ORCH",
        "mf_halfgap_IOWA_ORCH",
        "ff_halfgap_IOWA_ORCH",
        "MIDI",
        "ln Media pp",
        "ln Media mf",
        "ln Media ff",
        "reporting_status",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    names = sheet_names or collection_sheet_names(instrument)
    iowa_pp, orch_pp = _sheet_ref(names[("IOWA", "pp")]), _sheet_ref(names[("ORCH", "pp")])
    iowa_mf, orch_mf = _sheet_ref(names[("IOWA", "mf")]), _sheet_ref(names[("ORCH", "mf")])
    iowa_ff, orch_ff = _sheet_ref(names[("IOWA", "ff")]), _sheet_ref(names[("ORCH", "ff")])

    def _sub(a: Optional[float], b: Optional[float]) -> Optional[float]:
        if a is None or b is None:
            return None
        return a - b

    def _half(a: Optional[float], b: Optional[float]) -> Optional[float]:
        if a is None or b is None:
            return None
        return abs(a - b) / 2.0

    def _ln(v: Optional[float]) -> Optional[float]:
        if v is None or v <= 0:
            return None
        return math.log(v)

    for i, midi in enumerate(midis, start=2):
        r = str(i)
        if layers is not None:
            ipp = _layer_cdm(layers, "IOWA", "pp", midi)
            opp = _layer_cdm(layers, "ORCH", "pp", midi)
            imf = _layer_cdm(layers, "IOWA", "mf", midi)
            omf = _layer_cdm(layers, "ORCH", "mf", midi)
            iff_ = _layer_cdm(layers, "IOWA", "ff", midi)
            off = _layer_cdm(layers, "ORCH", "ff", midi)
            mpp = _resolved_media(ipp, opp, instrument, technique, cfg)
            mmf = _resolved_media(imf, omf, instrument, technique, cfg)
            mff = _resolved_media(iff_, off, instrument, technique, cfg)
            row = [
                midi_to_label(midi),
                ipp,
                opp,
                mpp,
                imf,
                omf,
                mmf,
                iff_,
                off,
                mff,
                None,
                midi_to_label(midi),
                mpp,
                mmf,
                mff,
                _sub(imf, ipp),
                _sub(iff_, imf),
                _sub(iff_, ipp),
                _sub(omf, opp),
                _sub(off, omf),
                _sub(off, opp),
                _sub(mmf, mpp),
                _sub(mff, mmf),
                _sub(mff, mpp),
                _half(opp, ipp),
                _half(omf, imf),
                _half(off, iff_),
                midi,
                _ln(mpp),
                _ln(mmf),
                _ln(mff),
                _status(midi, ceiling),
            ]
        else:
            media_pp = _media_sheet_formula("B", "C", r, instrument, technique, cfg)
            media_mf = _media_sheet_formula("E", "F", r, instrument, technique, cfg)
            media_ff = _media_sheet_formula("H", "I", r, instrument, technique, cfg)
            row = [
                midi_to_label(midi),
                f'=IF({iowa_pp}!{CDM_CELL[("IOWA","pp")]}{r}="","",{iowa_pp}!{CDM_CELL[("IOWA","pp")]}{r})',
                f'=IF({orch_pp}!{CDM_CELL[("ORCH","pp")]}{r}="","",{orch_pp}!{CDM_CELL[("ORCH","pp")]}{r})',
                media_pp,
                f'=IF({iowa_mf}!{CDM_CELL[("IOWA","mf")]}{r}="","",{iowa_mf}!{CDM_CELL[("IOWA","mf")]}{r})',
                f'=IF({orch_mf}!{CDM_CELL[("ORCH","mf")]}{r}="","",{orch_mf}!{CDM_CELL[("ORCH","mf")]}{r})',
                media_mf,
                f'=IF({iowa_ff}!{CDM_CELL[("IOWA","ff")]}{r}="","",{iowa_ff}!{CDM_CELL[("IOWA","ff")]}{r})',
                f'=IF({orch_ff}!{CDM_CELL[("ORCH","ff")]}{r}="","",{orch_ff}!{CDM_CELL[("ORCH","ff")]}{r})',
                media_ff,
                None,
                f"=A{r}",
                f"=D{r}",
                f"=G{r}",
                f"=J{r}",
                f'=IFERROR(E{r}-B{r},"")',
                f'=IFERROR(H{r}-E{r},"")',
                f'=IFERROR(H{r}-B{r},"")',
                f'=IFERROR(F{r}-C{r},"")',
                f'=IFERROR(I{r}-F{r},"")',
                f'=IFERROR(I{r}-C{r},"")',
                f'=IFERROR(N{r}-M{r},"")',
                f'=IFERROR(O{r}-N{r},"")',
                f'=IFERROR(O{r}-M{r},"")',
                f'=IFERROR(ABS(C{r}-B{r})/2,"")',
                f'=IFERROR(ABS(F{r}-E{r})/2,"")',
                f'=IFERROR(ABS(I{r}-H{r})/2,"")',
                midi,
                f'=IFERROR(LN(D{r}),"")',
                f'=IFERROR(LN(G{r}),"")',
                f'=IFERROR(LN(J{r}),"")',
                _status(midi, ceiling),
            ]
        ws.append(row)
        if midi > ceiling:
            ws.cell(ws.max_row, 32).fill = CEILING_FILL
    if midis:
        # Média pp/mf/ff (D, G, J) and Media pp/mf/ff (M, N, O)
        apply_media_family_conditional_format(ws, [4, 7, 10, 13, 14, 15], 2, 1 + len(midis), instrument)
    _autosize(ws, 36)
    if midis:
        chart = LineChart()
        chart.title = f"{resolve_instrument(instrument).display_name if resolve_instrument(instrument) else instrument.title()} Combined Density Metric — Media by dynamic"
        chart.style = 10
        chart.y_axis.title = "Combined Density Metric (Media)"
        chart.x_axis.title = "Note (chromatic scale)"
        chart.height = 10
        chart.width = 18
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        for col in (13, 14, 15):
            chart.add_data(Reference(ws, min_col=col, min_row=1, max_row=ws.max_row), titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, "A" + str(ws.max_row + 3))


def _write_media_uncertainty(
    ws: Worksheet,
    midis: list[int],
    layers: dict[tuple[str, str], Layer],
    technique: str,
    ceiling: int,
    arco: Optional[dict],
    media_name: str = "Violin_Media",
) -> None:
    short = _short_tech(technique)
    headers = [
        "Note",
        "MIDI",
        "Dynamic",
        f"IOWA {short}",
        "IOWA PI95 low",
        "IOWA PI95 high",
        "IOWA value_kind",
        "IOWA anchor_source",
        f"ORCH {short}",
        "ORCH PI95 low",
        "ORCH PI95 high",
        "ORCH value_kind",
        "ORCH anchor_source",
        "Media",
        "Media envelope low",
        "Media envelope high",
        "Media composition",
        "reporting_status",
        "IOWA ordinario (anchor)",
        "ORCH ordinario (anchor)",
        f"ORCH measured {short}",
        f"ratio ORCH ({short}/ord)",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    for i, midi in enumerate(midis):
        note_row = i + 2
        for dyn in CORE_DYNS:
            r = ws.max_row + 1
            iowa = layers.get(("IOWA", dyn))
            orch = layers.get(("ORCH", dyn))
            ic = iowa.cells.get(midi) if iowa else None
            oc = orch.cells.get(midi) if orch else None
            iowa_kind = ic.origin if ic else None
            orch_kind = oc.origin if oc else None
            parts = [k for k in (iowa_kind, orch_kind) if k]
            composition = " + ".join(parts) if parts else None
            media_ref = MEDIA_COL[dyn]
            iowa_ord = None
            orch_ord = None
            if arco:
                iowa_ord = (arco.get("IOWA") or {}).get(dyn, {}).get(midi)
                orch_ord = (arco.get("ORCH") or {}).get(dyn, {}).get(midi)
            orch_measured = oc.value if oc and (oc.origin or "").lower() == "measured" else None
            ws.append(
                [
                    midi_to_label(midi),
                    midi,
                    dyn,
                    ic.value if ic else None,
                    ic.pi95_low if ic else None,
                    ic.pi95_high if ic else None,
                    iowa_kind,
                    ic.anchor_source if ic else None,
                    oc.value if oc else None,
                    oc.pi95_low if oc else None,
                    oc.pi95_high if oc else None,
                    orch_kind,
                    oc.anchor_source if oc else None,
                    f'=IF({_sheet_ref(media_name)}!{media_ref}{note_row}="","",{_sheet_ref(media_name)}!{media_ref}{note_row})',
                    f'=IFERROR(MIN(E{r},J{r}),"")',
                    f'=IFERROR(MAX(F{r},K{r}),"")',
                    composition,
                    _status(midi, ceiling),
                    iowa_ord,
                    orch_ord,
                    orch_measured,
                    f'=IFERROR(U{r}/T{r},"")',
                ]
            )
            if iowa_kind:
                _paint_estimate(ws.cell(r, 7), iowa_kind)
            if orch_kind:
                _paint_estimate(ws.cell(r, 12), orch_kind)
            if midi > ceiling:
                ws.cell(r, 18).fill = CEILING_FILL
    _autosize(ws, 36)


def _write_empirical_orch(ws: Worksheet, project: Project, technique: str) -> int:
    """Principal evidence: measured Orchidea, or measured Iowa on woodwind ordinario."""
    rows = principal_empirical_rows(project, technique)
    ws["A1"] = principal_sheet_title(rows)
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = principal_sheet_subtitle(rows)
    if not rows:
        if project_is_woodwind(project) and not is_effect(technique):
            ws["A4"] = (
                f"No measured Iowa cells for '{technique}' on this woodwind. "
                "Empirical_ORCH stays empty until those recordings exist."
            )
        else:
            ws["A4"] = (
                f"No Orchidea recordings for '{technique}' in this project. "
                "The sheet is empty by design — treat the technique as prediction only."
            )
        ws["A4"].font = Font(name="Calibri", italic=True, color="C45911")
        return 0
    headers = [
        "Instrument",
        "Collection",
        "Technique",
        "Dynamic",
        "Note",
        "MIDI",
        "Combined density metric",
        "Estimate",
        "reporting_status",
        "source_workbook",
        "evidence_role",
    ]
    for col, name in enumerate(headers, 1):
        cell = ws.cell(4, col, name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws.freeze_panes = "A5"
    for rec in rows:
        ws.append(
            [
                rec["instrument"],
                rec["collection"],
                rec["technique"],
                rec["dynamic"],
                rec["note"],
                rec["midi"],
                rec["value"],
                rec["origin"],
                rec["reporting_status"],
                rec["source_workbook"],
                rec["evidence_role"],
            ]
        )
        _paint_estimate(ws.cell(ws.max_row, 8), "measured")
    _autosize(ws, 42)
    return len(rows)


def _write_evidence_map(ws: Worksheet, project: Project, technique: str) -> None:
    ws.append(
        [
            "collection",
            "dynamic",
            "n_cells",
            "n_measured",
            "n_modelled",
            "evidence_role",
            "l_donor_dynamic",
            "principal_evidence",
            "note",
        ]
    )
    _style_header(ws, 9)
    for rec in evidence_map_rows(project, technique):
        ws.append(
            [
                rec["collection"],
                rec["dynamic"],
                rec["n_cells"],
                rec["n_measured"],
                rec["n_modelled"],
                rec["evidence_role"],
                rec["l_donor_dynamic"],
                rec["principal_evidence"],
                rec["note"],
            ]
        )
        role = rec["evidence_role"]
        if rec["principal_evidence"] == "yes":
            ws.cell(ws.max_row, 8).fill = MEASURED_FILL
        elif role in {"prediction", "inherited_dynamic"}:
            ws.cell(ws.max_row, 8).fill = CEILING_FILL
        else:
            ws.cell(ws.max_row, 8).fill = MODELLED_FILL
    _autosize(ws, 80)


def _write_summary_empirical(ws: Worksheet, project: Project, technique: str) -> None:
    """Geometric means of Empirical_ORCH only — the principal published statistic."""
    rows = principal_empirical_rows(project, technique)
    iowa = principal_rows_are_iowa(rows)
    src = "IOWA" if iowa else "ORCH"
    ws["A1"] = f"Summary — Empirical_ORCH (principal evidence) · {technique}"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:D1")
    ws["A2"] = (
        (
            "Geometric means of measured Iowa cells only (no Orchidea on this woodwind). "
            "Family-transferred cousin Orchidea is omitted. "
            "Do not read the Media sheet / Summary_Measured_Range as a stronger result than this sheet."
        )
        if iowa
        else (
            "Geometric means of measured Orchidea cells only. "
            "Blank dynamic = no recordings (prediction or inherited L). "
            "Do not read the Media sheet / Summary_Measured_Range as a stronger result than this sheet."
        )
    )
    ws["A4"] = "Statistic"
    ws["B4"] = "pp"
    ws["C4"] = "mf"
    ws["D4"] = "ff"
    for col in range(1, 5):
        ws.cell(4, col).fill = HEADER_FILL
        ws.cell(4, col).font = HEADER_FONT
    def gmean(vals):
        vals = [v for v in vals if v and v > 0]
        if not vals:
            return None
        return math.exp(sum(math.log(v) for v in vals) / len(vals))

    by = {d: [] for d in CORE_DYNS}
    for rec in rows:
        if rec["dynamic"] in by:
            by[rec["dynamic"]].append(rec["value"])
    ws["A5"] = f"n measured ({src})"
    ws["A6"] = f"Geometric mean F-061 (measured {src})"
    for col, dyn in enumerate(CORE_DYNS, 2):
        ws.cell(5, col, len(by[dyn]))
        gm = gmean(by[dyn])
        ws.cell(6, col, gm if gm is not None else "—")
    ws["A8"] = SHARED_L_MEDIA_CAVEAT
    ws["A8"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A8:D10")
    ws.column_dimensions["A"].width = 55
    for letter in "BCD":
        ws.column_dimensions[letter].width = 18


def _write_summary(
    ws: Worksheet,
    n_notes: int,
    n_measured: int,
    technique: str,
    media_name: str = "Violin_Media",
    instrument: str = "violin",
    literal_stats: Optional[dict] = None,
    ceiling: int = DEFAULT_MEASURED_CEILING_MIDI,
) -> None:
    last = 1 + n_measured
    ws["A1"] = f"Summary — measured range only ({media_name} rows 2:{last}, up to MIDI {ceiling})"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:D1")
    ws["A2"] = (
        f"Every figure below excludes the {n_notes - n_measured} pitches "
        "whose ordinario anchor sits above the measured ceiling (MIDI 100)."
    )
    ws["A4"] = "Statistic"
    ws["B4"] = "pp"
    ws["C4"] = "mf"
    ws["D4"] = "ff"
    _style_header(ws, 4)
    ws["A1"] = f"Summary — measured range only ({media_name} rows 2:{last}, up to MIDI {ceiling})"
    ws["A1"].font = TITLE_FONT
    if literal_stats:
        ws["A5"] = "n measured (Iowa / empirical_only Media)"
        ws["A6"] = "Geometric mean F-061 (empirical_only)"
        for col, dyn in enumerate(CORE_DYNS, 2):
            blob = literal_stats.get(dyn) or {}
            ws.cell(5, col, blob.get("n", 0))
            gm = blob.get("gmean")
            ws.cell(6, col, gm if gm is not None else "—")
        ws["A8"] = (
            "Literal values (not Excel formulas). Media is empirical_only: a measured "
            "target cell is never replaced by a transfer estimate."
        )
        ws.column_dimensions["A"].width = 70
        return
    rows = [
        (
            "Geometric mean, Media (F-061)",
            f"=GEOMEAN({media_name}!D2:D{last})",
            f"=GEOMEAN({media_name}!G2:G{last})",
            f"=GEOMEAN({media_name}!J2:J{last})",
        ),
        (
            "Registral slope (ln per semitone)",
            f"=SLOPE({media_name}!AC2:AC{last},{media_name}!$AB$2:$AB${last})",
            f"=SLOPE({media_name}!AD2:AD{last},{media_name}!$AB$2:$AB${last})",
            f"=SLOPE({media_name}!AE2:AE{last},{media_name}!$AB$2:$AB${last})",
        ),
        (
            "Percent change per octave",
            "=(EXP(12*B6)-1)",
            "=(EXP(12*C6)-1)",
            "=(EXP(12*D6)-1)",
        ),
        (
            "Correlation of ln(density) with MIDI",
            f"=CORREL({media_name}!AC2:AC{last},{media_name}!$AB$2:$AB${last})",
            f"=CORREL({media_name}!AD2:AD{last},{media_name}!$AB$2:$AB${last})",
            f"=CORREL({media_name}!AE2:AE{last},{media_name}!$AB$2:$AB${last})",
        ),
        (
            "n notes with a Media value",
            f"=COUNT({media_name}!D2:D{last})",
            f"=COUNT({media_name}!G2:G{last})",
            f"=COUNT({media_name}!J2:J{last})",
        ),
    ]
    for i, row in enumerate(rows):
        for col, value in enumerate(row, 1):
            ws.cell(5 + i, col, value)
    ws["A11"] = "Interpretation"
    ws["A11"].font = Font(bold=True, name="Calibri")
    notes = [
        "Media is AVERAGE of the collections that have a value on that pitch (never halved when only one exists).",
        "IOWA has no recorded effect samples: every IOWA cell is IOWA ordinario × exp(L).",
        "Philharmonia and McGill are not in this file; they remain context-only in the STE Lab workbooks.",
    ]
    if technique == "harmonics":
        iid = (instrument or "").lower()
        if iid == "viola":
            notes.append("Harmonics are published from C5 (MIDI 72) only. Nothing below the first practical chromatic harmonic is modelled.")
        elif iid == "cello":
            notes.append("Harmonics are published from C4 (MIDI 60) only. Nothing below the first practical chromatic harmonic is modelled.")
        elif iid == "double_bass":
            notes.append("Harmonics are published from E3 (MIDI 52) only. Nothing below the first measured chromatic harmonic is modelled.")
        else:
            notes.append("Harmonics are published from G5 (MIDI 79) only. Nothing below the first practical sounding harmonic is modelled.")
    if technique == "sul ponticello":
        notes.append("Orchidea has measured sul ponticello at mf only. pp and ff are L-from-mf transfers onto those dynamics' arco curves.")
    if technique == "sul tasto":
        notes.append("No Orchidea sul tasto recordings. L is Philharmonia piano tasto / Philharmonia piano ordinario (context-grade).")
    if technique == "con sordino":
        if (instrument or "").lower() == "viola":
            notes.append("No Orchidea viola sordino. L is McGill muted / McGill non-vibrato (context-grade, not Empirical_ORCH).")
        else:
            notes.append("Orchidea con sordino is measured where a research workbook exists; remaining ORCH cells are modelled_ORCHIDEA_anchored.")
    for i, line in enumerate(notes):
        ws.cell(12 + i, 1, "— " + line)
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18


def _acoustic_table_value(ic, oc, qa_rec, instrument_id, technique, cfg=None):
    if qa_rec is not None and "final_value" in qa_rec:
        return qa_rec.get("final_value")
    iowa = ic.value if ic is not None else None
    orch = oc.value if oc is not None else None
    return _resolved_media(iowa, orch, instrument_id, technique, cfg)


def _write_acoustic_table(
    ws: Worksheet,
    midis: list[int],
    layers: dict[tuple[str, str], Layer],
    technique: str,
    ceiling: int,
    filename: str,
    instrument_id: str = "violin",
    media_name: str = "Violin_Media",
    final_by_cell: Optional[dict] = None,
    cfg: Optional[CalibrationConfig] = None,
) -> tuple[int, int, int]:
    headers = [
        "instrument_id",
        "note_sounding",
        "midi_sounding",
        "dynamic",
        "value",
        "value_kind",
        "unit",
        "cell_status",
        "source_system",
        "source_file",
        "source_column",
        "source_hash",
        "transform_policy",
        "uncertainty",
        "validation_status",
        "notes",
        "note_written_optional",
        "midi_written_optional",
        "transposition_semitones_optional",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    accepted = review_required = rejected = 0
    mu_row = 2
    for i, midi in enumerate(midis):
        note_row = i + 2
        for dyn in CORE_DYNS:
            iowa = layers.get(("IOWA", dyn))
            orch = layers.get(("ORCH", dyn))
            ic = iowa.cells.get(midi) if iowa else None
            oc = orch.cells.get(midi) if orch else None
            if not ic and not oc:
                mu_row += 1
                continue
            above = midi > ceiling
            kinds = " + ".join(k for k in ((ic.origin if ic else None), (oc.origin if oc else None)) if k)
            qa_rec = (final_by_cell or {}).get((dyn, int(midi)))
            if qa_rec:
                status = qa_rec.get("validation_status") or (
                    "rejected" if above else "accepted"
                )
            else:
                status = "review_required" if above else "accepted"
            notes = f"{technique}; {kinds}; see Media_Uncertainty row {mu_row}"
            if qa_rec and qa_rec.get("provenance"):
                notes = f"{notes}; provenance={qa_rec['provenance']}"
            ws.append(
                [
                    instrument_id,
                    midi_to_label(midi),
                    midi,
                    dyn,
                    _acoustic_table_value(ic, oc, qa_rec, instrument_id, technique, cfg),
                    "combined_density_metric",
                    "cdm_technique_sustain_v1",
                    "register_extrapolated" if above else "media_unique",
                    "IOWA+ORCH technique",
                    filename,
                    f"{media_name}!{MEDIA_COL[dyn]}{note_row}",
                    None,
                    "identity_v1",
                    "low" if above else "medium",
                    status,
                    notes,
                    None,
                    None,
                    None,
                ]
            )
            if status == "accepted":
                accepted += 1
            elif status == "review_required":
                review_required += 1
            else:
                rejected += 1
                if above:
                    ws.cell(ws.max_row, 8).fill = CEILING_FILL
            mu_row += 1
    _autosize(ws, 48)
    return accepted, review_required, rejected


def export_zenodo_workbook(
    project: Project,
    path: Path,
    technique: str,
    arco: Optional[dict] = None,
    ceiling_midi: int = DEFAULT_MEASURED_CEILING_MIDI,
    notes: str = "",
) -> Path:
    path = Path(path)
    layers = _iowa_orch_layers(project, technique)
    instrument = next((lg.instrument for lg in project.layers if lg.instrument), "violin")
    midis = _note_grid(layers, technique, instrument)
    if ceiling_midi == DEFAULT_MEASURED_CEILING_MIDI:
        ceiling_midi = measured_ceiling_midi(layers, instrument)
    n_measured = sum(1 for m in midis if m <= ceiling_midi)
    generated = datetime.now(timezone.utc)
    filename = path.name
    spec = resolve_instrument(instrument) or resolve_instrument("violin")
    display = spec.display_name if spec else "Violin"
    iid = spec.instrument_id if spec else "violin"
    names = collection_sheet_names(iid)
    media_name = media_sheet_name(iid)
    ww_ordinario = orchestral_group(iid) == "woodwinds" and not is_effect(technique)
    if ww_ordinario:
        arco_book = f"{display.replace(' ', '_')}_Zenodo_collections_media.xlsx"
    else:
        arco_book = {
            "viola": "VIOLA_Zenodo_collections_Arco_normal.xlsx",
            "cello": "CELLO_Zenodo_collections_media.xlsx",
            "double_bass": "DOUBLEBASS_Zenodo_collections_media.xlsx",
        }.get(iid, "VIOLIN_Zenodo_collections_Arco_normal.xlsx")

    wb = Workbook()
    first = True
    for coll, dyn in (("IOWA", "pp"), ("IOWA", "mf"), ("IOWA", "ff"), ("ORCH", "pp"), ("ORCH", "mf"), ("ORCH", "ff")):
        name = names[(coll, dyn)]
        ws = wb.active if first else wb.create_sheet(name)
        if first:
            ws.title = name
            first = False
        if coll == "IOWA":
            _write_iowa_sheet(
                ws, layers.get((coll, dyn)), midis, technique, dyn, ceiling_midi,
                "pp" if dyn == "pp" else "mf", display_name=display,
            )
        else:
            _write_orch_sheet(
                ws, layers.get((coll, dyn)), midis, technique, dyn, ceiling_midi,
                display_name=display,
            )

    _write_empirical_orch(wb.create_sheet("Empirical_ORCH"), project, technique)
    _write_evidence_map(wb.create_sheet("Evidence_Map"), project, technique)
    _write_summary_empirical(wb.create_sheet("Summary_Empirical_ORCH"), project, technique)

    media = wb.create_sheet(media_name)
    cfg = load_config() if ww_ordinario else None
    _write_media(
        media,
        midis,
        ceiling_midi,
        instrument,
        sheet_names=names,
        layers=layers,
        technique=technique,
        cfg=cfg,
    )

    orch_measured = sum(
        1
        for dyn in CORE_DYNS
        for cell in (layers.get(("ORCH", dyn)).cells.values() if layers.get(("ORCH", dyn)) else [])
        if (cell.origin or "").lower() == "measured"
    )
    iowa_n = {d: len(layers[("IOWA", d)].cells) if ("IOWA", d) in layers else 0 for d in CORE_DYNS}
    orch_n = {d: len(layers[("ORCH", d)].cells) if ("ORCH", d) in layers else 0 for d in CORE_DYNS}
    ww_ordinario = orchestral_group(iid) == "woodwinds" and not is_effect(technique)

    _write_media_uncertainty(
        wb.create_sheet("Media_Uncertainty"), midis, layers, technique, ceiling_midi, arco,
        media_name=media_name,
    )
    literal_stats = None
    if ww_ordinario:
        literal_stats = {}
        for dyn in CORE_DYNS:
            iowa = layers.get(("IOWA", dyn))
            vals = [
                c.value
                for c in (iowa.cells.values() if iowa else [])
                if c.value and c.value > 0 and (c.origin or "").lower() == "measured"
            ]
            if vals:
                literal_stats[dyn] = {
                    "n": len(vals),
                    "gmean": math.exp(sum(math.log(v) for v in vals) / len(vals)),
                }
            else:
                literal_stats[dyn] = {"n": 0, "gmean": None}
    _write_summary(
        wb.create_sheet("Summary_Measured_Range"), len(midis), n_measured, technique,
        media_name=media_name, instrument=iid, literal_stats=literal_stats, ceiling=ceiling_midi,
    )

    readme = wb.create_sheet("README")
    _write_kv(
        readme,
        ("field", "value"),
        [
            (f"{display} {technique} — empirical sheet first, then completed grid", None),
            ("Generated (UTC)", generated.isoformat()),
            ("Source workbooks", f"STE Lab batch from compiled_density_metrics_research.xlsx + {arco_book}"),
            (
                "Principal evidence",
                (
                    "Sheet Empirical_ORCH and Summary_Empirical_ORCH: measured Iowa cells on this woodwind. "
                    "Family-transferred cousin Orchidea is omitted there."
                    if ww_ordinario
                    else "Sheet Empirical_ORCH and Summary_Empirical_ORCH: Orchidea cells tagged measured only."
                ),
            ),
            (
                "Purpose",
                (
                    f"{display} {technique}: IOWA = measured Iowa (principal evidence). "
                    "ORCH = family transfer from the same-family donor (not principal; not a recording of this instrument). "
                    f"Completed-grid Media remains for operational use. Counts IOWA {iowa_n}; ORCH {orch_n}."
                    if ww_ordinario
                    else (
                        f"{display} {technique}: IOWA = IOWA_ordinario × exp(L) (model, not replication). "
                        "ORCH = measured F-061 where a research workbook exists, else ORCH_ordinario × exp(L). "
                        f"Completed-grid Media remains for operational use. Counts IOWA {iowa_n}; ORCH {orch_n}."
                    )
                ),
            ),
            ("Statistical rule", SHARED_L_MEDIA_CAVEAT if not ww_ordinario else WOODWIND_PRINCIPAL_SUBTITLE),
            (
                "IOWA column",
                (
                    "Measured Iowa where compiled research exists. This is the principal evidence for woodwind ordinario."
                    if ww_ordinario
                    else "Modelled throughout. IOWA holds no effect recordings."
                ),
            ),
            (
                "ORCH column",
                (
                    "Family-transferred cousin Orchidea (tagged family_transfer). "
                    "Not measured on this instrument. Kept off Empirical_ORCH."
                    if ww_ordinario
                    else (
                        f"{orch_measured} measured cells; remaining cells modelled_ORCHIDEA_anchored."
                        + (" Sul tasto has no Orchidea measurements: L is Philharmonia piano (context-grade)." if technique == "sul tasto" else "")
                        + (
                            " Con sordino has no Orchidea viola measurements: L is McGill muted (context-grade)."
                            if iid == "viola" and technique == "con sordino"
                            else ""
                        )
                    )
                ),
            ),
            ("Measured ceiling", f"MIDI {ceiling_midi}. Notes above this are above_measured_ceiling / register_extrapolated."),
            ("Sheet alignment", f"All data sheets carry the same {len(midis)} notes in the same MIDI order, so {media_name} row references stay aligned."),
            ("Recommended Zenodo role", "Deposit as internal dataset metadata / summary statistics, alongside the ZIP packages for the note-level outputs."),
            ("Do not use as", "A replacement for raw note-level outputs, logs, segmentation metadata, or source audio."),
            ("Notes", notes or project.notes),
        ],
    )
    readme["A1"].font = TITLE_FONT

    _write_kv(
        wb.create_sheet("Zenodo_File_metadata"),
        ("field", "value"),
        [
            ("recommended_filename", filename),
            (
                "title",
                (
                    f"{display} {technique} Combined Density Metric (IOWA measured + family-transfer ORCH)"
                    if ww_ordinario
                    else f"{display} {technique} Combined Density Metric (IOWA modelled + ORCHIDEA measured/modelled)"
                ),
            ),
            ("resource_role", "internal dataset metadata / summary statistics"),
            ("instrument", display),
            ("technique", technique),
            ("sample_type", "sustains"),
            ("dynamics", "pp; mf; ff"),
            ("source_collections", "IOWA + ORCHIDEA"),
            ("metric", "Combined Density Metric (spectral density, F-061 spectral_mass)"),
            (
                "statistic_principal",
                (
                    "Empirical_ORCH geometric mean of measured Iowa cells."
                    if ww_ordinario
                    else "Empirical_ORCH geometric mean of measured Orchidea cells."
                ),
            ),
            (
                "statistic_completed_grid",
                (
                    f"{media_name}: combination={cfg.combination_method if cfg else 'empirical_only'}; "
                    "default empirical_only keeps Iowa when present and uses transfer only in a gap. "
                    "A blend is COMBINED_ESTIMATE, not a measurement."
                    if ww_ordinario
                    else f"{media_name} = AVERAGE of available IOWA and ORCH values (shared L — not independent replication)."
                ),
            ),
            ("notes_covered", len(midis)),
            ("notes_measured_anchor", n_measured),
            ("notes_extrapolated_anchor", len(midis) - n_measured),
            ("measured_ceiling_midi", ceiling_midi),
            ("orch_cells_measured", orch_measured),
            ("source_workbook_A", arco_book),
            ("source_workbook_B", "compiled_density_metrics_research.xlsx (Orchidea / Philharmonia / McGill trees)"),
            ("evidence_tier", "STE Lab transfer — not the Extrapol v7.1 PRED pages"),
            ("generated_utc", generated.isoformat()),
        ],
    )

    vocab = wb.create_sheet("Controlled Vocabulary")
    vocab.append(["Category", "Value", "Meaning / use", "Recommended in this file?"])
    _style_header(vocab, 4)
    for row in [
        ("Sheet type", "source collection sheet", "One sheet for one sound collection and one dynamic level.", "Yes"),
        ("Sheet type", "derived mean sheet", "Sheet where ORCH and IOWA values are averaged by formula.", "Yes"),
        ("Sheet type", "identification sheet", "README or Zenodo metadata sheet documenting the file.", "Yes"),
        (
            "Aggregation",
            (
                f"{cfg.combination_method if cfg else 'empirical_only'} / resolved Media"
                if ww_ordinario
                else "arithmetic mean"
            ),
            (
                "Media follows the selected combination method. empirical_only keeps a valid Iowa cell; "
                "equal_weight / legacy_equal_weight mix both inputs (w_T = 1 - w_E). "
                "Unsupported modes are rejected."
                if ww_ordinario
                else "(ORCH + IOWA) / 2 for each note and dynamic level."
            ),
            "Yes",
        ),
        (
            "Aggregation",
            "median",
            "Equals the midpoint with two collections; not the stated aim here.",
            "No — woodwind Media uses the selected combination method"
            if ww_ordinario
            else "No — this file uses AVERAGE()",
        ),
        ("Dynamic", "pp", "Pianissimo dynamic level.", "Yes"),
        ("Dynamic", "mf", "Mezzo-forte dynamic level.", "Yes"),
        ("Dynamic", "ff", "Fortissimo dynamic level.", "Yes"),
        ("Collection", "ORCH", "Orchidea collection.", "Yes"),
        ("Collection", "IOWA", "Iowa collection.", "Yes"),
        ("Technique/state", technique, "The technique documented by this file.", "Yes"),
        ("Zenodo status", "internal metadata", "Detailed file-level or note-level information deposited inside a Zenodo record.", "Yes"),
        ("Estimate status", "register_extrapolated", "Ordinario anchor above the measured ceiling; excluded from aggregates.", "Yes — flagged, never silently pooled"),
    ]:
        vocab.append(list(row))
    _autosize(vocab, 70)

    _write_kv(
        wb.create_sheet("WorkbookMeta"),
        ("key", "value"),
        [
            ("acoustic_pitch_basis", "sounding_concert"),
            ("schema_version", "1.2.0"),
            ("template", "instrument_profiles_template.xlsx"),
            ("principal_evidence_sheet", "Empirical_ORCH"),
            ("source_media_sheet", f"{media_name} (completed grid, not principal evidence)"),
            ("curated_for", "Textural Density Phase 1a importer"),
            ("generated_at", generated.isoformat()),
            ("technique", technique),
            (
                "iowa_source",
                (
                    "STE Lab: measured Iowa compiled spectral_mass on this instrument"
                    if ww_ordinario
                    else "STE Lab: IOWA ordinario × exp(L)"
                ),
            ),
            (
                "orch_source",
                (
                    "STE Lab: orchidea_family_transfer_estimate (two_log_ratio); not an Orchidea recording of this instrument"
                    if ww_ordinario
                    else "STE Lab: Orchidea spectral_mass where measured, else ORCH ordinario × exp(L)"
                ),
            ),
            ("anchor_provenance", f"{arco_book} {media_name}"),
            (
                "acoustic_table_binding",
                (
                    f"literal values copied from {media_name}"
                    if ww_ordinario
                    else f"live formulas referencing {media_name}"
                ),
            ),
            ("importer_filter", "skip rows where cell_status = register_extrapolated, or validation_status = review_required"),
        ],
    )

    registry = wb.create_sheet("Registry")
    registry.append(
        [
            "instrument_id",
            "display_name",
            "family",
            "subfamily",
            "transposition",
            "sounding_range_low_midi",
            "sounding_range_high_midi",
            "comfortable_range_low_midi",
            "comfortable_range_high_midi",
            "profile_status",
            "source_type",
            "uncertainty",
            "source_notes",
            "supported_techniques",
            "aliases",
        ]
    )
    _style_header(registry, 15)
    registry.append(
        [
            iid,
            display,
            orchestral_group(iid),
            spec.family if spec else "",
            spec.transposition if spec else 0,
            spec.sounding_low if spec else 55,
            spec.sounding_high if spec else 93,
            spec.comfortable_low if spec else 55,
            spec.comfortable_high if spec else 88,
            "empirical_profile",
            "external_acoustic_metadata",
            "medium",
            (
                f"Sparse {technique} Combined Density Metric table for {display} at pp/mf/ff, sounding pitch. "
                "IOWA values are measured; ORCH is family transfer."
                if ww_ordinario
                else f"Sparse {technique} Combined Density Metric table for {display} at pp/mf/ff, sounding pitch. IOWA values are modelled, not measured."
            ),
            ("ordinario" if ww_ordinario else "arco|pizzicato|tremolo|harmonic|mute"),
            iid,
        ]
    )
    _autosize(registry)

    final_by_cell = None
    woodwind_calib = None
    if ww_ordinario:
        from .calibration import (
            build_final_rows,
            build_validation_tables,
            curves_from_project,
            origins_from_project,
            qa_index,
            transfer_estimates_from_layers,
        )

        cfg = load_config()
        measured_curves = curves_from_project(project, technique)
        layer_origins = origins_from_project(project, technique)
        transfer_by_dyn = transfer_estimates_from_layers(layers, cfg)
        val_rows, offsets, _excl = build_validation_tables(measured_curves, transfer_by_dyn, cfg)
        final_rows = []
        for dyn in CORE_DYNS:
            final_rows.extend(
                build_final_rows(
                    instrument=iid,
                    dynamic=dyn,
                    empirical=measured_curves.get(("IOWA", dyn), {}),
                    transfer=transfer_by_dyn[dyn].curve,
                    cfg=cfg,
                    transfer_method=transfer_by_dyn[dyn].method,
                )
            )
        final_by_cell = qa_index(final_rows)
        woodwind_calib = (measured_curves, transfer_by_dyn, val_rows, offsets, final_rows, cfg, layer_origins)

    accepted, review_required, rejected = _write_acoustic_table(
        wb.create_sheet("AcousticTable"),
        midis,
        layers,
        technique,
        ceiling_midi,
        filename,
        instrument_id=iid,
        media_name=media_name,
        final_by_cell=final_by_cell,
        cfg=cfg if ww_ordinario else None,
    )

    prov = wb.create_sheet("Provenance")
    prov.append(
        [
            "instrument_id",
            "source_type",
            "citation",
            "source_url_or_identifier",
            "upstream_system",
            "upstream_version",
            "analysis_profile_hash",
            "import_run_id",
            "import_date",
            "operator",
            "transform_policy",
            "transform_parameters",
            "rows_accepted",
            "rows_review_required",
            "rows_rejected",
            "notes",
        ]
    )
    _style_header(prov, 16)
    media_note = (
        f"{display} {technique}: principal evidence = Empirical_ORCH (measured Iowa). "
        "ORCH is family transfer, not a recording of this instrument. "
        "Media is the resolved empirical-priority calibration curve, not generally an average."
        if ww_ordinario
        else f"{display} {technique}: principal evidence = Empirical_ORCH. IOWA modelled; ORCH measured where available. Media is a completed grid, not a second experiment."
    )
    prov.append(
        [
            iid,
            "external_acoustic_metadata",
            media_note,
            str(path),
            (
                "IOWA+ORCH woodwind analysis corpus / STE Lab"
                if ww_ordinario
                else "IOWA+ORCH string analysis corpus / STE Lab"
            ),
            "STE Lab 1.0",
            None,
            None,
            generated.date().isoformat(),
            project.operator or "STE Lab batch",
            "identity_v1",
            "CDM midpoint pass-through; no rescaling",
            accepted,
            review_required,
            rejected,
            notes or project.notes,
        ]
    )
    _autosize(prov, 48)

    aliases = wb.create_sheet("Aliases")
    aliases.append(["instrument_id", "alias", "alias_kind"])
    _style_header(aliases, 3)
    aliases.append([iid, iid, "registry"])
    _autosize(aliases)

    if woodwind_calib:
        from .calibration_export import attach_calibration_sheets

        measured_curves, transfer_by_dyn, val_rows, offsets, final_rows, cfg, layer_origins = woodwind_calib
        attach_calibration_sheets(
            wb,
            instrument=display,
            measured=measured_curves,
            transfer_by_dyn=transfer_by_dyn,
            final_rows=final_rows,
            validation_rows=val_rows,
            offsets=offsets,
            cfg=cfg,
            ceiling=ceiling_midi,
            grid_lo=midis[0] if midis else 0,
            grid_hi=midis[-1] if midis else 0,
            notes=notes or project.notes,
            origins=layer_origins,
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _project_from_zenodo_sheets(wb, technique: str) -> Project:
    """Rebuild IOWA/ORCH layers from an existing deposit workbook (values unchanged)."""
    from .notes import parse_pitch
    from .session import layer_from_identity

    if any(n.startswith("VIOLA") for n in wb.sheetnames):
        instrument = "viola"
    elif any(n.startswith("Cello") or n.startswith("CELLO") for n in wb.sheetnames):
        instrument = "cello"
    else:
        instrument = "violin"
    names = collection_sheet_names(instrument)
    project = Project(title=f"{instrument.title()} {technique}", operator="STE Lab evidence refresh")
    for (coll, dyn), sheet in names.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        hmap = {str(h).strip().lower() if h else "": i + 1 for i, h in enumerate(headers)}

        def find(*names):
            for name in names:
                for key, idx in hmap.items():
                    if name in key:
                        return idx
            return None

        cdm = find("cdm - media", "combined density metric", "cdm")
        est = find("estimate", "value_kind")
        midi_c = find("midi")
        note_c = find("notes", "source note", "note")
        if cdm is None:
            continue
        layer = layer_from_identity(instrument, coll, technique, dyn)
        for r in range(2, ws.max_row + 1):
            raw = ws.cell(r, cdm).value
            if raw is None or raw == "" or (isinstance(raw, str) and raw.startswith("=")):
                continue
            try:
                value = float(raw)
            except (TypeError, ValueError):
                continue
            if value <= 0:
                continue
            midi = ws.cell(r, midi_c).value if midi_c else None
            try:
                midi = int(float(midi)) if midi is not None else None
            except (TypeError, ValueError):
                midi = None
            note = ws.cell(r, note_c).value if note_c else None
            pitch = parse_pitch(f"midi{midi}") if midi is not None else parse_pitch(note)
            if pitch is None:
                continue
            origin = ws.cell(r, est).value if est else None
            layer.place(pitch, value, origin or "modelled", overwrite="overwrite")
        project.add_layer(layer)
    return project


def attach_evidence_sheets(path: Path, technique: str) -> Path:
    """Add Empirical_ORCH / Evidence_Map / Summary_Empirical_ORCH without rewriting CDM values."""
    from openpyxl import load_workbook

    path = Path(path)
    wb = load_workbook(path)
    project = _project_from_zenodo_sheets(wb, technique)
    for name in ("Empirical_ORCH", "Evidence_Map", "Summary_Empirical_ORCH"):
        if name in wb.sheetnames:
            del wb[name]
    _write_empirical_orch(wb.create_sheet("Empirical_ORCH", 6), project, technique)
    _write_evidence_map(wb.create_sheet("Evidence_Map", 7), project, technique)
    _write_summary_empirical(wb.create_sheet("Summary_Empirical_ORCH", 8), project, technique)
    if "README" in wb.sheetnames:
        ws = wb["README"]
        ws["A14"] = "principal_evidence"
        ws["B14"] = (
            "Empirical_ORCH / Summary_Empirical_ORCH (measured Iowa on this woodwind). Media is a completed grid."
            if project_is_woodwind(project) and not is_effect(technique)
            else "Empirical_ORCH / Summary_Empirical_ORCH (measured Orchidea only). Media is a completed grid."
        )
    wb.save(path)
    return path
