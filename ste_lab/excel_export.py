"""Write a STE-style workbook: collection sheets, Media, AcousticTable, provenance."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .catalog import ORIGINS, TECHNIQUES, normalize_origin
from .evidence import (
    SHARED_L_MEDIA_CAVEAT,
    evidence_map_rows,
    is_effect,
    principal_empirical_rows,
    principal_rows_are_iowa,
    principal_sheet_subtitle,
    principal_sheet_title,
    project_is_woodwind,
)
from .media_format import apply_media_family_conditional_format
from .notes import midi_to_label, parse_pitch
from .qa import Flag
from .session import Layer, Project, layer_from_identity

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
TITLE_FONT = Font(bold=True, name="Calibri", size=14, color="1F4E79")
THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

ORIGIN_FILL = {
    "measured": PatternFill("solid", fgColor="C6EFCE"),
    "generated": PatternFill("solid", fgColor="BDD7EE"),
    "interpolated": PatternFill("solid", fgColor="C5D9F1"),
    "technique_transfer": PatternFill("solid", fgColor="E2D5F1"),
    "family_transfer": PatternFill("solid", fgColor="D9EAD3"),
    "collection_anchored": PatternFill("solid", fgColor="FFF2CC"),
    "modelled": PatternFill("solid", fgColor="FCE4D6"),
    "modelled_iowa_anchored": PatternFill("solid", fgColor="FCE4D6"),
    "modelled_orchidea_anchored": PatternFill("solid", fgColor="FCE4D6"),
    "modelled_on_extrapolated_anchor": PatternFill("solid", fgColor="F4B183"),
    "extrapolated_ridge": PatternFill("solid", fgColor="F8CBAD"),
    "extrapolated_polynomial": PatternFill("solid", fgColor="F4B183"),
    "manual": PatternFill("solid", fgColor="D9D9D9"),
}

SEV_FILL = {
    "critical": PatternFill("solid", fgColor="FF6B6B"),
    "high": PatternFill("solid", fgColor="F4B183"),
    "medium": PatternFill("solid", fgColor="FFE599"),
    "low": PatternFill("solid", fgColor="D9EAD3"),
}


def _style_header(ws: Worksheet, ncol: int) -> None:
    for col in range(1, ncol + 1):
        cell = ws.cell(1, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncol)}1"


def _autosize(ws: Worksheet, max_width: int = 42) -> None:
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = 10
        for cell in col[:80]:
            if cell.value is not None:
                width = max(width, min(max_width, len(str(cell.value)) + 2))
        ws.column_dimensions[letter].width = width


def _write_kv(ws: Worksheet, rows: list[tuple[str, object]]) -> None:
    ws.append(["field", "value"])
    _style_header(ws, 2)
    for key, value in rows:
        ws.append([key, value])
    _autosize(ws)


def _paint_origin(cell, origin: str) -> None:
    fill = ORIGIN_FILL.get((origin or "").lower().replace(" ", "_"))
    if fill:
        cell.fill = fill
    cell.border = THIN


def write_layer_sheet(ws: Worksheet, layer: Layer) -> None:
    headers = [
        "Instrument",
        "Collection",
        "Technique/state",
        "Dynamic",
        "Source note",
        "MIDI",
        "Combined density metric",
        "Estimate",
        "PI95 low",
        "PI95 high",
        "value_kind",
        "anchor_source",
        "reporting_status",
        "source_workbook",
        "Comment",
        "Flags",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    for cell in layer.sorted_cells():
        ws.append(
            [
                layer.instrument,
                layer.collection,
                layer.technique,
                layer.dynamic,
                cell.note_label,
                cell.midi,
                cell.value,
                cell.origin,
                cell.pi95_low,
                cell.pi95_high,
                cell.origin,
                cell.anchor_source,
                cell.reporting_status,
                cell.source_workbook,
                cell.comment,
                "; ".join(cell.flags),
            ]
        )
        _paint_origin(ws.cell(ws.max_row, 8), cell.origin)
        if cell.reporting_status == "above_measured_ceiling":
            ws.cell(ws.max_row, 13).fill = PatternFill("solid", fgColor="F4B183")
    _autosize(ws)


def write_media_sheet(ws: Worksheet, layers: list[Layer], media_layers: list[Layer]) -> None:
    """Wide Media table: one row per MIDI, collections side by side, origin beside each media."""
    dynamics = sorted({lg.dynamic for lg in layers})
    collections = sorted({lg.collection for lg in layers})
    midis = sorted({m for lg in layers for m in lg.cells})
    headers = ["Note", "MIDI"]
    for dyn in dynamics:
        for coll in collections:
            headers.append(f"{coll} {dyn}")
            headers.append(f"{coll} {dyn} origin")
        headers.append(f"Media {dyn}")
        headers.append(f"Media {dyn} origin")
    headers.append("reporting_status")
    ws.append(headers)
    _style_header(ws, len(headers))

    index = {(lg.collection, lg.dynamic, midi): lg.cells[midi] for lg in layers for midi in lg.cells}
    media_index = {(lg.dynamic, midi): lg.cells[midi] for lg in media_layers for midi in lg.cells}

    for midi in midis:
        row: list[object] = [None, midi]
        note = None
        for dyn in dynamics:
            for coll in collections:
                cell = index.get((coll, dyn, midi))
                if cell:
                    note = note or cell.note_label
                    row.extend([cell.value, cell.origin])
                else:
                    row.extend([None, None])
            mcell = media_index.get((dyn, midi))
            if mcell:
                note = note or mcell.note_label
                row.extend([mcell.value, mcell.origin])
            else:
                row.extend([None, None])
        statuses = []
        for dyn in dynamics:
            for coll in collections:
                cell = index.get((coll, dyn, midi))
                if cell and cell.reporting_status:
                    statuses.append(cell.reporting_status)
        row.append("above_measured_ceiling" if "above_measured_ceiling" in statuses else ("measured_range" if statuses else ""))
        row[0] = note
        ws.append(row)
        r = ws.max_row
        # origin columns are even-ish; paint any header containing 'origin'
        for c, header in enumerate(headers, 1):
            if "origin" in header.lower() and ws.cell(r, c).value:
                _paint_origin(ws.cell(r, c), str(ws.cell(r, c).value))
    media_cols = [i + 1 for i, h in enumerate(headers) if h.startswith("Media ") and "origin" not in h.lower()]
    if midis and media_cols and (layers or media_layers):
        instrument = (layers or media_layers)[0].instrument
        apply_media_family_conditional_format(ws, media_cols, 2, 1 + len(midis), instrument)
    _autosize(ws)

    if midis and media_layers:
        chart = LineChart()
        chart.title = f"{layers[0].instrument} CDM — Media by dynamic"
        chart.style = 10
        chart.y_axis.title = "Combined Density Metric"
        chart.x_axis.title = "Note (chromatic)"
        chart.height = 10
        chart.width = 18
        # Plot Media columns
        media_cols = [i + 1 for i, h in enumerate(headers) if h.startswith("Media ") and "origin" not in h.lower()]
        cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        for col in media_cols[:3]:
            data = Reference(ws, min_col=col, min_row=1, max_row=ws.max_row)
            chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, "A" + str(ws.max_row + 3))


def write_acoustic_table(ws: Worksheet, layers: list[Layer], final_by_cell: dict | None = None) -> None:
    headers = [
        "instrument_id",
        "note_sounding",
        "midi_sounding",
        "dynamic",
        "technique",
        "collection",
        "value",
        "value_kind",
        "unit",
        "origin",
        "cell_status",
        "uncertainty",
        "validation_status",
        "notes",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    for layer in layers:
        for cell in layer.sorted_cells():
            extrapolated = cell.reporting_status == "above_measured_ceiling" or "extrapolated_anchor" in (cell.origin or "")
            modelled = cell.origin.lower() != "measured"
            qa_rec = (final_by_cell or {}).get((layer.dynamic, int(cell.midi)))
            if qa_rec:
                status = qa_rec.get("validation_status") or ("review_required" if extrapolated else "accepted")
            else:
                status = "review_required" if extrapolated else "accepted"
            ws.append(
                [
                    layer.instrument,
                    cell.note_label,
                    cell.midi,
                    layer.dynamic,
                    layer.technique,
                    layer.collection,
                    cell.value,
                    "combined_density_metric",
                    "cdm_technique_sustain_v1",
                    cell.origin,
                    "register_extrapolated" if extrapolated else "media_unique",
                    cell.uncertainty or ("low" if extrapolated else ("medium" if modelled else "low")),
                    status,
                    cell.comment,
                ]
            )
            _paint_origin(ws.cell(ws.max_row, 10), cell.origin)
    _autosize(ws)


def write_media_uncertainty(ws: Worksheet, layers: list[Layer]) -> None:
    headers = [
        "Note",
        "MIDI",
        "Dynamic",
        "value",
        "PI95 low",
        "PI95 high",
        "value_kind",
        "anchor_source",
        "reporting_status",
        "collection",
        "technique",
    ]
    ws.append(headers)
    _style_header(ws, len(headers))
    for layer in layers:
        for cell in layer.sorted_cells():
            ws.append(
                [
                    cell.note_label,
                    cell.midi,
                    layer.dynamic,
                    cell.value,
                    cell.pi95_low,
                    cell.pi95_high,
                    cell.origin,
                    cell.anchor_source,
                    cell.reporting_status,
                    layer.collection,
                    layer.technique,
                ]
            )
            _paint_origin(ws.cell(ws.max_row, 7), cell.origin)
    _autosize(ws)


def write_summary_measured_range(ws: Worksheet, layers: list[Layer]) -> None:
    """Geometric mean and registral slope on measured_range cells only (F-061 style)."""
    import math

    import numpy as np

    ws["A1"] = "Summary — measured range only (excludes above_measured_ceiling / extrapolated anchors)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Every figure below excludes pitches whose ordinario anchor was filled by extrapol_data."
    ws.append([])
    ws.append(["Statistic", "pp", "mf", "ff"])
    _style_header(ws, 4)
    # header was written on row 1 by _style_header — rewrite title
    ws["A1"] = "Summary — measured range only (excludes above_measured_ceiling / extrapolated anchors)"
    ws["A1"].font = TITLE_FONT

    by_dyn: dict[str, list[tuple[int, float]]] = {"pp": [], "mf": [], "ff": []}
    for layer in layers:
        if "media" not in layer.name.lower() and not layer.collection.startswith("+"):
            continue
        for cell in layer.sorted_cells():
            if cell.reporting_status == "above_measured_ceiling" or cell.value <= 0:
                continue
            if layer.dynamic in by_dyn:
                by_dyn[layer.dynamic].append((cell.midi, cell.value))
    if not any(by_dyn.values()):
        # fall back: use all non-ceiling cells grouped by dynamic
        for layer in layers:
            for cell in layer.sorted_cells():
                if cell.reporting_status == "above_measured_ceiling" or cell.value <= 0:
                    continue
                if layer.dynamic in by_dyn:
                    by_dyn[layer.dynamic].append((cell.midi, cell.value))

    def _stats(pairs: list[tuple[int, float]]) -> dict:
        if not pairs:
            return {"gmean": None, "slope": None, "oct": None, "corr": None, "n": 0}
        vals = np.array([v for _, v in pairs], dtype=float)
        midis = np.array([m for m, _ in pairs], dtype=float)
        logs = np.log(vals)
        gmean = float(np.exp(np.mean(logs)))
        if len(pairs) < 3:
            return {"gmean": gmean, "slope": None, "oct": None, "corr": None, "n": len(pairs)}
        slope = float(np.polyfit(midis, logs, 1)[0])
        oct_chg = float(math.exp(slope * 12) - 1)
        corr = float(np.corrcoef(midis, logs)[0, 1])
        return {"gmean": gmean, "slope": slope, "oct": oct_chg, "corr": corr, "n": len(pairs)}

    stats = {d: _stats(by_dyn[d]) for d in ("pp", "mf", "ff")}

    def row(label: str, key: str) -> None:
        ws.append([label, stats["pp"][key], stats["mf"][key], stats["ff"][key]])

    row("Geometric mean, Media (F-061)", "gmean")
    row("Registral slope (ln per semitone)", "slope")
    row("Percent change per octave", "oct")
    row("Correlation of ln(density) with MIDI", "corr")
    row("n notes with a Media value (measured range)", "n")
    ws.append([])
    ws.append(
        [
            "Interpretation",
            "Exclude register_extrapolated rows from aggregates. "
            "A mute ratio or dynamic order that only appears after those tails are included is not a physical result.",
        ]
    )
    _autosize(ws, 70)


def write_empirical_orch_sheet(ws: Worksheet, project: Project, technique: str) -> None:
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
                f"No Orchidea recordings for '{technique}'. "
                "Empty by design — prediction only (e.g. sul tasto)."
            )
        return
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
        _paint_origin(ws.cell(ws.max_row, 8), "measured")
    _autosize(ws, 42)


def write_evidence_map_sheet(ws: Worksheet, project: Project, technique: str) -> None:
    headers = [
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
    ws.append(headers)
    _style_header(ws, len(headers))
    for rec in list(evidence_map_rows(project, technique)) + list(getattr(project, "extra_evidence", None) or []):
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
        if rec["principal_evidence"] == "yes":
            ws.cell(ws.max_row, 8).fill = ORIGIN_FILL.get("measured")
        elif rec["evidence_role"] in {"prediction", "inherited_dynamic"}:
            ws.cell(ws.max_row, 8).fill = SEV_FILL["high"]
        elif rec["evidence_role"] == "context":
            ws.cell(ws.max_row, 8).fill = ORIGIN_FILL.get("manual")
    _autosize(ws, 80)


def write_qa_sheet(ws: Worksheet, flags: list[Flag]) -> None:
    headers = ["severity", "code", "layer", "note", "midi", "approach", "message", "recommendation"]
    ws.append(headers)
    _style_header(ws, len(headers))
    for flag in flags:
        ws.append(
            [
                flag.severity,
                flag.code,
                flag.layer_name,
                flag.note,
                flag.midi,
                flag.approach,
                flag.message,
                flag.recommendation,
            ]
        )
        fill = SEV_FILL.get(flag.severity)
        if fill:
            ws.cell(ws.max_row, 1).fill = fill
    _autosize(ws, 70)


def export_workbook(project: Project, path: Path, flags: list[Flag], media_layers: list[Layer] | None = None) -> Path:
    path = Path(path)
    wb = Workbook()

    log = wb.active
    log.title = "Session_Log"
    log.append(["Turn", "Timestamp", "Event"])
    _style_header(log, 3)
    for i, line in enumerate(project.log, 1):
        log.append([i, datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"), line])
    _autosize(log)

    readme = wb.create_sheet("README")
    _write_kv(
        readme,
        [
            ("title", project.title),
            ("Generated", datetime.now(timezone.utc).isoformat()),
            ("Operator", project.operator),
            ("Purpose", "Derived Combined Density Metric layers with explicit origin tags."),
            ("Statistical rule", SHARED_L_MEDIA_CAVEAT),
            (
                "Principal evidence",
                (
                    "Sheet Empirical_ORCH: measured Iowa cells on this woodwind. "
                    "Family-transferred cousin Orchidea is omitted there."
                    if project.layers and project_is_woodwind(project)
                    else "Sheet Empirical_ORCH: Orchidea cells tagged measured. Sul tasto and inherited pp/ff are omitted there."
                ),
            ),
            ("Do not use as", "Replacement for raw note-level outputs, logs, or source audio."),
            ("Quality", "See QA_Flags. modelled_on_extrapolated_anchor / above_measured_ceiling rows are excluded from Summary_Measured_Range."),
            (
                "Caveat 1",
                (
                    "Woodwind Media is the resolved empirical-priority calibration (empirical_only): "
                    "a measured Iowa cell is kept; an acceptable transfer is used only where Iowa is absent. "
                    "It is not generally (IOWA + ORCH) / 2. legacy_equal_weight reproduces the old average."
                    if project.layers and project_is_woodwind(project)
                    else "Equal-weight Media of a measurement and an IOWA extrapolation is not a statistical interval and is not independent replication of L."
                ),
            ),
            ("Caveat 2", "If the IOWA transfer was calibrated on a META set that already contains the other collection, the average is partly circular."),
            ("Caveat 3", "Registral and technique-level trends are supportable; individual note values are not when collections disagree more than the technique effect."),
            ("Notes", project.notes),
            (
                "Transfer relations",
                "L_Validation / L_Spread / Summary_Validation check the transfer assumption across collections. "
                "Default transfer_field.mode=single is unchanged production L. See manual §§5.10–5.12.",
            ),
        ],
    )
    readme["A1"].font = TITLE_FONT

    for layer in project.layers:
        name = layer.sheet_name()
        base = name
        n = 2
        while name in wb.sheetnames:
            name = f"{base[:28]}_{n}"
            n += 1
        ws = wb.create_sheet(name)
        write_layer_sheet(ws, layer)

    tech = next((lg.technique for lg in project.layers if lg.technique), "")
    write_empirical_orch_sheet(wb.create_sheet("Empirical_ORCH"), project, tech)
    write_evidence_map_sheet(wb.create_sheet("Evidence_Map"), project, tech)

    media_layers = media_layers or []
    if project.layers:
        media_ws = wb.create_sheet(f"{project.layers[0].instrument[:12]}_Media"[:31])
        write_media_sheet(media_ws, project.layers, media_layers)

    if media_layers:
        for layer in media_layers:
            name = layer.sheet_name()
            if name in wb.sheetnames:
                name = f"Media_{layer.dynamic}"[:31]
            write_layer_sheet(wb.create_sheet(name), layer)

    from .relation_export import attach_relation_sheets

    attach_relation_sheets(
        wb,
        pairwise=getattr(project, "validation_pairwise", None),
        spread=getattr(project, "validation_spread", None),
        summary_lines=getattr(project, "validation_summary", None) or None,
    )
    write_media_uncertainty(wb.create_sheet("Media_Uncertainty"), project.layers + media_layers)
    write_summary_measured_range(wb.create_sheet("Summary_Measured_Range"), media_layers or project.layers)
    woodwind_calib = None
    final_by_cell = None
    if project.layers and project_is_woodwind(project):
        from .calibration import (
            acceptance_counts,
            build_final_rows,
            build_validation_tables,
            curves_from_project,
            load_config,
            origins_from_project,
            qa_index,
            transfer_estimates_from_layers,
        )
        from .zenodo_export import _iowa_orch_layers, measured_ceiling_midi, _note_grid

        cfg = load_config()
        tech = next((lg.technique for lg in project.layers if lg.technique), "ordinario")
        layers = _iowa_orch_layers(project, tech)
        inst = project.layers[0].instrument
        midis = _note_grid(layers, tech, inst)
        ceiling = measured_ceiling_midi(layers, inst)
        measured_curves = curves_from_project(project, tech)
        layer_origins = origins_from_project(project, tech)
        transfer_by_dyn = transfer_estimates_from_layers(layers, cfg)
        val_rows, offsets, _excl = build_validation_tables(measured_curves, transfer_by_dyn, cfg)
        final_rows = []
        for dyn in ("pp", "mf", "ff"):
            final_rows.extend(
                build_final_rows(
                    instrument=inst,
                    dynamic=dyn,
                    empirical=measured_curves.get(("IOWA", dyn), {}),
                    transfer=transfer_by_dyn[dyn].curve,
                    cfg=cfg,
                    transfer_method=transfer_by_dyn[dyn].method,
                )
            )
        final_by_cell = qa_index(final_rows)
        woodwind_counts = acceptance_counts(final_rows)
        woodwind_calib = (inst, measured_curves, transfer_by_dyn, val_rows, offsets, final_rows, cfg, ceiling, midis, woodwind_counts, layer_origins)
    write_acoustic_table(wb.create_sheet("AcousticTable"), project.layers + media_layers, final_by_cell)
    write_qa_sheet(wb.create_sheet("QA_Flags"), flags)

    vocab = wb.create_sheet("Controlled Vocabulary")
    vocab.append(["Category", "Value", "Meaning / use", "Recommended"])
    _style_header(vocab, 4)
    for tech in TECHNIQUES:
        vocab.append(["Technique", tech, "Playing state / technique label", "Yes"])
    for origin in ORIGINS:
        vocab.append(["Origin", origin, "How the number was obtained", "Yes — never relabel fill as measured"])
    vocab.append(
        [
            "Aggregation",
            "empirical_only / resolved Media",
            "Woodwind Media keeps a target measurement; transfer fills Iowa-absent cells only. (A+B)/2 is legacy_equal_weight only.",
            "Yes for woodwind ordinario",
        ]
    )
    vocab.append(["Aggregation", "arithmetic mean", "(A+B)/2 when both collections exist", "String books / legacy_equal_weight"])
    vocab.append(["Aggregation", "median", "Equals the midpoint with two values; not the aim here", "No unless chosen"])
    vocab.append(["Pitch basis", "sounding_concert", "Working MIDI is concert pitch", "Yes"])
    vocab.append(["Estimate status", "register_extrapolated", "Ordinario anchor filled above the measured ceiling", "Flagged, never silently pooled"])
    vocab.append(["Reporting", "measured_range", "At or below the measured-anchor ceiling", "Yes — use in aggregates"])
    vocab.append(["Reporting", "above_measured_ceiling", "Above MIDI ceiling or modelled_on_extrapolated_anchor", "Keep visible; exclude from Summary"])
    _autosize(vocab)

    meta = wb.create_sheet("WorkbookMeta")
    _write_kv(
        meta,
        [
            ("schema_version", "1.4.0"),
            ("acoustic_pitch_basis", "sounding_concert"),
            ("template", "STE Lab"),
            ("curated_for", "technique and family-register extrapolation"),
            ("generated_at", datetime.now(timezone.utc).isoformat()),
            ("operator", project.operator),
            ("title", project.title),
        ],
    )

    prov = wb.create_sheet("Provenance")
    _write_kv(
        prov,
        [
            ("title", project.title),
            ("operator", project.operator),
            ("created_at", project.created_at),
            ("layers", len(project.layers)),
            ("transform_policy", "identity for measured; tagged fill/transfer otherwise"),
            *(
                [
                    ("rows_accepted", woodwind_calib[9]["rows_accepted"]),
                    ("rows_review_required", woodwind_calib[9]["rows_review_required"]),
                    ("rows_rejected", woodwind_calib[9]["rows_rejected"]),
                ]
                if woodwind_calib
                else []
            ),
            ("notes", project.notes),
            *(
                [("transfer_field", project.pooled_provenance)]
                if getattr(project, "pooled_provenance", "")
                else []
            ),
        ],
    )

    if woodwind_calib:
        from .calibration_export import attach_calibration_sheets

        inst, measured_curves, transfer_by_dyn, val_rows, offsets, final_rows, cfg, ceiling, midis, _counts, layer_origins = woodwind_calib
        attach_calibration_sheets(
            wb,
            instrument=inst,
            measured=measured_curves,
            transfer_by_dyn=transfer_by_dyn,
            final_rows=final_rows,
            validation_rows=val_rows,
            offsets=offsets,
            cfg=cfg,
            ceiling=ceiling,
            grid_lo=midis[0] if midis else 0,
            grid_hi=midis[-1] if midis else 0,
            notes=project.notes,
            origins=layer_origins,
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def import_workbook(path: Path, project: Project) -> int:
    """Load STE-style or viola-style collection sheets into layers."""
    import pandas as pd

    path = Path(path)
    xl = pd.ExcelFile(path)
    skip = {
        "readme",
        "claude log",
        "session_log",
        "controlled vocabulary",
        "workbookmeta",
        "registry",
        "provenance",
        "aliases",
        "zenodo_file_metadata",
        "qa_flags",
        "acoustictable",
        "media_uncertainty",
        "summary_measured_range",
        "measured_data",
        "transfer_estimates",
        "validation",
        "final_calibration",
        "methodology",
        "l_validation",
        "l_spread",
        "summary_validation",
        "relations_inventory",
    }
    loaded = 0
    for sheet in xl.sheet_names:
        key = sheet.lower().replace(" ", "_")
        if key in skip or "media" in key:
            continue
        df = pd.read_excel(path, sheet_name=sheet)
        cols = {str(c).strip().lower(): c for c in df.columns}
        note_col = next((cols[k] for k in cols if k in {"notes", "source note", "note"} or (k.startswith("note") and "midi" not in k)), None)
        value_col = next(
            (
                cols[k]
                for k in cols
                if k in {"cdm - media", "combined density metric"} or k.startswith("cdm")
            ),
            None,
        )
        if value_col is None:
            value_col = next(
                (
                    cols[k]
                    for k in cols
                    if "density" in k or (k == "value")
                ),
                None,
            )
        if note_col is None or value_col is None:
            continue

        def _series(key: str, default: str = "") -> str:
            col = next((cols[k] for k in cols if key in k), None)
            if col is None or df.empty:
                return default
            val = df[col].dropna()
            return str(val.iloc[0]) if len(val) else default

        layer = layer_from_identity(
            _series("instrument", "custom"),
            _series("collection", sheet),
            _series("technique", "custom"),
            _series("dynamic", "mf"),
            name=sheet,
        )
        origin_col = next((cols[k] for k in cols if k in {"origin", "estimate", "value_kind"}), None)
        pi_lo = next((cols[k] for k in cols if "pi95" in k and "low" in k), None)
        pi_hi = next((cols[k] for k in cols if "pi95" in k and "high" in k), None)
        status_col = next((cols[k] for k in cols if k == "reporting_status"), None)
        anchor_col = next((cols[k] for k in cols if "anchor_source" in k), None)
        src_col = next((cols[k] for k in cols if k == "source_workbook"), None)
        for _, row in df.iterrows():
            pitch = parse_pitch(row.get(note_col))
            try:
                value = float(row.get(value_col))
            except (TypeError, ValueError):
                continue
            if not pitch:
                continue
            origin = "collection_anchored"
            if origin_col is not None and not pd.isna(row.get(origin_col)):
                origin = normalize_origin(str(row.get(origin_col)))
            def _opt(col):
                if col is None or pd.isna(row.get(col)):
                    return None
                try:
                    return float(row.get(col))
                except (TypeError, ValueError):
                    return None

            layer.place(
                pitch,
                value,
                origin,
                overwrite="overwrite",
                pi95_low=_opt(pi_lo),
                pi95_high=_opt(pi_hi),
                reporting_status="" if status_col is None or pd.isna(row.get(status_col)) else str(row.get(status_col)),
                anchor_source="" if anchor_col is None or pd.isna(row.get(anchor_col)) else str(row.get(anchor_col)),
                source_workbook="" if src_col is None or pd.isna(row.get(src_col)) else str(row.get(src_col)),
            )
        if layer.cells:
            layer.range_low = min(layer.cells)
            layer.range_high = max(layer.cells)
            project.add_layer(layer)
            loaded += 1
    return loaded
