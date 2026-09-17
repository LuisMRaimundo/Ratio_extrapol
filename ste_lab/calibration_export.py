"""Excel sheets for measured / transfer / validation / final calibration / methodology."""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Optional

from openpyxl.styles import Font
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from . import __version__
from .calibration import (
    COLLECTION_CODEC,
    ORIGIN_TO_PROVENANCE,
    PROVENANCE_MEASURED,
    PROVENANCE_TRANSFERRED,
    CalibrationConfig,
    TransferEstimate,
    acceptance_counts,
    agreement_metrics,
    collection_offset,
    default_config,
)
from .catalog import family_donor_instrument, orchestral_group, resolve_instrument
from .notes import midi_to_label

TITLE_FONT = Font(bold=True, name="Calibri", size=14, color="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)


def _header(ws: Worksheet, titles: list[str]) -> None:
    from openpyxl.styles import PatternFill

    ws.append(titles)
    fill = PatternFill("solid", fgColor="1F4E79")
    for col, _ in enumerate(titles, 1):
        cell = ws.cell(1, col)
        cell.fill = fill
        cell.font = HEADER_FONT
    ws.freeze_panes = "A2"


def _autosize(ws: Worksheet, cap: int = 42) -> None:
    from openpyxl.utils import get_column_letter

    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = min(cap, max((len(str(c.value or "")) for c in col), default=8) + 2)
        ws.column_dimensions[letter].width = max(10, width)


def _provenance_for_collection(
    coll: str,
    instrument: str,
    origins: Optional[dict[str, str]] = None,
) -> tuple[str, str]:
    """Return (Measured_Data provenance, layer origin) for one collection."""
    origin = (origins or {}).get(coll.upper()) or (origins or {}).get(coll)
    if origin:
        key = str(origin).strip().lower().replace(" ", "_")
        if key in ORIGIN_TO_PROVENANCE:
            return ORIGIN_TO_PROVENANCE[key], origin
        if "transfer" in key:
            return PROVENANCE_TRANSFERRED, origin
        return PROVENANCE_MEASURED, origin
    if str(coll).upper() == "ORCH" and orchestral_group(instrument) == "woodwinds":
        return PROVENANCE_TRANSFERRED, "orchidea_family_transfer_estimate"
    return PROVENANCE_MEASURED, "measured"


def write_measured_data(
    ws: Worksheet,
    measured: dict[tuple[str, str], dict[int, float]],
    instrument: str,
    origins: Optional[dict[str, str]] = None,
) -> None:
    _header(
        ws,
        [
            "instrument",
            "collection",
            "dynamic",
            "note",
            "midi",
            "value",
            "provenance",
            "source_type",
            "source_collection",
            "source_instrument",
            "target_instrument",
            "codec",
        ],
    )
    spec = resolve_instrument(instrument)
    host_id = spec.instrument_id if spec else instrument
    donor = family_donor_instrument(instrument)
    for (coll, dyn), curve in sorted(measured.items()):
        provenance, origin = _provenance_for_collection(coll, instrument, origins)
        transferred = provenance != PROVENANCE_MEASURED
        source_instrument = donor if transferred and donor else host_id
        source_type = "observation" if not transferred else origin
        for midi in sorted(curve):
            val = curve[midi]
            if val is None or val <= 0:
                continue
            ws.append(
                [
                    instrument,
                    coll,
                    dyn,
                    midi_to_label(midi),
                    midi,
                    val,
                    provenance,
                    source_type,
                    coll,
                    source_instrument,
                    host_id,
                    COLLECTION_CODEC.get(str(coll).upper(), ""),
                ]
            )
    _autosize(ws)


def write_transfer_estimates(
    ws: Worksheet,
    instrument: str,
    by_dyn: dict[str, TransferEstimate],
    empirical_by_dyn: dict[str, dict[int, float]],
) -> None:
    _header(
        ws,
        [
            "instrument",
            "dynamic",
            "note",
            "midi",
            "transfer_estimate",
            "provenance",
            "transfer_method",
            "L_instr",
            "L_coll",
            "legacy_scale",
            "extrapolation_distance_semitones",
            "source_pitch_min",
            "source_pitch_max",
            "L_instr_status",
            "fitted_on",
            "validation_excluded",
            "note_on_method",
        ],
    )
    for dyn, est in sorted(by_dyn.items()):
        emp = empirical_by_dyn.get(dyn) or {}
        lo = min(emp) if emp else None
        hi = max(emp) if emp else None
        for midi, val in sorted(est.curve.items()):
            if lo is None or hi is None:
                dist = 0
            elif lo <= midi <= hi:
                dist = 0
            else:
                dist = lo - midi if midi < lo else midi - hi
            if est.method == "legacy_target_calibrated":
                prov = "COMBINED_ESTIMATE"
            elif dist > 0:
                prov = "TRANSFERRED_EXTRAPOLATED"
            else:
                prov = "TRANSFERRED"
            ws.append(
                [
                    instrument,
                    dyn,
                    midi_to_label(midi),
                    midi,
                    val,
                    prov,
                    est.method,
                    est.L_instr.get(midi),
                    est.L_coll.get(midi),
                    est.scale_legacy,
                    dist,
                    lo,
                    hi,
                    est.L_instr_status,
                    ";".join(est.fitted_on),
                    ";".join(est.validation_excluded),
                    est.notes,
                ]
            )
    _autosize(ws, 70)


def write_validation(
    ws: Worksheet,
    rows: list[dict],
    offsets: list[dict],
    excluded: list[str],
) -> None:
    ws["A1"] = "Validation — genuine independent target-instrument measurements only"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (
        "Collections used to fit transfer parameters are excluded: "
        + (", ".join(excluded) if excluded else "(none)")
        + ". Measured-vs-measured collection offset is reported separately "
        "and is not a transfer-quality score. Philharmonia sources are MP3. "
        "ln_RMSE = sqrt(mean((ln(predicted)-ln(observed))**2)). "
        "Offset-normalised metrics are diagnostic measures of pitch-dependent "
        "curve shape. The estimated collection offset is NOT applied to the "
        "production calibration."
    )
    ws["A4"] = "Model comparison (per dynamic, not pooled)"
    headers = [
        "dynamic",
        "validation_collection",
        "model",
        "n",
        "r",
        "mae",
        "rmse",
        "ln_rmse",
        "collection_log_offset",
        "collection_scale_factor",
        "shape_ln_rmse",
        "shape_ln_mae",
    ]
    for col, name in enumerate(headers, 1):
        ws.cell(5, col, name)
        ws.cell(5, col).font = HEADER_FONT
    r = 6
    for rec in rows:
        for col, key in enumerate(headers, 1):
            ws.cell(r, col, rec.get(key))
        r += 1
    skipped = [s for rec in rows for s in (rec.get("skip_reasons") or [])]
    if skipped:
        r += 1
        ws.cell(r, 1, "Skipped non-positive or invalid overlap pairs (not silently dropped)")
        r += 1
        ws.cell(r, 1, f"n_skipped_nonpositive={len(skipped)}")
        r += 1
    r += 1
    ws.cell(r, 1, "Collection offset (median A/B on measured overlap) — not transfer quality")
    r += 2
    for col, name in enumerate(["dynamic", "collection_a", "collection_b", "n", "median_ratio_a_over_b"], 1):
        ws.cell(r, col, name)
        ws.cell(r, col).font = HEADER_FONT
    r += 1
    for rec in offsets:
        for col, key in enumerate(["dynamic", "collection_a", "collection_b", "n", "median_ratio"], 1):
            ws.cell(r, col, rec.get(key))
        r += 1
    _autosize(ws, 48)


def write_final_calibration(ws: Worksheet, rows: list[dict]) -> None:
    headers = [
        "Instrument",
        "Dynamic",
        "Note",
        "MIDI",
        "Empirical_Value",
        "Empirical_Source",
        "Transfer_Estimate",
        "Transfer_Source",
        "Final_Value",
        "Combination_Method",
        "Empirical_Weight",
        "Transfer_Weight",
        "Provenance",
        "Extrapolation_Distance_Semitones",
        "Validation_N",
        "Validation_MAE",
        "Validation_RMSE",
        "Validation_r",
        "Confidence",
        "QA_Flag",
        "Automatic_Acceptance",
        "Acceptance_Override",
        "Accepted_Final",
        "Final_Acceptance",
        "Validation_Status",
        "Absolute_Difference",
        "Relative_Difference",
        "Configured_Combination_Method",
        "Configured_Empirical_Weight",
        "Configured_Transfer_Weight",
    ]
    _header(ws, headers)
    for rec in rows:
        ws.append(
            [
                rec.get("instrument"),
                rec.get("dynamic"),
                rec.get("note"),
                rec.get("midi"),
                rec.get("empirical_value"),
                rec.get("empirical_source"),
                rec.get("transfer_estimate"),
                rec.get("transfer_source"),
                rec.get("final_value"),
                rec.get("combination_method"),
                rec.get("empirical_weight"),
                rec.get("transfer_weight"),
                rec.get("provenance"),
                rec.get("extrapolation_distance_semitones"),
                rec.get("validation_n"),
                rec.get("validation_mae"),
                rec.get("validation_rmse"),
                rec.get("validation_r"),
                rec.get("confidence"),
                rec.get("qa_flag"),
                rec.get("automatic_acceptance"),
                rec.get("acceptance_override"),
                rec.get("accepted_final"),
                rec.get("final_acceptance"),
                rec.get("validation_status"),
                rec.get("absolute_difference"),
                rec.get("relative_difference"),
                rec.get("configured_combination_method"),
                rec.get("configured_empirical_weight"),
                rec.get("configured_transfer_weight"),
            ]
        )
    _autosize(ws)


def write_methodology(
    ws: Worksheet,
    *,
    instrument: str,
    cfg: CalibrationConfig,
    transfer_by_dyn: dict[str, TransferEstimate],
    validation_rows: list[dict],
    excluded: list[str],
    ceiling: int,
    grid_lo: int,
    grid_hi: int,
    notes: str = "",
    final_rows: Optional[list[dict]] = None,
) -> None:
    ws["A1"] = f"Methodology — {instrument}"
    ws["A1"].font = TITLE_FONT
    lines = [
        ("Generated (UTC)", datetime.now(timezone.utc).isoformat()),
        ("STE Lab version", __version__),
        ("Instrument", instrument),
        (
            "Genuine measurements",
            "Iowa, Philharmonia, and McGill compiled spectral_mass on this instrument. "
            "There is no Orchidea recording of this woodwind.",
        ),
        (
            "ORCH / orchidea_family_transfer_estimate",
            "Model-derived estimate. Not an Orchidea measurement of the target instrument. "
            "The legacy column name ORCH is deprecated and maps to this estimate.",
        ),
        ("Transfer method", cfg.transfer_method),
        ("Combination method", cfg.combination_method),
        ("Empirical weight", cfg.empirical_weight),
        ("Transfer weight", cfg.transfer_weight),
        (
            "Default combination rule",
            "Supported methods: empirical_only, equal_weight, legacy_equal_weight. "
            "empirical_only never replaces a valid target measurement; transfer fills only a gap. "
            "equal_weight mixes 0.5/0.5 when both exist. "
            "legacy_equal_weight mixes with w_T = 1 - w_E. "
            "fixed_weight and validation_optimised are rejected until specified. "
            "Per-row Empirical_Weight / Transfer_Weight are the effective weights; "
            "Configured_* columns are the yaml defaults. "
            "A blended output is COMBINED_ESTIMATE, not MEASURED.",
        ),
        (
            "Interpolation / Fill",
            "fill_missing interpolates inside the measured MIDI span and may continue "
            "outside that span up to max_extrap_semitones (default 12). "
            "PCHIP requires ≥3 known points and labels exterior cells extrapolated_pchip "
            "(not ridge regression). Linear holds endpoints outside the span and labels "
            "those cells extrapolated_hold, not interpolated. "
            "Technique-transfer L still holds the edge L; that is a different operator. "
            "Fill is not used in the two-ratio transfer.",
        ),
        (
            "Two-ratio transfer",
            "L_coll(midi) = ln(ORCH_clarinet / IOWA_clarinet); "
            "L_instr(midi) = ln(IOWA_target / IOWA_clarinet); "
            "ORCH_target(midi) = ORCH_clarinet(midi) × exp(clip(L_instr, ±3)). "
            "L is held at the edge outside the anchor span.",
        ),
        (
            "Legacy scalar (deprecated)",
            "scale = median(IOWA_target / ORCH_donor) on overlap. "
            "Outputs are COMBINED_ESTIMATE. Reproduction only.",
        ),
        (
            "ln_RMSE",
            "sqrt(mean((ln(predicted) - ln(observed))**2)). "
            "This is RMSE in log space, not the mean squared log error.",
        ),
        (
            "Shape-normalised validation",
            "collection_log_offset = median(ln(y)-ln(y_hat)); "
            "collection_scale_factor = exp(offset); "
            "shape_ln_rmse = sqrt(mean((ln(y)-ln(y_hat)-offset)**2)). "
            "Diagnostic only — the offset is not applied to production values.",
        ),
        ("Extrapolation accept / review (st)", f"{cfg.accept_max_semitones} / {cfg.review_max_semitones}"),
        ("Allow long extrapolation", cfg.allow_long_extrapolation),
        ("Allow REVIEW_REQUIRED into production", cfg.allow_review_required),
        (
            "Production acceptance",
            "HIGH and MODERATE are production-eligible. "
            "REVIEW_REQUIRED remains visible but Accepted_Final is false unless "
            "allow_review_required is true (QA flag is preserved). REJECTED is never accepted.",
        ),
        ("Measured ceiling MIDI", ceiling),
        ("Note grid", f"MIDI {grid_lo}–{grid_hi} (instrument spec ∪ measured extrema)"),
        ("Validation excluded collections", ", ".join(excluded) or "(none)"),
        ("Validation min_n", cfg.validation_min_n),
        ("Philharmonia codec", COLLECTION_CODEC["PHIL"]),
        ("McGill codec", COLLECTION_CODEC["MCGILL"]),
        ("Iowa codec", COLLECTION_CODEC["IOWA"]),
        (
            "Family-transfer statement",
            "Family-transfer estimates are model-derived estimates and do not constitute "
            "measurements from the target collection.",
        ),
        ("Notes", notes),
    ]
    for dyn, est in sorted(transfer_by_dyn.items()):
        lines.append((f"Transfer {dyn}", f"method={est.method}; L_instr={est.L_instr_status}; anchors L_coll={est.n_L_coll_anchors} L_instr={est.n_L_instr_anchors}; n_cells={len(est.curve)}; {est.notes}"))
    if validation_rows:
        lines.append(("Validation rows", len(validation_rows)))
    if final_rows:
        counts = acceptance_counts(final_rows)
        lines.append(("rows_accepted", counts["rows_accepted"]))
        lines.append(("rows_review_required", counts["rows_review_required"]))
        lines.append(("rows_rejected", counts["rows_rejected"]))
        lines.append(("rows_total", counts["rows_total"]))
    ws["A3"] = "field"
    ws["B3"] = "value"
    for i, (k, v) in enumerate(lines, 4):
        ws.cell(i, 1, k)
        ws.cell(i, 2, v)
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 110


def attach_calibration_sheets(
    wb: Workbook,
    *,
    instrument: str,
    measured: dict[tuple[str, str], dict[int, float]],
    transfer_by_dyn: dict[str, TransferEstimate],
    final_rows: list[dict],
    validation_rows: list[dict],
    offsets: list[dict],
    cfg: Optional[CalibrationConfig] = None,
    ceiling: int,
    grid_lo: int,
    grid_hi: int,
    notes: str = "",
    origins: Optional[dict[str, str]] = None,
) -> None:
    cfg = cfg or default_config()
    excluded = sorted({c for est in transfer_by_dyn.values() for c in est.validation_excluded})
    empirical = {( "IOWA", d): measured.get(("IOWA", d), {}) for d in ("pp", "mf", "ff")}
    write_measured_data(wb.create_sheet("Measured_Data"), measured, instrument, origins=origins)
    write_transfer_estimates(
        wb.create_sheet("Transfer_Estimates"),
        instrument,
        transfer_by_dyn,
        {d: measured.get(("IOWA", d), {}) for d in ("pp", "mf", "ff")},
    )
    write_validation(wb.create_sheet("Validation"), validation_rows, offsets, excluded)
    write_final_calibration(wb.create_sheet("Final_Calibration"), final_rows)
    write_methodology(
        wb.create_sheet("Methodology"),
        instrument=instrument,
        cfg=cfg,
        transfer_by_dyn=transfer_by_dyn,
        validation_rows=validation_rows,
        excluded=excluded,
        ceiling=ceiling,
        grid_lo=grid_lo,
        grid_hi=grid_hi,
        notes=notes,
        final_rows=final_rows,
    )
    unused = empirical  # keep linter quiet if measured already written
    del unused
