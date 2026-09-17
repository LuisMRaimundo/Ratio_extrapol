"""Technique transfer, family-register mapping, and missing-note fill."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import numpy as np
from scipy.interpolate import PchipInterpolator

from .catalog import resolve_instrument
from .notes import midi_to_label, parse_pitch
from .session import Cell, Layer, layer_from_identity


class InsufficientFillData(ValueError):
    """Fill method rejected because the layer has too few known points."""


@dataclass
class Anchor:
    midi: int
    source_value: float
    target_value: float


@dataclass
class TransferResult:
    layer: Layer
    n_copied: int = 0
    n_interpolated: int = 0
    n_extrapolated: int = 0
    n_transferred: int = 0
    messages: list[str] | None = None

    def __post_init__(self) -> None:
        if self.messages is None:
            self.messages = []


def _safe_log_ratio(target: float, source: float) -> Optional[float]:
    if source <= 0 or target <= 0:
        return None
    return float(np.log(target / source))


def _interp_log_ratio(anchors: list[Anchor], midis: list[int]) -> dict[int, float]:
    """L(midi) inside the overlap; constant hold outside.

    PCHIP with extrapolate=True follows the last cubic segment off the
    measured pairs. A small wiggle near the top of the instrument then
    dives to L ≈ −20…−40; after the safety clip that becomes
    source × exp(−3) ≈ 0.05 × arco — the near-zero high notes.
    Holding the edge L keeps the last observed effect/ordinario ratio.
    """
    usable = [a for a in anchors if _safe_log_ratio(a.target_value, a.source_value) is not None]
    if not usable:
        return {}
    xs = np.array([a.midi for a in usable], dtype=float)
    ys = np.array([_safe_log_ratio(a.target_value, a.source_value) for a in usable], dtype=float)
    order = np.argsort(xs)
    xs, ys = xs[order], ys[order]
    if len(xs) == 1:
        return {m: float(ys[0]) for m in midis}
    lo, hi = float(xs[0]), float(xs[-1])
    y_lo, y_hi = float(ys[0]), float(ys[-1])
    if len(xs) == 2:
        interior = {m: float(np.interp(m, xs, ys)) for m in midis if lo <= m <= hi}
    else:
        fn = PchipInterpolator(xs, ys, extrapolate=False)
        interior = {m: float(fn(m)) for m in midis if lo <= m <= hi}
    out: dict[int, float] = {}
    for m in midis:
        if m < lo:
            out[m] = y_lo
        elif m > hi:
            out[m] = y_hi
        else:
            out[m] = interior[m]
    return out


def technique_transfer(
    source: Layer,
    anchors: list[Anchor],
    target_technique: str,
    method: str = "log_ratio",
    collection: Optional[str] = None,
    copy_anchors: bool = True,
    log_ratio_field: Optional[dict[int, float]] = None,
    origin_override: Optional[str] = None,
    field_comment: str = "",
) -> TransferResult:
    """
    Map a source-technique curve onto a target technique.
    log_ratio implements the workbook rule: target = source × exp(L(midi)),
    with L interpolated from measured overlap pairs.
    """
    target = layer_from_identity(
        source.instrument,
        collection or source.collection,
        target_technique,
        source.dynamic,
        source.family,
        source.pitch_basis,
        name=f"{source.instrument}_{collection or source.collection}_{target_technique}_{source.dynamic}",
    )
    target.range_low, target.range_high = source.range_low, source.range_high
    target.labels = dict(source.labels)
    target.labels["transfer"] = f"{source.technique}->{target_technique}:{method}"

    result = TransferResult(layer=target)
    if method == "log_ratio":
        L = log_ratio_field if log_ratio_field is not None else _interp_log_ratio(anchors, list(source.cells))
        for midi, cell in source.cells.items():
            if midi not in L:
                continue
            ratio = float(np.exp(np.clip(L[midi], -3.0, 3.0)))
            value = cell.value * ratio
            if not np.isfinite(value) or value <= 0:
                continue
            origin = origin_override or "technique_transfer"
            comment = f"from {source.technique}"
            if field_comment:
                comment += f"; {field_comment}"
            anchor_midis = [a.midi for a in anchors]
            for a in anchors:
                if a.midi == midi and copy_anchors:
                    origin = "measured"
                    value = a.target_value
                    result.n_copied += 1
                    break
            else:
                result.n_transferred += 1
                if anchor_midis and (midi < min(anchor_midis) or midi > max(anchor_midis)):
                    comment += "; L held at nearest measured pair (no cubic extrapolation)"
            target.cells[midi] = Cell(midi, cell.note_label, float(value), origin, comment)
        result.messages.append(
            f"Technique transfer {source.technique} → {target_technique} via log-ratio "
            f"on {len(anchors)} anchors; {result.n_transferred} mapped, {result.n_copied} kept as measured."
        )
        return result

    if method == "linear_ratio":
        ratios = [a.target_value / a.source_value for a in anchors if a.source_value > 0]
        scale = float(np.median(ratios)) if ratios else 1.0
        for midi, cell in source.cells.items():
            hit = next((a for a in anchors if a.midi == midi), None)
            if hit:
                target.cells[midi] = Cell(midi, cell.note_label, hit.target_value, "measured", "anchor")
                result.n_copied += 1
            else:
                target.cells[midi] = Cell(
                    midi, cell.note_label, cell.value * scale, "technique_transfer", f"global ratio {scale:.4f}"
                )
                result.n_transferred += 1
        result.messages.append(f"Global linear ratio {scale:.4f} from {len(ratios)} anchors.")
        return result

    raise ValueError(f"Unknown technique method: {method}")


def family_transfer(
    source: Layer,
    target_instrument: str,
    mode: str = "concert_pitch",
    anchors: Optional[list[Anchor]] = None,
) -> TransferResult:
    """
    concert_pitch: copy/scale values that share sounding MIDI.
    register_relative: align lowest sounding notes, then scale along the shared index.
    """
    spec = resolve_instrument(target_instrument)
    target = layer_from_identity(
        target_instrument,
        source.collection,
        source.technique,
        source.dynamic,
        spec.family if spec else source.family,
        source.pitch_basis,
        name=f"{target_instrument}_{source.collection}_{source.technique}_{source.dynamic}",
    )
    if spec:
        target.range_low, target.range_high = spec.sounding_low, spec.sounding_high
        if source.technique == "harmonics" and spec.harmonics_sounding_low:
            target.range_low = spec.harmonics_sounding_low
            target.range_high = min(108, spec.sounding_high + 24)

    result = TransferResult(layer=target)
    anchors = anchors or []
    scale = 1.0
    ratios = [a.target_value / a.source_value for a in anchors if a.source_value > 0]
    if ratios:
        scale = float(np.median(ratios))

    if mode == "concert_pitch":
        for midi, cell in source.cells.items():
            if midi < target.range_low or midi > target.range_high:
                continue
            hit = next((a for a in anchors if a.midi == midi), None)
            if hit:
                target.cells[midi] = Cell(midi, midi_to_label(midi), hit.target_value, "measured", "overlap anchor")
                result.n_copied += 1
            else:
                target.cells[midi] = Cell(
                    midi, midi_to_label(midi), cell.value * scale, "family_transfer", f"concert × {scale:.4f}"
                )
                result.n_transferred += 1
        result.messages.append(
            f"Family concert-pitch map {source.instrument} → {target_instrument}, scale={scale:.4f}."
        )
        return result

    if mode == "register_relative":
        src_spec = resolve_instrument(source.instrument)
        src_low = src_spec.sounding_low if src_spec else min(source.cells)
        tgt_low = spec.sounding_low if spec else src_low
        for midi, cell in source.cells.items():
            offset = midi - src_low
            tgt_midi = tgt_low + offset
            if tgt_midi < target.range_low or tgt_midi > target.range_high:
                continue
            target.cells[tgt_midi] = Cell(
                tgt_midi,
                midi_to_label(tgt_midi),
                cell.value * scale,
                "family_transfer",
                f"register index {offset} from {source.instrument}",
            )
            result.n_transferred += 1
        result.messages.append(
            f"Register-relative map {source.instrument} → {target_instrument} "
            f"(align {midi_to_label(src_low)} → {midi_to_label(tgt_low)}), scale={scale:.4f}."
        )
        return result

    raise ValueError(f"Unknown family mode: {mode}")


def fill_missing(
    layer: Layer,
    method: str = "pchip",
    max_extrap_semitones: int = 12,
    poly_degree: int = 3,
    origin_interpolated: str = "interpolated",
    origin_extrapolated: str = "extrapolated_polynomial",
) -> TransferResult:
    """Fill empty chromatic slots. Interpolation inside measured span; capped extrapolation outside."""
    known = layer.sorted_cells()
    result = TransferResult(layer=layer)
    if len(known) < 2:
        result.messages.append("Need at least two known values to fill.")
        return result

    xs = np.array([c.midi for c in known], dtype=float)
    ys = np.array([c.value for c in known], dtype=float)
    lo, hi = int(xs.min()), int(xs.max())
    grid_lo = max(layer.range_low, lo - max_extrap_semitones)
    grid_hi = min(layer.range_high, hi + max_extrap_semitones)

    if method == "pchip":
        if len(known) < 3:
            raise InsufficientFillData(
                f"PCHIP fill needs at least 3 known values; this layer has {len(known)}. "
                "Choose linear fill for two points, or add another measured note. "
                "No other method is substituted automatically."
            )
        fn = PchipInterpolator(xs, ys, extrapolate=True)

        def predict(m: int) -> float:
            return float(fn(m))

        extra_origin = "extrapolated_pchip"
    elif method == "linear":
        def predict(m: int) -> float:
            return float(np.interp(m, xs, ys))

        extra_origin = "extrapolated_hold"
    elif method == "polynomial":
        deg = max(1, min(poly_degree, len(known) - 1))
        coef = np.polyfit(xs, ys, deg)

        def predict(m: int) -> float:
            return float(np.polyval(coef, m))

        extra_origin = origin_extrapolated
    else:
        raise ValueError(f"Unknown fill method: {method}")

    for midi in range(grid_lo, grid_hi + 1):
        if midi in layer.cells:
            continue
        value = predict(midi)
        if value <= 0:
            continue
        inside = lo <= midi <= hi
        origin = origin_interpolated if inside else extra_origin
        layer.cells[midi] = Cell(
            midi,
            midi_to_label(midi),
            value,
            origin,
            f"fill:{method}",
        )
        if inside:
            result.n_interpolated += 1
        else:
            result.n_extrapolated += 1
    result.messages.append(
        f"Fill {method}: +{result.n_interpolated} interpolated, "
        f"+{result.n_extrapolated} extrapolated (cap {max_extrap_semitones} st)."
    )
    return result


def media_of(layers: list[Layer], name: str = "", prefer_measured: bool = False) -> Layer:
    if not layers:
        raise ValueError("No layers to average.")
    first = layers[0]
    media = layer_from_identity(
        first.instrument,
        "+".join(sorted({lg.collection for lg in layers})),
        first.technique,
        first.dynamic,
        first.family,
        first.pitch_basis,
        name=name or f"{first.instrument}_Media_{first.technique}_{first.dynamic}",
    )
    media.range_low = min(lg.range_low for lg in layers)
    media.range_high = max(lg.range_high for lg in layers)
    midis = set()
    for layer in layers:
        midis.update(layer.cells)
    for midi in sorted(midis):
        present = [lg.cells[midi] for lg in layers if midi in lg.cells]
        if not present:
            continue
        measured = [c for c in present if (c.origin or "").strip().lower() == "measured"]
        chosen = measured if (prefer_measured and measured) else present
        value = float(np.mean([c.value for c in chosen]))
        origins = sorted({c.origin for c in chosen})
        if prefer_measured and measured:
            origin = "measured"
            comment = "empirical_only (woodwind): target measurement kept; string Media still averages available peers"
        else:
            origin = "generated" if len(chosen) > 1 else chosen[0].origin
            comment = f"average of {len(chosen)}: {', '.join(origins)}"
        media.cells[midi] = Cell(
            midi,
            midi_to_label(midi),
            value,
            origin,
            comment,
        )
    return media


def parse_anchor_block(block: str) -> list[Anchor]:
    """Accept 'C5 16.8 18.3' or '72, 16.8, 18.3' lines."""
    anchors: list[Anchor] = []
    for line in (block or "").splitlines():
        parts = [p for p in line.replace(";", " ").replace(",", " ").split() if p]
        if len(parts) < 3:
            continue
        pitch = parse_pitch(parts[0])
        try:
            src = float(parts[1].replace(",", "."))
            tgt = float(parts[2].replace(",", "."))
        except ValueError:
            continue
        if pitch:
            anchors.append(Anchor(pitch.midi, src, tgt))
    return anchors
