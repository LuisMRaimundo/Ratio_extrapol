"""In-memory layers: one collection × technique × dynamic curve."""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Iterable, Optional

from .catalog import DEFAULT_MEASURED_CEILING_MIDI, normalize_origin, origin_rank, resolve_instrument
from .notes import Pitch, chromatic_range, midi_to_label, parse_pitch


@dataclass
class Cell:
    midi: int
    note_label: str
    value: float
    origin: str
    comment: str = ""
    flags: list[str] = field(default_factory=list)
    pi95_low: Optional[float] = None
    pi95_high: Optional[float] = None
    reporting_status: str = ""
    anchor_source: str = ""
    source_workbook: str = ""
    uncertainty: str = ""

    def as_row(self) -> dict:
        return {
            "midi": self.midi,
            "note": self.note_label,
            "value": self.value,
            "origin": self.origin,
            "comment": self.comment,
            "flags": "; ".join(self.flags),
            "pi95_low": self.pi95_low,
            "pi95_high": self.pi95_high,
            "reporting_status": self.reporting_status,
            "anchor_source": self.anchor_source,
            "uncertainty": self.uncertainty,
        }


@dataclass
class Layer:
    layer_id: str
    name: str
    instrument: str
    family: str
    collection: str
    technique: str
    dynamic: str
    pitch_basis: str = "sounding_concert"
    transposition: int = 0
    range_low: int = 36
    range_high: int = 96
    cells: dict[int, Cell] = field(default_factory=dict)
    labels: dict[str, str] = field(default_factory=dict)
    measured_ceiling_midi: int = DEFAULT_MEASURED_CEILING_MIDI

    def display_name(self) -> str:
        return self.name or f"{self.instrument}_{self.collection}_{self.technique}_{self.dynamic}"

    def sheet_name(self) -> str:
        raw = f"{self.instrument}_{self.collection}_{self.dynamic}"
        safe = "".join(ch if ch.isalnum() or ch in "_-" else "_" for ch in raw)
        return safe[:31] or "Layer"

    def sorted_cells(self) -> list[Cell]:
        return [self.cells[m] for m in sorted(self.cells)]

    def measured_midis(self) -> list[int]:
        return [m for m, c in self.cells.items() if (c.origin or "").lower() == "measured"]

    def place(
        self,
        pitch: Pitch,
        value: float,
        origin: str,
        comment: str = "",
        overwrite: str = "prefer_measured",
        pi95_low: Optional[float] = None,
        pi95_high: Optional[float] = None,
        reporting_status: str = "",
        anchor_source: str = "",
        source_workbook: str = "",
    ) -> str:
        """
        Place a value on the chromatic slot.
        overwrite: prefer_measured | overwrite | skip
        Returns: placed | overwritten | skipped | merged_kept
        """
        origin = normalize_origin(origin)
        if not reporting_status:
            if origin == "modelled_on_extrapolated_anchor" or pitch.midi > self.measured_ceiling_midi:
                reporting_status = "above_measured_ceiling"
            else:
                reporting_status = "measured_range"
        new_cell = Cell(
            pitch.midi,
            pitch.label,
            float(value),
            origin,
            comment,
            [],
            pi95_low,
            pi95_high,
            reporting_status,
            anchor_source,
            source_workbook,
        )
        existing = self.cells.get(pitch.midi)
        if existing is None:
            self.cells[pitch.midi] = new_cell
            return "placed"
        if overwrite == "skip":
            return "skipped"
        if overwrite == "overwrite":
            new_cell.flags = list(existing.flags)
            self.cells[pitch.midi] = new_cell
            return "overwritten"
        if origin_rank(existing.origin) <= origin_rank(origin):
            if existing.origin != origin and abs(existing.value - float(value)) > 1e-9:
                existing.flags.append("duplicate_conflict_kept")
            return "merged_kept"
        new_cell.flags = list(existing.flags) + ["duplicate_overwritten"]
        new_cell.comment = comment or existing.comment
        self.cells[pitch.midi] = new_cell
        return "overwritten"

    def ensure_grid(self) -> list[Pitch]:
        return chromatic_range(self.range_low, self.range_high)


@dataclass
class Project:
    title: str = "STE Lab session"
    operator: str = ""
    created_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    notes: str = ""
    layers: list[Layer] = field(default_factory=list)
    log: list[str] = field(default_factory=list)
    measured_ceiling_midi: int = DEFAULT_MEASURED_CEILING_MIDI
    extra_evidence: list[dict] = field(default_factory=list)
    validation_pairwise: object = None
    validation_spread: object = None
    validation_summary: list[str] = field(default_factory=list)
    pooled_provenance: str = ""

    def add_layer(self, layer: Layer) -> Layer:
        self.layers.append(layer)
        self.log.append(f"Added layer {layer.display_name()}")
        return layer

    def get_layer(self, layer_id: str) -> Optional[Layer]:
        for layer in self.layers:
            if layer.layer_id == layer_id:
                return layer
        return None

    def remove_layer(self, layer_id: str) -> None:
        self.layers = [layer for layer in self.layers if layer.layer_id != layer_id]


def new_layer_id() -> str:
    return uuid.uuid4().hex[:10]


def layer_from_identity(
    instrument: str,
    collection: str,
    technique: str,
    dynamic: str,
    family: str = "",
    pitch_basis: str = "sounding_concert",
    name: str = "",
    extra_labels: Optional[dict] = None,
) -> Layer:
    spec = resolve_instrument(instrument)
    family = family or (spec.family if spec else "other")
    low = spec.sounding_low if spec else 36
    high = spec.sounding_high if spec else 96
    # Harmonic maps need headroom above the stopped range.
    if technique == "harmonics" and spec and spec.harmonics_sounding_low:
        low = spec.harmonics_sounding_low
        high = min(108, spec.sounding_high + 24)
    return Layer(
        layer_id=new_layer_id(),
        name=name or f"{instrument}_{collection}_{technique}_{dynamic}",
        instrument=instrument,
        family=family,
        collection=collection,
        technique=technique,
        dynamic=dynamic,
        pitch_basis=pitch_basis,
        transposition=spec.transposition if spec else 0,
        range_low=low,
        range_high=high,
        labels=dict(extra_labels or {}),
    )


def cells_to_curve(cells: Iterable[Cell]) -> dict[int, float]:
    return {c.midi: c.value for c in cells}
