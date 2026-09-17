"""Single resolver for IOWA/ORCH Media and woodwind combination."""

from __future__ import annotations

from typing import Optional

from .calibration import (
    CalibrationConfig,
    ConfigurationError,
    SUPPORTED_COMBINATION_METHODS,
    combine_empirical_and_modelled,
    load_config,
)
from .catalog import orchestral_group
from .evidence import is_effect
from .notes import midi_to_label
from .session import Cell, Layer, layer_from_identity

PRODUCTION_COLLECTIONS = frozenset({"IOWA", "ORCH", "ORCHIDEA", "ORCH."})

__all__ = [
    "ConfigurationError",
    "PRODUCTION_COLLECTIONS",
    "SUPPORTED_COMBINATION_METHODS",
    "is_production_collection",
    "pair_policy",
    "matching_identity_peers",
    "production_average_for_layer",
    "production_media_of",
    "production_peers",
    "resolve_iowa_orch_pair",
    "uses_woodwind_combination",
]


def normalize_collection(collection: str) -> str:
    coll = (collection or "").upper().strip()
    if coll in {"ORCHIDEA", "ORCH."}:
        return "ORCH"
    return coll


def is_production_collection(collection: str) -> bool:
    return normalize_collection(collection) in {"IOWA", "ORCH"}


def production_peers(layers: list[Layer]) -> list[Layer]:
    return [lg for lg in layers if is_production_collection(lg.collection)]


def uses_woodwind_combination(instrument: str, technique: str) -> bool:
    return orchestral_group(instrument) == "woodwinds" and not is_effect(technique)


def pair_policy(instrument: str, technique: str, cfg: Optional[CalibrationConfig] = None) -> str:
    if uses_woodwind_combination(instrument, technique):
        cfg = cfg or load_config()
        return (cfg.combination_method or "empirical_only").strip().lower()
    return "string_mean"


def resolve_iowa_orch_pair(
    iowa: Optional[float],
    orch: Optional[float],
    *,
    instrument: str,
    technique: str = "ordinario",
    cfg: Optional[CalibrationConfig] = None,
) -> tuple[Optional[float], str, float, float]:
    """Return (value, method, w_E, w_T) for one MIDI.

    string_mean: arithmetic mean of the present production collections; one
    present value is kept; neither present is missing.
    Woodwind ordinario follows the calibrated combination method.
    """
    cfg = cfg or load_config()
    policy = pair_policy(instrument, technique, cfg)
    if policy == "string_mean":
        vals = [v for v in (iowa, orch) if v is not None]
        if not vals:
            return None, "string_mean", 0.0, 0.0
        if len(vals) == 1:
            w_e = 1.0 if iowa is not None else 0.0
            return float(vals[0]), "string_mean", w_e, 1.0 - w_e
        return float(sum(vals) / len(vals)), "string_mean", 0.5, 0.5
    return combine_empirical_and_modelled(iowa, orch, cfg)


def pair_origin(
    iowa_cell: Optional[Cell],
    orch_cell: Optional[Cell],
    method: str,
    w_e: float,
    w_t: float,
) -> str:
    blended = (
        method in {"equal_weight", "legacy_equal_weight", "string_mean"}
        and iowa_cell is not None
        and orch_cell is not None
        and 0.0 < w_e < 1.0
        and 0.0 < w_t < 1.0
    )
    if blended and method != "string_mean":
        return "combined_estimate"
    if blended and method == "string_mean":
        return "generated"
    if method in {"empirical_only", "string_mean"} and iowa_cell is not None and w_e == 1.0:
        return iowa_cell.origin
    if orch_cell is not None and w_t == 1.0:
        return orch_cell.origin
    if iowa_cell is not None:
        return iowa_cell.origin
    if orch_cell is not None:
        return orch_cell.origin
    return "generated"


def matching_identity_peers(layers: list[Layer], layer: Layer) -> list[Layer]:
    return [
        lg
        for lg in layers
        if lg.technique == layer.technique
        and lg.dynamic == layer.dynamic
        and lg.instrument == layer.instrument
    ]


def production_average_for_layer(
    layers: list[Layer],
    layer: Layer,
    cfg: Optional[CalibrationConfig] = None,
) -> Layer:
    """GUI/batch Average: same identity set, production IOWA/ORCH policy."""
    peers = matching_identity_peers(layers, layer)
    if len(peers) < 2:
        raise ValueError("Need at least two layers with the same instrument, technique and dynamic.")
    return production_media_of(
        peers,
        instrument=layer.instrument,
        technique=layer.technique,
        cfg=cfg,
    )


def production_media_of(
    layers: list[Layer],
    *,
    name: str = "",
    instrument: str = "",
    technique: str = "",
    cfg: Optional[CalibrationConfig] = None,
) -> Layer:
    """IOWA/ORCH Media under the family-specific production policy."""
    cfg = cfg or load_config()
    peers = production_peers(layers)
    if not peers:
        raise ValueError("No IOWA/ORCH production layers to average.")
    first = peers[0]
    instrument = instrument or first.instrument
    technique = technique or first.technique
    media = layer_from_identity(
        instrument,
        "+".join(sorted({normalize_collection(lg.collection) for lg in peers})),
        technique,
        first.dynamic,
        first.family,
        first.pitch_basis,
        name=name or f"{instrument}_Media_{technique}_{first.dynamic}",
    )
    media.range_low = min(lg.range_low for lg in peers)
    media.range_high = max(lg.range_high for lg in peers)
    iowa_layers = [lg for lg in peers if normalize_collection(lg.collection) == "IOWA"]
    orch_layers = [lg for lg in peers if normalize_collection(lg.collection) == "ORCH"]
    midis = set()
    for layer in peers:
        midis.update(layer.cells)
    for midi in sorted(midis):
        iowa_cell = next((lg.cells[midi] for lg in iowa_layers if midi in lg.cells), None)
        orch_cell = next((lg.cells[midi] for lg in orch_layers if midi in lg.cells), None)
        iowa = iowa_cell.value if iowa_cell is not None else None
        orch = orch_cell.value if orch_cell is not None else None
        value, method, w_e, w_t = resolve_iowa_orch_pair(
            iowa, orch, instrument=instrument, technique=technique, cfg=cfg
        )
        if value is None:
            continue
        origin = pair_origin(iowa_cell, orch_cell, method, w_e, w_t)
        comment = f"{method} w_E={w_e:.4f} w_T={w_t:.4f}"
        media.cells[midi] = Cell(
            midi,
            midi_to_label(midi),
            float(value),
            origin,
            comment,
        )
    return media
