"""Merge genuine compiled observations. Longest-file-wins is prohibited."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from run_ste_effects_batch import load_spectral_mass
from ste_lab.notes import midi_to_label

REL_TOL = 1e-6
ABS_TOL = 1e-9


@dataclass
class MeasuredObservation:
    midi: int
    value: float
    collection: str
    technique: str
    dynamic: str
    source_path: Path
    label: str = ""


@dataclass
class MergeReport:
    n_source_observations: int = 0
    n_unique_empirical_midis: int = 0
    n_exported_empirical_midis: int = 0
    n_conflicts: int = 0
    n_dropped: int = 0
    dropped: list[dict] = field(default_factory=list)
    sources: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return {
            "n_source_observations": self.n_source_observations,
            "n_unique_empirical_midis": self.n_unique_empirical_midis,
            "n_exported_empirical_midis": self.n_exported_empirical_midis,
            "n_conflicts": self.n_conflicts,
            "n_dropped": self.n_dropped,
            "dropped": list(self.dropped),
            "sources": list(self.sources),
        }


def values_equivalent(a: float, b: float, *, rel: float = REL_TOL, abs_tol: float = ABS_TOL) -> bool:
    return abs(a - b) <= max(abs_tol, rel * max(abs(a), abs(b), 1.0))


def assert_not_source(output: Path, sources: set[Path]) -> None:
    """Refuse to write a generated book over any file used as input."""
    out = Path(output).resolve()
    for raw in sources:
        src = Path(raw).resolve()
        if out == src:
            raise RuntimeError(f"Refusing to overwrite source workbook: {out}")


def _payload_value(payload) -> float:
    return float(payload[1] if isinstance(payload, tuple) else payload)


def _payload_label(midi: int, payload=None) -> str:
    if isinstance(payload, tuple) and payload[0]:
        return str(payload[0])
    return midi_to_label(midi)


def merge_measured_specs(
    specs: list[dict],
    *,
    title: str = "FILES USED  (merged compiled books per collection × technique × dynamic)",
    extra_observations: Optional[list[MeasuredObservation]] = None,
) -> tuple[dict[tuple[str, str, str], dict], dict[tuple[str, str, str], MergeReport]]:
    """Union every genuine compiled workbook. Conflicts are flagged, not averaged."""
    buckets: dict[tuple[str, str, str], dict[int, list[MeasuredObservation]]] = {}
    reports: dict[tuple[str, str, str], MergeReport] = {}
    for spec in specs:
        mass = load_spectral_mass(spec["path"])
        if not mass:
            continue
        key = (spec["collection"], spec["technique"], spec["dynamic"])
        bucket = buckets.setdefault(key, {})
        for midi, payload in mass.items():
            val = _payload_value(payload)
            if val <= 0:
                continue
            obs = MeasuredObservation(
                midi=int(midi),
                value=val,
                collection=spec["collection"],
                technique=spec["technique"],
                dynamic=spec["dynamic"],
                source_path=Path(spec["path"]),
                label=_payload_label(int(midi), payload),
            )
            bucket.setdefault(int(midi), []).append(obs)
    for obs in extra_observations or []:
        if obs.value <= 0:
            continue
        key = (obs.collection, obs.technique, obs.dynamic)
        buckets.setdefault(key, {}).setdefault(int(obs.midi), []).append(obs)

    measured: dict[tuple[str, str, str], dict] = {}
    for key, by_midi in sorted(buckets.items()):
        report = MergeReport()
        curve: dict[int, tuple[str, float]] = {}
        paths: list[Path] = []
        conflicts: list[dict] = []
        for midi, group in by_midi.items():
            report.n_source_observations += len(group)
            report.n_unique_empirical_midis += 1
            for item in group:
                if item.source_path not in paths:
                    paths.append(item.source_path)
            values = [item.value for item in group]
            if all(values_equivalent(values[0], v) for v in values[1:]):
                curve[midi] = (group[0].label or midi_to_label(midi), values[0])
                report.n_exported_empirical_midis += 1
                continue
            report.n_conflicts += 1
            report.n_dropped += 1
            reason = {
                "midi": midi,
                "reason": "EMPIRICAL_CONFLICT",
                "values": values,
                "sources": [str(item.source_path) for item in group],
            }
            report.dropped.append(reason)
            conflicts.append(reason)
            print(
                f"  EMPIRICAL_CONFLICT  {key[0]} {key[1]} {key[2]}  MIDI {midi}  "
                f"values={values}  — not silently resolved"
            )
        report.sources = [str(p) for p in paths]
        measured[key] = {
            "curve": curve,
            "path": paths[0] if paths else None,
            "paths": paths,
            "conflicts": conflicts,
            "report": report.as_dict(),
        }
        reports[key] = report

    print(f"\n{title}")
    if not measured:
        print("  (none)")
    for (coll, tech, dyn), blob in sorted(measured.items()):
        n = len(blob["curve"])
        srcs = "; ".join(p.name for p in blob["paths"])
        print(f"  {coll:6} {tech:16} {dyn:3}  n={n}  sources={srcs}")
        rep = blob["report"]
        print(
            f"         n_source_observations={rep['n_source_observations']}  "
            f"n_unique_empirical_midis={rep['n_unique_empirical_midis']}  "
            f"n_exported_empirical_midis={rep['n_exported_empirical_midis']}  "
            f"n_conflicts={rep['n_conflicts']}  n_dropped={rep['n_dropped']}"
        )
    return measured, reports


def iowa_observations_from_media(
    curves: dict[str, dict[str, dict[int, float]]],
    source_path: Path,
    technique: str = "ordinario",
) -> list[MeasuredObservation]:
    out: list[MeasuredObservation] = []
    iowa = curves.get("IOWA") or {}
    for dyn, curve in iowa.items():
        for midi, value in curve.items():
            if value and value > 0:
                out.append(
                    MeasuredObservation(
                        midi=int(midi),
                        value=float(value),
                        collection="IOWA",
                        technique=technique,
                        dynamic=dyn,
                        source_path=Path(source_path),
                        label=midi_to_label(int(midi)),
                    )
                )
    return out
