"""Flag notes and methods that can scramble or misrepresent CDM results."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .catalog import GENERATED_ORIGINS, resolve_instrument, same_family
from .evidence import SHARED_L_MEDIA_CAVEAT, evidence_role, is_prediction_technique, n_measured
from .notes import midi_to_label
from .session import Layer, Project

SEVERITY_ORDER = {"critical": 0, "high": 1, "medium": 2, "low": 3}


@dataclass
class Flag:
    severity: str
    code: str
    message: str
    recommendation: str
    layer_id: str = ""
    layer_name: str = ""
    midi: Optional[int] = None
    note: str = ""
    approach: str = ""

    def row_key(self) -> str:
        return f"{self.severity}|{self.code}|{self.note}|{self.layer_name}"


def audit_layer(layer: Layer) -> list[Flag]:
    flags: list[Flag] = []
    spec = resolve_instrument(layer.instrument)
    cells = layer.sorted_cells()
    if not cells:
        flags.append(
            Flag(
                "medium",
                "empty_layer",
                f"Layer '{layer.display_name()}' has no numerical values.",
                "Paste or type measured anchors before any transfer or export.",
                layer.layer_id,
                layer.display_name(),
            )
        )
        return flags

    measured = [c for c in cells if (c.origin or "").lower() == "measured"]
    generated = [c for c in cells if (c.origin or "").lower() in GENERATED_ORIGINS]
    n = len(cells)
    measured_share = len(measured) / n

    if measured_share < 0.25 and n >= 8:
        flags.append(
            Flag(
                "high",
                "sparse_measured",
                f"Only {len(measured)}/{n} cells ({measured_share:.0%}) are measured.",
                "Do not treat this curve as an empirical profile. Keep origin tags visible and avoid averaging it with a second modelled collection as if both were observations.",
                layer.layer_id,
                layer.display_name(),
                approach="sparse observation + dense fill",
            )
        )

    if cells:
        midis = [c.midi for c in cells]
        span = max(midis) - min(midis)
        if measured:
            mspan = max(c.midi for c in measured) - min(c.midi for c in measured)
            beyond = max(
                max(midis) - max(c.midi for c in measured),
                min(c.midi for c in measured) - min(midis),
            )
            if beyond > 12:
                flags.append(
                    Flag(
                        "high",
                        "wide_extrapolation",
                        f"Fill extends {beyond} semitones beyond the last measured pitch (span {span}).",
                        "Cap extrapolation at one octave unless new measured anchors are added. Polynomial/ridge tails scramble Media midpoints.",
                        layer.layer_id,
                        layer.display_name(),
                        approach="register extrapolation",
                    )
                )
            if mspan < 12 and span > 24:
                flags.append(
                    Flag(
                        "critical",
                        "thin_anchor_span",
                        f"Measured anchors cover only {mspan} semitones while the published curve spans {span}.",
                        "Refuse family or technique transfer from this layer until the measured core is wider.",
                        layer.layer_id,
                        layer.display_name(),
                        approach="under-anchored curve",
                    )
                )

    poly = [c for c in cells if "polynomial" in (c.origin or "").lower()]
    if poly and len(measured) < 8:
        flags.append(
            Flag(
                "critical",
                "polynomial_underdetermined",
                f"Polynomial fill on {len(poly)} notes with only {len(measured)} measured points.",
                "Degree-3 (or any global polynomial) on a short measured core invents turning points. Use PCHIP interpolation inside the measured span and leave the exterior blank.",
                layer.layer_id,
                layer.display_name(),
                approach="global polynomial",
            )
        )

    if spec:
        for cell in cells:
            if cell.midi < spec.sounding_low - 2 or cell.midi > spec.sounding_high + 12:
                flags.append(
                    Flag(
                        "high",
                        "outside_instrument_range",
                        f"{cell.note_label} (MIDI {cell.midi}) is outside the catalogue sounding range "
                        f"{midi_to_label(spec.sounding_low)}–{midi_to_label(spec.sounding_high)} for {spec.display_name}.",
                        "Check octave labels before export. A one-octave slip (e.g. B5 paired with B4) silently wrecks Media.",
                        layer.layer_id,
                        layer.display_name(),
                        cell.midi,
                        cell.note_label,
                        "range",
                    )
                )
            if (
                layer.technique == "harmonics"
                and spec.harmonics_sounding_low is not None
                and cell.midi < spec.harmonics_sounding_low
            ):
                flags.append(
                    Flag(
                        "critical",
                        "harmonic_below_first",
                        f"{cell.note_label} is below the first practical sounding harmonic "
                        f"({midi_to_label(spec.harmonics_sounding_low)}) of {spec.display_name}.",
                        "Do not publish that pitch as a sounding harmonic. If the value is a written/fingered artificial-harmonic pitch, switch pitch basis to written and keep sounding empty — mixing the two bases is how C3 'harmonics' appeared on the viola AcousticTable.",
                        layer.layer_id,
                        layer.display_name(),
                        cell.midi,
                        cell.note_label,
                        "harmonics as sounding_concert",
                    )
                )

    if layer.technique == "harmonics" and layer.family not in {"bowed_strings", "other"}:
        flags.append(
            Flag(
                "medium",
                "wind_harmonics_semantics",
                "Harmonics on winds are overblowing / register changes, not string nodes.",
                "Do not reuse a string-harmonic transfer function (ordinario × exp(L)) on clarinet-family instruments without a dedicated wind model.",
                layer.layer_id,
                layer.display_name(),
                approach="string-harmonic model on winds",
            )
        )

    by_midi: dict[int, list] = {}
    for cell in cells:
        by_midi.setdefault(cell.midi, []).append(cell)
        if cell.value <= 0:
            flags.append(
                Flag(
                    "critical",
                    "non_positive_cdm",
                    f"{cell.note_label} has CDM {cell.value}.",
                    "Combined density must be positive. Check the paste (percentages, dB, or a dashed placeholder).",
                    layer.layer_id,
                    layer.display_name(),
                    cell.midi,
                    cell.note_label,
                )
            )
        if (cell.origin or "").lower() == "measured" and cell.midi > 103:
            flags.append(
                Flag(
                    "medium",
                    "extreme_register_claimed_measured",
                    f"{cell.note_label} is marked measured in an extreme high register.",
                    "Confirm the source file. High-harmonic tails in the viola workbook were modelled, then later averaged as if they were observations.",
                    layer.layer_id,
                    layer.display_name(),
                    cell.midi,
                    cell.note_label,
                    "origin hygiene",
                )
            )

    values = [c.value for c in cells if 0 < c.value < 1e6]
    if len(values) >= 6:
        mean = sum(values) / len(values)
        var = sum((v - mean) ** 2 for v in values) / len(values)
        sd = var ** 0.5
        if sd > 0:
            for cell in cells:
                if abs(cell.value - mean) > 3.5 * sd:
                    flags.append(
                        Flag(
                            "medium",
                            "outlier",
                            f"{cell.note_label} = {cell.value:.3f} is a 3.5σ outlier on this layer.",
                            "Inspect before it enters a Media midpoint; one wild cell tilts every downstream gap column.",
                            layer.layer_id,
                            layer.display_name(),
                            cell.midi,
                            cell.note_label,
                        )
                    )

    ceiling = getattr(layer, "measured_ceiling_midi", 100)
    above = [c for c in cells if c.reporting_status == "above_measured_ceiling" or "extrapolated_anchor" in (c.origin or "")]
    if above:
        flags.append(
            Flag(
                "critical",
                "above_measured_ceiling",
                f"{len(above)} pitches sit above the measured-anchor ceiling (MIDI {ceiling}) "
                "or rest on an ordinario anchor that was itself filled (modelled_on_extrapolated_anchor).",
                "The violin con sordino workbook excludes these from every aggregate (Summary_Measured_Range). "
                "Do not let them into geometric means, mute ratios, or AcousticTable rows marked accepted. "
                "Importer filter: skip cell_status=register_extrapolated / validation_status=review_required.",
                layer.layer_id,
                layer.display_name(),
                approach="register_extrapolated pooled into Media",
            )
        )
        for cell in above:
            flags.append(
                Flag(
                    "high",
                    "register_extrapolated_note",
                    f"{cell.note_label} is {cell.origin} / {cell.reporting_status}.",
                    "Shade, keep, but never silently average into a published statistic.",
                    layer.layer_id,
                    layer.display_name(),
                    cell.midi,
                    cell.note_label,
                    "modelled_on_extrapolated_anchor",
                )
            )

    iowa_modelled = any("iowa" in (c.origin or "").lower() and "modelled" in (c.origin or "").lower() for c in cells)
    orch_measured = any((c.origin or "").lower() == "measured" for c in cells)
    if iowa_modelled and layer.collection.upper() in {"IOWA"} and layer.technique.replace("_", " ") == "con sordino":
        flags.append(
            Flag(
                "high",
                "iowa_has_no_sordino",
                "IOWA holds no con sordino recordings. Every IOWA value is ordinario × exp(L_hat).",
                "Never label those cells measured. Carry PI95 with them. Equal-weight Media with a live Orchidea measurement is Caveat 1 of the v7.1 sordino workbook.",
                layer.layer_id,
                layer.display_name(),
                approach="IOWA sordino as if observed",
            )
        )
    if iowa_modelled and orch_measured:
        flags.append(
            Flag(
                "medium",
                "circular_meta_average",
                "Averaging an IOWA extrapolation (calibrated on a META set that already contains Orchidea sordino rows) with the Orchidea measurement is partly circular.",
                "Do not treat Media as principal evidence. Use Empirical_ORCH. " + SHARED_L_MEDIA_CAVEAT,
                approach="IOWA+ORCH Media at equal weight",
            )
        )

    role = evidence_role(layer)
    if is_prediction_technique(layer.technique) and layer.collection.upper() in {"IOWA", "ORCH", "ORCHIDEA"}:
        flags.append(
            Flag(
                "high",
                "prediction_only_technique",
                f"{layer.technique} has no Orchidea recordings in this pipeline. The layer is a contextual prediction.",
                "Keep it off Empirical_ORCH. Do not put it in a table of measured technique effects.",
                layer.layer_id,
                layer.display_name(),
                approach="prediction",
            )
        )
    if role == "inherited_dynamic":
        flags.append(
            Flag(
                "high",
                "inherited_dynamic",
                f"{layer.dynamic} reuses L from another dynamic; it is not an independent calibration.",
                "Omit this dynamic from Empirical_ORCH and from confirmatory tests. PCA/clustering of pp/mf/ff is descriptive of the completed grid only.",
                layer.layer_id,
                layer.display_name(),
                approach="donor L copied across dynamics",
            )
        )
    if layer.collection.upper() == "IOWA" and n_measured(layer) == 0 and "ordinario" not in (layer.technique or "").lower():
        flags.append(
            Flag(
                "medium",
                "iowa_shared_L",
                "IOWA effect curve shares L with ORCH by construction.",
                "Not independent replication. Do not average it with ORCH to 'strengthen' the relative effect.",
                layer.layer_id,
                layer.display_name(),
                approach="shared-L transfer",
            )
        )
    return flags


def audit_project(project: Project) -> list[Flag]:
    flags: list[Flag] = []
    for layer in project.layers:
        flags.extend(audit_layer(layer))

    groups: dict[tuple, list[Layer]] = {}
    for layer in project.layers:
        key = (layer.instrument.lower(), layer.technique.lower(), layer.dynamic.lower())
        groups.setdefault(key, []).append(layer)

    for key, layers in groups.items():
        if len(layers) < 2:
            continue
        curves = [{c.midi: c for c in layer.sorted_cells()} for layer in layers]
        midis = set(curves[0]).intersection(*[set(c) for c in curves[1:]])
        for midi in sorted(midis):
            labels = [curves[i][midi].note_label for i in range(len(layers))]
            if len({lb.replace("#", "s") for lb in labels}) > 1:
                # enharmonic spelling only — not a pitch error
                if len({c[midi].midi for c in curves}) == 1:
                    flags.append(
                        Flag(
                            "low",
                            "enharmonic_spelling",
                            f"MIDI {midi} is spelled {', '.join(labels)} across collections.",
                            "Joins on the printed label will break (A#6 vs Bb6). Join on MIDI or a normalised pitch.",
                            note=labels[0],
                            midi=midi,
                            approach="label join",
                        )
                    )

    # Dynamic ladder on generated-only pairs
    by_note: dict[tuple, dict[str, Layer]] = {}
    for layer in project.layers:
        by_note.setdefault(
            (layer.instrument.lower(), layer.collection.lower(), layer.technique.lower()),
            {},
        )[layer.dynamic.lower()] = layer
    for ident, dyns in by_note.items():
        if "pp" in dyns and "ff" in dyns:
            pp, ff = dyns["pp"], dyns["ff"]
            shared = set(pp.cells).intersection(ff.cells)
            inversions = 0
            both_modelled = 0
            for midi in shared:
                a, b = pp.cells[midi], ff.cells[midi]
                if b.value + 0.25 < a.value:
                    inversions += 1
                    if (a.origin.lower() in GENERATED_ORIGINS) and (b.origin.lower() in GENERATED_ORIGINS):
                        both_modelled += 1
            if both_modelled >= 4:
                flags.append(
                    Flag(
                        "medium",
                        "modelled_dynamic_inversion",
                        f"{ident[0]} {ident[2]}: ff < pp on {both_modelled} modelled pitches.",
                        "A measured mf against modelled pp/ff (as in the viola harmonics file) can manufacture a false dynamic peak at mf. Do not interpret that peak as bow-force physics until pp and ff are measured.",
                        approach="asymmetric evidence across dynamics",
                    )
                )

    return sorted(flags, key=lambda f: (SEVERITY_ORDER.get(f.severity, 9), f.code, f.note))


def audit_transfer(
    source: Layer,
    target_instrument: str,
    target_technique: str,
    method: str,
    n_anchors: int,
    max_extrap: int,
    poly_degree: int,
) -> list[Flag]:
    flags: list[Flag] = []
    if not same_family(source.instrument, target_instrument) and source.instrument.lower() != target_instrument.lower():
        flags.append(
            Flag(
                "critical",
                "cross_family_transfer",
                f"{source.instrument} → {target_instrument} crosses instrument families.",
                "Refuse the transfer. Spectral-mass / CDM is not portable from strings to clarinets (or the reverse). Stay inside one family (clarinet → bass clarinet, violin → viola).",
                source.layer_id,
                source.display_name(),
                approach="cross-family mapping",
            )
        )
    if target_technique != source.technique and n_anchors < 3:
        flags.append(
            Flag(
                "critical",
                "technique_without_anchors",
                f"Technique transfer {source.technique} → {target_technique} has {n_anchors} overlap anchors.",
                "Need at least three shared pitches with both source and target measured. The workbook rule IOWA_ordinario × exp(L) is only safe when L is interpolated from real pairs, not guessed.",
                source.layer_id,
                source.display_name(),
                approach="log-ratio technique transfer",
            )
        )
    if method == "polynomial" and poly_degree >= 3:
        flags.append(
            Flag(
                "high",
                "high_degree_polynomial",
                f"Polynomial degree {poly_degree} will invent turning points outside the measured core.",
                "Prefer PCHIP inside the span. If a polynomial is required for a methods appendix, keep it on a scratch layer and never copy it into AcousticTable as 'accepted'.",
                approach="polynomial regression",
            )
        )
    if max_extrap > 12:
        flags.append(
            Flag(
                "high",
                "extrap_cap_too_wide",
                f"Extrapolation cap is {max_extrap} semitones.",
                "Keep the cap at 12 unless new measurements exist. The viola C#7–B7 tail was the main source of Media/AcousticTable drift.",
                approach="register fill",
            )
        )
    if method in {"family_concert", "family_register"} and n_anchors < 2:
        flags.append(
            Flag(
                "high",
                "family_without_overlap",
                "Family-register transfer has fewer than two overlap anchors.",
                "Align clarinet and bass clarinet on at least two shared concert pitches (or two register-relative indexes) before scaling the rest.",
                approach="family transfer",
            )
        )
    return flags
