"""Evidence classes: measured recordings vs completed / inherited / predicted cells.

Principal empirical evidence is ORCH cells tagged measured.
IOWA effect curves share L with ORCH by construction — they are not a
second experiment. Media of that pair is a completed-grid convenience,
not the principal effect statistic.
"""

from __future__ import annotations

from typing import Iterable, Optional

from .catalog import orchestral_group
from .session import Cell, Layer, Project

PREDICTION_TECHNIQUES = frozenset({"sul tasto"})
EFFECT_TECHNIQUES = frozenset(
    {"con sordino", "sul ponticello", "sul tasto", "harmonics", "mute"}
)
CONTEXT_COLLECTIONS = frozenset({"PHIL", "MCGILL", "PHILHARMONIA"})
ORCH_COLLECTIONS = frozenset({"ORCH", "ORCHIDEA", "ORCH."})


def is_context_collection(collection: str) -> bool:
    return (collection or "").upper() in CONTEXT_COLLECTIONS


def is_orch_collection(collection: str) -> bool:
    return (collection or "").upper() in ORCH_COLLECTIONS


def is_measured(cell: Cell) -> bool:
    return (cell.origin or "").strip().lower() == "measured"


def is_effect(technique: str) -> bool:
    t = (technique or "").strip().lower().replace("_", " ")
    return t in EFFECT_TECHNIQUES or t not in {"ordinario", "arco", ""}


def is_prediction_technique(technique: str) -> bool:
    return (technique or "").strip().lower().replace("_", " ") in PREDICTION_TECHNIQUES


def n_measured(layer: Layer) -> int:
    return sum(1 for c in layer.cells.values() if is_measured(c))


def measured_cells(layer: Layer) -> list[Cell]:
    return [c for c in layer.sorted_cells() if is_measured(c)]


def donor_dynamic(layer: Layer) -> str:
    return (layer.labels.get("l_donor_dynamic") or "").strip()


def dynamic_was_inherited(layer: Layer) -> bool:
    donor = donor_dynamic(layer)
    return bool(donor and donor != layer.dynamic)


def evidence_role(layer: Layer) -> str:
    """prediction | inherited_dynamic | empirical | modelled_completion | context."""
    tagged = (layer.labels.get("evidence_role") or "").strip()
    if tagged:
        return tagged
    tech = (layer.technique or "").strip().lower().replace("_", " ")
    coll = (layer.collection or "").upper()
    if is_context_collection(coll):
        return "context"
    if is_prediction_technique(tech):
        return "prediction"
    if dynamic_was_inherited(layer):
        return "inherited_dynamic"
    if coll in {"IOWA"} and is_effect(tech):
        return "modelled_completion"
    if n_measured(layer):
        return "empirical"
    if (
        coll.startswith("ORCH")
        and is_effect(tech)
        and layer.dynamic != "mf"
        and tech in {"sul ponticello", "harmonics"}
    ):
        return "inherited_dynamic"
    return "modelled_completion"


def stamp_evidence(
    layer: Layer,
    *,
    inherited_from: Optional[str] = None,
    prediction: bool = False,
    l_invariance: Optional[str] = None,
    l_invariance_spread: Optional[float] = None,
    invariance_report=None,
) -> Layer:
    """Write evidence labels onto a layer after it is built."""
    if inherited_from:
        layer.labels["l_donor_dynamic"] = inherited_from
    if invariance_report is not None:
        layer.labels.update(
            {key: value for key, value in invariance_report.as_labels().items() if value}
        )
    else:
        if l_invariance:
            layer.labels["l_invariance"] = l_invariance
        if l_invariance_spread is not None:
            layer.labels["l_invariance_spread"] = f"{float(l_invariance_spread):.4f}"
    if is_context_collection(layer.collection):
        layer.labels["evidence_role"] = "context"
    elif prediction or is_prediction_technique(layer.technique):
        layer.labels["evidence_role"] = "prediction"
    elif inherited_from and inherited_from != layer.dynamic:
        layer.labels["evidence_role"] = "inherited_dynamic"
    elif (layer.collection or "").upper() == "IOWA" and is_effect(layer.technique):
        layer.labels["evidence_role"] = "modelled_completion"
        layer.labels["iowa_no_recordings"] = "true"
    elif n_measured(layer) and is_orch_collection(layer.collection):
        layer.labels["evidence_role"] = "empirical"
    elif (
        n_measured(layer)
        and (layer.collection or "").upper() == "IOWA"
        and orchestral_group(layer.instrument) == "woodwinds"
        and not is_effect(layer.technique)
    ):
        layer.labels["evidence_role"] = "empirical"
    else:
        layer.labels["evidence_role"] = "modelled_completion"
    layer.labels["n_measured"] = str(n_measured(layer))
    return layer


def _sort_empirical(rows: list[dict]) -> list[dict]:
    order = ["pp", "mf", "ff", "p", "mp", "f"]
    rows.sort(key=lambda r: (order.index(r["dynamic"]) if r["dynamic"] in order else 9, r["midi"]))
    return rows


def orch_empirical_rows(layers: Iterable[Layer], technique: str) -> list[dict]:
    """ORCH measured cells only — principal evidence for an effect."""
    rows = []
    for layer in layers:
        if not is_orch_collection(layer.collection):
            continue
        if (layer.technique or "").strip().lower() != (technique or "").strip().lower():
            continue
        if is_prediction_technique(layer.technique):
            continue
        if dynamic_was_inherited(layer) and not n_measured(layer):
            continue
        for cell in measured_cells(layer):
            rows.append(
                dict(
                    instrument=layer.instrument,
                    collection="ORCH",
                    technique=layer.technique,
                    dynamic=layer.dynamic,
                    note=cell.note_label,
                    midi=cell.midi,
                    value=cell.value,
                    origin=cell.origin,
                    reporting_status=cell.reporting_status,
                    source_workbook=cell.source_workbook,
                    comment=cell.comment,
                    evidence_role="empirical",
                )
            )
    return _sort_empirical(rows)


def iowa_measured_rows(layers: Iterable[Layer], technique: str) -> list[dict]:
    """Iowa cells tagged measured — principal evidence for woodwind ordinario."""
    rows = []
    for layer in layers:
        if (layer.collection or "").upper() != "IOWA":
            continue
        if (layer.technique or "").strip().lower() != (technique or "").strip().lower():
            continue
        if is_effect(layer.technique):
            continue
        for cell in measured_cells(layer):
            rows.append(
                dict(
                    instrument=layer.instrument,
                    collection="IOWA",
                    technique=layer.technique,
                    dynamic=layer.dynamic,
                    note=cell.note_label,
                    midi=cell.midi,
                    value=cell.value,
                    origin=cell.origin,
                    reporting_status=cell.reporting_status,
                    source_workbook=cell.source_workbook,
                    comment=cell.comment,
                    evidence_role="empirical",
                )
            )
    return _sort_empirical(rows)


def project_instrument(project: Project) -> str:
    if project.layers:
        return project.layers[0].instrument
    return ""


def project_is_woodwind(project: Project) -> bool:
    return orchestral_group(project_instrument(project)) == "woodwinds"


def principal_empirical_rows(project: Project, technique: str) -> list[dict]:
    """Orchidea measured cells when they exist; else Iowa measured on woodwind ordinario."""
    orch = orch_empirical_rows(project.layers, technique)
    if orch:
        return orch
    if project_is_woodwind(project) and not is_effect(technique):
        return iowa_measured_rows(project.layers, technique)
    return []


def principal_rows_are_iowa(rows: list[dict]) -> bool:
    return bool(rows) and (rows[0].get("collection") or "").upper() == "IOWA"


ORCH_PRINCIPAL_TITLE = "Empirical_ORCH — principal evidence (measured Orchidea cells only)"
WOODWIND_PRINCIPAL_TITLE = (
    "Empirical_ORCH — principal evidence (measured Iowa cells; no Orchidea on this woodwind)"
)
WOODWIND_PRINCIPAL_SUBTITLE = (
    "This sheet is the inferential dataset for woodwind ordinario. "
    "Family-transferred cousin Orchidea stays on the ORCH layers and is not listed here. "
    "Philharmonia and McGill remain context."
)


def principal_sheet_title(rows: list[dict]) -> str:
    if principal_rows_are_iowa(rows):
        return WOODWIND_PRINCIPAL_TITLE
    return ORCH_PRINCIPAL_TITLE


def principal_sheet_subtitle(rows: list[dict]) -> str:
    if principal_rows_are_iowa(rows):
        return WOODWIND_PRINCIPAL_SUBTITLE
    return (
        "This sheet is the inferential dataset for the technique. "
        "Completed-grid Media, IOWA (shared L), sul tasto, and inherited pp/ff are excluded. "
        + SHARED_L_MEDIA_CAVEAT
    )


def evidence_map_rows(project: Project, technique: str) -> list[dict]:
    rows = []
    for layer in project.layers:
        if (layer.technique or "").strip().lower() != (technique or "").strip().lower():
            continue
        if (layer.collection or "").upper() not in {"IOWA", "ORCH", "ORCHIDEA", "ORCH.", "PHIL", "MCGILL"}:
            continue
        n_all = len(layer.cells)
        n_m = n_measured(layer)
        role = evidence_role(layer)
        principal = "no"
        if role == "empirical" and n_m:
            if is_orch_collection(layer.collection):
                principal = "yes"
            elif (
                (layer.collection or "").upper() == "IOWA"
                and orchestral_group(layer.instrument) == "woodwinds"
                and not is_effect(layer.technique)
            ):
                principal = "yes"
        rows.append(
            dict(
                collection=layer.collection,
                dynamic=layer.dynamic,
                n_cells=n_all,
                n_measured=n_m,
                n_modelled=n_all - n_m,
                evidence_role=role,
                l_donor_dynamic=donor_dynamic(layer) or "",
                l_invariance=(layer.labels.get("l_invariance") or ""),
                l_invariance_heuristic=(layer.labels.get("l_invariance_heuristic") or ""),
                l_invariance_n_shared=(layer.labels.get("l_invariance_n_shared") or ""),
                l_invariance_mean_abs=(layer.labels.get("l_invariance_mean_abs") or ""),
                l_invariance_max_abs=(layer.labels.get("l_invariance_max_abs") or ""),
                l_mean_L_spread=(layer.labels.get("l_mean_L_spread") or ""),
                l_invariance_pair=(layer.labels.get("l_invariance_pair") or ""),
                l_invariance_collection=(layer.labels.get("l_invariance_collection") or ""),
                l_invariance_spread=(layer.labels.get("l_invariance_spread") or ""),
                principal_evidence=principal,
                note=_role_note(layer, role, n_m),
            )
        )
    return rows


def _role_note(layer: Layer, role: str, n_m: int) -> str:
    if role == "validation":
        return (
            f"{n_m} {layer.collection} anchors used to check the transfer assumption. "
            "Not Media and not a measured production cell."
        )
    if role == "context":
        return (
            f"{n_m} measured {layer.collection} cells. "
            "Context only — not Empirical_ORCH and not Media."
        )
    inv = (layer.labels.get("l_invariance") or "").strip()
    inv_bit = f" L invariance {inv}." if inv else ""
    if role == "prediction":
        donor = donor_dynamic(layer)
        if donor and donor != (layer.dynamic or ""):
            return (
                f"No Orchidea recordings. L from {donor} transposed via ordinario "
                f"dynamic ratios onto {layer.dynamic}.{inv_bit} Not principal evidence."
            )
        return (
            "No Orchidea recordings. L is contextual (e.g. Philharmonia)."
            f"{inv_bit} Not principal evidence."
        )
    if role == "inherited_dynamic":
        return (
            f"Relative effect inherited from {donor_dynamic(layer) or 'a donor dynamic'}."
            f"{inv_bit} Not an independent calibration."
        )
    if (layer.collection or "").upper() == "IOWA" and is_effect(layer.technique):
        return "IOWA has no effect recordings. Same-dynamic L when those anchors exist; otherwise donor L. Not replication."
    if role == "empirical":
        if (layer.collection or "").upper() == "IOWA":
            return (
                f"{n_m} measured Iowa cells. "
                "Principal evidence for this woodwind ordinario (no Orchidea on the target instrument)."
            )
        return f"{n_m} measured Orchidea cells. Principal evidence for this technique × dynamic."
    return "Completed / modelled cells. Use for exploration, not confirmatory tests."


SHARED_L_MEDIA_CAVEAT = (
    "Media of IOWA and ORCH on an effect is not two independent estimates of L. "
    "The batch teaches one L (usually from Orchidea) and applies it to both arco spines. "
    "Use Empirical_ORCH (measured cells only) as the principal effect statistic. "
    "Violin_Media is a completed-grid convenience."
)
