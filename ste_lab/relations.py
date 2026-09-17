"""Transfer relations: one abstraction for technique, instrument, and collection L.

A relation is a pair (source condition, target condition) measured inside one
collection at shared sounding MIDI, giving L(m) = ln(target/source). Future
kinds (register, dynamic) use the same record.
"""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Iterable, Optional

import pandas as pd

from .catalog import family_donor_instrument, orchestral_group
from .notes import midi_to_label
from .transfer import Anchor, _interp_log_ratio, _safe_log_ratio


ORDINARIO_ALIASES = frozenset({"ordinario", "arco"})
PRODUCTION_DYNAMICS = frozenset({"pp", "mf", "ff"})
DYNAMIC_ORDER = ("pp", "p", "mp", "mf", "f", "ff", "fff")
L_INVARIANCE_HOLD = 0.40

INVENTORY_COLUMNS = [
    "kind",
    "instrument",
    "dynamic",
    "collection",
    "source_cond",
    "target_cond",
    "n_anchors",
    "grade",
    "used_in_production",
]


@dataclass
class Relation:
    kind: str
    instrument: str
    source_cond: str
    target_cond: str
    collection: str
    dynamic: str
    anchors: list[Anchor] = field(default_factory=list)
    grade: str = ""
    used_in_production: bool = False

    @property
    def n_anchors(self) -> int:
        return len(self.anchors)

    @property
    def key(self) -> tuple[str, str, str, str, str]:
        return (self.kind, self.instrument, self.source_cond, self.target_cond, self.dynamic)

    def L_at_anchors(self) -> dict[int, float]:
        out: dict[int, float] = {}
        for a in self.anchors:
            L = _safe_log_ratio(a.target_value, a.source_value)
            if L is not None:
                out[int(a.midi)] = L
        return out


def _payload_value(payload: Any) -> Optional[float]:
    raw = payload[1] if isinstance(payload, tuple) else payload
    try:
        val = float(raw)
    except (TypeError, ValueError):
        return None
    if val <= 0 or not math.isfinite(val):
        return None
    return val


def _payload_label(midi: int, payload: Any = None) -> str:
    if isinstance(payload, tuple) and payload and payload[0]:
        return str(payload[0])
    return midi_to_label(int(midi))


def _as_curve(blob: Any) -> dict[int, Any]:
    if not blob:
        return {}
    if isinstance(blob, dict) and "curve" in blob:
        raw = blob.get("curve") or {}
    else:
        raw = blob
    if not isinstance(raw, dict):
        return {}
    out: dict[int, Any] = {}
    for midi, payload in raw.items():
        try:
            out[int(midi)] = payload
        except (TypeError, ValueError):
            continue
    return out


def _curve_values(blob: Any) -> dict[int, float]:
    out: dict[int, float] = {}
    for midi, payload in _as_curve(blob).items():
        val = _payload_value(payload)
        if val is not None:
            out[int(midi)] = val
    return out


def _norm_coll(name: str) -> str:
    raw = (name or "").strip()
    up = raw.upper()
    if up in {"ORCHIDEA", "ORCH.", "ORCH"}:
        return "ORCH"
    if up in {"PHILHARMONIA", "PHIL"}:
        return "PHIL"
    if up in {"MCGILL", "MCGILL_U"}:
        return "MCGILL"
    return raw or up


def _is_orch(name: str) -> bool:
    return _norm_coll(name) == "ORCH"


def _teacher_coll_match(left: Relation, right: Relation) -> bool:
    if _is_orch(left.collection) and _is_orch(right.collection):
        return True
    return _norm_coll(left.collection) == _norm_coll(right.collection)


def dynamic_rank(dynamic: str) -> Optional[int]:
    key = (dynamic or "").strip().lower()
    try:
        return DYNAMIC_ORDER.index(key)
    except ValueError:
        return None


def closest_production_relation(relations: Iterable[Relation], dynamic: str) -> Optional[Relation]:
    """Exact dynamic if present; else nearest on DYNAMIC_ORDER (more anchors, then ORCH/PHIL/McGill)."""
    cands = [rel for rel in relations if rel.anchors]
    if not cands:
        return None
    want = (dynamic or "").strip().lower()
    exact = [rel for rel in cands if (rel.dynamic or "").strip().lower() == want]
    if exact:
        return max(exact, key=lambda rel: rel.n_anchors)
    target = dynamic_rank(dynamic)
    coll_rank = {"ORCH": 0, "PHIL": 1, "MCGILL": 2}

    def _key(rel: Relation) -> tuple:
        rank = dynamic_rank(rel.dynamic)
        dist = 99 if target is None or rank is None else abs(rank - target)
        return (dist, -rel.n_anchors, coll_rank.get(_norm_coll(rel.collection), 3))

    return min(cands, key=_key)


def mean_anchor_L(rel: Relation) -> Optional[float]:
    vals = list(rel.L_at_anchors().values())
    if not vals:
        return None
    return sum(vals) / len(vals)


def l_invariance_status(relations: Iterable[Relation]) -> tuple[str, Optional[float]]:
    """held | wide | untested, plus max−min mean L across teacher dynamics (nats)."""
    means: list[float] = []
    dyns: set[str] = set()
    for rel in relations:
        mid = mean_anchor_L(rel)
        if mid is None:
            continue
        dyns.add((rel.dynamic or "").strip().lower())
        means.append(mid)
    if len(dyns) < 2 or not means:
        return "untested", None
    spread = max(means) - min(means)
    status = "held" if spread <= L_INVARIANCE_HOLD else "wide"
    return status, spread


def _is_tasto(technique: str) -> bool:
    return "tasto" in (technique or "").strip().lower()


def _is_ordinario(technique: str) -> bool:
    return (technique or "").strip().lower() in ORDINARIO_ALIASES


def _collections_in(measured: dict) -> set[str]:
    found: set[str] = set()
    for key in measured:
        if isinstance(key, tuple) and key:
            found.add(str(key[0]))
    return found


def _lookup(measured: dict, collection: str, technique: str, dynamic: str) -> dict[int, Any]:
    """Return the raw curve blob for a cell, trying collection / technique aliases."""
    colls = {collection, _norm_coll(collection)}
    if _norm_coll(collection) == "ORCH":
        colls.update({"ORCH", "ORCHIDEA", "Orchidea"})
    techs = {technique}
    if _is_ordinario(technique):
        techs.update(ORDINARIO_ALIASES)
    for coll in colls:
        for tech in techs:
            blob = measured.get((coll, tech, dynamic))
            if blob:
                curve = _as_curve(blob)
                if curve:
                    return curve
    return {}


def _arco_curve(catalog: dict, collection: str, dynamic: str) -> dict[int, float]:
    arco = catalog.get("arco") or {}
    for coll_key in (collection, _norm_coll(collection)):
        block = arco.get(coll_key)
        if not isinstance(block, dict):
            continue
        curve = block.get(dynamic) or {}
        if curve:
            return {int(m): float(v) for m, v in curve.items() if _payload_value(v) is not None}
    return {}


def _source_curve(
    measured: dict,
    catalog: dict,
    collection: str,
    technique: str,
    dynamic: str,
) -> dict[int, float]:
    if _is_ordinario(technique):
        from_arco = _arco_curve(catalog, collection, dynamic)
        if from_arco:
            return from_arco
    return _curve_values(_lookup(measured, collection, technique, dynamic))


def anchors_from_curves(
    source: dict[int, float],
    target: dict,
    *,
    min_midi: int | None = None,
) -> list[Anchor]:
    """Shared-MIDI pairs with T>0 and S>0. Harmonics floor applied when given."""
    out: list[Anchor] = []
    tgt_curve = target if target and not isinstance(next(iter(target.values()), None), tuple) else target
    if target and any(isinstance(v, tuple) for v in target.values()):
        tgt_curve = _curve_values(target)
    for midi in sorted(set(source) & set(tgt_curve)):
        if min_midi is not None and int(midi) < int(min_midi):
            continue
        src = source.get(midi)
        tgt = tgt_curve.get(midi)
        tgt_val = _payload_value(tgt) if not isinstance(tgt, (int, float)) else float(tgt)
        if src is None or tgt_val is None:
            continue
        src_val = float(src)
        if src_val <= 0 or tgt_val <= 0 or not math.isfinite(src_val) or not math.isfinite(tgt_val):
            continue
        if _safe_log_ratio(tgt_val, src_val) is None:
            continue
        out.append(Anchor(int(midi), src_val, tgt_val))
    return out


def relation_anchor_frame(rel: Relation, target_payloads: dict[int, Any] | None = None) -> pd.DataFrame:
    """Anchors_all row shape used by the preparatory workbook."""
    rows = []
    payloads = target_payloads or {}
    for a in rel.anchors:
        L = _safe_log_ratio(a.target_value, a.source_value)
        if L is None:
            continue
        note = _payload_label(a.midi, payloads.get(a.midi))
        rows.append(
            {
                "effect": rel.target_cond if rel.kind == "technique" else rel.target_cond,
                "evidence_grade": rel.grade,
                "dynamic": rel.dynamic,
                "note": note,
                "midi": int(a.midi),
                "sourceCDM": float(a.source_value),
                "targetCDM": float(a.target_value),
                "L": float(L),
                "STE_Lab_paste": f"{note}  {a.source_value:.6f}  {a.target_value:.6f}",
            }
        )
    return pd.DataFrame(rows)


def inventory_dataframe(relations: Iterable[Relation]) -> pd.DataFrame:
    rows = []
    for rel in relations:
        rows.append(
            {
                "kind": rel.kind,
                "instrument": rel.instrument,
                "dynamic": rel.dynamic,
                "collection": rel.collection,
                "source_cond": rel.source_cond,
                "target_cond": rel.target_cond,
                "n_anchors": rel.n_anchors,
                "grade": rel.grade,
                "used_in_production": "yes" if rel.used_in_production else "no",
            }
        )
    df = pd.DataFrame(rows, columns=INVENTORY_COLUMNS)
    if df.empty:
        return df
    return df.sort_values(
        ["kind", "instrument", "target_cond", "dynamic", "collection"],
        kind="mergesort",
    ).reset_index(drop=True)


def _default_grade(kind: str, collection: str) -> str:
    coll = _norm_coll(collection)
    if kind == "technique" and _is_orch(coll):
        return "Empirical_ORCH"
    if kind == "technique" and coll == "MCGILL":
        return "prediction_McGill"
    if kind == "technique" and coll == "PHIL":
        return "prediction_Philharmonia"
    if kind == "instrument":
        return f"L_instr_{coll}"
    if kind == "collection":
        return f"L_coll_{collection}"
    return f"Empirical_{coll or collection}"


def _techniques_in(measured: dict) -> set[str]:
    techs: set[str] = set()
    for key in measured:
        if isinstance(key, tuple) and len(key) >= 2:
            techs.add(str(key[1]))
    return techs


def _dynamics_in(measured: dict, catalog: dict) -> set[str]:
    dyns: set[str] = set()
    for key in measured:
        if isinstance(key, tuple) and len(key) >= 3:
            dyns.add(str(key[2]))
    for block in (catalog.get("arco") or {}).values():
        if isinstance(block, dict):
            dyns.update(str(d) for d in block)
    return {d for d in dyns if d and d.lower() != "nan"}


def _discover_technique(
    measured: dict,
    instrument: str,
    catalog: dict,
) -> list[Relation]:
    relations: list[Relation] = []
    collections = _collections_in(measured)
    for coll_key, block in (catalog.get("arco") or {}).items():
        if block:
            collections.add(str(coll_key))
    techniques = _techniques_in(measured)
    dynamics = _dynamics_in(measured, catalog)
    floor = catalog.get("harm_floor")
    for collection in collections:
        for tech in techniques:
            if _is_ordinario(tech):
                continue
            for dyn in dynamics:
                source = _source_curve(measured, catalog, collection, "ordinario", dyn)
                target = _lookup(measured, collection, tech, dyn)
                if not source or not target:
                    continue
                min_midi = int(floor) if floor is not None and tech == "harmonics" else None
                anchors = anchors_from_curves(source, target, min_midi=min_midi)
                if not anchors:
                    continue
                relations.append(
                    Relation(
                        kind="technique",
                        instrument=instrument,
                        source_cond="ordinario",
                        target_cond=tech,
                        collection=_norm_coll(collection) or collection,
                        dynamic=dyn,
                        anchors=anchors,
                        grade=_default_grade("technique", collection),
                    )
                )
    return relations


def _discover_collection(
    measured: dict,
    instrument: str,
    catalog: dict,
) -> list[Relation]:
    relations: list[Relation] = []
    collections = sorted(_collections_in(measured) | set((catalog.get("arco") or {})), key=lambda c: _norm_coll(c))
    techniques = _techniques_in(measured) | {"ordinario"}
    dynamics = _dynamics_in(measured, catalog)
    floor = catalog.get("harm_floor")
    seen: set[tuple] = set()
    for tech in techniques:
        for dyn in dynamics:
            present: list[tuple[str, dict[int, float]]] = []
            for collection in collections:
                curve = _source_curve(measured, catalog, collection, tech, dyn)
                if curve:
                    present.append((_norm_coll(collection) or collection, curve))
            # unique by normalised collection id, first curve wins
            by_id: dict[str, dict[int, float]] = {}
            for coll, curve in present:
                by_id.setdefault(coll, curve)
            ids = sorted(by_id)
            for i, src_id in enumerate(ids):
                for tgt_id in ids[i + 1 :]:
                    min_midi = int(floor) if floor is not None and tech == "harmonics" else None
                    anchors = anchors_from_curves(by_id[src_id], by_id[tgt_id], min_midi=min_midi)
                    if not anchors:
                        continue
                    pair = (instrument, tech, dyn, src_id, tgt_id)
                    if pair in seen:
                        continue
                    seen.add(pair)
                    relations.append(
                        Relation(
                            kind="collection",
                            instrument=instrument,
                            source_cond=src_id,
                            target_cond=tgt_id,
                            collection=f"{src_id}->{tgt_id}",
                            dynamic=dyn,
                            anchors=anchors,
                            grade=_default_grade("collection", f"{src_id}->{tgt_id}"),
                        )
                    )
    return relations


def _discover_instrument(
    target_measured: dict,
    donor_measured: dict,
    catalog: dict,
) -> list[Relation]:
    target_id = catalog.get("instrument") or ""
    donor_id = catalog.get("donor_instrument") or family_donor_instrument(target_id)
    if not donor_id:
        return []
    collections = _collections_in(target_measured) | _collections_in(donor_measured)
    dynamics = _dynamics_in(target_measured, catalog) | _dynamics_in(donor_measured, {})
    if not dynamics:
        dynamics = {"pp", "mf", "ff"}
    relations: list[Relation] = []
    empty_cat: dict = {}
    for collection in collections:
        for dyn in dynamics:
            sib = _source_curve(donor_measured, empty_cat, collection, "ordinario", dyn)
            if not sib:
                sib = _source_curve(donor_measured, empty_cat, collection, "ordinario", "mf")
            tgt = _source_curve(target_measured, catalog, collection, "ordinario", dyn)
            if not tgt:
                tgt = _source_curve(target_measured, catalog, collection, "ordinario", "mf")
            if not sib or not tgt:
                continue
            anchors = anchors_from_curves(sib, tgt)
            if not anchors:
                continue
            relations.append(
                Relation(
                    kind="instrument",
                    instrument=target_id,
                    source_cond=donor_id,
                    target_cond=target_id,
                    collection=_norm_coll(collection) or collection,
                    dynamic=dyn,
                    anchors=anchors,
                    grade=_default_grade("instrument", collection),
                )
            )
    return relations


def catalog_for_discovery(
    *,
    instrument: str,
    arco: Optional[dict] = None,
    donor_measured: Optional[dict] = None,
    donor_instrument: Optional[str] = None,
    harm_floor: int | None = None,
) -> dict:
    return {
        "instrument": instrument,
        "arco": arco or {},
        "donor_measured": donor_measured or {},
        "donor_instrument": donor_instrument or family_donor_instrument(instrument),
        "harm_floor": harm_floor,
        "family": orchestral_group(instrument),
    }


def discover_relations(measured: dict, catalog: dict) -> list[Relation]:
    """Enumerate every (kind, instrument, dynamic) pair a collection can teach.

    Collection ids are taken from the trees — they are not a closed list.
    Extra collections may teach a technique that IOWA/ORCH lack.
    Production uses only pp/mf/ff. Harmonics honour catalog['harm_floor']. Pairs with
    no positive overlap are omitted; relations with fewer than 3 anchors are
    kept for inventory (production still inherits mf when n < 3).
    """
    instrument = str(catalog.get("instrument") or "")
    relations: list[Relation] = []
    relations.extend(_discover_technique(measured, instrument, catalog))
    relations.extend(_discover_collection(measured, instrument, catalog))
    donor_measured = catalog.get("donor_measured") or {}
    donor_id = catalog.get("donor_instrument") or family_donor_instrument(instrument)
    if donor_measured:
        donor_cat = dict(catalog)
        donor_cat["instrument"] = donor_id or instrument
        donor_cat["arco"] = {}
        relations.extend(_discover_technique(donor_measured, donor_cat["instrument"], donor_cat))
        relations.extend(_discover_collection(donor_measured, donor_cat["instrument"], donor_cat))
        relations.extend(_discover_instrument(measured, donor_measured, catalog))
    return relations


def _mark(rel: Relation) -> Relation:
    rel.used_in_production = True
    return rel


def select_production_relation(relations: list[Relation], config: Optional[dict] = None) -> list[Relation]:
    """Today's teacher choice, expressed as a policy on Relation objects.

    Strings: ORCH technique teacher at pp/mf/ff; else Philharmonia, then McGill.
    Extra-collection L is how a missing technique is transferred onto Media.
    CORE teachers stay ORCH then Philharmonia then McGill. Same-collection
    leftover dynamics stay available so a missing CORE layer can take the
    closest donor. If a technique has no CORE teacher, every dynamic of the
    winning extra collection is kept and the runner picks the closest.
    Woodwinds: Iowa L_instr when present, else Philharmonia then McGill;
    L_coll is the IOWA→ORCH pair on the sibling (stored, not multiplied).
    Mutates used_in_production on the selected rows and returns them.
    """
    config = config or {}
    for rel in relations:
        rel.used_in_production = False
    instrument = str(config.get("instrument") or (relations[0].instrument if relations else ""))
    family = str(config.get("family") or orchestral_group(instrument))
    selected: list[Relation] = []

    if family != "woodwinds":
        tech = [r for r in relations if r.kind == "technique" and r.instrument == instrument]
        groups: dict[tuple[str, str], list[Relation]] = defaultdict(list)
        for rel in tech:
            groups[(rel.target_cond, rel.dynamic)].append(rel)
        for (target, dyn), cands in sorted(groups.items()):
            if dyn not in PRODUCTION_DYNAMICS:
                continue
            orch = [r for r in cands if _is_orch(r.collection)]
            if orch:
                chosen = max(orch, key=lambda r: (r.n_anchors, r.collection))
                selected.append(_mark(chosen))
                continue
            for coll in ("PHIL", "MCGILL"):
                hits = [r for r in cands if _norm_coll(r.collection) == coll]
                if hits:
                    selected.append(_mark(max(hits, key=lambda r: r.n_anchors)))
                    break
        by_target: dict[str, list[Relation]] = defaultdict(list)
        for rel in selected:
            by_target[rel.target_cond].append(rel)
        extras: list[Relation] = []
        for rel in tech:
            if rel.dynamic in PRODUCTION_DYNAMICS:
                continue
            peers = by_target.get(rel.target_cond) or []
            if peers and any(_teacher_coll_match(rel, peer) for peer in peers):
                extras.append(_mark(rel))
        selected.extend(extras)
        have_core = {r.target_cond for r in selected if r.dynamic in PRODUCTION_DYNAMICS}
        leftover: dict[str, list[Relation]] = defaultdict(list)
        for rel in tech:
            if rel.target_cond in have_core:
                continue
            leftover[rel.target_cond].append(rel)
        for _target, cands in sorted(leftover.items()):
            chosen = None
            for coll in ("PHIL", "MCGILL"):
                hits = [r for r in cands if _norm_coll(r.collection) == coll]
                if hits:
                    chosen = max(hits, key=lambda r: r.n_anchors)
                    break
            if chosen is None and cands:
                chosen = max(cands, key=lambda r: r.n_anchors)
            if chosen is None:
                continue
            win = _norm_coll(chosen.collection)
            for rel in cands:
                if _norm_coll(rel.collection) == win:
                    selected.append(_mark(rel))
        return selected

    instr_rels = [r for r in relations if r.kind == "instrument"]
    by_dyn: dict[str, list[Relation]] = defaultdict(list)
    for rel in instr_rels:
        by_dyn[rel.dynamic].append(rel)
    for dyn, cands in sorted(by_dyn.items()):
        prefer = ["IOWA", "PHIL", "MCGILL"]
        chosen = None
        for coll in prefer:
            hits = [r for r in cands if _norm_coll(r.collection) == coll]
            if hits:
                chosen = max(hits, key=lambda r: r.n_anchors)
                break
        if chosen is None and cands:
            chosen = max(cands, key=lambda r: r.n_anchors)
        if chosen:
            selected.append(_mark(chosen))

    coll_rels = [r for r in relations if r.kind == "collection"]
    donor_id = config.get("donor_instrument") or family_donor_instrument(instrument)
    for rel in coll_rels:
        if rel.source_cond == "IOWA" and rel.target_cond == "ORCH" and rel.dynamic in {"pp", "mf", "ff"}:
            if donor_id and rel.instrument not in {instrument, donor_id}:
                continue
            if rel.instrument == donor_id or (not donor_id and rel.instrument == instrument):
                selected.append(_mark(rel))
    return selected


def production_relations_from_effects(
    effects: list[dict],
    instrument: str,
) -> list[Relation]:
    """Wrap Anchors_all blocks so the runner transfers through Relation objects."""
    out: list[Relation] = []
    for spec in effects:
        name = str(spec.get("name") or "").strip()
        if not name:
            continue
        grade = str(spec.get("grade") or "")
        if grade.lower().startswith("empirical"):
            collection = "ORCH"
        elif "phil" in grade.lower():
            collection = "PHIL"
        else:
            collection = "MCGILL"
        by_dyn = spec.get("anchors_by_dyn") or {"mf": spec.get("anchors") or []}
        for dyn, anchors in by_dyn.items():
            out.append(
                Relation(
                    kind="technique",
                    instrument=instrument,
                    source_cond="ordinario",
                    target_cond=name,
                    collection=collection,
                    dynamic=str(dyn),
                    anchors=list(anchors or []),
                    grade=grade,
                    used_in_production=True,
                )
            )
    return out


def relations_from_pack(pack: dict, instrument: str, harm_floor: int | None = None) -> list[Relation]:
    """Rebuild a measured tree from a loaded preparatory pack and discover."""
    measured: dict = {}

    def _put(coll: str, tech: str, dyn: str, curve: dict) -> None:
        if not curve:
            return
        values: dict[int, Any] = {}
        for midi, payload in curve.items():
            values[int(midi)] = payload
        measured[(coll, tech, dyn)] = {"curve": values}

    for coll, dyns in (pack.get("arco") or {}).items():
        for dyn, curve in (dyns or {}).items():
            _put(str(coll), "ordinario", str(dyn), curve)
    for key, curve in (pack.get("measured") or {}).items():
        if isinstance(key, tuple) and len(key) >= 3:
            _put(str(key[0]), str(key[1]), str(key[2]), curve)
    for key, curve in (pack.get("context") or {}).items():
        if isinstance(key, tuple) and len(key) >= 3:
            _put(str(key[0]), str(key[1]), str(key[2]), curve)
    catalog = catalog_for_discovery(
        instrument=instrument,
        arco=pack.get("arco") or {},
        harm_floor=harm_floor,
    )
    return discover_relations(measured, catalog)


def group_relations(relations: Iterable[Relation]) -> dict[tuple, list[Relation]]:
    groups: dict[tuple, list[Relation]] = defaultdict(list)
    for rel in relations:
        groups[rel.key].append(rel)
    return dict(groups)


def interpolate_relation(rel: Relation, midis: list[int]) -> dict[int, float]:
    return _interp_log_ratio(rel.anchors, midis)


def pooled_log_ratio_field(
    relations: list[Relation],
    midis: list[int],
    *,
    production: Optional[Relation] = None,
    weight: str = "anchors",
    min_collections: int = 2,
) -> tuple[dict[int, float], dict[int, list[tuple[str, float, float]]], list[str]]:
    """Weighted mean of per-collection *anchor* L; single field elsewhere.

    A collection is present at MIDI *m* only when it has an anchor there
    (T>0, S>0). Edge-hold of one collection's interpolator is not an overlap.
    Overlap points become the pooled anchor list; that list is interpolated
    (PCHIP / linear / hold) inside its own span. Outside it, or where fewer
    than ``min_collections`` anchors overlap, the production (single) field
    is used and the fallback is logged.
    """
    messages: list[str] = []
    if not relations:
        return {}, {}, ["pooled: no relations; empty field"]
    per_coll: dict[str, dict[int, float]] = {}
    weights: dict[str, float] = {}
    for rel in relations:
        per_coll[rel.collection] = rel.L_at_anchors()
        weights[rel.collection] = float(rel.n_anchors if weight == "anchors" else 1.0)
    production = production or max(relations, key=lambda r: (r.used_in_production, r.n_anchors))
    single = interpolate_relation(production, midis)
    overlap_L: dict[int, float] = {}
    overlap_contrib: dict[int, list[tuple[str, float, float]]] = {}
    for midi in sorted({m for vals in per_coll.values() for m in vals}):
        present = [
            (coll, vals[midi], weights[coll])
            for coll, vals in per_coll.items()
            if midi in vals
        ]
        if len(present) < int(min_collections):
            continue
        wsum = sum(w for _, _, w in present)
        if wsum <= 0:
            continue
        overlap_L[midi] = sum(val * w for _, val, w in present) / wsum
        overlap_contrib[midi] = present
    pooled_interp: dict[int, float] = {}
    if overlap_L:
        dummy = [Anchor(m, 1.0, math.exp(L)) for m, L in sorted(overlap_L.items())]
        pooled_interp = _interp_log_ratio(dummy, midis)
    lo = min(overlap_L) if overlap_L else None
    hi = max(overlap_L) if overlap_L else None
    field: dict[int, float] = {}
    contributors: dict[int, list[tuple[str, float, float]]] = {}
    for midi in midis:
        inside = lo is not None and hi is not None and lo <= midi <= hi
        if midi in overlap_L:
            field[midi] = overlap_L[midi]
            contributors[midi] = overlap_contrib[midi]
        elif inside and midi in pooled_interp:
            field[midi] = pooled_interp[midi]
            contributors[midi] = overlap_contrib.get(midi, [])
        else:
            if midi in single:
                field[midi] = single[midi]
            contributors[midi] = (
                [(production.collection, single[midi], float(production.n_anchors))]
                if midi in single
                else []
            )
            messages.append(
                f"pooled MIDI {midi}: fewer than min_collections={min_collections} "
                f"anchor overlaps; fell back to single ({production.collection})"
            )
    return field, contributors, messages


def field_for_transfer(
    relations: list[Relation],
    midis: list[int],
    *,
    production: Optional[Relation],
    mode: str = "single",
    weight: str = "anchors",
    min_collections: int = 2,
) -> tuple[dict[int, float], dict[int, list[tuple[str, float, float]]], list[str]]:
    """Dispatch single (today's interpolator) vs optional pooled field."""
    if (mode or "single").strip().lower() != "pooled" or production is None:
        src = production or (relations[0] if relations else None)
        if src is None:
            return {}, {}, []
        L = interpolate_relation(src, midis)
        contrib = {
            m: [(src.collection, L[m], float(src.n_anchors))] for m in L
        }
        return L, contrib, []
    peers = [r for r in relations if r.key == production.key]
    if not peers:
        peers = [production]
    return pooled_log_ratio_field(
        peers,
        midis,
        production=production,
        weight=weight,
        min_collections=min_collections,
    )


def pooled_origin_tag(collections: Iterable[str]) -> str:
    ids = sorted({_norm_coll(c) or c for c in collections if c and "->" not in str(c)})
    if not ids:
        return "modelled_pooled"
    return f"modelled_{'+'.join(ids)}_pooled"
