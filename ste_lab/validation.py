"""Cross-collection checks of a transfer relation (always on, non-destructive)."""

from __future__ import annotations

import math
from collections import defaultdict
from typing import Iterable, Optional

import pandas as pd

from .notes import midi_to_label
from .relations import Relation, group_relations
from .transfer import _safe_log_ratio


PAIRWISE_COLUMNS = [
    "kind",
    "instrument",
    "source_cond",
    "target_cond",
    "dynamic",
    "collection_a",
    "collection_b",
    "n_overlap",
    "bias",
    "mae",
    "sd",
    "pearson_r",
    "bias_pct",
]

SPREAD_COLUMNS = [
    "kind",
    "instrument",
    "source_cond",
    "target_cond",
    "dynamic",
    "midi",
    "note",
    "n_collections",
    "L_min",
    "L_max",
    "L_range",
    "sd_L",
    "sd_diff",
    "collections",
]


def _pearson(xs: list[float], ys: list[float]) -> Optional[float]:
    n = len(xs)
    if n < 2:
        return None
    mx = sum(xs) / n
    my = sum(ys) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    denx = math.sqrt(sum((x - mx) ** 2 for x in xs))
    deny = math.sqrt(sum((y - my) ** 2 for y in ys))
    if denx == 0.0 or deny == 0.0:
        return None
    return float(num / (denx * deny))


def _sample_sd(values: list[float]) -> Optional[float]:
    n = len(values)
    if n < 2:
        return None
    mean = sum(values) / n
    var = sum((v - mean) ** 2 for v in values) / (n - 1)
    return float(math.sqrt(var))


def compare_relations(relations: list[Relation]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Pairwise L agreement and per-note spread where ≥2 collections overlap.

    Pairwise metrics use the MIDI intersection of the two anchor sets and
    only pairs with T>0 and S>0 (_safe_log_ratio). bias is mean(L_a − L_b)
    in ln units; bias_pct = 100*(exp(bias)−1).
    """
    pairwise_rows: list[dict] = []
    spread_rows: list[dict] = []
    groups = group_relations(relations)
    for key, peers in sorted(groups.items()):
        if len(peers) < 2:
            continue
        kind, instrument, source_cond, target_cond, dynamic = key
        by_coll: dict[str, Relation] = {}
        for rel in peers:
            by_coll.setdefault(rel.collection, rel)
        if len(by_coll) < 2:
            continue
        ids = sorted(by_coll)
        pair_sd: dict[tuple[str, str], float] = {}
        for i, a_id in enumerate(ids):
            for b_id in ids[i + 1 :]:
                La = by_coll[a_id].L_at_anchors()
                Lb = by_coll[b_id].L_at_anchors()
                overlap = sorted(set(La) & set(Lb))
                xs = [La[m] for m in overlap]
                ys = [Lb[m] for m in overlap]
                # Re-filter through _safe_log_ratio identity (already applied).
                usable = [
                    (x, y)
                    for x, y in zip(xs, ys)
                    if math.isfinite(x) and math.isfinite(y)
                ]
                n = len(usable)
                if n == 0:
                    continue
                diffs = [x - y for x, y in usable]
                bias = sum(diffs) / n
                mae = sum(abs(d) for d in diffs) / n
                sd = _sample_sd(diffs) if n >= 2 else 0.0
                r = _pearson([x for x, _ in usable], [y for _, y in usable])
                pairwise_rows.append(
                    {
                        "kind": kind,
                        "instrument": instrument,
                        "source_cond": source_cond,
                        "target_cond": target_cond,
                        "dynamic": dynamic,
                        "collection_a": a_id,
                        "collection_b": b_id,
                        "n_overlap": n,
                        "bias": bias,
                        "mae": mae,
                        "sd": sd,
                        "pearson_r": r,
                        "bias_pct": 100.0 * (math.exp(bias) - 1.0),
                    }
                )
                if sd is not None:
                    pair_sd[(a_id, b_id)] = float(sd)

        values_at: dict[int, list[tuple[str, float]]] = defaultdict(list)
        for coll, rel in by_coll.items():
            for midi, L in rel.L_at_anchors().items():
                values_at[int(midi)].append((coll, L))
        for midi, hits in sorted(values_at.items()):
            if len(hits) < 2:
                continue
            Ls = [L for _, L in hits]
            colls = [c for c, _ in hits]
            sd_L = _sample_sd(Ls)
            if len(hits) == 2:
                a_id, b_id = sorted(colls)
                sd_diff = pair_sd.get((a_id, b_id), pair_sd.get((b_id, a_id)))
                if sd_diff is None:
                    sd_diff = abs(Ls[0] - Ls[1])
            else:
                sd_diff = sd_L
            spread_rows.append(
                {
                    "kind": kind,
                    "instrument": instrument,
                    "source_cond": source_cond,
                    "target_cond": target_cond,
                    "dynamic": dynamic,
                    "midi": int(midi),
                    "note": midi_to_label(int(midi)),
                    "n_collections": len(hits),
                    "L_min": min(Ls),
                    "L_max": max(Ls),
                    "L_range": max(Ls) - min(Ls),
                    "sd_L": sd_L,
                    "sd_diff": sd_diff,
                    "collections": ";".join(sorted(colls)),
                }
            )

    pairwise = pd.DataFrame(pairwise_rows, columns=PAIRWISE_COLUMNS)
    spread = pd.DataFrame(spread_rows, columns=SPREAD_COLUMNS)
    return pairwise, spread


def relation_level_sd(pairwise: pd.DataFrame, rel: Relation) -> Optional[float]:
    if pairwise is None or pairwise.empty:
        return None
    mask = (
        (pairwise["kind"] == rel.kind)
        & (pairwise["instrument"] == rel.instrument)
        & (pairwise["source_cond"] == rel.source_cond)
        & (pairwise["target_cond"] == rel.target_cond)
        & (pairwise["dynamic"] == rel.dynamic)
    )
    rows = pairwise.loc[mask]
    if rows.empty:
        return None
    vals = [float(v) for v in rows["sd"].tolist() if v is not None and math.isfinite(float(v))]
    if not vals:
        return None
    return sum(vals) / len(vals)


def spread_sd_at(
    spread: pd.DataFrame,
    *,
    kind: str,
    instrument: str,
    source_cond: str,
    target_cond: str,
    dynamic: str,
    midi: int,
    fallback_sd: Optional[float] = None,
) -> Optional[float]:
    if spread is None or spread.empty:
        return None
    mask = (
        (spread["kind"] == kind)
        & (spread["instrument"] == instrument)
        & (spread["source_cond"] == source_cond)
        & (spread["target_cond"] == target_cond)
        & (spread["dynamic"] == dynamic)
        & (spread["midi"] == int(midi))
    )
    rows = spread.loc[mask]
    if rows.empty:
        return None
    rec = rows.iloc[0]
    n = int(rec["n_collections"])
    if n < 2:
        return None
    if n == 2:
        raw = rec.get("sd_diff")
        if raw is None or (isinstance(raw, float) and not math.isfinite(raw)):
            raw = fallback_sd
        return None if raw is None else float(raw)
    raw = rec.get("sd_L")
    if raw is None or (isinstance(raw, float) and not math.isfinite(raw)):
        raw = rec.get("sd_diff", fallback_sd)
    return None if raw is None else float(raw)


def summary_validation_lines(pairwise: pd.DataFrame) -> list[str]:
    """One human-readable line per pairwise relation comparison."""
    lines: list[str] = []
    if pairwise is None or pairwise.empty:
        return lines
    for rec in pairwise.itertuples(index=False):
        label = rec.target_cond
        if rec.kind == "instrument":
            label = f"{rec.target_cond} vs {rec.source_cond}"
        elif rec.kind == "collection":
            label = f"{rec.source_cond}→{rec.target_cond}"
        r = rec.pearson_r
        r_txt = "NA" if r is None or (isinstance(r, float) and not math.isfinite(r)) else f"{float(r):.3f}"
        lines.append(
            f"Transfer assumption check — {label} {rec.dynamic}: "
            f"{rec.collection_a} vs {rec.collection_b} n={int(rec.n_overlap)}, "
            f"bias={float(rec.bias):.4f}, MAE={float(rec.mae):.4f}, r={r_txt}"
        )
    return lines


def validation_evidence_rows(
    relations: list[Relation],
    *,
    technique: str,
    production: Optional[Iterable[Relation]] = None,
) -> list[dict]:
    """Evidence_Map extras for non-production collections that contributed."""
    prod_ids = set()
    for rel in production or []:
        if rel.target_cond == technique or rel.kind != "technique":
            prod_ids.add((rel.kind, rel.collection, rel.dynamic, rel.target_cond))
    rows: list[dict] = []
    seen: set[tuple] = set()
    groups = group_relations(relations)
    for key, peers in groups.items():
        kind, _instrument, _src, target, dynamic = key
        if kind == "technique" and target != technique:
            continue
        if len({r.collection for r in peers}) < 2:
            continue
        for rel in peers:
            if rel.used_in_production:
                continue
            tag = (rel.collection, rel.dynamic, rel.kind, rel.target_cond)
            if tag in seen:
                continue
            if (rel.kind, rel.collection, rel.dynamic, rel.target_cond) in prod_ids:
                continue
            seen.add(tag)
            rows.append(
                {
                    "collection": rel.collection,
                    "dynamic": rel.dynamic,
                    "n_cells": rel.n_anchors,
                    "n_measured": rel.n_anchors,
                    "n_modelled": 0,
                    "evidence_role": "validation",
                    "l_donor_dynamic": "",
                    "l_invariance": "",
                    "l_invariance_spread": "",
                    "principal_evidence": "no",
                    "note": (
                        f"{rel.n_anchors} {rel.kind} anchors in {rel.collection}. "
                        "Validation of the transfer assumption — not Media and not a measured production cell."
                    ),
                }
            )
    return rows
