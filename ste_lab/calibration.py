"""Empirical vs modelled calibration: provenance, transfer, combination, QA.

The numerical family-transfer used through 1.2.8 (scalar median of
IOWA_target / ORCH_donor) is retained only as
``transfer.method = legacy_target_calibrated``. Its outputs are
COMBINED_ESTIMATE, never TRANSFERRED.

Default production transfer is two register-dependent log-ratios
(L_coll on clarinet ORCH vs clarinet IOWA; L_instr on Iowa bass vs
Iowa clarinet) composed with ``transfer._interp_log_ratio``.
Iowa bass used in L_instr is excluded from validation of that transfer.
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable, Optional

from .notes import midi_to_label
from .transfer import Anchor, _interp_log_ratio

PROVENANCE_MEASURED = "MEASURED"
PROVENANCE_INTERPOLATED = "INTERPOLATED"
PROVENANCE_TRANSFERRED = "TRANSFERRED"
PROVENANCE_EXTRAPOLATED = "EXTRAPOLATED"
PROVENANCE_TRANSFERRED_EXTRAPOLATED = "TRANSFERRED_EXTRAPOLATED"
PROVENANCE_COMBINED_ESTIMATE = "COMBINED_ESTIMATE"

QA_HIGH = "HIGH"
QA_MODERATE = "MODERATE"
QA_LOW = "LOW"
QA_REVIEW = "REVIEW_REQUIRED"
QA_REJECTED = "REJECTED"

L_CLIP = 3.0
DEFAULT_REL_EPS = 1e-12
WEIGHT_TOL = 1e-12
SUPPORTED_COMBINATION_METHODS = ("empirical_only", "equal_weight", "legacy_equal_weight")
UNSUPPORTED_COMBINATION_METHODS = {
    "fixed_weight": (
        "fixed_weight has no implemented specification. "
        "Use empirical_only, equal_weight, or legacy_equal_weight."
    ),
    "validation_optimised": (
        "validation_optimised is not connected to a production optimisation workflow. "
        "optimal_weight() is a research helper only. "
        "Use empirical_only, equal_weight, or legacy_equal_weight."
    ),
    "validation_optimized": (
        "validation_optimised is not connected to a production optimisation workflow. "
        "optimal_weight() is a research helper only. "
        "Use empirical_only, equal_weight, or legacy_equal_weight."
    ),
}


class ConfigurationError(ValueError):
    """Unsupported combination mode or inconsistent weights."""
COLLECTION_CODEC = {
    "IOWA": "WAV / studio research compile",
    "ORCH": "Orchidea source (sibling instrument); not a target-instrument recording",
    "PHIL": "MP3 (Philharmonia)",
    "PHILHARMONIA": "MP3 (Philharmonia)",
    "MCGILL": "WAV / McGill research compile",
}

ORIGIN_TO_PROVENANCE = {
    "measured": PROVENANCE_MEASURED,
    "interpolated": PROVENANCE_INTERPOLATED,
    "family_transfer": PROVENANCE_TRANSFERRED,
    "orchidea_family_transfer_estimate": PROVENANCE_TRANSFERRED,
    "extrapolated_ridge": PROVENANCE_EXTRAPOLATED,
    "extrapolated_pchip": PROVENANCE_EXTRAPOLATED,
    "extrapolated_hold": PROVENANCE_EXTRAPOLATED,
    "extrapolated_polynomial": PROVENANCE_EXTRAPOLATED,
    "combined_estimate": PROVENANCE_COMBINED_ESTIMATE,
    "generated": PROVENANCE_COMBINED_ESTIMATE,
}


@dataclass
class CalibrationConfig:
    transfer_enabled: bool = True
    transfer_method: str = "two_log_ratio"
    combination_method: str = "empirical_only"
    empirical_weight: float = 0.5
    transfer_weight: float = 0.5
    validation_enabled: bool = True
    validation_metric: str = "mae"
    validation_min_n: int = 5
    accept_max_semitones: int = 3
    review_max_semitones: int = 6
    allow_long_extrapolation: bool = False
    allow_review_required: bool = False
    relative_epsilon: float = DEFAULT_REL_EPS
    seed: int = 0
    transfer_field_mode: str = "pooled"
    transfer_field_weight: str = "anchors"
    transfer_field_min_collections: int = 2

    def as_export_rows(self) -> list[tuple[str, object]]:
        return [(k, v) for k, v in asdict(self).items()]


def default_config() -> CalibrationConfig:
    return CalibrationConfig()


def load_config(path: Optional[Path] = None) -> CalibrationConfig:
    cfg = default_config()
    if path is None:
        path = Path(__file__).resolve().parents[1] / "calibration.yaml"
    if not path.exists():
        return cfg
    text = path.read_text(encoding="utf-8")
    data = _parse_simple_yaml(text)
    transfer = data.get("transfer") or {}
    combination = data.get("combination") or {}
    validation = data.get("validation") or {}
    extra = data.get("extrapolation") or {}
    disagree = data.get("disagreement") or {}
    cfg.transfer_enabled = bool(transfer.get("enabled", cfg.transfer_enabled))
    cfg.transfer_method = str(transfer.get("method", cfg.transfer_method))
    cfg.combination_method = str(combination.get("method", cfg.combination_method))
    cfg.empirical_weight = float(combination.get("empirical_weight", cfg.empirical_weight))
    cfg.transfer_weight = float(combination.get("transfer_weight", cfg.transfer_weight))
    cfg.validation_enabled = bool(validation.get("enabled", cfg.validation_enabled))
    cfg.validation_metric = str(validation.get("metric", cfg.validation_metric))
    cfg.validation_min_n = int(validation.get("min_n", cfg.validation_min_n))
    cfg.accept_max_semitones = int(extra.get("accept_max_semitones", cfg.accept_max_semitones))
    cfg.review_max_semitones = int(extra.get("review_max_semitones", cfg.review_max_semitones))
    cfg.allow_long_extrapolation = bool(extra.get("allow_long_extrapolation", cfg.allow_long_extrapolation))
    cfg.allow_review_required = bool(extra.get("allow_review_required", cfg.allow_review_required))
    cfg.relative_epsilon = float(disagree.get("relative_epsilon", cfg.relative_epsilon))
    field = data.get("transfer_field") or {}
    cfg.transfer_field_mode = str(field.get("mode", cfg.transfer_field_mode)).strip().lower()
    cfg.transfer_field_weight = str(field.get("weight", cfg.transfer_field_weight)).strip().lower()
    cfg.transfer_field_min_collections = int(field.get("min_collections", cfg.transfer_field_min_collections))
    if cfg.transfer_field_mode not in {"single", "pooled"}:
        raise ConfigurationError(
            f"transfer_field.mode={cfg.transfer_field_mode!r} must be 'single' or 'pooled'."
        )
    if cfg.transfer_field_weight not in {"anchors", "equal"}:
        raise ConfigurationError(
            f"transfer_field.weight={cfg.transfer_field_weight!r} must be 'anchors' or 'equal'."
        )
    if cfg.transfer_field_min_collections < 1:
        raise ConfigurationError("transfer_field.min_collections must be >= 1.")
    validate_combination_config(cfg)
    return cfg


def normalize_combination_method(name: str) -> str:
    method = (name or "empirical_only").strip().lower()
    if method in SUPPORTED_COMBINATION_METHODS:
        return method
    if method in UNSUPPORTED_COMBINATION_METHODS:
        raise ConfigurationError(UNSUPPORTED_COMBINATION_METHODS[method])
    raise ConfigurationError(
        f"Unknown combination method {name!r}. "
        f"Supported: {', '.join(SUPPORTED_COMBINATION_METHODS)}."
    )


def validate_combination_config(cfg: CalibrationConfig) -> CalibrationConfig:
    cfg.combination_method = normalize_combination_method(cfg.combination_method)
    for label, raw in (("empirical_weight", cfg.empirical_weight), ("transfer_weight", cfg.transfer_weight)):
        try:
            val = float(raw)
        except (TypeError, ValueError) as exc:
            raise ConfigurationError(f"{label} must be a finite number in [0, 1].") from exc
        if not math.isfinite(val) or val < 0.0 or val > 1.0:
            raise ConfigurationError(f"{label}={raw!r} must be finite and in [0, 1].")
    expected_t = 1.0 - float(cfg.empirical_weight)
    if abs(float(cfg.transfer_weight) - expected_t) > WEIGHT_TOL:
        raise ConfigurationError(
            f"transfer_weight ({cfg.transfer_weight}) must equal "
            f"1 - empirical_weight ({cfg.empirical_weight}) = {expected_t}. "
            "The authoritative legacy rule is w_T = 1 - w_E."
        )
    cfg.empirical_weight = float(cfg.empirical_weight)
    cfg.transfer_weight = expected_t
    return cfg


def _parse_simple_yaml(text: str) -> dict:
    """Minimal indented mapping parser (no nested lists). Avoids a PyYAML dependency."""
    root: dict = {}
    section: Optional[str] = None
    for raw in text.splitlines():
        line = raw.split("#", 1)[0].rstrip()
        if not line.strip():
            continue
        if not line.startswith(" ") and line.endswith(":"):
            section = line[:-1].strip()
            root[section] = {}
            continue
        if ":" not in line:
            continue
        key, val = line.split(":", 1)
        key, val = key.strip(), val.strip()
        parsed: object
        if val.lower() in {"true", "yes"}:
            parsed = True
        elif val.lower() in {"false", "no"}:
            parsed = False
        else:
            try:
                parsed = int(val)
            except ValueError:
                try:
                    parsed = float(val)
                except ValueError:
                    parsed = val
        if section and line.startswith(" "):
            root[section][key] = parsed
        else:
            root[key] = parsed
            section = None
    return root


def provenance_of_origin(origin: str, *, outside_measured: bool = False, legacy_scalar: bool = False) -> str:
    if legacy_scalar:
        return PROVENANCE_COMBINED_ESTIMATE
    base = ORIGIN_TO_PROVENANCE.get((origin or "").strip().lower(), PROVENANCE_TRANSFERRED)
    if base == PROVENANCE_TRANSFERRED and outside_measured:
        return PROVENANCE_TRANSFERRED_EXTRAPOLATED
    if base == PROVENANCE_INTERPOLATED and outside_measured:
        return PROVENANCE_EXTRAPOLATED
    return base


def measured_span(curve: dict[int, float]) -> tuple[Optional[int], Optional[int]]:
    keys = [m for m, v in curve.items() if v is not None and v > 0 and math.isfinite(v)]
    if not keys:
        return None, None
    return min(keys), max(keys)


def extrapolation_distance_semitones(midi: int, lo: Optional[int], hi: Optional[int]) -> int:
    if lo is None or hi is None:
        return 0
    if lo <= midi <= hi:
        return 0
    if midi < lo:
        return lo - midi
    return midi - hi


def legacy_target_scale(donor: dict[int, float], target: dict[int, float]) -> float:
    """Deprecated: median(target/donor) on overlap. Calibrates the donor on the target."""
    ratios = [
        target[m] / val
        for m, val in donor.items()
        if val > 0 and target.get(m, 0) > 0 and math.isfinite(val) and math.isfinite(target[m])
    ]
    if not ratios:
        return 1.0
    ratios.sort()
    mid = len(ratios) // 2
    if len(ratios) % 2:
        return float(ratios[mid])
    return float(0.5 * (ratios[mid - 1] + ratios[mid]))


def _anchors(source: dict[int, float], target: dict[int, float]) -> list[Anchor]:
    out = []
    for midi, src in source.items():
        tgt = target.get(midi)
        if src and src > 0 and tgt and tgt > 0 and math.isfinite(src) and math.isfinite(tgt):
            out.append(Anchor(int(midi), float(src), float(tgt)))
    return out


def apply_log_ratio(source_curve: dict[int, float], L: dict[int, float], midis: Iterable[int]) -> dict[int, float]:
    out = {}
    for midi in midis:
        src = source_curve.get(midi)
        if src is None or src <= 0 or not math.isfinite(src):
            continue
        if midi not in L:
            continue
        value = src * math.exp(min(L_CLIP, max(-L_CLIP, L[midi])))
        if value > 0 and math.isfinite(value):
            out[int(midi)] = float(value)
    return out


@dataclass
class TransferEstimate:
    curve: dict[int, float]
    method: str
    L_coll: dict[int, float] = field(default_factory=dict)
    L_instr: dict[int, float] = field(default_factory=dict)
    n_L_coll_anchors: int = 0
    n_L_instr_anchors: int = 0
    L_instr_status: str = "UNDETERMINED"
    fitted_on: list[str] = field(default_factory=list)
    validation_excluded: list[str] = field(default_factory=list)
    scale_legacy: Optional[float] = None
    notes: str = ""

    def origin_tag(self) -> str:
        if self.method == "legacy_target_calibrated":
            return "combined_estimate"
        return "orchidea_family_transfer_estimate"


def two_log_ratio_transfer(
    orch_clarinet: dict[int, float],
    iowa_clarinet: dict[int, float],
    iowa_target: dict[int, float],
    midis: list[int],
) -> TransferEstimate:
    """ORCH_target(m) = ORCH_clarinet(m) × exp(L_instr(m)).

    L_coll(m) = ln(ORCH_clarinet / IOWA_clarinet)  (same instrument)
    L_instr(m) = ln(IOWA_target / IOWA_clarinet)   (same collection)

    L_instr uses the target Iowa curve, so IOWA is excluded from validation.
    """
    coll_anchors = _anchors(iowa_clarinet, orch_clarinet)
    instr_anchors = _anchors(iowa_clarinet, iowa_target)
    L_coll = _interp_log_ratio(coll_anchors, midis) if coll_anchors else {}
    L_instr = _interp_log_ratio(instr_anchors, midis) if instr_anchors else {}
    curve = apply_log_ratio(orch_clarinet, L_instr, midis) if L_instr else {}
    status = "estimated" if instr_anchors else "UNDETERMINED"
    excluded = ["IOWA"] if instr_anchors else []
    note = (
        "L_instr from Iowa target vs Iowa sibling clarinet; "
        "ORCH_target = ORCH_clarinet × exp(L_instr). "
        "Iowa is excluded from validation of this transfer."
        if instr_anchors
        else "L_instr UNDETERMINED: no same-collection sibling overlap. No transfer cells."
    )
    return TransferEstimate(
        curve=curve,
        method="two_log_ratio",
        L_coll=L_coll,
        L_instr=L_instr,
        n_L_coll_anchors=len(coll_anchors),
        n_L_instr_anchors=len(instr_anchors),
        L_instr_status=status,
        fitted_on=["IOWA_clarinet", "ORCH_clarinet"] + (["IOWA_target"] if instr_anchors else []),
        validation_excluded=excluded,
        notes=note,
    )


def legacy_scalar_transfer(
    orch_clarinet: dict[int, float],
    iowa_target: dict[int, float],
    midis: list[int],
    lo: int,
    hi: int,
) -> TransferEstimate:
    scale = legacy_target_scale(orch_clarinet, iowa_target)
    curve = {
        int(m): float(orch_clarinet[m] * scale)
        for m in midis
        if lo <= m <= hi and orch_clarinet.get(m, 0) > 0
    }
    return TransferEstimate(
        curve=curve,
        method="legacy_target_calibrated",
        scale_legacy=scale,
        fitted_on=["IOWA_target", "ORCH_clarinet"],
        validation_excluded=["IOWA"],
        notes=(
            "DEPRECATED scalar median(IOWA_target / ORCH_donor). "
            "Centres the donor on the target. Labelled COMBINED_ESTIMATE, not TRANSFERRED."
        ),
    )


def combine_empirical_and_modelled(
    empirical: Optional[float],
    transfer: Optional[float],
    cfg: CalibrationConfig,
    *,
    empirical_weight: Optional[float] = None,
) -> tuple[Optional[float], str, float, float]:
    """Return (value, method, w_E, w_T).

    empirical_only: a valid target empirical value always wins.
    equal_weight / legacy_equal_weight: arithmetic mix when both exist.
    Effective transfer weight is always 1 - w_E.
    """
    validate_combination_config(cfg)
    method = cfg.combination_method
    w_e = cfg.empirical_weight if empirical_weight is None else float(empirical_weight)
    if not math.isfinite(w_e) or w_e < 0.0 or w_e > 1.0:
        raise ConfigurationError(f"empirical_weight={w_e!r} must be finite and in [0, 1].")
    w_t = 1.0 - w_e
    has_e = empirical is not None and empirical > 0 and math.isfinite(empirical)
    has_t = transfer is not None and transfer > 0 and math.isfinite(transfer)
    if method == "legacy_equal_weight" or method == "equal_weight":
        if method == "equal_weight":
            w_e, w_t = 0.5, 0.5
        if has_e and has_t:
            return w_e * empirical + w_t * transfer, method, w_e, w_t
        if has_e:
            return empirical, method, 1.0, 0.0
        if has_t:
            return transfer, method, 0.0, 1.0
        return None, method, w_e, w_t
    if has_e:
        return empirical, "empirical_only", 1.0, 0.0
    if has_t:
        return transfer, "transfer_only", 0.0, 1.0
    return None, method, w_e, w_t


def finite_pairs(a: dict[int, float], b: dict[int, float]) -> list[tuple[int, float, float]]:
    out = []
    for midi, va in a.items():
        vb = b.get(midi)
        if va and vb and va > 0 and vb > 0 and math.isfinite(va) and math.isfinite(vb):
            out.append((int(midi), float(va), float(vb)))
    return out


def agreement_metrics(pred: dict[int, float], truth: dict[int, float]) -> dict:
    """Agreement on positive finite overlap.

    ln_rmse is RMSE in log space, not MSE:

        sqrt( mean( (ln(predicted) - ln(observed))**2 ) )
    """
    pairs = finite_pairs(pred, truth)
    n = len(pairs)
    empty = {"n": 0, "r": None, "mae": None, "rmse": None, "medae": None, "mape": None, "ln_rmse": None}
    if n == 0:
        return empty
    err = [p - t for _, p, t in pairs]
    abs_err = [abs(e) for e in err]
    mae = sum(abs_err) / n
    rmse = math.sqrt(sum(e * e for e in err) / n)
    ln_err = [math.log(p) - math.log(t) for _, p, t in pairs]
    ln_rmse = math.sqrt(sum(e * e for e in ln_err) / n)
    medae = sorted(abs_err)[n // 2]
    mape = sum(abs(p - t) / t for _, p, t in pairs) / n
    r = None
    if n >= 3:
        mp = sum(p for _, p, _ in pairs) / n
        mt = sum(t for _, _, t in pairs) / n
        num = sum((p - mp) * (t - mt) for _, p, t in pairs)
        denp = math.sqrt(sum((p - mp) ** 2 for _, p, _ in pairs))
        dent = math.sqrt(sum((t - mt) ** 2 for _, _, t in pairs))
        if denp > 0 and dent > 0:
            r = num / (denp * dent)
    return {"n": n, "r": r, "mae": mae, "rmse": rmse, "medae": medae, "mape": mape, "ln_rmse": ln_rmse}


def _median(values: list[float]) -> Optional[float]:
    if not values:
        return None
    ordered = sorted(values)
    mid = len(ordered) // 2
    if len(ordered) % 2:
        return float(ordered[mid])
    return float(0.5 * (ordered[mid - 1] + ordered[mid]))


def shape_normalised_metrics(pred: dict[int, float], truth: dict[int, float]) -> dict:
    """Collection-offset-normalised shape diagnostics. Validation only.

    c = median(ln(observed) - ln(predicted))
    scale_factor = exp(c)
    e_i = ln(observed_i) - (ln(predicted_i) + c)
    shape_ln_rmse = sqrt(mean(e_i**2))

    Does not modify production calibration values.
    """
    skipped: list[dict] = []
    logs: list[tuple[int, float, float]] = []
    for midi in sorted(set(pred) | set(truth)):
        if midi not in pred or midi not in truth:
            continue
        y_hat, y = pred[midi], truth[midi]
        if y is None or not math.isfinite(float(y)) or float(y) <= 0:
            skipped.append({"midi": int(midi), "reason": "nonpositive_or_invalid_observed"})
            continue
        if y_hat is None or not math.isfinite(float(y_hat)) or float(y_hat) <= 0:
            skipped.append({"midi": int(midi), "reason": "nonpositive_or_invalid_prediction"})
            continue
        logs.append((int(midi), math.log(float(y)), math.log(float(y_hat))))
    empty = {
        "n": 0,
        "collection_log_offset": None,
        "collection_scale_factor": None,
        "shape_ln_rmse": None,
        "shape_ln_mae": None,
        "n_skipped_nonpositive": len(skipped),
        "skip_reasons": skipped,
    }
    if not logs:
        return empty
    c = _median([ly - lh for _, ly, lh in logs])
    if c is None:
        return empty
    residuals = [ly - (lh + c) for _, ly, lh in logs]
    n = len(residuals)
    return {
        "n": n,
        "collection_log_offset": c,
        "collection_scale_factor": math.exp(c),
        "shape_ln_rmse": math.sqrt(sum(e * e for e in residuals) / n),
        "shape_ln_mae": sum(abs(e) for e in residuals) / n,
        "n_skipped_nonpositive": len(skipped),
        "skip_reasons": skipped,
    }


def collection_offset(a: dict[int, float], b: dict[int, float]) -> Optional[float]:
    """Median a/b on overlap — collection offset, not a transfer quality score."""
    ratios = [p / t for _, p, t in finite_pairs(a, b)]
    if not ratios:
        return None
    ratios.sort()
    mid = len(ratios) // 2
    if len(ratios) % 2:
        return float(ratios[mid])
    return float(0.5 * (ratios[mid - 1] + ratios[mid]))


def optimal_weight(
    empirical: dict[int, float],
    transfer: dict[int, float],
    truth: dict[int, float],
    metric: str = "mae",
    min_n: int = 5,
    step: float = 0.01,
) -> Optional[dict]:
    midis = [
        m
        for m in empirical
        if m in transfer
        and m in truth
        and empirical[m] > 0
        and transfer[m] > 0
        and truth[m] > 0
    ]
    if len(midis) < min_n:
        return None
    best_w, best_score, best_metrics = None, None, None
    w = 0.0
    while w <= 1.0 + 1e-12:
        pred = {m: w * empirical[m] + (1.0 - w) * transfer[m] for m in midis}
        mets = agreement_metrics(pred, {m: truth[m] for m in midis})
        score = mets["mae"] if metric == "mae" else mets["rmse"]
        if score is not None and (best_score is None or score < best_score):
            best_w, best_score, best_metrics = w, score, mets
        w = round(w + step, 6)
    if best_w is None:
        return None
    return {"weight": best_w, "score": best_score, "metric": metric, **best_metrics}


def disagreement(empirical: Optional[float], transfer: Optional[float], eps: float = DEFAULT_REL_EPS) -> tuple[Optional[float], Optional[float]]:
    if empirical is None or transfer is None:
        return None, None
    if not (math.isfinite(empirical) and math.isfinite(transfer)):
        return None, None
    delta = abs(empirical - transfer)
    rel = delta / max(abs(empirical), eps)
    return delta, rel


def qa_status(
    *,
    provenance: str,
    extrap_st: int,
    cfg: CalibrationConfig,
    has_empirical: bool,
    has_transfer: bool,
    relative_diff: Optional[float] = None,
    value: Optional[float] = None,
) -> str:
    if value is None or not math.isfinite(value) or value <= 0:
        return QA_REJECTED
    if provenance == PROVENANCE_MEASURED:
        return QA_HIGH
    if provenance == PROVENANCE_INTERPOLATED and extrap_st == 0:
        return QA_HIGH
    if extrap_st > cfg.review_max_semitones:
        return QA_REJECTED if not cfg.allow_long_extrapolation else QA_LOW
    if extrap_st > cfg.accept_max_semitones:
        return QA_REVIEW
    if provenance == PROVENANCE_TRANSFERRED_EXTRAPOLATED:
        return QA_REVIEW if extrap_st > cfg.accept_max_semitones else QA_MODERATE
    if provenance == PROVENANCE_TRANSFERRED:
        return QA_MODERATE
    if provenance == PROVENANCE_COMBINED_ESTIMATE:
        if relative_diff is not None and relative_diff > 0.5:
            return QA_REVIEW
        return QA_MODERATE
    if has_empirical:
        return QA_HIGH
    if has_transfer:
        return QA_MODERATE
    return QA_LOW


def automatic_acceptance(status: str, cfg: CalibrationConfig) -> bool:
    """HIGH and MODERATE are production-eligible. REVIEW_REQUIRED and REJECTED are not."""
    if status in {QA_HIGH, QA_MODERATE}:
        return True
    if status == QA_LOW and cfg.allow_long_extrapolation:
        return True
    return False


def accepted_for_final(status: str, cfg: CalibrationConfig) -> bool:
    if automatic_acceptance(status, cfg):
        return True
    if status == QA_REVIEW and cfg.allow_review_required:
        return True
    return False


def validation_status_of(qa_flag: str, *, accepted_final: bool) -> str:
    if qa_flag == QA_REJECTED:
        return "rejected"
    if qa_flag == QA_REVIEW and not accepted_final:
        return "review_required"
    if accepted_final:
        return "accepted"
    if qa_flag == QA_REVIEW:
        return "review_required"
    return "rejected"


def acceptance_counts(rows: list[dict]) -> dict[str, int]:
    accepted = review = rejected = 0
    for rec in rows:
        flag = rec.get("qa_flag")
        if rec.get("accepted_final"):
            accepted += 1
        elif flag == QA_REVIEW:
            review += 1
        else:
            rejected += 1
    return {
        "rows_accepted": accepted,
        "rows_review_required": review,
        "rows_rejected": rejected,
        "rows_total": accepted + review + rejected,
    }


def qa_index(rows: list[dict]) -> dict[tuple[str, int], dict]:
    return {(str(rec["dynamic"]), int(rec["midi"])): rec for rec in rows if rec.get("midi") is not None}


_LAST_ESTIMATES: dict[str, TransferEstimate] = {}


def remember_transfer_estimates(estimates: Optional[dict[str, TransferEstimate]]) -> None:
    """Keep the last two-ratio (or legacy) fit so export sheets retain L_* metadata."""
    global _LAST_ESTIMATES
    _LAST_ESTIMATES = dict(estimates or {})


def transfer_estimates_from_layers(
    layers: dict,
    cfg: CalibrationConfig,
    stored: Optional[dict[str, TransferEstimate]] = None,
) -> dict[str, TransferEstimate]:
    remembered = stored if stored is not None else _LAST_ESTIMATES
    out: dict[str, TransferEstimate] = {}
    for dyn in ("pp", "mf", "ff"):
        if remembered.get(dyn) and remembered[dyn].curve:
            out[dyn] = remembered[dyn]
            continue
        layer = layers.get(("ORCH", dyn)) if layers else None
        curve = {}
        origin = ""
        if layer:
            curve = {
                int(m): float(c.value)
                for m, c in layer.cells.items()
                if c.value and c.value > 0 and math.isfinite(c.value)
            }
            if layer.cells:
                origin = next(iter(layer.cells.values())).origin or ""
        method = (
            "legacy_target_calibrated"
            if origin == "combined_estimate" or cfg.transfer_method == "legacy_target_calibrated"
            else "two_log_ratio"
        )
        out[dyn] = TransferEstimate(
            curve=curve,
            method=method,
            L_instr_status="estimated" if curve else "UNDETERMINED",
            fitted_on=["IOWA_target", "ORCH_clarinet"] if method == "legacy_target_calibrated" else ["IOWA_clarinet", "ORCH_clarinet", "IOWA_target"],
            validation_excluded=["IOWA"],
            notes="ORCH layer at export; Iowa excluded from transfer validation.",
        )
    return out


def _collection_key(collection: str) -> Optional[str]:
    coll = (collection or "").upper().strip()
    if coll in {"ORCHIDEA", "ORCH."}:
        coll = "ORCH"
    if coll == "PHILHARMONIA":
        coll = "PHIL"
    if coll not in {"IOWA", "ORCH", "PHIL", "MCGILL"}:
        return None
    return coll


def curves_from_project(project, technique: str = "ordinario") -> dict[tuple[str, str], dict[int, float]]:
    out: dict[tuple[str, str], dict[int, float]] = {}
    for layer in project.layers:
        if (layer.technique or "").strip().lower() != technique:
            continue
        coll = _collection_key(layer.collection)
        if not coll:
            continue
        curve = {
            int(m): float(c.value)
            for m, c in layer.cells.items()
            if c.value and c.value > 0 and math.isfinite(c.value)
        }
        if curve:
            out[(coll, layer.dynamic)] = curve
    return out


def origins_from_project(project, technique: str = "ordinario") -> dict[str, str]:
    """Collection → layer-cell origin for Measured_Data provenance."""
    out: dict[str, str] = {}
    for layer in project.layers:
        if (layer.technique or "").strip().lower() != technique:
            continue
        coll = _collection_key(layer.collection)
        if not coll or coll in out:
            continue
        for cell in layer.cells.values():
            if cell.origin:
                out[coll] = cell.origin
                break
    return out


def build_validation_tables(
    measured: dict[tuple[str, str], dict[int, float]],
    transfer_by_dyn: dict[str, TransferEstimate],
    cfg: CalibrationConfig,
) -> tuple[list[dict], list[dict], list[str]]:
    excluded = sorted({c for est in transfer_by_dyn.values() for c in est.validation_excluded})
    rows = []
    offsets = []
    validators = ["PHIL", "MCGILL"]
    for dyn in ("pp", "mf", "ff"):
        iowa = measured.get(("IOWA", dyn), {})
        trans = transfer_by_dyn[dyn].curve if dyn in transfer_by_dyn else {}
        for other in validators:
            if other in excluded:
                continue
            truth = measured.get((other, dyn), {})
            if not truth:
                continue
            for model, pred in (("IOWA", iowa), ("transfer", trans), ("combined_empirical_only", iowa)):
                mets = agreement_metrics(pred, truth)
                shape = shape_normalised_metrics(pred, truth)
                rows.append(
                    {
                        "dynamic": dyn,
                        "validation_collection": other,
                        "model": model,
                        "n": mets["n"],
                        "r": mets["r"],
                        "mae": mets["mae"],
                        "rmse": mets["rmse"],
                        "ln_rmse": mets["ln_rmse"],
                        "collection_log_offset": shape["collection_log_offset"],
                        "collection_scale_factor": shape["collection_scale_factor"],
                        "shape_ln_rmse": shape["shape_ln_rmse"],
                        "shape_ln_mae": shape["shape_ln_mae"],
                        "n_skipped_nonpositive": shape["n_skipped_nonpositive"],
                        "skip_reasons": shape["skip_reasons"],
                    }
                )
            pairs = finite_pairs(iowa, truth)
            offsets.append(
                {
                    "dynamic": dyn,
                    "collection_a": "IOWA",
                    "collection_b": other,
                    "n": len(pairs),
                    "median_ratio": collection_offset(iowa, truth),
                }
            )
    return rows, offsets, excluded


def build_final_rows(
    *,
    instrument: str,
    dynamic: str,
    empirical: dict[int, float],
    transfer: dict[int, float],
    cfg: CalibrationConfig,
    empirical_source: str = "IOWA",
    transfer_source: str = "orchidea_family_transfer_estimate",
    transfer_method: str = "two_log_ratio",
    validation_n: Optional[int] = None,
    validation_mae: Optional[float] = None,
    validation_rmse: Optional[float] = None,
    validation_r: Optional[float] = None,
    weight_e: Optional[float] = None,
) -> list[dict]:
    lo, hi = measured_span(empirical)
    midis = sorted(set(empirical) | set(transfer))
    rows = []
    for midi in midis:
        e = empirical.get(midi)
        t = transfer.get(midi)
        if e is not None and not (e > 0 and math.isfinite(e)):
            e = None
        if t is not None and not (t > 0 and math.isfinite(t)):
            t = None
        dist = extrapolation_distance_semitones(midi, lo, hi)
        outside = dist > 0
        final, method, we, wt = combine_empirical_and_modelled(e, t, cfg, empirical_weight=weight_e)
        blended = (
            e is not None
            and t is not None
            and method in {"equal_weight", "legacy_equal_weight"}
            and 0.0 < we < 1.0
            and 0.0 < wt < 1.0
        )
        if blended:
            prov = PROVENANCE_COMBINED_ESTIMATE
        elif e is not None:
            prov = PROVENANCE_MEASURED
        elif t is not None and transfer_method == "legacy_target_calibrated":
            prov = PROVENANCE_COMBINED_ESTIMATE
        elif t is not None and outside:
            prov = PROVENANCE_TRANSFERRED_EXTRAPOLATED
        elif t is not None:
            prov = PROVENANCE_TRANSFERRED
        else:
            continue
        delta, rel = disagreement(e, t, cfg.relative_epsilon)
        status = qa_status(
            provenance=prov,
            extrap_st=dist,
            cfg=cfg,
            has_empirical=e is not None,
            has_transfer=t is not None,
            relative_diff=rel,
            value=final,
        )
        auto = automatic_acceptance(status, cfg)
        override = bool(status == QA_REVIEW and cfg.allow_review_required)
        include = accepted_for_final(status, cfg)
        keep_value = final if status != QA_REJECTED else None
        rows.append(
            {
                "instrument": instrument,
                "dynamic": dynamic,
                "note": midi_to_label(midi),
                "midi": midi,
                "empirical_value": e,
                "empirical_source": empirical_source if e is not None else None,
                "transfer_estimate": t,
                "transfer_source": transfer_source if t is not None else None,
                "final_value": keep_value,
                "combination_method": method,
                "empirical_weight": we,
                "transfer_weight": wt,
                "provenance": prov,
                "extrapolation_distance_semitones": dist,
                "absolute_difference": delta,
                "relative_difference": rel,
                "validation_n": validation_n,
                "validation_mae": validation_mae,
                "validation_rmse": validation_rmse,
                "validation_r": validation_r,
                "confidence": status,
                "qa_flag": status,
                "automatic_acceptance": auto,
                "acceptance_override": override,
                "accepted_final": include,
                "final_acceptance": include,
                "validation_status": validation_status_of(status, accepted_final=include),
                "source_pitch_min": lo,
                "source_pitch_max": hi,
                "configured_combination_method": cfg.combination_method,
                "configured_empirical_weight": cfg.empirical_weight,
                "configured_transfer_weight": cfg.transfer_weight,
            }
        )
    return rows
