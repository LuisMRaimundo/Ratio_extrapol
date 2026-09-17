"""Family-specific STE dispatch. Strings and woodwinds are different workflows."""

from __future__ import annotations

from pathlib import Path

from ste_lab.catalog import orchestral_group, resolve_instrument

STRING_EFFECT_TECHNIQUES = frozenset({"harmonics", "sul ponticello", "con sordino"})


class DataDiscoveryError(RuntimeError):
    """Source trees exist but the preparatory pack lost the effects."""


def pipeline_for_instrument(instrument: str) -> str:
    """Return 'strings', 'woodwinds', 'brass', or 'other'."""
    spec = resolve_instrument(instrument)
    key = spec.instrument_id if spec else (instrument or "").strip().lower().replace(" ", "_")
    if key in {"violin", "viola", "cello", "double_bass"}:
        return "strings"
    group = orchestral_group(instrument)
    if group == "strings":
        return "strings"
    if group == "woodwinds":
        return "woodwinds"
    if group == "brass":
        return "brass"
    return "other"


def discovered_orch_effect_techniques(root: Path, instrument: str) -> set[str]:
    from build_ste_preparatory import collect_compiled_specs, find_research_trees

    trees = find_research_trees(Path(root), instrument)
    specs, _ = collect_compiled_specs(trees)
    return {
        spec["technique"]
        for spec in specs
        if str(spec.get("collection", "")).upper() == "ORCH"
        and spec.get("technique") not in {None, "", "ordinario", "arco"}
    }


def assert_string_effects_survived(root: Path, instrument: str, prep: Path) -> None:
    from run_ste_from_preparatory import load_preparatory

    found = discovered_orch_effect_techniques(root, instrument)
    if not found:
        return
    pack = load_preparatory(Path(prep))
    names = {str(spec.get("name", "")).strip() for spec in pack.get("effects") or []}
    if not names:
        raise DataDiscoveryError(
            "REGRESSION/INPUT DISCOVERY ERROR: source effect trees exist but Anchors_all is empty. "
            f"Discovered Orchidea effects: {sorted(found)}. "
            "A string project must not silently degrade to ordinario-only."
        )
    missing = found - names
    if missing:
        raise DataDiscoveryError(
            "REGRESSION/INPUT DISCOVERY ERROR: source effect trees exist but were not "
            f"written to Anchors_all: {sorted(missing)}."
        )


def _rebuild_prep(rebuild_root: Path, instrument: str, fallback: Path) -> Path:
    """Rebuild and return the workbook path written by this call.

    `from build_ste_preparatory import last_path` binds a snapshot and stays
    stale after `build()` assigns the module attribute. Always read it from
    the module.
    """
    import build_ste_preparatory as bsp

    bsp.build(Path(rebuild_root), instrument)
    return Path(bsp.last_path) if bsp.last_path else fallback


def run_string_pipeline(
    prep: Path,
    out: Path,
    instrument: str,
    *,
    rebuild: bool = False,
    rebuild_root: Path | None = None,
    operator: str = "",
    notes: str = "",
) -> list[Path]:
    """Bowed-string workflow: dedicated effect discovery + one STE/Zenodo pair per effect.

    String Media still averages IOWA/ORCH peers (historical rule). Woodwind
    combination (`empirical_only` by default) is not applied to string effect books.
    Dedicated deposit CLIs (`run_ste_effects_batch_*.py`) remain available
    and are not deleted.
    """
    from run_ste_from_preparatory import run

    prep = Path(prep)
    if rebuild and rebuild_root is not None:
        print(f"STRING PIPELINE  rebuild {instrument} from {rebuild_root}")
        prep = _rebuild_prep(Path(rebuild_root), instrument, prep)
        assert_string_effects_survived(Path(rebuild_root), instrument, prep)
    print(f"STRING PIPELINE  run {instrument}  prep={prep}")
    return run(prep, Path(out), instrument, operator=operator, notes=notes)


def run_woodwind_pipeline(
    prep: Path,
    out: Path,
    instrument: str,
    *,
    rebuild: bool = False,
    rebuild_root: Path | None = None,
    operator: str = "",
    notes: str = "",
) -> list[Path]:
    """Woodwind ordinario + family-transfer calibration. No invented string techniques."""
    from run_ste_from_preparatory import run

    prep = Path(prep)
    if rebuild and rebuild_root is not None:
        print(f"WOODWIND PIPELINE  rebuild {instrument} from {rebuild_root}")
        prep = _rebuild_prep(Path(rebuild_root), instrument, prep)
    print(f"WOODWIND PIPELINE  run {instrument}  prep={prep}")
    return run(prep, Path(out), instrument, operator=operator, notes=notes)


def run_selected_pipeline(
    instrument: str,
    prep: Path,
    out: Path,
    *,
    rebuild: bool = False,
    rebuild_root: Path | None = None,
    operator: str = "",
    notes: str = "",
) -> list[Path]:
    kind = pipeline_for_instrument(instrument)
    runner = run_woodwind_pipeline if kind == "woodwinds" else run_string_pipeline
    if kind not in {"strings", "woodwinds"}:
        from run_ste_from_preparatory import run

        prep = Path(prep)
        if rebuild and rebuild_root is not None:
            prep = _rebuild_prep(Path(rebuild_root), instrument, prep)
        return run(prep, Path(out), instrument, operator=operator, notes=notes)
    return runner(
        prep,
        out,
        instrument,
        rebuild=rebuild,
        rebuild_root=rebuild_root,
        operator=operator,
        notes=notes,
    )
