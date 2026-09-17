# -*- coding: utf-8 -*-
"""Batch STE Lab run: IOWA + ORCHIDEA for every violin effect we can ground in data."""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
WORKSPACE = ROOT.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ste_lab.evidence import stamp_evidence
from ste_lab.excel_export import export_workbook
from ste_lab.notes import midi_to_label, parse_pitch
from ste_lab.qa import audit_project
from ste_lab.session import Project, layer_from_identity
from ste_lab.media_policy import production_media_of
from ste_lab.transfer import Anchor, technique_transfer
from ste_lab.zenodo_export import export_zenodo_workbook

ARCO = Path(r"D:\CORDAS_2\VIOLIN\VIOLIN_Zenodo_collections_Arco_normal.xlsx")
OUT = WORKSPACE / "VIOLIN_3"
DYN_MAP = {
    "pianissimo": "pp",
    "piano": "p",
    "mezzo-piano": "mp",
    "mezzo-forte": "mf",
    "forte": "f",
    "fortissimo": "ff",
}
CORE_DYNS = ("pp", "mf", "ff")
COMPILED_RESEARCH = "compiled_density_metrics_research.xlsx"
STABLE_TOKEN = "_sustains_stable"


def _compiled_research_files(root: Path) -> list[Path]:
    return [
        path
        for path in root.rglob(f"*{COMPILED_RESEARCH}")
        if path.is_file() and not path.name.startswith("~$")
    ]


def _stable_dirs(root: Path) -> list[Path]:
    return [
        path
        for path in root.rglob("*")
        if path.is_dir() and STABLE_TOKEN in path.name.lower()
    ]


def _skipped_raw_sustains(path: Path) -> bool:
    text = str(path).lower()
    return "_sustains" in text and STABLE_TOKEN not in text


def report_research_tree(tree: Path, *, collection: str = "", technique: str = "") -> list[str]:
    """Print every folder searched and every compiled book found. Also warn if empty."""
    tag = f"{collection} {technique}".strip() or tree.name
    warnings: list[str] = []

    def note(msg: str) -> None:
        line = f"WARNING: {msg}"
        print(line)
        warnings.append(line)

    print(f"\nSEARCH  {tag}")
    print(f"  folder  {tree}")
    if not tree.exists():
        print("  result  MISSING")
        note(f"missing research folder ({tag}): {tree}")
        return warnings

    print("  result  found")
    stables = _stable_dirs(tree)
    print(f"  _Sustains_Stable folders  {len(stables)}")
    for folder in stables:
        inside = _compiled_research_files(folder)
        print(f"    {folder}")
        print(f"      compiled_books={len(inside)}")
        if not inside:
            note(f"_Sustains_Stable has no *{COMPILED_RESEARCH}: {folder}")

    compiled = _compiled_research_files(tree)
    usable = [path for path in compiled if not _skipped_raw_sustains(path)]
    print(f"  compiled research xlsx  {len(compiled)}  (usable {len(usable)})")
    for path in compiled:
        if _skipped_raw_sustains(path):
            print(f"    SKIP  raw _Sustains  {path}")
        else:
            print(f"    FOUND {path}")

    if not stables:
        note(f"no _Sustains_Stable folder under {tag}: {tree}")
    if not compiled:
        note(f"no *{COMPILED_RESEARCH} under {tag}: {tree}")
    elif not usable:
        note(
            f"compiled research books under {tag} sit only in raw _Sustains "
            f"(skipped). Need a _Sustains_Stable copy: {tree}"
        )
    return warnings


def audit_research_tree(tree: Path, *, collection: str = "", technique: str = "") -> list[str]:
    """Warn if a research tree has no _Sustains_Stable folder or compiled book."""
    return report_research_tree(tree, collection=collection, technique=technique)


def log_discovered_specs(specs: list[dict]) -> None:
    print("\nFILES MATCHED  (technique / dynamic from folder names)")
    if not specs:
        print("  (none)")
        return
    for spec in specs:
        print(
            f"  {spec['collection']:6} {spec['technique']:16} {spec['dynamic']:3}  "
            f"{spec['path']}"
        )


def _extract_mass(df: pd.DataFrame) -> dict[int, tuple[str, float]]:
    if df is None or df.empty:
        return {}
    note_col = next((c for c in df.columns if str(c).strip().lower() == "note"), None)
    mass_col = next((c for c in df.columns if str(c).strip().lower() == "spectral_mass"), None)
    if note_col is None or mass_col is None:
        return {}
    out: dict[int, tuple[str, float]] = {}
    for _, row in df.iterrows():
        pitch = parse_pitch(row.get(note_col))
        try:
            value = float(row.get(mass_col))
        except (TypeError, ValueError):
            continue
        if pitch is None or value <= 0:
            continue
        out[pitch.midi] = (pitch.label, value)
    return out


def load_spectral_mass(path: Path) -> dict[int, tuple[str, float]]:
    with pd.ExcelFile(path) as xl:
        names = list(xl.sheet_names)
    for sheet in ("Research_Core", "Spectral_Density_Metrics"):
        if sheet not in names:
            continue
        got = _extract_mass(pd.read_excel(path, sheet_name=sheet))
        if got:
            print(f"  READ   {path}")
            print(f"         sheet={sheet}  columns=note,spectral_mass  n={len(got)}")
            return got
    print(f"  READ   {path}")
    print("         no note+spectral_mass on Research_Core / Spectral_Density_Metrics")
    return {}


def resolve_media_sheet(sheet_names: list[str], preferred: str | None = None) -> str | None:
    if preferred and preferred in sheet_names:
        return preferred
    for name in sheet_names:
        low = name.lower()
        if low.endswith("_media") or low == "media":
            return name
    return preferred


def _iowa_dynamic_sheets(sheet_names: list[str]) -> dict[str, str]:
    """Zenodo IOWA sheets keep compiled CDM by string (C/G/D/A…), not in _Sustains_Stable."""
    found: dict[str, str] = {}
    for dyn in CORE_DYNS:
        hits = []
        for name in sheet_names:
            low = name.lower().replace(" ", "_")
            if "iowa" not in low or "media" in low or "meta" in low:
                continue
            if low.endswith("_" + dyn) or low.endswith(dyn):
                hits.append(name)
        if hits:
            found[dyn] = sorted(hits, key=len)[0]
    return found


def _curve_from_iowa_string_sheet(path: Path, sheet: str) -> dict[int, float]:
    df = pd.read_excel(path, sheet_name=sheet)
    note_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"source note", "note"}),
        None,
    )
    val_col = next(
        (c for c in df.columns if str(c).strip().lower() in {"cdm - media", "cdm-media"}),
        None,
    )
    if val_col is None:
        val_col = next((c for c in df.columns if str(c).strip().lower() == "cdm"), None)
    if note_col is None or val_col is None:
        return {}
    out: dict[int, float] = {}
    for _, row in df.iterrows():
        pitch = parse_pitch(row.get(note_col))
        try:
            value = float(row.get(val_col))
        except (TypeError, ValueError):
            continue
        if pitch is None or value <= 0:
            continue
        out[pitch.midi] = value
    return out


def audit_iowa_zenodo_workbook(path: Path, media_sheet: str | None = None) -> list[str]:
    """Warn if the string-instrument Zenodo book is missing IOWA (by-string) data.

    IOWA compiled metrics are not in _Sustains_Stable trees. They live on the
    Zenodo collections workbook: IOWA pp/mf/ff sheets (CDM by string) and the
    Media sheet columns IOWA pp / IOWA mf / IOWA ff.
    """
    warnings: list[str] = []

    def note(msg: str) -> None:
        line = f"WARNING: {msg}"
        print(line)
        warnings.append(line)

    print("\nSEARCH  IOWA Zenodo collections (by-string compiled metrics)")
    print(f"  file    {path}")
    if not path.exists():
        print("  result  MISSING")
        note(f"missing IOWA Zenodo collections workbook: {path}")
        return warnings

    print("  result  found")
    with pd.ExcelFile(path) as xl:
        sheets = list(xl.sheet_names)
    media = resolve_media_sheet(sheets, media_sheet)
    iowa_sheets = _iowa_dynamic_sheets(sheets)
    print(f"  Media sheet  {media or '(none)'}")
    print(f"  by-string IOWA sheets  {iowa_sheets or '(none)'}")

    counts: dict[str, int] = {}
    if media is None:
        note(f"IOWA Zenodo book has no Media sheet: {path.name}")
    else:
        df = pd.read_excel(path, sheet_name=media)
        for dyn in CORE_DYNS:
            col = f"IOWA {dyn}"
            if col not in df.columns:
                note(f"{path.name} / {media}: missing column '{col}'")
                continue
            n = 0
            for raw in df[col]:
                try:
                    if float(raw) > 0:
                        n += 1
                except (TypeError, ValueError):
                    continue
            counts[dyn] = n
            if n == 0:
                note(f"{path.name} / {media}: column '{col}' has no IOWA values")
        print(f"  USE    {path}")
        print(f"         sheet={media}  columns IOWA pp/mf/ff  n={counts}")

    media_has_iowa = bool(media) and any(int(counts.get(dyn, 0) or 0) > 0 for dyn in CORE_DYNS)
    for dyn in CORE_DYNS:
        sheet = iowa_sheets.get(dyn)
        if not sheet:
            if media_has_iowa:
                continue
            note(f"{path.name}: missing IOWA {dyn} by-string sheet (compiled metrics kept by string)")
            continue
        curve = _curve_from_iowa_string_sheet(path, sheet)
        print(f"  USE    {path}")
        print(f"         sheet={sheet}  Source note + CDM - Media  n={len(curve)}")
        if not curve and not media_has_iowa:
            note(f"{path.name} / {sheet}: IOWA {dyn} by-string sheet has no CDM values")
    return warnings


def load_arco_media(path: Path, media_sheet: str | None = None) -> dict[str, dict[str, dict[int, float]]]:
    """Read IOWA + ORCH arco from a Zenodo collections workbook.

    IOWA is taken from the Media columns when present, otherwise from the
    by-string IOWA sheets (CDM - Media). Orchidea effects still come from
    compiled_density_metrics_research trees.
    """
    curves = {"IOWA": {d: {} for d in CORE_DYNS}, "ORCH": {d: {} for d in CORE_DYNS}}
    print("\nSEARCH  arco Media (IOWA + ORCH spines)")
    print(f"  file    {path}")
    if not path.exists():
        print("  result  MISSING")
        return curves
    print("  result  found")
    with pd.ExcelFile(path) as xl:
        sheets = list(xl.sheet_names)
    media_name = resolve_media_sheet(sheets, media_sheet)
    if media_name:
        media = pd.read_excel(path, sheet_name=media_name)
        for _, row in media.iterrows():
            pitch = parse_pitch(row.get("Note"))
            if not pitch:
                continue
            for coll, prefix in (("IOWA", "IOWA"), ("ORCH", "ORCH")):
                for dyn in CORE_DYNS:
                    col = f"{prefix} {dyn}"
                    if col not in media.columns:
                        continue
                    try:
                        value = float(row[col])
                    except (TypeError, ValueError):
                        continue
                    if value > 0:
                        curves[coll][dyn][pitch.midi] = value
    iowa_sheets = _iowa_dynamic_sheets(sheets)
    for dyn, sheet in iowa_sheets.items():
        if curves["IOWA"][dyn]:
            continue
        curves["IOWA"][dyn] = _curve_from_iowa_string_sheet(path, sheet)
        print(f"  FALLBACK IOWA {dyn} from by-string sheet {sheet}  n={len(curves['IOWA'][dyn])}")
    print(f"  USE    {path}")
    print(f"         sheet={media_name}  IOWA { {d: len(curves['IOWA'][d]) for d in CORE_DYNS} }")
    print(f"         sheet={media_name}  ORCH { {d: len(curves['ORCH'][d]) for d in CORE_DYNS} }")
    return curves


def discover_research_files() -> list[dict]:
    specs = []
    warnings: list[str] = []
    warnings.extend(audit_iowa_zenodo_workbook(ARCO, media_sheet="Violin_Media"))
    trees = [
        (Path(r"D:\CORDAS_2\VIOLIN\Orchidea_violin\con-sord"), "ORCH", "con sordino"),
        (Path(r"D:\CORDAS_2\VIOLIN\Orchidea_violin\sul-ponticello"), "ORCH", "sul ponticello"),
        (Path(r"D:\CORDAS_2\VIOLIN\Orchidea_violin\harmonics"), "ORCH", "harmonics"),
        (Path(r"D:\CORDAS_2\VIOLIN\Philharmonia_violin"), "PHIL", None),
        (Path(r"D:\CORDAS_2\VIOLIN\McGill_violin"), "MCGILL", None),
    ]
    for tree, collection, forced_tech in trees:
        warnings.extend(
            audit_research_tree(tree, collection=collection, technique=forced_tech or "")
        )
        if not tree.exists():
            continue
        for path in tree.rglob("*compiled_density_metrics_research.xlsx"):
            parts = [p.lower() for p in path.parts]
            tech = forced_tech
            if tech is None:
                if "con-sord" in parts or "muted" in parts:
                    tech = "con sordino"
                elif "sul-ponticello" in parts or "arco-sul-ponticello" in parts:
                    tech = "sul ponticello"
                elif "sul-tasto" in parts or "arco-sul-tasto" in parts:
                    tech = "sul tasto"
                elif "harmonics" in parts:
                    tech = "harmonics"
                elif "arco-normal" in parts or "non-vibrato" in parts:
                    tech = "ordinario"
                else:
                    continue
            dyn = None
            for folder, code in DYN_MAP.items():
                if folder in parts:
                    dyn = code
                    break
            if dyn is None and collection == "MCGILL":
                dyn = "mf"
            if dyn is None:
                continue
            if "_sustains\\" in str(path).lower() and "_sustains_stable" not in str(path).lower():
                continue
            specs.append(
                {
                    "path": path,
                    "collection": collection,
                    "technique": tech,
                    "dynamic": dyn,
                }
            )
    log_discovered_specs(specs)
    discover_research_files.last_warnings = warnings
    return specs


def curve_from_cells(layer) -> dict[int, float]:
    return {m: c.value for m, c in layer.cells.items()}


def build_layer(instrument, collection, technique, dynamic, curve, origin, comment="", source=""):
    layer = layer_from_identity(instrument, collection, technique, dynamic)
    for midi, payload in curve.items():
        if isinstance(payload, tuple):
            label, value = payload
            pitch = parse_pitch(label) or parse_pitch(f"midi{midi}")
        else:
            value = payload
            pitch = parse_pitch(f"midi{midi}")
        if not pitch or value <= 0:
            continue
        layer.place(
            pitch,
            float(value),
            origin,
            comment,
            overwrite="overwrite",
            source_workbook=source,
        )
    if layer.cells:
        layer.range_low = min(layer.cells)
        layer.range_high = max(layer.cells)
    return layer


def anchors_from_pair(source: dict[int, float], target: dict[int, tuple[str, float] | float]) -> list[Anchor]:
    out = []
    for midi, src in source.items():
        if midi not in target or src <= 0:
            continue
        tgt = target[midi][1] if isinstance(target[midi], tuple) else target[midi]
        if tgt > 0:
            out.append(Anchor(midi, float(src), float(tgt)))
    return out


def transfer_or_copy(
    source_layer,
    anchors,
    technique,
    collection,
    origin_fallback="technique_transfer",
    copy_anchors: bool = True,
    min_midi: int | None = None,
):
    if len(anchors) < 1:
        return None
    result = technique_transfer(
        source_layer, anchors, technique, "log_ratio", collection, copy_anchors=copy_anchors
    )
    drop = []
    for midi, cell in result.layer.cells.items():
        if min_midi is not None and midi < min_midi:
            drop.append(midi)
            continue
        if cell.origin != "measured":
            cell.origin = origin_fallback
            cell.comment = (cell.comment or "") + "; L from measured effect/ordinario pairs"
    for midi in drop:
        result.layer.cells.pop(midi, None)
    return result.layer


def main() -> None:
    print("Loading arco-normal Media…")
    arco = load_arco_media(ARCO)
    print(
        "  IOWA", {d: len(arco["IOWA"][d]) for d in CORE_DYNS},
        "ORCH", {d: len(arco["ORCH"][d]) for d in CORE_DYNS},
    )

    print("Discovering research workbooks…")
    specs = discover_research_files()
    measured: dict[tuple[str, str, str], dict] = {}
    for spec in specs:
        mass = load_spectral_mass(spec["path"])
        key = (spec["collection"], spec["technique"], spec["dynamic"])
        print(f"  {spec['collection']:6} {spec['technique']:16} {spec['dynamic']:3} n={len(mass)}  {spec['path'].name}")
        if not mass:
            continue
        if key not in measured or len(mass) > len(measured[key]["curve"]):
            measured[key] = {"curve": mass, "path": spec["path"]}

    effects = ["con sordino", "sul ponticello", "harmonics", "sul tasto"]
    written = []

    for effect in effects:
        print(f"\n=== {effect} ===")
        project = Project(
            title=f"Violin {effect} — STE Lab IOWA + ORCHIDEA",
            operator="STE Lab batch",
            notes=(
                "Measured spectral_mass (F-061) from compiled_density_metrics_research.xlsx. "
                "Principal evidence = Empirical_ORCH (measured Orchidea only). "
                "IOWA has no effect recordings: IOWA = IOWA_ordinario × exp(L), L from Orchidea "
                "effect/ordinario pairs when available — shared L, not replication. "
                "Philharmonia/McGill are context layers, not Media. "
                "Sul tasto and inherited pp/ff are prediction, not Empirical_ORCH."
            ),
        )

        orch_meas = {d: measured.get(("ORCH", effect, d), {}).get("curve", {}) for d in CORE_DYNS + ("p", "mp", "f")}
        phil_meas = {d: measured.get(("PHIL", effect, d), {}).get("curve", {}) for d in list(DYN_MAP.values())}
        mcgill_meas = {d: measured.get(("MCGILL", effect, d), {}).get("curve", {}) for d in CORE_DYNS + ("p",)}

        teacher = {}
        for d in CORE_DYNS:
            teacher[d] = orch_meas.get(d) or {}
        if not any(teacher.values()):
            # sul tasto: teach L from Philharmonia piano vs Phil/IOWA/ORCH arco
            if any(phil_meas.values()):
                print("  No Orchidea measurements; using Philharmonia as L teacher (context-grade).")
                for d, curve in phil_meas.items():
                    if curve:
                        teacher[d] = curve
            else:
                print("  No teacher curve — skip")
                continue

        # Prefer mf as the L donor when a core dynamic has no measured effect.
        # L must be same-dynamic: effect_D / ordinario_D. Never pair effect_mf with arco_pp.
        donor_dyn = next((d for d in ("mf", "pp", "ff", "p") if teacher.get(d)), "mf")
        donor_curve = teacher[donor_dyn]
        if effect == "sul tasto" and phil_meas.get("p"):
            phil_arco_p = measured.get(("PHIL", "ordinario", "p"), {}).get("curve", {})
            donor_arco_curve = {m: v[1] if isinstance(v, tuple) else v for m, v in phil_arco_p.items()}
            donor_source_label = "PHIL ordinario p"
        else:
            donor_arco_curve = arco["ORCH"].get(donor_dyn) or arco["ORCH"]["mf"]
            donor_source_label = f"ORCH ordinario {donor_dyn}"
        donor_anchors = anchors_from_pair(donor_arco_curve, donor_curve)
        print(f"  L donor {donor_dyn} vs {donor_source_label}: {len(donor_anchors)} anchors")

        iowa_layers = []
        orch_layers = []

        for dyn in CORE_DYNS:
            iowa_arco_layer = build_layer(
                "violin", "IOWA", "ordinario", dyn, arco["IOWA"][dyn], "measured",
                "Violin_Media arco-normal", str(ARCO.name),
            )
            orch_arco_layer = build_layer(
                "violin", "ORCH", "ordinario", dyn, arco["ORCH"][dyn], "measured",
                "Violin_Media arco-normal", str(ARCO.name),
            )

            used_donor = True
            if teacher.get(dyn) and orch_meas.get(dyn):
                local_arco = arco["ORCH"].get(dyn) or arco["ORCH"]["mf"]
                anchors = anchors_from_pair(local_arco, teacher[dyn])
                used_donor = False
            else:
                anchors = donor_anchors
            if len(anchors) < 3:
                anchors = donor_anchors
                used_donor = True
            inherited_from = donor_dyn if used_donor else None
            prediction = effect == "sul tasto"
            has_orch_meas = bool(orch_meas.get(dyn))
            min_midi = 79 if effect == "harmonics" else None

            # ORCHIDEA effect
            if has_orch_meas:
                orch_eff = build_layer(
                    "violin",
                    "ORCH",
                    effect,
                    dyn,
                    orch_meas[dyn],
                    "measured",
                    "Orchidea compiled spectral_mass F-061",
                    str(measured[("ORCH", effect, dyn)]["path"].name),
                )
                if min_midi is not None:
                    for midi in [m for m in orch_eff.cells if m < min_midi]:
                        orch_eff.cells.pop(midi, None)
                transferred = transfer_or_copy(
                    orch_arco_layer,
                    anchors,
                    effect,
                    "ORCH",
                    copy_anchors=False,
                    min_midi=min_midi,
                )
                if transferred:
                    for midi, cell in transferred.cells.items():
                        if midi not in orch_eff.cells:
                            orch_eff.place(
                                parse_pitch(f"midi{midi}"),
                                cell.value,
                                "modelled_ORCHIDEA_anchored",
                                "filled from ordinario × exp(L)",
                                overwrite="overwrite",
                            )
                stamp_evidence(orch_eff, inherited_from=inherited_from, prediction=prediction)
                project.add_layer(orch_eff)
                orch_layers.append(orch_eff)
            else:
                transferred = transfer_or_copy(
                    orch_arco_layer,
                    anchors,
                    effect,
                    "ORCH",
                    "modelled_ORCHIDEA_anchored",
                    copy_anchors=False,
                    min_midi=min_midi,
                )
                if transferred:
                    stamp_evidence(transferred, inherited_from=inherited_from or donor_dyn, prediction=prediction)
                    project.add_layer(transferred)
                    orch_layers.append(transferred)

            # IOWA effect — never measured
            iowa_origin = "modelled_IOWA_anchored"
            iowa_eff = transfer_or_copy(
                iowa_arco_layer,
                anchors,
                effect,
                "IOWA",
                iowa_origin,
                copy_anchors=False,
                min_midi=min_midi,
            )
            if iowa_eff:
                if not orch_meas.get(dyn):
                    for cell in iowa_eff.cells.values():
                        if cell.midi > 100:
                            cell.reporting_status = "above_measured_ceiling"
                stamp_evidence(iowa_eff, inherited_from=inherited_from, prediction=prediction)
                project.add_layer(iowa_eff)
                iowa_layers.append(iowa_eff)

        # context layers
        for dyn, curve in phil_meas.items():
            if not curve:
                continue
            src = measured.get(("PHIL", effect, dyn), {}).get("path")
            project.add_layer(
                build_layer(
                    "violin", "PHIL", effect, dyn, curve, "measured",
                    "Philharmonia context — not pooled into Media",
                    src.name if src else "",
                )
            )
        for dyn, curve in mcgill_meas.items():
            if not curve:
                continue
            src = measured.get(("MCGILL", effect, dyn), {}).get("path")
            project.add_layer(
                build_layer(
                    "violin", "MCGILL", effect, dyn, curve, "measured",
                    "McGill context — not pooled into Media",
                    src.name if src else "",
                )
            )

        media_layers = []
        for dyn in CORE_DYNS:
            peers = [
                lg
                for lg in project.layers
                if lg.technique == effect
                and lg.dynamic == dyn
                and lg.collection in {"IOWA", "ORCH"}
                and lg.cells
            ]
            if len(peers) >= 1:
                media = production_media_of(
                    peers,
                    name=f"violin_Media_{effect}_{dyn}",
                    instrument="violin",
                    technique=effect,
                )
                project.add_layer(media)
                media_layers.append(media)

        flags = audit_project(project)
        slug = effect.replace(" ", "_")
        out = OUT / f"Violin_STE_{slug}_IOWA_ORCH.xlsx"
        export_workbook(project, out, flags, media_layers)
        written.append(out)
        print(f"  wrote {out.name}  layers={len(project.layers)}  flags={len(flags)}  size={out.stat().st_size}")

        zenodo = OUT / f"Violin_Zenodo_collections_{slug}.xlsx"
        export_zenodo_workbook(
            project,
            zenodo,
            technique=effect,
            arco=arco,
            notes=project.notes,
        )
        written.append(zenodo)
        print(f"  wrote {zenodo.name}  size={zenodo.stat().st_size}")

    print("\nDone:")
    for p in written:
        print(" ", p)


if __name__ == "__main__":
    main()
