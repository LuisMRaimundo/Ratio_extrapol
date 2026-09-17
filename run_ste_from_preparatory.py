# -*- coding: utf-8 -*-
"""Run all STE effect books from one preparatory workbook.

The GUI is not required. The Excel produced by build_viola_ste_preparatory.py
already holds arco spines, measured effects and Tab-3 anchors. This script
reads those sheets and writes STE + Zenodo books for every effect at once.

    python run_ste_from_preparatory.py
    python run_ste_from_preparatory.py --gui
    python run_ste_from_preparatory.py --prep "D:\\CORDAS_3\\VIOLA 4\\Viola_STE_preparatory.xlsx"
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from run_ste_effects_batch import build_layer, transfer_or_copy
from ste_lab.catalog import orchestral_group, resolve_instrument

STRING_ONLY_EFFECTS = frozenset(
    {"sul ponticello", "con sordino", "sul tasto", "harmonics"}
)
from ste_lab.evidence import stamp_evidence
from ste_lab.excel_export import export_workbook
from ste_lab.notes import parse_pitch
from ste_lab.qa import audit_project
from ste_lab.session import Project
from ste_lab.media_policy import production_media_of
from ste_lab.transfer import Anchor
from ste_lab.zenodo_export import export_zenodo_workbook

DEFAULT_PREP = Path(r"D:\CORDAS_3\VIOLA 4\Viola_STE_preparatory.xlsx")
DEFAULT_OUT = Path(r"D:\CORDAS_3\VIOLA 4\Viola_Extrapoled_effects")
CORE = ("pp", "mf", "ff")
HARM_LO = 72
HARM_FLOOR = {
    "violin": 79,
    "viola": 72,
    "cello": 60,
    "double_bass": 52,
}

# Known slugs for ORCH_{slug}_{dynamic} columns. Unknown slugs map back by replacing _.
EFFECT_SLUGS = {
    "harmonics": "harmonics",
    "sul ponticello": "ponticello",
    "con sordino": "sordino",
}
SLUG_TO_EFFECT = {slug: effect for effect, slug in EFFECT_SLUGS.items()}
# Backward-compatible aliases (mf only). Loader also scans every ORCH_*_{dyn} column.
EFFECT_MEASURED_COL = {
    "harmonics": "ORCH_harmonics_mf",
    "sul ponticello": "ORCH_ponticello_mf",
    "con sordino": "ORCH_sordino_mf",
}
CONTEXT_COLS = {
    "PHIL_harmonics_mf": ("PHIL", "harmonics"),
    "McGill_harmonics_mf": ("MCGILL", "harmonics"),
    "McGill_sordino_mf": ("MCGILL", "con sordino"),
    "McGill_ordinario_mf": ("MCGILL", "ordinario"),
}


def parse_context_column(col: str) -> tuple[str, str, str] | None:
    """PHIL_ordinario_pp / McGill_ordinario_mf → (PHIL, ordinario, pp)."""
    name = str(col).strip()
    low = name.lower()
    dyn = None
    for code in (*CORE, "p", "mp", "f"):
        if low.endswith("_" + code):
            dyn = code
            stem = name[: -len(code) - 1]
            break
    if dyn is None:
        return None
    stem_l = stem.lower()
    if stem_l.startswith("phil_"):
        tech = stem[5:].replace("_", " ").strip() or "ordinario"
        return "PHIL", tech, dyn
    if stem_l.startswith("mcgill_"):
        tech = stem[7:].replace("_", " ").strip() or "ordinario"
        return "MCGILL", tech, dyn
    return None


def parse_orch_effect_column(col: str) -> tuple[str, str] | None:
    """Map ORCH_ponticello_ff → ('sul ponticello', 'ff'). Skip arco, tasto."""
    name = str(col).strip()
    if not name.upper().startswith("ORCH_"):
        return None
    rest = name[5:]
    low = rest.lower()
    if low.startswith("arco_") or "tasto" in low:
        return None
    for dyn in (*CORE, "p", "mp", "f"):
        suffix = f"_{dyn}"
        if rest.endswith(suffix):
            slug = rest[: -len(suffix)]
            if slug.lower() in {"ordinario", "arco"}:
                return None
            effect = SLUG_TO_EFFECT.get(slug, slug.replace("_", " "))
            return effect, dyn
    return None


def orchidea_recorded_layer(
    instrument: str,
    effect: str,
    dyn: str,
    curve: dict,
    source: str,
    min_midi: int | None,
):
    """Routine for every instrument, effect and dynamic.

    If Orchidea recorded this effect at this dynamic, those cells *are* the
    layer (origin measured). L is not applied. Do not invent sul tasto.
    """
    if effect.lower() == "sul tasto" or not curve:
        return None
    layer = build_layer(
        instrument, "ORCH", effect, dyn, curve, "measured",
        "Paste_effects recorded Orchidea — not modelled", source,
    )
    if min_midi is not None:
        for midi in [m for m in layer.cells if m < min_midi]:
            layer.cells.pop(midi, None)
    if not layer.cells:
        return None
    stamp_evidence(layer, inherited_from=None, prediction=False)
    return layer


def _finite(v) -> bool:
    try:
        x = float(v)
    except (TypeError, ValueError):
        return False
    return math.isfinite(x) and x > 0


def _series_map(df: pd.DataFrame, value_col: str) -> dict[int, tuple[str, float]]:
    out: dict[int, tuple[str, float]] = {}
    if value_col not in df.columns:
        return out
    for rec in df.itertuples(index=False):
        raw = getattr(rec, value_col, None)
        if not _finite(raw):
            continue
        midi = int(rec.midi) if hasattr(rec, "midi") and pd.notna(rec.midi) else None
        note = str(rec.note) if hasattr(rec, "note") and pd.notna(rec.note) else ""
        pitch = parse_pitch(note) or (parse_pitch(f"midi{midi}") if midi else None)
        if pitch is None:
            continue
        out[pitch.midi] = (pitch.label, float(raw))
    return out


def load_preparatory(path: Path) -> dict:
    with pd.ExcelFile(path) as xl:
        need = {"Paste_arco", "Paste_effects_mf", "Anchors_all"}
        missing = need - set(xl.sheet_names)
    if missing:
        raise SystemExit(f"{path} is missing sheets: {sorted(missing)}")
    def _read_sheet(sheet: str, required: str) -> pd.DataFrame:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        header_at = None
        for i, row in raw.iterrows():
            vals = [str(v).strip().lower() for v in row.tolist() if pd.notna(v)]
            if required in vals:
                header_at = int(i)
                break
        if header_at is None:
            raise SystemExit(f"{path.name} / {sheet}: no header row containing '{required}'")
        df = raw.iloc[header_at + 1 :].copy()
        df.columns = [str(c).strip() for c in raw.iloc[header_at].tolist()]
        return df.reset_index(drop=True)

    arco_df = _read_sheet("Paste_arco", "note")
    eff_df = _read_sheet("Paste_effects_mf", "note")
    anc_df = _read_sheet("Anchors_all", "effect")

    arco_df = arco_df.dropna(how="all")
    eff_df = eff_df.dropna(how="all")
    anc_df = anc_df.dropna(how="all")
    arco = {"IOWA": {d: {} for d in CORE}, "ORCH": {d: {} for d in CORE}}
    for rec in arco_df.itertuples(index=False):
        midi = int(rec.midi) if pd.notna(rec.midi) else None
        if midi is None:
            continue
        for coll in ("IOWA", "ORCH"):
            for dyn in CORE:
                col = f"{coll}_{dyn}"
                if col not in arco_df.columns:
                    continue
                v = getattr(rec, col, None)
                if _finite(v):
                    arco[coll][dyn][midi] = float(v)

    measured = {}
    for col in eff_df.columns:
        parsed = parse_orch_effect_column(col)
        if not parsed:
            continue
        effect, dyn = parsed
        curve = _series_map(eff_df, col)
        if curve:
            measured[("ORCH", effect, dyn)] = curve
    context = {}
    for col, (coll, tech) in CONTEXT_COLS.items():
        mapped = _series_map(eff_df, col)
        if mapped:
            context[(coll, tech, "mf")] = mapped
    for col in eff_df.columns:
        parsed = parse_context_column(str(col))
        if not parsed:
            continue
        coll, tech, dyn = parsed
        mapped = _series_map(eff_df, str(col))
        if mapped:
            context[(coll, tech, dyn)] = mapped

    def _anchors_from(group) -> list:
        out = []
        for rec in group.itertuples(index=False):
            if not _finite(rec.sourceCDM) or not _finite(rec.targetCDM):
                continue
            midi = int(rec.midi) if pd.notna(rec.midi) else None
            if midi is None:
                pitch = parse_pitch(str(rec.note))
                midi = pitch.midi if pitch else None
            if midi is None:
                continue
            out.append(Anchor(midi, float(rec.sourceCDM), float(rec.targetCDM)))
        return out

    origins = {"IOWA": "measured", "ORCH": "measured"}
    with pd.ExcelFile(path) as xl:
        has_prov = "Spine_provenance" in xl.sheet_names
    if has_prov:
        try:
            prov = _read_sheet("Spine_provenance", "collection")
            for rec in prov.itertuples(index=False):
                coll = str(getattr(rec, "collection", "")).strip().upper()
                origin = str(getattr(rec, "origin", "")).strip()
                if coll in origins and origin:
                    origins[coll] = origin
        except SystemExit:
            pass

    effects = []
    if "effect" not in anc_df.columns or anc_df.empty:
        print("  Anchors_all has no effects (ordinario-only instrument)")
        return {
            "arco": arco,
            "measured": measured,
            "context": context,
            "effects": effects,
            "path": path,
            "origins": origins,
        }
    for effect, g in anc_df.groupby(anc_df["effect"].astype(str).str.strip(), sort=False):
        effect = str(effect).strip()
        if not effect or effect.lower() == "sul tasto":
            continue
        grade = str(g["evidence_grade"].iloc[0])
        anchors_by_dyn: dict[str, list] = {}
        if "dynamic" in g.columns:
            for dyn, sg in g.groupby(g["dynamic"].astype(str).str.strip()):
                dyn = str(dyn).strip()
                if dyn and dyn.lower() != "nan":
                    anchors_by_dyn[dyn] = _anchors_from(sg)
        if not anchors_by_dyn:
            anchors_by_dyn["mf"] = _anchors_from(g)
        anchors = anchors_by_dyn.get("mf") or next(iter(anchors_by_dyn.values()), [])
        effects.append(
            {
                "name": effect,
                "grade": grade,
                "anchors": anchors,
                "anchors_by_dyn": anchors_by_dyn,
            }
        )
    return {
        "arco": arco,
        "measured": measured,
        "context": context,
        "effects": effects,
        "path": path,
        "origins": origins,
    }


def harm_floor(instrument: str) -> int:
    spec = resolve_instrument(instrument)
    key = spec.instrument_id if spec else (instrument or "viola").strip().lower().replace(" ", "_")
    return HARM_FLOOR.get(key, HARM_LO)


def _instrument_names(instrument: str) -> tuple[str, str]:
    spec = resolve_instrument(instrument)
    if spec:
        return spec.instrument_id, spec.display_name.replace(" ", "_")
    raw = (instrument or "viola").strip() or "viola"
    return raw, raw.replace(" ", "_").title()


def run(
    prep: Path,
    out: Path,
    instrument: str = "viola",
    operator: str = "",
    notes: str = "",
) -> list[Path]:
    pack = load_preparatory(prep)
    out.mkdir(parents=True, exist_ok=True)
    arco = pack["arco"]
    written: list[Path] = []
    instr_id, instr_file = _instrument_names(instrument)
    woodwind = orchestral_group(instr_id) == "woodwinds"
    if woodwind:
        pack["context"] = {
            key: curve
            for key, curve in (pack.get("context") or {}).items()
            if str(key[1]).lower() not in STRING_ONLY_EFFECTS
        }
    op = (operator or "").strip() or "STE from preparatory"
    extra = (notes or "").strip()
    spine = "ordinario" if woodwind else "arco"
    print("Preparatory", prep)
    print("  instrument", instr_id)
    print(f"  IOWA {spine}", {d: len(arco["IOWA"][d]) for d in CORE})
    print(f"  ORCH {spine}", {d: len(arco["ORCH"][d]) for d in CORE})
    prep_l = str(prep).lower()
    if instr_id != "viola" and "viola" in prep_l:
        print(
            "WARNING: Instrument is "
            f"{instr_id} but the preparatory Excel is a viola pack.\n"
            f"  Data still come from {prep}\n"
            "  Changing the Instrument box only relabels the output. "
            "It does not load cello/violin/bass trees."
        )
    with pd.ExcelFile(prep) as xl:
        has_inv = "Measured_inventory" in xl.sheet_names
    if has_inv:
        inv = pd.read_excel(prep, sheet_name="Measured_inventory", header=None)
        header_at = None
        for i, row in inv.iterrows():
            vals = [str(v).strip().lower() for v in row.tolist() if pd.notna(v)]
            if "collection" in vals and "source" in vals:
                header_at = int(i)
                break
        print("\nFILES RECORDED IN Measured_inventory")
        if header_at is not None:
            inv.columns = [str(c).strip() for c in inv.iloc[header_at].tolist()]
            inv = inv.iloc[header_at + 1 :].reset_index(drop=True)
            for rec in inv.itertuples(index=False):
                print(
                    f"  {getattr(rec, 'collection', '')}  {getattr(rec, 'technique', '')}  "
                    f"{getattr(rec, 'dynamic', '')}  n={getattr(rec, 'n', '')}  "
                    f"{getattr(rec, 'source', '')}"
                )
    print(f"  USE    {prep}  sheets=Paste_arco, Paste_effects_mf, Anchors_all")
    origins = pack.get("origins") or {"IOWA": "measured", "ORCH": "measured"}
    if woodwind:
        print("  woodwind run: string techniques are not invented")

    for spec in pack["effects"]:
        effect = spec["name"]
        if woodwind and effect.lower() in STRING_ONLY_EFFECTS:
            print(f"\n=== {effect} ===  skip: string technique, not invented for woodwinds")
            continue
        anchors = spec["anchors"]
        anchors_by_dyn = spec.get("anchors_by_dyn") or {"mf": anchors}
        empirical = str(spec["grade"]).startswith("Empirical")
        prediction = not empirical
        min_midi = harm_floor(instrument) if effect == "harmonics" else None
        recorded = {d: pack["measured"].get(("ORCH", effect, d), {}) for d in CORE}
        print(
            f"\n=== {effect} ===  anchors={len(anchors)}  grade={spec['grade']}  "
            f"ORCH recorded { {d: len(recorded[d]) for d in CORE} }"
        )
        if len(anchors) < 3 and not any(recorded.values()):
            print("  skip: fewer than 3 anchors and no Orchidea recording")
            continue

        project = Project(
            title=f"{instr_file.replace('_', ' ')} {effect} — STE from preparatory workbook",
            operator=op,
            notes=(
                f"Driven by {prep.name}. Anchors_all grade={spec['grade']}. "
                "Orchidea recorded cells stay measured; "
                "L is not applied to a recorded (effect, dynamic). "
                "IOWA = IOWA_ordinario × exp(L). Shared L is not a second experiment. "
                "Philharmonia/McGill context columns are not Media. "
                "No sul tasto unless a future preparatory sheet supplies measured pairs."
                + (f" {extra}" if extra else "")
            ),
        )

        for dyn in CORE:
            dyn_anchors = anchors_by_dyn.get(dyn) or anchors
            same_dyn = bool(anchors_by_dyn.get(dyn)) and dyn in anchors_by_dyn
            inherited = None if (same_dyn or (dyn == "mf" and empirical)) else "mf"
            iowa_arco = build_layer(
                instr_id, "IOWA", "ordinario", dyn, arco["IOWA"][dyn],
                origins.get("IOWA", "measured"),
                "Paste_arco", prep.name,
            )
            orch_arco = build_layer(
                instr_id, "ORCH", "ordinario", dyn, arco["ORCH"][dyn],
                origins.get("ORCH", "measured"),
                "Paste_arco", prep.name,
            )

            if recorded[dyn]:
                orch_eff = orchidea_recorded_layer(
                    instr_id, effect, dyn, recorded[dyn], prep.name, min_midi,
                )
                if orch_eff:
                    print(f"  ORCH {dyn} recorded n={len(orch_eff.cells)}  (no L)")
                    project.add_layer(orch_eff)
            elif len(dyn_anchors) >= 3:
                transferred = transfer_or_copy(
                    orch_arco,
                    dyn_anchors,
                    effect,
                    "ORCH",
                    "modelled_ORCHIDEA_anchored",
                    copy_anchors=False,
                    min_midi=min_midi,
                )
                if transferred:
                    stamp_evidence(
                        transferred,
                        inherited_from=inherited,
                        prediction=prediction,
                    )
                    project.add_layer(transferred)

            if len(dyn_anchors) >= 3:
                iowa_eff = transfer_or_copy(
                    iowa_arco,
                    dyn_anchors,
                    effect,
                    "IOWA",
                    "modelled_IOWA_anchored",
                    copy_anchors=False,
                    min_midi=min_midi,
                )
                if iowa_eff:
                    if prediction:
                        for cell in iowa_eff.cells.values():
                            if cell.midi > 88:
                                cell.reporting_status = "above_measured_ceiling"
                    stamp_evidence(iowa_eff, inherited_from=inherited, prediction=prediction)
                    project.add_layer(iowa_eff)

        if effect == "harmonics":
            for (coll, tech, dyn), curve in pack["context"].items():
                if tech != "harmonics" or not curve:
                    continue
                ctx = build_layer(
                    instr_id, coll, tech, dyn, curve, "measured",
                    "Paste_effects_mf context — not Media", prep.name,
                )
                stamp_evidence(ctx, inherited_from=None, prediction=False)
                project.add_layer(ctx)
        if effect == "con sordino":
            curve = pack["context"].get(("MCGILL", "con sordino", "mf"), {})
            if curve:
                ctx = build_layer(
                    instr_id, "MCGILL", "con sordino", "mf", curve, "measured",
                    "Paste_effects_mf prediction teacher — not Media", prep.name,
                )
                stamp_evidence(ctx, inherited_from=None, prediction=False)
                project.add_layer(ctx)

        media_layers = []
        for dyn in CORE:
            peers = [
                lg
                for lg in project.layers
                if lg.technique == effect
                and lg.dynamic == dyn
                and (lg.collection or "").upper() in {"IOWA", "ORCH"}
                and lg.cells
            ]
            if peers:
                media = production_media_of(
                    peers,
                    name=f"{instr_id}_Media_{effect}_{dyn}",
                    instrument=instr_id,
                    technique=effect,
                )
                project.add_layer(media)
                media_layers.append(media)

        flags = audit_project(project)
        slug = effect.replace(" ", "_")
        ste = out / f"{instr_file}_STE_{slug}_IOWA_ORCH.xlsx"
        zen = out / f"{instr_file}_Zenodo_collections_{slug}.xlsx"
        export_workbook(project, ste, flags, media_layers)
        export_zenodo_workbook(project, zen, technique=effect, arco=arco, notes=project.notes)
        written.extend([ste, zen])
        print(f"  wrote {ste.name}  layers={len(project.layers)}  flags={len(flags)}")
        print(f"  wrote {zen.name}")

    written.extend(
        _export_ordinario_spine(
            pack, out, instr_id, instr_file, arco, op, extra, prep,
        )
    )
    return written


def _export_ordinario_spine(pack, out, instr_id, instr_file, arco, op, extra, prep) -> list[Path]:
    """Always write the ordinario book. Woodwinds often have no string effects."""
    woodwind = orchestral_group(instr_id) == "woodwinds"
    project = Project(
        title=f"{instr_file.replace('_', ' ')} ordinario — STE from preparatory workbook",
        operator=op or "STE from preparatory",
        notes=(
            f"Driven by {prep.name}. Ordinario spine. "
            "Philharmonia and McGill are measured context when those columns exist. "
            "IOWA is measured. ORCH is a deprecated alias for orchidea_family_transfer_estimate."
            + ("" if woodwind else " Sul tasto is not invented.")
            + (f" {extra}" if extra else "")
        ),
    )
    origins = pack.get("origins") or {"IOWA": "measured", "ORCH": "measured"}
    for dyn in CORE:
        iowa = build_layer(
            instr_id, "IOWA", "ordinario", dyn, arco["IOWA"][dyn],
            origins.get("IOWA", "measured"),
            "Paste_arco", prep.name,
        )
        orch = build_layer(
            instr_id, "ORCH", "ordinario", dyn, arco["ORCH"][dyn],
            origins.get("ORCH", "measured"),
            "Paste_arco", prep.name,
        )
        if iowa.cells:
            stamp_evidence(iowa)
            project.add_layer(iowa)
        if orch.cells:
            stamp_evidence(orch)
            project.add_layer(orch)
    for (coll, tech, dyn), curve in pack.get("context", {}).items():
        if tech != "ordinario" or not curve:
            continue
        ctx = build_layer(
            instr_id, coll, tech, dyn, curve, "measured",
            "Paste_effects_mf context — not Media", prep.name,
        )
        stamp_evidence(ctx)
        project.add_layer(ctx)
    if not project.layers:
        print("\n=== ordinario ===  skip: no Paste_arco and no PHIL/McGill ordinario")
        return []
    media_layers = []
    for dyn in CORE:
        peers = [
            lg
            for lg in project.layers
            if lg.technique == "ordinario"
            and lg.dynamic == dyn
            and (lg.collection or "").upper() in {"IOWA", "ORCH"}
            and lg.cells
        ]
        if len(peers) >= 1:
            media = production_media_of(
                peers,
                name=f"{instr_id}_Media_ordinario_{dyn}",
                instrument=instr_id,
                technique="ordinario",
            )
            project.add_layer(media)
            media_layers.append(media)
    flags = audit_project(project)
    ste = out / f"{instr_file}_STE_ordinario_IOWA_ORCH.xlsx"
    zen = out / f"{instr_file}_Zenodo_collections_ordinario.xlsx"
    export_workbook(project, ste, flags, media_layers)
    export_zenodo_workbook(project, zen, technique="ordinario", arco=arco, notes=project.notes)
    print(f"\n=== ordinario ===  layers={len(project.layers)}  flags={len(flags)}")
    print(f"  wrote {ste.name}")
    print(f"  wrote {zen.name}")
    return [ste, zen]


def main() -> None:
    p = argparse.ArgumentParser(description="STE all effects from one preparatory Excel")
    p.add_argument("--prep", type=Path, default=DEFAULT_PREP, help="Viola_STE_preparatory.xlsx")
    p.add_argument("--out", type=Path, default=DEFAULT_OUT, help="Output folder for STE/Zenodo books")
    p.add_argument("--instrument", default="viola")
    p.add_argument("--operator", default="")
    p.add_argument("--notes", default="")
    p.add_argument("--gui", action="store_true", help="Open the small preparatory runner window")
    p.add_argument(
        "--rebuild-prep",
        action="store_true",
        help="Regenerate the preparatory workbook first from --rebuild-root.",
    )
    p.add_argument(
        "--rebuild-root",
        type=Path,
        default=None,
        help="Folder to hunt (Zenodo arco book + Orchidea/Phil/McGill trees).",
    )
    args = p.parse_args()
    if args.gui:
        from ste_lab.prep_gui import main as gui_main

        gui_main()
        return
    from ste_lab.pipeline import run_selected_pipeline

    if args.rebuild_prep:
        from build_ste_preparatory import DEPOSIT_DEFAULTS, deposit_defaults
        import build_ste_preparatory as bsp

        hint = deposit_defaults(args.instrument) or DEPOSIT_DEFAULTS["viola"]
        root = (args.rebuild_root or hint["root"]).resolve()
        written = run_selected_pipeline(
            args.instrument,
            args.prep,
            args.out.resolve(),
            rebuild=True,
            rebuild_root=root,
            operator=args.operator,
            notes=args.notes,
        )
        if bsp.last_path and (args.prep == DEFAULT_PREP or not args.prep.exists()):
            args.prep = bsp.last_path
        print("\nDone:")
        for path in written:
            print(" ", path)
        return
    if not args.prep.exists():
        raise SystemExit(f"missing preparatory workbook: {args.prep}")
    written = run_selected_pipeline(
        args.instrument,
        args.prep.resolve(),
        args.out.resolve(),
        operator=args.operator,
        notes=args.notes,
    )
    print("\nDone:")
    for path in written:
        print(" ", path)


if __name__ == "__main__":
    main()
